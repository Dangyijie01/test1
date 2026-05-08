#!/usr/bin/env python3
import argparse
import html
import json
import math
import os
import re
import statistics
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

FIELD_ALIASES = {
    "src": ("源IP(sip)", "sip", "源IP", "源ip", "src_ip", "source_ip", "src", "source"),
    "dst": ("目的IP(dip)", "dip", "目的IP", "目的ip", "dst_ip", "dest_ip", "destination_ip", "dst", "dest", "destination"),
    "port": ("目的端口(dport)", "dport", "目的端口", "dst_port", "dest_port", "port"),
    "protocol": ("应用层协议(app_protocol)", "app_protocol", "应用层协议", "协议", "protocol", "proto", "app"),
    "count": ("counts", "count", "次数", "访问次数", "流量数", "hits", "total"),
}


def column_letters(cell_ref):
    return re.sub(r"\d+", "", cell_ref or "")


def column_index(cell_ref):
    letters = column_letters(cell_ref)
    index = 0
    for ch in letters:
        index = index * 26 + ord(ch.upper()) - ord("A") + 1
    return index - 1


def text_from_rich_text(node):
    parts = []
    for text_node in node.findall(".//main:t", NS):
        parts.append(text_node.text or "")
    return "".join(parts)


def read_shared_strings(zf):
    try:
        raw = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []

    root = ET.fromstring(raw)
    return [text_from_rich_text(si) for si in root.findall("main:si", NS)]


def workbook_first_sheet_path(zf, requested_sheet=None):
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall("pkgrel:Relationship", NS)
    }

    sheets = workbook.findall("main:sheets/main:sheet", NS)
    if not sheets:
        raise ValueError("工作簿里没有可读取的 sheet")

    chosen = None
    if requested_sheet:
        for sheet in sheets:
            if sheet.attrib.get("name") == requested_sheet:
                chosen = sheet
                break
        if chosen is None:
            names = ", ".join(sheet.attrib.get("name", "") for sheet in sheets)
            raise ValueError(f"找不到 sheet: {requested_sheet}，可用 sheet: {names}")
    else:
        chosen = sheets[0]

    rel_id = chosen.attrib.get(f"{{{NS['rel']}}}id")
    target = rel_targets.get(rel_id)
    if not target:
        raise ValueError("无法解析 sheet 路径")
    if target.startswith("/"):
        target = target.lstrip("/")
    elif not target.startswith("xl/"):
        target = "xl/" + target
    return target, chosen.attrib.get("name", "")


def cell_value(cell, shared_strings):
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return text_from_rich_text(cell)

    value_node = cell.find("main:v", NS)
    value = value_node.text if value_node is not None else ""
    if cell_type == "s" and value != "":
        return shared_strings[int(value)]
    if cell_type == "b":
        return "TRUE" if value == "1" else "FALSE"
    return value or ""


def read_xlsx_rows(path, sheet_name=None):
    with zipfile.ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        sheet_path, resolved_sheet_name = workbook_first_sheet_path(zf, sheet_name)
        root = ET.fromstring(zf.read(sheet_path))

    rows = []
    for row in root.findall(".//main:sheetData/main:row", NS):
        values = []
        for cell in row.findall("main:c", NS):
            idx = column_index(cell.attrib.get("r", ""))
            while len(values) <= idx:
                values.append("")
            values[idx] = cell_value(cell, shared_strings).strip()
        if any(v != "" for v in values):
            rows.append(values)
    return rows, resolved_sheet_name


def normalize_header(value):
    return re.sub(r"[\s_\-（）()]+", "", str(value or "").strip().lower())


def find_column(headers, field, explicit=None):
    if explicit:
        normalized = normalize_header(explicit)
        for i, header in enumerate(headers):
            if normalize_header(header) == normalized:
                return i
        raise ValueError(f"找不到指定列 {explicit!r}")

    aliases = {normalize_header(v) for v in FIELD_ALIASES[field]}
    for i, header in enumerate(headers):
        if normalize_header(header) in aliases:
            return i
    for i, header in enumerate(headers):
        normalized = normalize_header(header)
        if any(alias in normalized or normalized in alias for alias in aliases):
            return i
    return None


def parse_count(value):
    if value in (None, ""):
        return 1
    text = str(value).replace(",", "").strip()
    try:
        return int(float(text))
    except ValueError:
        return 1


def get(row, index, default=""):
    if index is None or index >= len(row):
        return default
    return str(row[index]).strip()


def build_graph(rows, columns):
    edge_map = {}
    node_totals = defaultdict(lambda: {"in": 0, "out": 0, "total": 0})
    protocols = Counter()
    ports = Counter()
    raw_count = 0

    for row in rows:
        src = get(row, columns["src"])
        dst = get(row, columns["dst"])
        if not src or not dst:
            continue

        port = get(row, columns["port"], "UNKNOWN") or "UNKNOWN"
        protocol = get(row, columns["protocol"], "UNKNOWN") or "UNKNOWN"
        count = parse_count(get(row, columns["count"], "1"))
        raw_count += 1

        key = (src, dst, port, protocol)
        if key not in edge_map:
            edge_map[key] = {
                "source": src,
                "target": dst,
                "port": port,
                "protocol": protocol,
                "count": 0,
            }
        edge_map[key]["count"] += count
        node_totals[src]["out"] += count
        node_totals[src]["total"] += count
        node_totals[dst]["in"] += count
        node_totals[dst]["total"] += count
        protocols[protocol] += count
        ports[port] += count

    edges = sorted(edge_map.values(), key=lambda item: item["count"], reverse=True)
    nodes = []
    for node_id, total in node_totals.items():
        if total["in"] and total["out"]:
            role = "both"
        elif total["out"]:
            role = "source"
        else:
            role = "target"
        nodes.append({
            "id": node_id,
            "label": node_id,
            "role": role,
            "in": total["in"],
            "out": total["out"],
            "total": total["total"],
        })
    nodes.sort(key=lambda item: item["total"], reverse=True)
    return nodes, edges, protocols, ports, raw_count


def percentile(values, pct):
    if not values:
        return 0
    values = sorted(values)
    idx = min(len(values) - 1, max(0, round((len(values) - 1) * pct)))
    return values[idx]


def render_html(payload):
    data_json = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    title = html.escape(payload["title"])
    generated_at = html.escape(payload["generated_at"])
    source_name = html.escape(payload["source_name"])

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{title}</title>
  <script src="https://cdn.jsdelivr.net/npm/echarts@5.5.1/dist/echarts.min.js"></script>
  <style>
    :root {{
      color-scheme: light;
      --bg: #f5f7fb;
      --panel: #ffffff;
      --line: #d7deea;
      --muted: #607087;
      --text: #172235;
      --source: #1e88e5;
      --target: #00a884;
      --both: #9a5cff;
      --accent: #ff8a3d;
      --danger: #d44950;
      --shadow: 0 16px 42px rgba(35, 48, 71, .12);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      min-height: 100vh;
      background:
        radial-gradient(circle at top left, rgba(30, 136, 229, .08), transparent 28%),
        radial-gradient(circle at top right, rgba(0, 168, 132, .08), transparent 24%),
        linear-gradient(180deg, #f7f9fd 0%, #eef3f8 100%);
      color: var(--text);
      font-family: Inter, "Segoe UI", "PingFang SC", "Microsoft YaHei", Arial, sans-serif;
    }}
    header {{
      padding: 24px 28px 16px;
      display: flex;
      align-items: flex-end;
      justify-content: space-between;
      gap: 18px;
      border-bottom: 1px solid rgba(215, 222, 234, .8);
      background: rgba(255, 255, 255, .72);
      backdrop-filter: blur(14px);
      position: sticky;
      top: 0;
      z-index: 5;
    }}
    h1 {{
      margin: 0 0 8px;
      font-size: 24px;
      line-height: 1.25;
      font-weight: 750;
      letter-spacing: 0;
    }}
    .meta {{
      color: var(--muted);
      font-size: 13px;
      display: flex;
      flex-wrap: wrap;
      gap: 8px 18px;
    }}
    .toolbar {{
      display: flex;
      gap: 10px;
      align-items: center;
      flex-wrap: wrap;
      justify-content: flex-end;
    }}
    button, select, input {{
      height: 36px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: #fff;
      color: var(--text);
      font: inherit;
      font-size: 13px;
      outline: none;
    }}
    button {{
      padding: 0 12px;
      cursor: pointer;
      box-shadow: 0 1px 0 rgba(23, 34, 53, .04);
    }}
    button:hover {{ border-color: #aebbd0; }}
    input, select {{ padding: 0 10px; }}
    main {{
      padding: 18px 28px 28px;
      display: grid;
      grid-template-columns: 320px minmax(0, 1fr);
      gap: 18px;
      min-height: calc(100vh - 92px);
    }}
    aside, .graph-panel, .table-panel {{
      background: rgba(255, 255, 255, .94);
      border: 1px solid rgba(215, 222, 234, .86);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }}
    aside {{
      padding: 16px;
      align-self: start;
      position: sticky;
      top: 96px;
      max-height: calc(100vh - 116px);
      overflow: auto;
    }}
    .section-title {{
      margin: 2px 0 12px;
      font-size: 13px;
      font-weight: 720;
      color: #2b3a50;
    }}
    .field {{ margin-bottom: 13px; }}
    .field label {{
      display: block;
      margin-bottom: 6px;
      color: var(--muted);
      font-size: 12px;
    }}
    .field input, .field select {{ width: 100%; }}
    .range-row {{
      display: grid;
      grid-template-columns: 1fr 72px;
      gap: 8px;
      align-items: center;
    }}
    .query-row {{
      display: grid;
      grid-template-columns: 1fr 68px;
      gap: 8px;
      align-items: center;
    }}
    .query-row input {{ width: 100%; }}
    .query-row button {{ width: 100%; padding: 0 10px; }}
    input[type="range"] {{ padding: 0; }}
    .stats {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
      margin-bottom: 16px;
    }}
    .stat {{
      border: 1px solid #e2e8f2;
      border-radius: 7px;
      padding: 10px;
      background: #fafcff;
    }}
    .stat strong {{
      display: block;
      font-size: 20px;
      line-height: 1.1;
      margin-bottom: 4px;
    }}
    .stat span {{ color: var(--muted); font-size: 12px; }}
    .legend {{
      display: grid;
      gap: 8px;
      margin-top: 16px;
      color: var(--muted);
      font-size: 12px;
    }}
    .legend-item {{
      display: flex;
      align-items: center;
      gap: 8px;
    }}
    .dot {{
      width: 10px;
      height: 10px;
      border-radius: 50%;
      display: inline-block;
    }}
    .content {{
      min-width: 0;
      display: grid;
      grid-template-rows: minmax(68vh, 1fr) auto;
      gap: 18px;
    }}
    .graph-panel {{
      position: relative;
      min-height: 68vh;
      overflow: hidden;
    }}
    #graph {{
      display: block;
      width: 100%;
      height: 100%;
      min-height: 68vh;
    }}
    .empty {{
      position: absolute;
      inset: 0;
      display: none;
      align-items: center;
      justify-content: center;
      color: var(--muted);
      font-size: 14px;
      background: rgba(255,255,255,.72);
    }}
    .tooltip {{
      position: fixed;
      display: none;
      pointer-events: none;
      max-width: 340px;
      padding: 10px 12px;
      border-radius: 7px;
      background: rgba(23, 34, 53, .94);
      color: #fff;
      font-size: 12px;
      line-height: 1.6;
      z-index: 20;
      box-shadow: 0 16px 36px rgba(0,0,0,.22);
    }}
    .table-panel {{ overflow: hidden; }}
    .table-head {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 12px 14px;
      border-bottom: 1px solid var(--line);
      font-size: 13px;
      color: var(--muted);
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
    }}
    th, td {{
      padding: 10px 12px;
      border-bottom: 1px solid #edf1f7;
      text-align: left;
      white-space: nowrap;
    }}
    th {{
      color: #526177;
      background: #fbfcff;
      font-weight: 650;
    }}
    tbody tr:hover {{ background: #f8fbff; }}
    .pill {{
      display: inline-flex;
      align-items: center;
      min-height: 22px;
      padding: 2px 8px;
      border-radius: 999px;
      background: #eef4fb;
      color: #33445c;
      border: 1px solid #dbe5f1;
    }}
    .footer-note {{
      color: var(--muted);
      font-size: 12px;
      padding: 10px 14px 14px;
    }}
    @media (max-width: 960px) {{
      header {{ align-items: flex-start; flex-direction: column; }}
      main {{ grid-template-columns: 1fr; padding: 14px; }}
      aside {{ position: static; max-height: none; }}
      .content {{ grid-template-rows: auto auto; }}
      .graph-panel, svg {{ min-height: 520px; }}
    }}
  </style>
</head>
<body>
  <header>
    <div>
      <h1>{title}</h1>
      <div class="meta">
        <span>数据源：{source_name}</span>
        <span>生成时间：{generated_at}</span>
      </div>
    </div>
    <div class="toolbar">
      <button id="fitBtn" title="重新布局并适配画布">适配视图</button>
      <button id="exportBtn" title="导出当前过滤后的边数据">导出 CSV</button>
    </div>
  </header>

  <main>
    <aside>
      <div class="stats">
        <div class="stat"><strong id="nodeCount">0</strong><span>节点</span></div>
        <div class="stat"><strong id="edgeCount">0</strong><span>链路</span></div>
        <div class="stat"><strong id="hitCount">0</strong><span>访问次数</span></div>
        <div class="stat"><strong id="portCount">0</strong><span>端口</span></div>
      </div>

      <div class="section-title">筛选条件</div>
      <div class="field">
        <label for="anchorInput">中心 IP</label>
        <div class="query-row">
          <input id="anchorInput" list="anchorList" placeholder="例如 192.168.201.10">
          <button id="queryBtn" title="按当前条件查询">查询</button>
        </div>
        <datalist id="anchorList"></datalist>
      </div>
      <div class="field">
        <label for="sourceInput">源 IP</label>
        <input id="sourceInput" list="sourceList" placeholder="例如 10.0.1.92">
        <datalist id="sourceList"></datalist>
      </div>
      <div class="field">
        <label for="targetInput">目的 IP</label>
        <input id="targetInput" list="targetList" placeholder="例如 10.0.1.43">
        <datalist id="targetList"></datalist>
      </div>
      <div class="field">
        <label for="portSelect">目的端口</label>
        <select id="portSelect"></select>
      </div>
      <div class="field">
        <label for="protocolSelect">应用层协议</label>
        <select id="protocolSelect"></select>
      </div>
      <div class="field">
        <label for="matchMode">匹配模式</label>
        <select id="matchMode">
          <option value="fuzzy" selected>模糊匹配</option>
          <option value="exact">精确匹配</option>
        </select>
      </div>
      <div class="field">
        <label for="expandDepth">层级展开</label>
        <select id="expandDepth">
          <option value="0">0 层</option>
          <option value="1">1 层</option>
          <option value="2" selected>2 层</option>
          <option value="3">3 层</option>
          <option value="4">4 层</option>
        </select>
      </div>
      <div class="field">
        <label for="minCount">最小访问次数</label>
        <div class="range-row">
          <input id="minCount" type="range" min="1" max="1" value="1">
          <input id="minCountNumber" type="number" min="1" value="1">
        </div>
      </div>
      <div class="field">
        <label for="topLimit">最多显示链路</label>
        <select id="topLimit">
          <option value="50">Top 50</option>
          <option value="100">Top 100</option>
          <option value="200">Top 200</option>
          <option value="500">Top 500</option>
          <option value="0" selected>全部</option>
        </select>
      </div>

      <div class="legend">
        <div class="legend-item"><span class="dot" style="background: var(--source)"></span>仅作为源 IP</div>
        <div class="legend-item"><span class="dot" style="background: var(--target)"></span>仅作为目的 IP</div>
        <div class="legend-item"><span class="dot" style="background: var(--both)"></span>同时为源和目的</div>
        <div class="legend-item"><span class="dot" style="background: var(--accent)"></span>线越粗，访问次数越高</div>
      </div>
    </aside>

    <section class="content">
      <div class="graph-panel">
        <div id="graph" role="img" aria-label="源目的 IP 关系链路图"></div>
        <div id="empty" class="empty">当前筛选条件下没有链路</div>
      </div>

      <div class="table-panel">
        <div class="table-head">
          <strong>链路明细</strong>
          <span id="tableSummary"></span>
        </div>
        <div style="overflow:auto; max-height: 320px;">
          <table>
            <thead>
              <tr>
                <th>源 IP</th>
                <th>目的 IP</th>
                <th>目的端口</th>
                <th>协议</th>
                <th>次数</th>
              </tr>
            </thead>
            <tbody id="edgeTable"></tbody>
          </table>
        </div>
        <div class="footer-note">提示：节点可拖动；鼠标悬停节点或链路可查看聚合信息。</div>
      </div>
    </section>
  </main>
  <div id="tooltip" class="tooltip"></div>

  <script>
    const DATA = {data_json};
    const state = {{
      filteredEdges: [],
      nodes: [],
      edges: [],
      chart: null
    }};

    const els = {{
      svg: document.getElementById('graph'),
      empty: document.getElementById('empty'),
      tooltip: document.getElementById('tooltip'),
      anchor: document.getElementById('anchorInput'),
      source: document.getElementById('sourceInput'),
      target: document.getElementById('targetInput'),
      anchorList: document.getElementById('anchorList'),
      port: document.getElementById('portSelect'),
      protocol: document.getElementById('protocolSelect'),
      matchMode: document.getElementById('matchMode'),
      expandDepth: document.getElementById('expandDepth'),
      minCount: document.getElementById('minCount'),
      minCountNumber: document.getElementById('minCountNumber'),
      topLimit: document.getElementById('topLimit'),
      sourceList: document.getElementById('sourceList'),
      targetList: document.getElementById('targetList'),
      nodeCount: document.getElementById('nodeCount'),
      edgeCount: document.getElementById('edgeCount'),
      hitCount: document.getElementById('hitCount'),
      portCount: document.getElementById('portCount'),
      table: document.getElementById('edgeTable'),
      tableSummary: document.getElementById('tableSummary'),
      fitBtn: document.getElementById('fitBtn'),
      queryBtn: document.getElementById('queryBtn'),
      exportBtn: document.getElementById('exportBtn')
    }};

    const fmt = new Intl.NumberFormat('zh-CN');
    const maxCount = Math.max(1, ...DATA.edges.map(e => e.count));
    const palette = ['#1e88e5', '#00a884', '#9a5cff', '#ff8a3d', '#d44950', '#607087', '#17a2b8', '#bf7d00'];

    function escapeHtml(value) {{
      return String(value).replace(/[&<>"']/g, ch => ({{
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
      }}[ch]));
    }}

    function populateSelect(select, values, label) {{
      select.innerHTML = '<option value="">全部' + label + '</option>' +
        values.map(v => '<option value="' + escapeHtml(v) + '">' + escapeHtml(v) + '</option>').join('');
    }}

    function populateDatalist(list, values) {{
      list.innerHTML = values.map(v => '<option value="' + escapeHtml(v) + '"></option>').join('');
    }}

    function initFilters() {{
      const sources = [...new Set(DATA.edges.map(e => e.source))].sort();
      const targets = [...new Set(DATA.edges.map(e => e.target))].sort();
      const anchors = [...new Set([...sources, ...targets])].sort();
      populateDatalist(els.anchorList, anchors);
      populateDatalist(els.sourceList, sources);
      populateDatalist(els.targetList, targets);
      populateSelect(els.port, DATA.ports.map(x => x.value), '端口');
      populateSelect(els.protocol, DATA.protocols.map(x => x.value), '协议');
      els.minCount.max = String(maxCount);
      els.minCountNumber.max = String(maxCount);
      els.minCount.value = String(Math.max(1, DATA.suggested_min_count));
      els.minCountNumber.value = els.minCount.value;
    }}

    function matchValue(value, query) {{
      if (!query) return true;
      const left = String(value || '').trim().toLowerCase();
      const right = String(query || '').trim().toLowerCase();
      if (!left || !right) return false;
      if (els.matchMode.value === 'exact') return left === right;
      return left.includes(right);
    }}

    function filteredEdges(baseEdges = DATA.edges) {{
      const source = els.source.value.trim();
      const target = els.target.value.trim();
      const port = els.port.value;
      const protocol = els.protocol.value;
      const minCount = Number(els.minCountNumber.value || 1);
      const limit = Number(els.topLimit.value || 0);

      let rows = baseEdges.filter(e => {{
        if (!matchValue(e.source, source)) return false;
        if (!matchValue(e.target, target)) return false;
        if (port && e.port !== port) return false;
        if (protocol && e.protocol !== protocol) return false;
        if (e.count < minCount) return false;
        return true;
      }});
      rows.sort((a, b) => b.count - a.count);
      if (limit > 0) rows = rows.slice(0, limit);
      return rows;
    }}

    function applyNonAnchorFilters(baseEdges) {{
      const source = els.source.value.trim();
      const target = els.target.value.trim();
      const port = els.port.value;
      const protocol = els.protocol.value;
      const minCount = Number(els.minCountNumber.value || 1);

      return baseEdges.filter(e => {{
        if (!matchValue(e.source, source)) return false;
        if (!matchValue(e.target, target)) return false;
        if (port && e.port !== port) return false;
        if (protocol && e.protocol !== protocol) return false;
        if (e.count < minCount) return false;
        return true;
      }});
    }}

    function buildVisibleGraph(edges) {{
      const totals = new Map();
      for (const edge of edges) {{
        if (!totals.has(edge.source)) totals.set(edge.source, {{id: edge.source, in: 0, out: 0, total: 0}});
        if (!totals.has(edge.target)) totals.set(edge.target, {{id: edge.target, in: 0, out: 0, total: 0}});
        totals.get(edge.source).out += edge.count;
        totals.get(edge.source).total += edge.count;
        totals.get(edge.target).in += edge.count;
        totals.get(edge.target).total += edge.count;
      }}
      const nodes = [...totals.values()].map(n => ({{
        ...n,
        role: n.in && n.out ? 'both' : (n.out ? 'source' : 'target')
      }})).sort((a, b) => b.total - a.total);
      return {{ nodes, edges }};
    }}

    function neighborhoodEdges(centerIds, depth) {{
      if (!centerIds.length || depth < 0) return [];
      const baseEdges = applyNonAnchorFilters(DATA.edges);
      const center = new Set(centerIds);
      if (depth === 0) return baseEdges.filter(edge => center.has(edge.source) || center.has(edge.target));

      const edgeMap = new Map();
      for (const edge of baseEdges) {{
        if (!edgeMap.has(edge.source)) edgeMap.set(edge.source, []);
        if (!edgeMap.has(edge.target)) edgeMap.set(edge.target, []);
        edgeMap.get(edge.source).push(edge);
        edgeMap.get(edge.target).push(edge);
      }}

      const visibleKeys = new Set();
      const visited = new Set(centerIds);
      let frontier = new Set(centerIds);
      for (let level = 0; level < depth; level++) {{
        const next = new Set();
        for (const id of frontier) {{
          const related = edgeMap.get(id) || [];
          for (const edge of related) {{
            visibleKeys.add(`${{edge.source}}|${{edge.target}}|${{edge.port}}|${{edge.protocol}}`);
            const neighbor = edge.source === id ? edge.target : edge.source;
            if (!visited.has(neighbor)) {{
              visited.add(neighbor);
              next.add(neighbor);
            }}
          }}
        }}
        frontier = next;
        if (!frontier.size) break;
      }}

      return baseEdges.filter(edge => visibleKeys.has(`${{edge.source}}|${{edge.target}}|${{edge.port}}|${{edge.protocol}}`));
    }}

    function collectCenterIds() {{
      const ids = new Set();
      const anchorQuery = els.anchor.value.trim();
      const sourceQuery = els.source.value.trim();
      const targetQuery = els.target.value.trim();

      for (const edge of DATA.edges) {{
        if (anchorQuery && matchValue(edge.source, anchorQuery)) ids.add(edge.source);
        if (anchorQuery && matchValue(edge.target, anchorQuery)) ids.add(edge.target);
        if (!anchorQuery && sourceQuery && matchValue(edge.source, sourceQuery)) ids.add(edge.source);
        if (!anchorQuery && targetQuery && matchValue(edge.target, targetQuery)) ids.add(edge.target);
      }}

      return [...ids];
    }}

    function colorByRole(role) {{
      if (role === 'source') return getComputedStyle(document.documentElement).getPropertyValue('--source').trim();
      if (role === 'target') return getComputedStyle(document.documentElement).getPropertyValue('--target').trim();
      return getComputedStyle(document.documentElement).getPropertyValue('--both').trim();
    }}

    function protocolColor(protocol) {{
      let hash = 0;
      for (const ch of String(protocol || '')) hash = (hash * 31 + ch.charCodeAt(0)) >>> 0;
      return palette[hash % palette.length];
    }}

    function edgeWidth(count) {{
      return 1.2 + Math.log(count + 1) / Math.log(maxCount + 1) * 8;
    }}

    function layoutLevels(nodes, edges, centerIds) {{
      if (!centerIds.length) return null;
      const nodeIds = new Set(nodes.map(node => node.id));
      const centers = centerIds.filter(id => nodeIds.has(id));
      if (!centers.length) return null;

      const inMap = new Map();
      const outMap = new Map();
      for (const edge of edges) {{
        if (!inMap.has(edge.target)) inMap.set(edge.target, new Set());
        if (!outMap.has(edge.source)) outMap.set(edge.source, new Set());
        inMap.get(edge.target).add(edge.source);
        outMap.get(edge.source).add(edge.target);
      }}

      const levels = new Map();
      centers.forEach(id => levels.set(id, 0));
      let frontier = new Set(centers);
      for (let depth = 1; depth <= 8; depth++) {{
        const next = new Set();
        for (const id of frontier) {{
          for (const parent of (inMap.get(id) || [])) {{
            if (!levels.has(parent)) {{
              levels.set(parent, -depth);
              next.add(parent);
            }}
          }}
        }}
        frontier = next;
        if (!frontier.size) break;
      }}

      frontier = new Set(centers);
      for (let depth = 1; depth <= 8; depth++) {{
        const next = new Set();
        for (const id of frontier) {{
          for (const child of (outMap.get(id) || [])) {{
            if (!levels.has(child)) {{
              levels.set(child, depth);
              next.add(child);
            }}
          }}
        }}
        frontier = next;
        if (!frontier.size) break;
      }}

      const groups = new Map();
      for (const node of nodes) {{
        const level = levels.has(node.id) ? levels.get(node.id) : 0;
        if (!groups.has(level)) groups.set(level, []);
        groups.get(level).push(node.id);
      }}

      const levelKeys = [...groups.keys()].sort((a, b) => a - b);
      const width = Math.max(900, els.svg.clientWidth || 900);
      const height = Math.max(560, els.svg.clientHeight || 560);
      const gapX = Math.min(260, Math.max(130, width / Math.max(3, levelKeys.length + 1)));
      const centerX = width / 2;
      const centerY = height / 2;
      const result = new Map();

      for (const level of levelKeys) {{
        const ids = groups.get(level).sort((a, b) => {{
          const na = nodes.find(n => n.id === a);
          const nb = nodes.find(n => n.id === b);
          return (nb?.total || 0) - (na?.total || 0);
        }});
        const usableHeight = Math.max(220, height - 140);
        const gapY = ids.length > 1 ? Math.min(92, usableHeight / (ids.length - 1)) : 0;
        const startY = centerY - gapY * (ids.length - 1) / 2;
        ids.forEach((id, idx) => {{
          result.set(id, {{
            x: centerX + level * gapX,
            y: ids.length === 1 ? centerY : startY + idx * gapY,
            fixed: centers.includes(id)
          }});
        }});
      }}

      return result;
    }}

    function showTip(event, content) {{
      els.tooltip.innerHTML = content;
      els.tooltip.style.display = 'block';
      els.tooltip.style.left = Math.min(window.innerWidth - 360, event.clientX + 14) + 'px';
      els.tooltip.style.top = Math.min(window.innerHeight - 160, event.clientY + 14) + 'px';
    }}

    function hideTip() {{
      els.tooltip.style.display = 'none';
    }}

    function renderTable(edges) {{
      els.table.innerHTML = edges.slice(0, 300).map(edge => `
        <tr>
          <td>${{escapeHtml(edge.source)}}</td>
          <td>${{escapeHtml(edge.target)}}</td>
          <td><span class="pill">${{escapeHtml(edge.port)}}</span></td>
          <td>${{escapeHtml(edge.protocol)}}</td>
          <td>${{fmt.format(edge.count)}}</td>
        </tr>
      `).join('');
      els.tableSummary.textContent = `显示 ${{fmt.format(Math.min(edges.length, 300))}} / ${{fmt.format(edges.length)}} 条`;
    }}

    function updateStats(nodes, edges) {{
      els.nodeCount.textContent = fmt.format(nodes.length);
      els.edgeCount.textContent = fmt.format(edges.length);
      els.hitCount.textContent = fmt.format(edges.reduce((sum, e) => sum + e.count, 0));
      els.portCount.textContent = fmt.format(new Set(edges.map(e => e.port)).size);
    }}

    function buildSeries(nodes, edges, centerIds = []) {{
      const categories = [
        {{ name: '源 IP', itemStyle: {{ color: getComputedStyle(document.documentElement).getPropertyValue('--source').trim() }} }},
        {{ name: '目的 IP', itemStyle: {{ color: getComputedStyle(document.documentElement).getPropertyValue('--target').trim() }} }},
        {{ name: '双向节点', itemStyle: {{ color: getComputedStyle(document.documentElement).getPropertyValue('--both').trim() }} }},
      ];

      const levelLayout = layoutLevels(nodes, edges, centerIds);
      const nodeData = nodes.map(node => {{
        const pos = levelLayout ? levelLayout.get(node.id) : null;
        return {{
        id: node.id,
        name: node.id,
        value: node.total,
        _in: node.in,
        _out: node.out,
        category: node.role === 'source' ? 0 : node.role === 'target' ? 1 : 2,
        symbolSize: 10 + Math.log(node.total + 1) / Math.log(maxCount + 1) * 26,
        itemStyle: {{ color: colorByRole(node.role) }},
        label: {{
          show: true,
          position: 'right',
          color: '#243246',
          fontSize: 12
        }},
        emphasis: {{
          focus: 'adjacency',
          label: {{ show: true }}
        }},
        ...(pos ? {{ x: pos.x, y: pos.y, fixed: pos.fixed }} : {{}})
      }};
      }});

      const links = edges.map(edge => ({{
        source: edge.source,
        target: edge.target,
        value: edge.count,
        port: edge.port,
        protocol: edge.protocol,
        lineStyle: {{
          width: edgeWidth(edge.count),
          color: protocolColor(edge.protocol),
          curveness: 0.18,
          opacity: 0.85
        }},
        symbol: ['none', 'arrow'],
        symbolSize: 8,
        label: {{
          show: edge.count >= Math.max(DATA.suggested_min_count, maxCount * 0.08),
          formatter: () => `${{edge.port}} / ${{edge.protocol}}`,
          color: '#516174',
          fontSize: 10,
          backgroundColor: 'rgba(255,255,255,.82)',
          padding: [2, 4],
          borderRadius: 4
        }},
      }}));

      return {{ categories, nodeData, links, isLayered: Boolean(levelLayout) }};
    }}

    function renderChart(nodes, edges, centerIds = []) {{
      const {{ categories, nodeData, links, isLayered }} = buildSeries(nodes, edges, centerIds);
      if (!state.chart) {{
        state.chart = echarts.init(els.svg, null, {{ renderer: 'canvas' }});
      }}

      const option = {{
        backgroundColor: 'transparent',
        animationDuration: 700,
        animationEasing: 'cubicOut',
        tooltip: {{
          trigger: 'item',
          backgroundColor: 'rgba(23, 34, 53, .95)',
          borderWidth: 0,
          textStyle: {{ color: '#fff' }},
          formatter: params => {{
            if (params.dataType === 'node') {{
              return `<b>${{escapeHtml(params.data.id || params.data.name)}}</b><br>流出：${{fmt.format(params.data._out || 0)}}<br>流入：${{fmt.format(params.data._in || 0)}}<br>总计：${{fmt.format(params.data.value || 0)}}`;
            }}
            if (params.dataType === 'edge') {{
              return `<b>${{escapeHtml(params.data.source)}} → ${{escapeHtml(params.data.target)}}</b><br>端口：${{escapeHtml(params.data.port || '')}}<br>协议：${{escapeHtml(params.data.protocol || '')}}<br>次数：${{fmt.format(params.data.value || 0)}}`;
            }}
            return '';
          }}
        }},
        legend: {{
          data: categories.map(c => c.name),
          top: 10,
          left: 14,
          itemWidth: 14,
          itemHeight: 10,
          textStyle: {{ color: '#516174' }}
        }},
        graphic: {{
          elements: edges.length ? [] : [{{
            type: 'text',
            left: 'center',
            top: 'middle',
            style: {{
              text: '当前筛选条件下没有链路',
              fill: '#607087',
              fontSize: 16,
              fontWeight: 600
            }}
          }}]
        }},
        series: [{{
          type: 'graph',
          layout: isLayered ? 'none' : 'force',
          roam: true,
          draggable: true,
          symbol: 'circle',
          categories,
          data: nodeData,
          links,
          edgeSymbol: ['none', 'arrow'],
          edgeSymbolSize: 8,
          label: {{
            show: true,
            position: 'right',
            formatter: '{{b}}',
            color: '#243246'
          }},
          lineStyle: {{
            color: 'source',
            curveness: 0.18
          }},
          emphasis: {{
            focus: 'adjacency',
            lineStyle: {{ width: 4, opacity: 1 }},
            itemStyle: {{ shadowBlur: 12, shadowColor: 'rgba(0,0,0,.2)' }}
          }},
          force: isLayered ? undefined : {{
            initLayout: 'circular',
            repulsion: 120,
            gravity: 0.24,
            edgeLength: [54, 112],
            friction: 0.62,
            layoutAnimation: true
          }}
        }}]
      }};

      state.chart.setOption(option, true);
      state.chart.resize();
    }}

    function applyFilters(resetLayout = true) {{
      const depth = Number(els.expandDepth.value || 0);
      const seeds = collectCenterIds();
      const scopedEdges = seeds.length ? neighborhoodEdges(seeds, depth) : DATA.edges;
      const edges = filteredEdges(scopedEdges);
      const graph = buildVisibleGraph(edges);
      state.filteredEdges = edges;
      state.nodes = graph.nodes;
      state.edges = graph.edges;
      els.empty.style.display = edges.length ? 'none' : 'flex';
      updateStats(graph.nodes, edges);
      renderTable(edges);
      renderChart(graph.nodes, graph.edges, seeds);
      if (resetLayout && state.chart) {{
        state.chart.resize();
      }}
    }}

    function exportCsv() {{
      const rows = [['源IP', '目的IP', '目的端口', '应用层协议', '次数'], ...state.filteredEdges.map(e => [e.source, e.target, e.port, e.protocol, e.count])];
      const csv = rows.map(row => row.map(value => `"${{String(value).replaceAll('"', '""')}}"`).join(',')).join('\\n');
      const blob = new Blob(['\\ufeff' + csv], {{type: 'text/csv;charset=utf-8'}});
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'filtered_link_edges.csv';
      a.click();
      URL.revokeObjectURL(url);
    }}

    function debounce(fn, wait = 120) {{
      let timer = null;
      return (...args) => {{
        clearTimeout(timer);
        timer = setTimeout(() => fn(...args), wait);
      }};
    }}

    initFilters();
    els.queryBtn.addEventListener('click', () => applyFilters(true));
    els.anchor.addEventListener('keydown', event => {{
      if (event.key === 'Enter') applyFilters(true);
    }});
    els.minCount.addEventListener('input', () => {{
      els.minCountNumber.value = els.minCount.value;
    }});
    els.minCountNumber.addEventListener('input', () => {{
      els.minCount.value = els.minCountNumber.value || 1;
    }});
    els.fitBtn.addEventListener('click', () => {{
      if (state.chart) state.chart.resize();
      applyFilters(true);
    }});
    els.exportBtn.addEventListener('click', exportCsv);
    window.addEventListener('resize', debounce(() => {{
      if (state.chart) state.chart.resize();
    }}, 180));
    applyFilters(true);
  </script>
</body>
</html>
"""


def write_html(input_path, output_path, rows, sheet_name, columns, title):
    headers = rows[0]
    data_rows = rows[1:]
    nodes, edges, protocols, ports, raw_count = build_graph(data_rows, columns)
    if not edges:
        raise ValueError("未能从 Excel 中解析出有效的源/目的 IP 链路")

    counts = [edge["count"] for edge in edges]
    payload = {
        "title": title,
        "source_name": os.path.basename(input_path),
        "sheet_name": sheet_name,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "headers": headers,
        "columns": {key: headers[idx] if idx is not None and idx < len(headers) else "" for key, idx in columns.items()},
        "summary": {
            "excel_rows": raw_count,
            "node_count": len(nodes),
            "edge_count": len(edges),
            "total_count": sum(counts),
        },
        "suggested_min_count": max(1, int(percentile(counts, 0.55))),
        "nodes": nodes,
        "edges": edges,
        "protocols": [{"value": key, "count": value} for key, value in protocols.most_common()],
        "ports": [{"value": key, "count": value} for key, value in ports.most_common()],
    }

    output_path.write_text(render_html(payload), encoding="utf-8")
    return payload


def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="从日志检索 Excel 生成源目的 IP 关系互访链路静态 HTML。"
    )
    parser.add_argument("excel", help="输入 xlsx 文件路径")
    parser.add_argument("-o", "--output", default="relationship_graph.html", help="输出 HTML 文件路径")
    parser.add_argument("--sheet", help="指定 sheet 名称，默认读取第一个 sheet")
    parser.add_argument("--title", default="关系互访链路图", help="HTML 页面标题")
    parser.add_argument("--src-col", help="源 IP 列名，默认自动识别")
    parser.add_argument("--dst-col", help="目的 IP 列名，默认自动识别")
    parser.add_argument("--port-col", help="目的端口列名，默认自动识别")
    parser.add_argument("--protocol-col", help="应用层协议列名，默认自动识别")
    parser.add_argument("--count-col", help="次数列名，默认自动识别")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv or sys.argv[1:])
    input_path = Path(args.excel).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    if not input_path.exists():
        raise SystemExit(f"输入文件不存在：{input_path}")
    if input_path.suffix.lower() != ".xlsx":
        raise SystemExit("当前脚本读取 .xlsx 文件；如是 .xls，请先另存为 .xlsx")

    rows, sheet_name = read_xlsx_rows(input_path, args.sheet)
    if not rows:
        raise SystemExit("Excel 中没有可读取的数据行")

    headers = rows[0]
    explicit = {
        "src": args.src_col,
        "dst": args.dst_col,
        "port": args.port_col,
        "protocol": args.protocol_col,
        "count": args.count_col,
    }
    columns = {key: find_column(headers, key, explicit[key]) for key in FIELD_ALIASES}
    missing = [key for key in ("src", "dst") if columns[key] is None]
    if missing:
        raise SystemExit(f"缺少必要列：{', '.join(missing)}。表头为：{headers}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload = write_html(input_path, output_path, rows, sheet_name, columns, args.title)
    summary = payload["summary"]
    print(f"已生成：{output_path}")
    print(f"节点：{summary['node_count']}，链路：{summary['edge_count']}，访问次数：{summary['total_count']}")


if __name__ == "__main__":
    main()
