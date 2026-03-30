#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import html
import json
import math
import re
import sys
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


HOME = Path.home()
GRAPH_SESSION_PATH = HOME / "Library/Application Support/plutus-investment-group-desktop/runtime-store/graph_session_v1.json"

COLOR_PALETTE = ["#818cf8", "#34d399", "#fbbf24", "#a78bfa", "#f472b6", "#22d3ee", "#fb7185"]
COMPOSITION_STAGE_ORDER = [
    "Target",
    "Contact Started",
    "Contacted / Meeting Done",
    "Waiting / Ongoing",
    "Replied / Moving Forward",
]
COMPOSITION_GROUPS = [
    {"key": "vc", "label": "VC Funds", "shortLabel": "VC", "color": "#4f46e5", "childColors": ["#312e81", "#3730a3", "#4338ca", "#6366f1", "#818cf8", "#a5b4fc"]},
    {"key": "hnwi", "label": "HNWI / Angels", "shortLabel": "HNWI", "color": "#10b981", "childColors": ["#065f46", "#047857", "#059669", "#10b981", "#34d399", "#6ee7b7"]},
    {"key": "fo", "label": "Family Offices", "shortLabel": "FO", "color": "#f59e0b", "childColors": ["#92400e", "#b45309", "#d97706", "#f59e0b", "#fbbf24", "#fcd34d"]},
]
SEARCH_FIELDS = [
    "Investor",
    "Investor Name",
    "Name",
    "Email",
    "Type",
    "Type of Client",
    "Description",
    "Size of Investment",
    "Investment Size",
    "Replied",
    "Moving Forward",
    "Meeting with Company",
    "Meeting",
    "Call/Meeting",
    "Call",
    "Contact",
]
CONTACT_FIELD_KEYS = ["Contact\xa0", "Contact", "Contact ", "Contact/Call", "Contact / Call"]
CALL_FIELD_KEYS = ["Call/Meeting", "Call"]
MEETING_FIELD_KEYS = ["Meeting\xa0", "Meeting with Company", "Meeting"]
FORWARD_FIELD_KEYS = ["Reply", "Replied", "Moving Forward"]
NAVRO_KEY_TYPE_LABELS = {
    "CVC": "Corporate Venture Capital",
    "VC": "Venture Capital",
    "PB": "Private Bank",
    "SFO": "Single Family Office",
    "MFO": "Multi Family Office",
    "AS": "Angel Syndicate",
    "PF": "Pension Funds",
    "AM": "Asset Manager",
    "INS": "Insurance Company",
    "HNW": "High Net Worth Individual",
}
NAVRO_CLIENT_TYPE_TO_KEY = {value.lower(): key for key, value in NAVRO_KEY_TYPE_LABELS.items()}
NAVRO_VC_KEYS = {"CVC", "VC", "PB", "PF", "INS", "AM"}
NAVRO_HNWI_KEYS = {"AS", "HNW"}

NS_MAIN = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}


def read_graph_token() -> str:
    raw = json.loads(GRAPH_SESSION_PATH.read_text())
    token = str(raw.get("accessToken") or "").strip()
    if not token:
        raise RuntimeError(f"No accessToken found in {GRAPH_SESSION_PATH}")
    return token


def encode_share_link(share_url: str) -> str:
    encoded = base64.b64encode(share_url.encode("utf-8")).decode("ascii")
    encoded = encoded.replace("+", "-").replace("/", "_").rstrip("=")
    return f"u!{encoded}"


def download_workbook(share_url: str, output_path: Path) -> None:
    token = read_graph_token()
    graph_url = f"https://graph.microsoft.com/v1.0/shares/{encode_share_link(share_url)}/driveItem/content"
    request = urllib.request.Request(graph_url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(request) as response:
        output_path.write_bytes(response.read())


def column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch.upper()) - 64)
    return value - 1


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        xml_bytes = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(xml_bytes)
    values = []
    for si in root.findall("main:si", NS_MAIN):
        text_parts = []
        for node in si.iter():
            if node.tag == f"{{{NS_MAIN['main']}}}t" and node.text:
                text_parts.append(node.text)
        values.append("".join(text_parts))
    return values


def load_sheet_targets(zf: zipfile.ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {}
    for rel in rels_root.findall("rel:Relationship", NS_REL):
        rel_targets[rel.attrib["Id"]] = rel.attrib["Target"]

    targets = {}
    for sheet in workbook_root.findall("main:sheets/main:sheet", NS_MAIN):
        name = sheet.attrib.get("name", "")
        rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
        target = rel_targets.get(rel_id)
        if target:
            targets[name] = "xl/" + target.lstrip("/")
    return targets


def read_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    value_node = cell.find("main:v", NS_MAIN)
    if cell_type == "inlineStr":
        return "".join(t.text or "" for t in cell.findall(".//main:t", NS_MAIN))
    if value_node is None or value_node.text is None:
        return ""
    raw = value_node.text
    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except (ValueError, IndexError):
            return ""
    return raw


def parse_sheet_rows(zf: zipfile.ZipFile, sheet_path: str, shared_strings: list[str]) -> list[dict[int, str]]:
    root = ET.fromstring(zf.read(sheet_path))
    rows: list[dict[int, str]] = []
    for row in root.findall("main:sheetData/main:row", NS_MAIN):
        values: dict[int, str] = {}
        for cell in row.findall("main:c", NS_MAIN):
            ref = cell.attrib.get("r", "")
            if not ref:
                continue
            values[column_index(ref)] = read_cell_value(cell, shared_strings)
        rows.append(values)
    return rows


def detect_header_row(rows: list[dict[int, str]], max_scan_rows: int = 20) -> int:
    for index, row in enumerate(rows[:max_scan_rows]):
        values = [str(value).strip().lower() for value in row.values() if str(value).strip()]
        if "investor" in values:
            return index
    return 3


def rows_to_objects(rows: list[dict[int, str]], header_index: int) -> list[dict[str, str]]:
    if not rows:
        return []
    headers_row = rows[header_index] if header_index < len(rows) else {}
    max_col = max((max(r.keys()) for r in rows if r), default=-1)
    headers = [str(headers_row.get(i, "") or "") for i in range(max_col + 1)]

    items: list[dict[str, str]] = []
    for row in rows[header_index + 1:]:
        item: dict[str, str] = {}
        has_data = False
        for idx, header in enumerate(headers):
            if not header:
                continue
            value = str(row.get(idx, "") or "")
            if value:
                has_data = True
            item[header] = value
        if has_data:
            items.append(item)
    return items


def normalize_text(value: str) -> str:
    return str(value or "").strip().lower()


def get_navro_client_type(row: dict[str, str]) -> str:
    for key in ("Type of Client", "Type of client", "Type Of Client", "type of client", "Type"):
        if key in row and str(row[key]).strip():
            return str(row[key]).strip()
    return ""


def get_dashboard_type_key(row: dict[str, str]) -> str:
    for key in ("KEY", "Key", "key"):
        if key in row and str(row[key]).strip():
            return str(row[key]).strip().upper()
    return ""


def normalize_navro_client_type_code(value: str) -> str:
    normalized = str(value or "").strip().upper()
    if normalized in NAVRO_KEY_TYPE_LABELS:
        return normalized
    return NAVRO_CLIENT_TYPE_TO_KEY.get(normalize_text(value), "")


def get_navro_client_type_code(row: dict[str, str]) -> str:
    code = normalize_navro_client_type_code(get_navro_client_type(row))
    if code:
        return code
    return normalize_navro_client_type_code(get_dashboard_type_key(row))


def get_investor_type_label(row: dict[str, str]) -> str:
    code = get_navro_client_type_code(row)
    if code and code in NAVRO_KEY_TYPE_LABELS:
        return NAVRO_KEY_TYPE_LABELS[code]
    client_type = get_navro_client_type(row)
    if client_type:
        return client_type
    return str(row.get("Type") or "Unknown").strip() or "Unknown"


def is_navro_vc_row(row: dict[str, str]) -> bool:
    return get_navro_client_type_code(row) in NAVRO_VC_KEYS


def is_hnwi_or_angel_row(row: dict[str, str]) -> bool:
    return get_navro_client_type_code(row) in NAVRO_HNWI_KEYS


def get_status_badge(item: dict[str, str]) -> str:
    forward = get_field_value(item, FORWARD_FIELD_KEYS)
    meeting = get_field_value(item, MEETING_FIELD_KEYS)
    call = get_field_value(item, [*CALL_FIELD_KEYS, *CONTACT_FIELD_KEYS])

    if forward == "yes":
        return '<span class="stage-pill badge-green">Replied / Mov. Forward</span>'
    if forward == "waiting" or meeting == "waiting" or call == "waiting":
        return '<span class="stage-pill badge-amber">Waiting / Ongoing</span>'
    if forward == "no":
        return ""
    if meeting == "yes":
        return '<span class="stage-pill badge-neutral">Contacted / Met</span>'
    if call == "yes":
        return '<span class="stage-pill badge-neutral">Contact Started</span>'
    return '<span class="stage-pill" style="color:var(--text-dim); opacity:0.5;">Target</span>'


def get_status_text(item: dict[str, str]) -> str:
    forward = get_field_value(item, FORWARD_FIELD_KEYS)
    meeting = get_field_value(item, MEETING_FIELD_KEYS)
    call = get_field_value(item, [*CALL_FIELD_KEYS, *CONTACT_FIELD_KEYS])

    if forward == "yes":
        return "Replied / Moving Forward"
    if forward == "waiting" or meeting == "waiting" or call == "waiting":
        return "Waiting / Ongoing"
    if forward == "no":
        return ""
    if meeting == "yes":
        return "Contacted / Meeting Done"
    if call == "yes":
        return "Contact Started"
    return "Target"


def build_search_index(item: dict[str, str]) -> str:
    values = []
    for field in SEARCH_FIELDS:
        value = str(item.get(field) or "").strip()
        if value:
            values.append(value)
    return " ".join(values).lower()


def prepare_rows(rows: list[dict[str, str]]) -> None:
    for item in rows:
        item["__dashboardName"] = str(item.get("Investor") or item.get("Investor Name") or item.get("Name") or "").strip()
        item["__dashboardSize"] = str(item.get("Size of Investment") or item.get("Investment Size") or "–").strip() or "–"
        item["__dashboardEmail"] = str(item.get("Email") or "–").strip() or "–"
        item["__dashboardType"] = get_investor_type_label(item)
        if not str(item.get("Type") or "").strip():
            item["Type"] = item["__dashboardType"]
        item["__dashboardStatusText"] = get_status_text(item)
        item["__dashboardStatusBadge"] = get_status_badge(item)
        item["__dashboardSearchIndex"] = build_search_index(item)


def get_field_value(row: dict[str, str], keys: list[str]) -> str:
    for key in keys:
        if key in row:
            value = str(row.get(key) or "").strip().lower()
            if value:
                return value
    return ""


def get_contact_status(row: dict[str, str]) -> str:
    return get_field_value(row, CONTACT_FIELD_KEYS)


def get_call_status(row: dict[str, str]) -> str:
    return get_field_value(row, CALL_FIELD_KEYS)


def stage_value(row: dict[str, str], columns: list[str]) -> str:
    for column in columns:
        value = row.get(column)
        if value is not None and str(value).strip() != "":
            return str(value).lower().strip()
    return ""


def get_min_numeric_value(raw_value: str) -> float:
    if not raw_value:
        return -1
    match = re.search(r"([0-9.]+)\s*([kKmM])", str(raw_value))
    if not match:
        return -1
    value = float(match.group(1))
    unit = match.group(2).upper()
    if unit == "K":
        value *= 1000
    if unit == "M":
        value *= 1000000
    return value


def polar_point(cx: float, cy: float, radius: float, angle: float) -> tuple[float, float]:
    radians = math.radians(angle - 90)
    return cx + radius * math.cos(radians), cy + radius * math.sin(radians)


def build_arc_path(cx: float, cy: float, inner_radius: float, outer_radius: float, start_angle: float, end_angle: float) -> str:
    outer_start = polar_point(cx, cy, outer_radius, start_angle)
    outer_end = polar_point(cx, cy, outer_radius, end_angle)
    inner_end = polar_point(cx, cy, inner_radius, end_angle)
    inner_start = polar_point(cx, cy, inner_radius, start_angle)
    large_arc = 1 if end_angle - start_angle > 180 else 0
    return (
        f"M {outer_start[0]:.3f} {outer_start[1]:.3f} "
        f"A {outer_radius} {outer_radius} 0 {large_arc} 1 {outer_end[0]:.3f} {outer_end[1]:.3f} "
        f"L {inner_end[0]:.3f} {inner_end[1]:.3f} "
        f"A {inner_radius} {inner_radius} 0 {large_arc} 0 {inner_start[0]:.3f} {inner_start[1]:.3f} Z"
    )


def build_composition_rows(raw_data: dict[str, list[dict[str, str]]], group_key: str) -> list[dict[str, str]]:
    if group_key == "vc":
        return list(raw_data["vc"])
    if group_key == "hnwi":
        return [row for row in raw_data["fo"] if is_hnwi_or_angel_row(row)]
    if group_key == "fo":
        return [row for row in raw_data["fo"] if not is_hnwi_or_angel_row(row)]
    return []


def build_composition_hierarchy(raw_data: dict[str, list[dict[str, str]]]) -> list[dict[str, object]]:
    groups = []
    for meta in COMPOSITION_GROUPS:
        rows = build_composition_rows(raw_data, meta["key"])
        stage_counts = [{"label": label, "count": 0} for label in COMPOSITION_STAGE_ORDER]
        for row in rows:
            status = get_status_text(row)
            for entry in stage_counts:
                if entry["label"] == status:
                    entry["count"] += 1
                    break
        children = []
        for index, entry in enumerate(stage_counts):
            if entry["count"] <= 0:
                continue
            children.append({
                "label": entry["label"],
                "count": entry["count"],
                "shortLabel": entry["label"].replace(" / Moving Forward", "").replace(" / Meeting Done", "").replace(" / Ongoing", ""),
                "color": meta["childColors"][index % len(meta["childColors"])],
            })
        if rows:
            groups.append({
                **meta,
                "count": len(rows),
                "children": children,
            })
    return groups


def build_composition_svg_and_legend(raw_data: dict[str, list[dict[str, str]]]) -> tuple[str, str]:
    hierarchy = build_composition_hierarchy(raw_data)
    total = sum(group["count"] for group in hierarchy)
    if not total:
        return "", ""

    start_angle = 0.0
    svg_parts: list[str] = []

    for group in hierarchy:
        span = (group["count"] / total) * 360.0
        end_angle = start_angle + span
        svg_parts.append(
            f'<path class="composition-segment is-clickable" data-composition-group="{group["key"]}" '
            f'd="{build_arc_path(50, 50, 18, 34, start_angle, end_angle)}" '
            f'fill="{group["color"]}" aria-label="{html.escape(group["label"])}"></path>'
        )
        if span >= 18:
            x, y = polar_point(50, 50, 26, start_angle + span / 2)
            font_size = "4.2" if span > 70 else "3.2"
            svg_parts.append(f'<text x="{x:.2f}" y="{y:.2f}" class="composition-label" font-size="{font_size}">{html.escape(group["shortLabel"])} ({group["count"]})</text>')

        child_start = start_angle
        for child in group["children"]:
            child_span = (child["count"] / total) * 360.0
            child_end = child_start + child_span
            svg_parts.append(
                f'<path class="composition-segment is-clickable" data-composition-group="{group["key"]}" '
                f'data-composition-stage="{html.escape(child["label"])}" '
                f'd="{build_arc_path(50, 50, 36, 49.5, child_start, child_end)}" '
                f'fill="{child["color"]}" aria-label="{html.escape(group["label"] + " " + child["label"])}"></path>'
            )
            if child_span >= 14:
                x, y = polar_point(50, 50, 42.5, child_start + child_span / 2)
                font_size = "2.8" if child_span > 28 else "2.2"
                svg_parts.append(f'<text x="{x:.2f}" y="{y:.2f}" class="composition-label" font-size="{font_size}">{html.escape(child["shortLabel"])}</text>')
            child_start = child_end

        start_angle = end_angle

    svg_parts.append('<circle cx="50" cy="50" r="11.5" fill="rgba(255,255,255,0.06)" stroke="rgba(148,163,184,0.16)"></circle>')
    svg_parts.append('<text x="50" y="48.6" class="composition-center-note">pool</text>')
    svg_parts.append(f'<text x="50" y="55.3" class="composition-center-total">{total}</text>')

    legend_parts = []
    for group in hierarchy:
        child_markup = []
        for child in group["children"]:
            child_markup.append(
                f'''
                                <div class="composition-legend-item composition-interactive" data-composition-group="{group["key"]}" data-composition-stage="{html.escape(child["label"])}" role="button" tabindex="0" aria-pressed="false">
                                    <span class="composition-legend-swatch" style="background:{child["color"]}"></span>
                                    <span>{html.escape(child["label"])}</span>
                                    <strong>{child["count"]}</strong>
                                </div>
                            '''
            )
        legend_parts.append(
            f'''
                        <div class="composition-legend-group">
                            <div class="composition-legend-head composition-interactive" data-composition-group="{group["key"]}" role="button" tabindex="0" aria-pressed="false">
                                <div class="composition-legend-title">
                                    <span class="composition-legend-dot" style="background:{group["color"]}"></span>
                                    <span>{html.escape(group["label"])}</span>
                                </div>
                                <span class="composition-legend-total">{group["count"]}</span>
                            </div>
                            {"".join(child_markup)}
                        </div>
                    '''
        )

    return "".join(svg_parts), "".join(legend_parts)


def build_range_bars_markup(combined: list[dict[str, str]]) -> str:
    buckets = [
        {"label": "< $100K", "min": 0, "max": 100000},
        {"label": "$100K - $500K", "min": 100000, "max": 500000},
        {"label": "$500K - $1M", "min": 500000, "max": 1000000},
        {"label": "$1M - $5M", "min": 1000000, "max": 5000000},
        {"label": "$5M - $10M", "min": 5000000, "max": 10000000},
        {"label": "> $10M", "min": 10000000, "max": float("inf")},
    ]
    counts = [0 for _ in buckets]

    for row in combined:
        raw = row.get("Size of Investment") or row.get("Investment Size") or ""
        value = get_min_numeric_value(raw)
        if value == -1:
            continue
        for index, bucket in enumerate(buckets):
            if bucket["min"] <= value < bucket["max"]:
                counts[index] += 1
                break

    max_count = max(max(counts), 1)
    parts = []
    for bucket, count in zip(buckets, counts):
        width = (count / max_count) * 100
        parts.append(
            f'''
                    <div>
                        <div style="display:flex; justify-content:space-between; font-size:0.85rem; margin-bottom:6px;">
                            <span style="font-weight:500;">{html.escape(bucket["label"])}</span>
                            <span style="color:var(--text-dim);">{count}</span>
                        </div>
                        <div style="height:6px; background:rgba(255,255,255,0.05); border-radius:10px; overflow:hidden;">
                            <div style="width:{width}%; height:100%; background:var(--accent); border-radius:10px;"></div>
                        </div>
                    </div>
                '''
        )
    return "".join(parts)


def replace_between(text: str, start_marker: str, end_marker: str, replacement: str) -> str:
    start = text.find(start_marker)
    if start == -1:
        raise RuntimeError(f"Could not find start marker: {start_marker}")
    end = text.find(end_marker, start)
    if end == -1:
        raise RuntimeError(f"Could not find end marker: {end_marker}")
    return text[: start + len(start_marker)] + replacement + text[end:]


def update_stat_value(text: str, element_id: str, value: str) -> str:
    pattern = rf'(<div class="stat-value(?: [^"]+)?" id="{re.escape(element_id)}">)(.*?)(</div>)'
    updated, count = re.subn(pattern, lambda match: f"{match.group(1)}{value}{match.group(3)}", text, count=1, flags=re.DOTALL)
    if count != 1:
        raise RuntimeError(f"Could not update stat value for {element_id}")
    return updated


def update_stat_label(text: str, element_id: str, value: str) -> str:
    pattern = rf'(<div class="stat-label" id="{re.escape(element_id)}">)(.*?)(</div>)'
    updated, count = re.subn(pattern, lambda match: f"{match.group(1)}{value}{match.group(3)}", text, count=1, flags=re.DOTALL)
    if count != 1:
        raise RuntimeError(f"Could not update stat label for {element_id}")
    return updated


def build_snapshot_state(raw_data: dict[str, list[dict[str, str]]], type_colors: dict[str, str]) -> dict[str, object]:
    def clone_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
        output = []
        for index, row in enumerate(rows):
            cloned = dict(row)
            cloned["__snapshotInvestorLabel"] = f"Investor {index + 1}"
            output.append(cloned)
        return output

    return {
        "rawData": {
            "vc": clone_rows(raw_data["vc"]),
            "fo": clone_rows(raw_data["fo"]),
        },
        "typeColors": type_colors,
        "activeType": "vc",
        "activeFilters": {"all": True, "calls": False, "meetings": False, "forward": False},
        "searchTerm": "",
        "exportOptions": {"includeInvestorDetails": True},
    }


def generate_snapshot(template_path: Path, workbook_path: Path, output_path: Path) -> dict[str, int]:
    with zipfile.ZipFile(workbook_path) as zf:
        shared_strings = load_shared_strings(zf)
        sheet_targets = load_sheet_targets(zf)

        sheet_name = "Sheet1" if "Sheet1" in sheet_targets else next(iter(sheet_targets))
        rows = parse_sheet_rows(zf, sheet_targets[sheet_name], shared_strings)
        header_index = detect_header_row(rows)
        sheet_rows = rows_to_objects(rows, header_index)

    navro_rows = [row for row in sheet_rows if str(row.get("Investor") or "").strip()]
    raw_data = {
        "vc": [row for row in navro_rows if is_navro_vc_row(row)],
        "fo": [row for row in navro_rows if not is_navro_vc_row(row)],
    }
    prepare_rows(raw_data["vc"])
    prepare_rows(raw_data["fo"])

    combined = [*raw_data["vc"], *raw_data["fo"]]
    unique_types = []
    seen_types: set[str] = set()
    for row in combined:
        label = get_investor_type_label(row)
        if label and label not in seen_types:
            unique_types.append(label)
            seen_types.add(label)
    type_colors = {label: COLOR_PALETTE[index % len(COLOR_PALETTE)] for index, label in enumerate(unique_types)}

    total_contact_count = sum(1 for row in combined if get_contact_status(row) in {"yes", "waiting"})
    ongoing_count = sum(1 for row in combined if get_call_status(row) in {"yes", "waiting"})
    replied_count = sum(1 for row in combined if stage_value(row, FORWARD_FIELD_KEYS) == "yes")

    snapshot_state = build_snapshot_state(raw_data, type_colors)
    svg_markup, legend_markup = build_composition_svg_and_legend(raw_data)
    range_bars_markup = build_range_bars_markup(combined)

    html_text = template_path.read_text()
    html_text = update_stat_value(html_text, "k-pool", str(len(combined)))
    html_text = update_stat_value(html_text, "k-contacted", str(total_contact_count))
    html_text = update_stat_value(html_text, "k-replied", str(replied_count))
    html_text = update_stat_label(html_text, "k-meetings-label", "Ongoing")
    html_text = update_stat_value(html_text, "k-meetings", str(ongoing_count))
    html_text = replace_between(
        html_text,
        '<svg id="composition-sunburst" viewBox="0 0 100 100" aria-label="Investor composition multi-level pie">',
        "</svg>",
        svg_markup,
    )
    html_text = replace_between(
        html_text,
        '<div class="composition-legend" id="composition-legend">',
        "</div>\n                            </div>",
        legend_markup,
    )
    html_text = replace_between(
        html_text,
        '<div id="range-bars" style="display: flex; flex-direction: column; gap: 15px; margin-top: 10px;">',
        "</div>\n                        </div>\n                    </div>\n\n                    <div class=\"tabs dashboard-tabs\">",
        range_bars_markup,
    )

    state_pattern = r"const state = \{.*?\};"
    state_replacement = f"const state = {json.dumps(snapshot_state, ensure_ascii=False, separators=(',', ':'))};"
    html_text, count = re.subn(state_pattern, state_replacement, html_text, count=1, flags=re.DOTALL)
    if count != 1:
        raise RuntimeError("Could not replace embedded snapshot state")

    html_text = html_text.replace(
        "<td>${statusBadge}</td>\n              <td><a href=\"mailto:${email}\" style=\"color:var(--accent); text-decoration:none;\">${email}</a></td>",
        "<td><a href=\"mailto:${email}\" style=\"color:var(--accent); text-decoration:none;\">${email}</a></td>",
    )
    html_text = html_text.replace(
        "<td>${statusBadge}</td>\n            </tr>`;",
        "</tr>`;",
    )
    html_text = html_text.replace(
        "<td>${statusBadge}</td>\n          </tr>`;",
        "</tr>`;",
    )
    html_text = html_text.replace(
        "<tr><th>VC Funds Name</th><th>Investment Size</th><th>Stage</th><th>Contact Email</th></tr>",
        "<tr><th>VC Funds Name</th><th>Investment Size</th><th>Contact Email</th></tr>",
    )
    html_text = html_text.replace(
        "<tr><th>Investor</th><th>Investment Size</th><th>Stage</th></tr>",
        "<tr><th>Investor</th><th>Investment Size</th></tr>",
    )
    html_text = html_text.replace(
        "<tr><th>Investor Name</th><th>Investor Type</th><th>Investment Size</th><th>Stage</th></tr>",
        "<tr><th>Investor Name</th><th>Investor Type</th><th>Investment Size</th></tr>",
    )
    html_text = html_text.replace(
        "<tr><th>Investor</th><th>Investor Type</th><th>Investment Size</th><th>Stage</th></tr>",
        "<tr><th>Investor</th><th>Investor Type</th><th>Investment Size</th></tr>",
    )
    html_text = html_text.replace(
        "    if (forward === 'no') return '<span class=\"stage-pill badge-red\">Passed</span>';",
        "    if (forward === 'no') return '';",
    )

    output_path.write_text(html_text)
    return {
        "pool": len(combined),
        "contacted": total_contact_count,
        "replied": replied_count,
        "ongoing": ongoing_count,
        "vc": len(raw_data["vc"]),
        "fo": len(raw_data["fo"]),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate a standalone Navro investor dashboard snapshot from a SharePoint workbook.")
    parser.add_argument("--template", required=True, type=Path, help="Path to the attached standalone HTML template")
    parser.add_argument("--share-url", required=True, help="SharePoint workbook URL")
    parser.add_argument("--output", required=True, type=Path, help="Output standalone HTML file")
    args = parser.parse_args()

    workbook_path = args.output.with_suffix(".tmp.xlsx")

    try:
        download_workbook(args.share_url, workbook_path)
        stats = generate_snapshot(args.template, workbook_path, args.output)
    finally:
        workbook_path.unlink(missing_ok=True)

    print(json.dumps({"output": str(args.output), **stats}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
