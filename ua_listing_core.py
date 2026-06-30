"""UA listing-data builder.

No third-party spreadsheet library is required. The module reads the supplied
XLSX files directly and creates an XLSX that keeps the structure and the visual
formatting of the uploaded reference listing.
"""
from __future__ import annotations

from collections import Counter, OrderedDict, defaultdict
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable
import re
import zipfile
from xml.etree import ElementTree as ET

def escape(value: str) -> str:
    """Escape text safely for XLSX XML without importing xml.sax."""
    return (str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;"))

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCREL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKGREL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN_NS, "r": DOCREL_NS, "pr": PKGREL_NS}

OUTPUT_HEADERS = [
    "Style", "Article", "SKU", "usdis", "Style Description", "Size UA",
    "Size US DIST.", "Size EUR", "Size Scale", "Selling Season", "EAN",
    "Product Division", "Gender Description", "Size Group Description", "End Use",
    "Signature Collections", "DTC exclusive", "GHL", "Story Tier", "Fit Type",
    "Merch Department", "Class Description", "Sub Class Description", "Period",
    "Shipment Start Date", "Shipment End Date", "Launch Date", "Color Group",
    "Primary Color", "Secondary Color", "Logo Colorway", "Product Ranking", "C/O",
    "Gross Weight (kg)", "Article Lenght", "Article Width", "Article Height",
    "Volume", "Dimensions (inch)", "Dimensions (cm)", "HTS Code", "COO COUNTRY",
    "FEDAS Code", "FEDAS Size Range", "Material", "Tech Platform", "Product photo",
]

# The strings are intentionally kept in one place, because UA occasionally makes
# small header changes between seasons.
OOB_HEADERS = {
    "ean": "Article Upc With Leading Zeros",
    "article": "Article Generic",
    "description": "Article Generic Description",
    "division": "Article Mh Product Division Description",
    "gender": "Article Gender Description",
    "end_use": "Article Sub Category Description",
    "size": "Sd Article Size",
}
MATERIAL_HEADERS = {
    "article": "Article Generic",
    "style": "Style Number",
    "style_name": "Style Name",
    "size": "Article Size",
    "status": "Article Status Desc",
    "fedas": "Article Fedas Code",
    "fedas_size": "Article Fedas Size",
    "hts": "Article Site Commodity Code Import Code For Foreign Trade",
    "ean": "Article Zupc",
    "division": "Ua Division",
    "merch_class": "Merch Class",
    "merch_department": "Merch Department",
    "merch_subclass": "Merch Sub Class",
    "tech": "Article Fabric Platform",
    "end_use": "Sub Category",
    "product_family": "Product Family",
    "gender": "Gender",
    "size_group": "Article Size Group Desc",
    "length": "Article Length",
    "width": "Article Width",
    "height": "Article Height",
    "weight": "Article Gross Weight",
    "color_group": "Article Color Group",
    "primary_color": "Article Primary Color",
    "secondary_color": "Article Secondary Color",
    "logo_color": "Article Logo Color",
    "ranking": "Article Fashion Grade Desc",
    "material": "Article Fiber Code Desc",
    "country": "Factory Country",
}
LINE_HEADERS = {
    "colorway": "Colorway Number",
    "style": "Style Number",
    "style_name": "Style Name",
    "division": "UA Division",
    "gender": "Gender",
    "size_group": "Size Group",
    "category": "Category",
    "end_use": "Sub Category",
    "fashion_grade": "Fashion Grade",
    "tech": "Tech Platform",
    "signature": "Signature Collection",
    "fit": "Fit Type",
    "size_scale": "Size Scale",
    "hard_launch": "Hard Launch Date",
    "ship_start": "Shipment Start Date",
    "ship_end": "Shipment End Date",
    "period": "Period",
    "storytier": "Storytier",
    "carryover_2": "Emeacarryoverseason2",
    "carryover_1": "Emeacarryoverseason1",
    "catalog_copy": "Product Catalog Copy",
    "hero_look_name": "Hero Look Name",
}
# Hero Look Name is deliberately optional in case UA removes or renames the
# field in a later season. In that case, GHL safely defaults to NE and the
# audit records the missing commercial field.
LINE_REQUIRED_HEADERS = tuple(
    header for key, header in LINE_HEADERS.items() if key != "hero_look_name"
)

# Centric Brands files are separate licensed-product master data.  Their material
# codes can be either standard UA codes (via UA Full Article Code) or Centric
# codes such as 25UJFJM07F-001-JPC.  The builder deliberately does not match
# them to the generic UA Material Data Report.
CENTRIC_HEADERS = {
    "division": "Division Name",
    "brand": "Brand Description",
    "range": "Range Segment",
    "gender": "Gender Description",
    "age_group": "Age Group Description",
    "category": "Product Category Description",
    "material_number": "Material Number",
    "description": "Material Description",
    "ua_style": "UA Style Code",
    "ua_color": "UA Colour Code",
    "ua_article": "UA Full Article Code",
    "ean": ("EAN Number", "UPC Number"),
    "size": "Size",
    "season": "Season",
    "color": "Color Description",
    "country": "Country of Origin for PO",
    "hts": ("Commodity Code", "HTS Code"),
    "material": "Fibre Composition",
    "fedas": "FEDAS Code",
    "fedas_size": "FEDAS Size Code",
}



def norm(value: Any) -> str:
    return str(value or "").strip().upper()


def clean(value: Any) -> str:
    return str(value or "").strip()


def excel_col(n: int) -> str:
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def col_number(col: str) -> int:
    value = 0
    for char in col:
        value = value * 26 + ord(char) - 64
    return value


def _shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    strings: list[str] = []
    for item in root.findall("m:si", NS):
        strings.append("".join(node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t")))
    return strings


def _sheet_paths(zf: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    paths: dict[str, str] = {}
    for sheet in workbook.findall("m:sheets/m:sheet", NS):
        rel_id = sheet.attrib[f"{{{DOCREL_NS}}}id"]
        target = rel_map[rel_id].lstrip("/")
        paths[sheet.attrib["name"]] = target if target.startswith("xl/") else f"xl/{target}"
    return paths


def _cell_value(cell: ET.Element, strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{{{MAIN_NS}}}t"))
    value = cell.find("m:v", NS)
    if value is None:
        return ""
    raw = value.text or ""
    if cell_type == "s":
        try:
            return strings[int(raw)]
        except (ValueError, IndexError):
            return raw
    return raw


def _read_sheet_from_xlsx(file_bytes: bytes, sheet_name: str) -> list[dict[str, str]]:
    """Read a sheet as records. Header whitespace is stripped deliberately."""
    with zipfile.ZipFile(BytesIO(file_bytes)) as zf:
        paths = _sheet_paths(zf)
        if sheet_name not in paths:
            available = ", ".join(paths)
            raise ValueError(f"Sheet '{sheet_name}' was not found. Available sheets: {available}")
        root = ET.fromstring(zf.read(paths[sheet_name]))
        strings = _shared_strings(zf)

    rows: list[dict[str, str]] = []
    for row in root.findall(".//m:sheetData/m:row", NS):
        current: dict[str, str] = {}
        for cell in row.findall("m:c", NS):
            ref = cell.attrib.get("r", "")
            col = "".join(character for character in ref if character.isalpha())
            if col:
                current[col] = _cell_value(cell, strings)
        rows.append(current)

    if not rows:
        return []
    header_row = rows[0]
    columns = sorted(header_row, key=col_number)
    headers = [clean(header_row[column]) for column in columns]
    records: list[dict[str, str]] = []
    for raw in rows[1:]:
        records.append({header: clean(raw.get(column, "")) for column, header in zip(columns, headers)})
    return records


def _require_headers(records: list[dict[str, str]], required: Iterable[str], source_name: str) -> None:
    if not records:
        raise ValueError(f"{source_name} is empty.")
    available = set(records[0])
    missing = [header for header in required if header not in available]
    if missing:
        raise ValueError(
            f"{source_name} has an unexpected structure. Missing: {', '.join(missing)}."
        )


def _parse_date(value: str) -> datetime:
    value = clean(value)
    for pattern in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, pattern)
        except ValueError:
            continue
    return datetime.min


def _date_string(value: str) -> str:
    date = _parse_date(value)
    return date.strftime("%m/%d/%Y") if date != datetime.min else clean(value)


def _latest_by_date(records: list[dict[str, str]], key_header: str, date_header: str) -> dict[str, dict[str, str]]:
    latest: dict[str, tuple[datetime, int, dict[str, str]]] = {}
    for row_number, row in enumerate(records):
        key = norm(row.get(key_header))
        if not key:
            continue
        candidate = (_parse_date(row.get(date_header, "")), row_number, row)
        existing = latest.get(key)
        if existing is None or candidate[:2] >= existing[:2]:
            latest[key] = candidate
    return {key: value[2] for key, value in latest.items()}


def _select_material(candidates: list[dict[str, str]], oob: dict[str, str]) -> tuple[dict[str, str] | None, bool]:
    """Pick a deterministic material row if the report contains duplicate UPCs."""
    if not candidates:
        return None, False
    article = norm(oob[OOB_HEADERS["article"]])
    size = norm(oob[OOB_HEADERS["size"]])
    scored: list[tuple[int, dict[str, str]]] = []
    for row in candidates:
        score = 0
        if norm(row.get(MATERIAL_HEADERS["article"])) == article:
            score += 4
        if norm(row.get(MATERIAL_HEADERS["size"])) == size:
            score += 2
        if norm(row.get(MATERIAL_HEADERS["status"])) == "ACTIVE":
            score += 1
        scored.append((score, row))
    scored.sort(key=lambda pair: pair[0], reverse=True)
    best_score = scored[0][0]
    ambiguous = len([row for score, row in scored if score == best_score]) > 1
    return scored[0][1], ambiguous


def _style_from_article(article: str) -> str:
    return clean(article).split("-")[0]


def _extract_volume_and_dimensions(catalog_copy: str) -> tuple[str, str, str]:
    """Best-effort extraction from Line List copy for bags and accessories."""
    copy = clean(catalog_copy)
    if not copy:
        return "", "", ""
    volume_match = re.search(r"\bVolume:\s*([^\n]+)", copy, flags=re.IGNORECASE)
    dimension_match = re.search(r"\bDimensions(?:\s+When\s+Full)?:\s*([^\n]+)", copy, flags=re.IGNORECASE)
    volume = clean(volume_match.group(1)) if volume_match else ""
    if "/" in volume:
        volume = clean(volume.rsplit("/", 1)[-1])
    dimensions_in = clean(dimension_match.group(1)) if dimension_match else ""
    dimensions_cm = ""
    if dimensions_in:
        # Converts familiar values such as 18"H x 14"L x 2"W. We do not infer
        # values from non-standard descriptions, which is safer than producing
        # plausible but incorrect dimensions.
        figures = re.findall(r"(\d+(?:\.\d+)?)\s*\"", dimensions_in)
        if len(figures) >= 3:
            dimensions_cm = " x ".join(
                f"{float(number) * 2.54:.1f} cm".replace(".", ",") for number in figures[:3]
            )
    return volume, dimensions_in, dimensions_cm


def _reference_indexes(template_records: list[dict[str, str]]) -> dict[str, Any]:
    _require_headers(
        template_records,
        ["Article", "Style", "Size UA", "Size EUR", "Size Scale", "Product Division", "Gender Description"],
        "Reference listing",
    )
    exact: dict[tuple[str, str], dict[str, str]] = {}
    style_scale: dict[str, set[str]] = defaultdict(set)
    article_scale: dict[str, set[str]] = defaultdict(set)
    general_eu: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    for row in template_records:
        article = norm(row.get("Article"))
        style = norm(row.get("Style"))
        size = norm(row.get("Size UA"))
        if article and size:
            exact.setdefault((article, size), row)
        if article and clean(row.get("Size Scale")):
            article_scale[article].add(clean(row.get("Size Scale")))
        if style and clean(row.get("Size Scale")):
            style_scale[style].add(clean(row.get("Size Scale")))
        eu = clean(row.get("Size EUR"))
        key = (size, norm(row.get("Product Division")), norm(row.get("Gender Description")))
        if all(key) and eu:
            general_eu[key].add(eu)
    return {
        "exact": exact,
        "article_scale": article_scale,
        "style_scale": style_scale,
        "general_eu": general_eu,
    }


def _infer_size_scale(sizes: list[str]) -> str:
    """Fallback only. Output is explicitly flagged in the audit report."""
    unique = []
    for value in sizes:
        value = clean(value)
        if value and value not in unique:
            unique.append(value)
    if not unique:
        return ""

    known_order = [
        "XXS", "XS", "SM", "S", "MD", "M", "LG", "L", "XL", "XXL", "2XL", "3XL", "4XL", "5XL",
        "YXS", "YSM", "YMD", "YLG", "YXL", "OSFA", "OSFM",
    ]
    order_index = {value: index for index, value in enumerate(known_order)}
    if all(value in order_index for value in unique):
        sorted_sizes = sorted(unique, key=lambda value: order_index[value])
        # UA uses both SM/MD/LG and S/M/L. We retain supplied labels rather than
        # rewriting them to another convention.
        return "-".join([sorted_sizes[0], sorted_sizes[-1]]) if len(sorted_sizes) > 2 else ", ".join(sorted_sizes)

    def numerical_key(value: str) -> tuple[int, float, str]:
        match = re.fullmatch(r"(\d+(?:\.\d+)?)", value)
        return (0, float(match.group(1)), "") if match else (1, 0, value)

    if all(re.fullmatch(r"\d+(?:\.\d+)?", value) for value in unique):
        ordered = sorted(unique, key=numerical_key)
        return f"{ordered[0]}-{ordered[-1]}" if len(ordered) > 2 else ", ".join(ordered)

    return ", ".join(unique)


def _style_sizes(oob_rows: list[dict[str, str]]) -> dict[str, list[str]]:
    grouped: dict[str, list[str]] = defaultdict(list)
    for row in oob_rows:
        grouped[_style_from_article(row[OOB_HEADERS["article"]])].append(row[OOB_HEADERS["size"]])
    return grouped


def _material_dimension_string(material: dict[str, str] | None) -> str:
    if material is None:
        return ""
    values = [
        clean(material.get(MATERIAL_HEADERS["length"])),
        clean(material.get(MATERIAL_HEADERS["width"])),
        clean(material.get(MATERIAL_HEADERS["height"])),
    ]
    if not any(values):
        return ""
    unit = "cm"
    return " x ".join(f"{value} {unit}" if value else "" for value in values).strip(" x")


def _size_order_key(value: str) -> tuple[int, float, str]:
    """Stable order for the output; text values remain deterministic."""
    known_order = {
        "XXS": 1, "XS": 2, "SM": 3, "S": 4, "MD": 5, "M": 6,
        "LG": 7, "L": 8, "XL": 9, "XXL": 10, "2XL": 11, "3XL": 12,
        "4XL": 13, "5XL": 14, "YXS": 20, "YSM": 21, "YMD": 22,
        "YLG": 23, "YXL": 24, "OSFA": 30, "OSFM": 31,
    }
    text = norm(value)
    if text in known_order:
        return (0, float(known_order[text]), text)
    number = re.fullmatch(r"(\d+(?:\.\d+)?)", text)
    if number:
        return (1, float(number.group(1)), text)
    return (2, 0.0, text)


def _scope_style_sizes(scope_records: list[dict[str, Any]]) -> dict[str, list[str]]:
    """Create Size Scale fallbacks from the complete active seasonal scope."""
    grouped: dict[str, list[str]] = defaultdict(list)
    for record in scope_records:
        article = _scope_article(record)
        size = _scope_size(record)
        if article and size:
            grouped[_scope_style(record)].append(size)
    return grouped


def _scope_style(record: dict[str, Any]) -> str:
    centric = record.get("centric")
    material = record.get("material")
    if centric:
        return clean(centric.get("style"))
    if material:
        return clean(material.get(MATERIAL_HEADERS["style"]))
    return _style_from_article(_scope_article(record))


def _scope_article(record: dict[str, Any]) -> str:
    centric = record.get("centric")
    material = record.get("material")
    oob = record.get("oob")
    if centric:
        return clean(centric.get("article"))
    return clean(material.get(MATERIAL_HEADERS["article"])) if material else clean(oob.get(OOB_HEADERS["article"])) if oob else ""


def _scope_ean(record: dict[str, Any]) -> str:
    centric = record.get("centric")
    material = record.get("material")
    oob = record.get("oob")
    if centric:
        return clean(centric.get("ean"))
    return clean(material.get(MATERIAL_HEADERS["ean"])) if material else clean(oob.get(OOB_HEADERS["ean"])) if oob else ""


def _scope_size(record: dict[str, Any]) -> str:
    centric = record.get("centric")
    material = record.get("material")
    oob = record.get("oob")
    if centric:
        return clean(centric.get("size"))
    return clean(material.get(MATERIAL_HEADERS["size"])) if material else clean(oob.get(OOB_HEADERS["size"])) if oob else ""


def _is_summary_row(row: dict[str, str], identifier_headers: Iterable[str]) -> bool:
    """Ignore Excel total rows accidentally included in source exports."""
    total_labels = {"GRAND TOTAL", "TOTAL", "SUBTOTAL"}
    return any(norm(row.get(header)) in total_labels for header in identifier_headers)


def _centric_value(row: dict[str, str], header: str | tuple[str, ...]) -> str:
    if isinstance(header, tuple):
        for candidate in header:
            value = clean(row.get(candidate))
            if value:
                return value
        return ""
    return clean(row.get(header))


def _centric_article_and_style(row: dict[str, str]) -> tuple[str, str]:
    """Return customer-facing article/style while retaining the raw Centric code.

    Examples:
    - UA Full Article Code 6011474-001 remains 6011474-001
    - 25UJFJM07F-001-JPC becomes article 25UJFJM07F-001, style 25UJFJM07F
    """
    supplied_article = _centric_value(row, CENTRIC_HEADERS["ua_article"])
    material_number = _centric_value(row, CENTRIC_HEADERS["material_number"])
    supplied_style = _centric_value(row, CENTRIC_HEADERS["ua_style"])
    supplied_color = _centric_value(row, CENTRIC_HEADERS["ua_color"])
    valid_style = supplied_style and norm(supplied_style) not in {"N/A", "NA", "TBC"}
    valid_color = supplied_color and norm(supplied_color) not in {"N/A", "NA", "TBC"}
    article = supplied_article
    if not article or norm(article) in {"N/A", "NA", "TBC"}:
        # Boys underwear often omits UA Full Article Code but supplies the
        # standard UA style and colour code. Prefer that canonical article so
        # it replaces the erroneous generic-UA master-data record.
        if valid_style and valid_color:
            article = f"{supplied_style}-{supplied_color}"
        else:
            parts = material_number.split("-")
            article = "-".join(parts[:-1]) if len(parts) >= 3 else material_number
    if valid_style:
        style = supplied_style
    else:
        style = article.rsplit("-", 1)[0] if "-" in article else article
    return clean(article), clean(style)


def _centric_ean(row: dict[str, str]) -> str:
    return _centric_value(row, CENTRIC_HEADERS["ean"])


def _valid_ean(value: str) -> bool:
    text = clean(value).replace(" ", "")
    return text.isdigit() and 8 <= len(text) <= 14


def _centric_record(row: dict[str, str], source: str) -> dict[str, str]:
    article, style = _centric_article_and_style(row)
    return {
        "article": article,
        "style": style,
        "ean": _centric_ean(row),
        "size": _centric_value(row, CENTRIC_HEADERS["size"]),
        "source": source,
        "raw_material_number": _centric_value(row, CENTRIC_HEADERS["material_number"]),
        "description": _centric_value(row, CENTRIC_HEADERS["description"]),
        "division": _centric_value(row, CENTRIC_HEADERS["division"]),
        "gender": _centric_value(row, CENTRIC_HEADERS["gender"]),
        "size_group": _centric_value(row, CENTRIC_HEADERS["age_group"]),
        "end_use": _centric_value(row, CENTRIC_HEADERS["category"]),
        "season": _centric_value(row, CENTRIC_HEADERS["season"]),
        "color": _centric_value(row, CENTRIC_HEADERS["color"]),
        "country": _centric_value(row, CENTRIC_HEADERS["country"]),
        "hts": _centric_value(row, CENTRIC_HEADERS["hts"]),
        "material": _centric_value(row, CENTRIC_HEADERS["material"]),
        "fedas": _centric_value(row, CENTRIC_HEADERS["fedas"]),
        "fedas_size": _centric_value(row, CENTRIC_HEADERS["fedas_size"]),
    }


def _centric_rows(
    records: list[dict[str, str]],
    source: str,
    issues: list[dict[str, str]],
    counters: Counter,
) -> list[dict[str, str]]:
    required = [
        CENTRIC_HEADERS["division"], CENTRIC_HEADERS["range"], CENTRIC_HEADERS["gender"],
        CENTRIC_HEADERS["age_group"], CENTRIC_HEADERS["category"], CENTRIC_HEADERS["material_number"],
        CENTRIC_HEADERS["description"], CENTRIC_HEADERS["size"], CENTRIC_HEADERS["color"],
    ]
    _require_headers(records, required, source)
    usable: list[dict[str, str]] = []
    seen: set[str] = set()
    for row in records:
        # Only current in-line product is included. This guards against a future
        # source file mixing MFO rows into the same sheet.
        if norm(row.get(CENTRIC_HEADERS["range"])) != "IN-LINE":
            continue
        item = _centric_record(row, source)
        ean = norm(item["ean"])
        if not _valid_ean(ean):
            counters["centric_invalid_ean"] += 1
            issues.append({
                "Severity": "Review", "Type": "Centric row without valid EAN",
                "EAN": clean(item["ean"]), "Article": item["article"],
                "Detail": f"{source}: raw Centric material number '{item['raw_material_number']}' has EAN/UPC '{item['ean'] or '(blank)'}'.",
                "Recommended action": "Do not send this product to customers until Centric supplies a numeric EAN/UPC.",
            })
            continue
        if ean in seen:
            counters["centric_duplicate_source_ean"] += 1
            issues.append({
                "Severity": "Warning", "Type": "Duplicate Centric EAN in source",
                "EAN": item["ean"], "Article": item["article"],
                "Detail": f"{source}: EAN appears more than once. The first row is used.",
                "Recommended action": "Check the Centric master-data export for duplicate UPC/EAN assignments.",
            })
            continue
        seen.add(ean)
        if not item["country"]:
            counters["centric_missing_coo"] += 1
            issues.append({
                "Severity": "Info", "Type": "Centric COO missing", "EAN": item["ean"], "Article": item["article"],
                "Detail": f"{source}: Country of Origin for PO is blank; output stays blank rather than borrowing data from UA Material Data.",
                "Recommended action": "Request a refreshed Centric master-data file if COO is required by the customer.",
            })
        usable.append(item)
    return usable


def build_listing(
    oob_bytes: bytes,
    material_bytes: bytes,
    line_list_bytes: bytes,
    changelog_bytes: bytes,
    template_bytes: bytes,
    centric_underwear_bytes: bytes,
    centric_outerwear_bytes: bytes,
    centric_sportswear_bytes: bytes,
    season: str,
) -> tuple[list[dict[str, str]], dict[str, Any], list[dict[str, str]], list[dict[str, str]]]:
    """Build the complete active-season customer listing.

    Scope hierarchy:
      1. Current Line List + Licensed List define the active seasonal portfolio.
      2. Change Log applies ADD/DROP corrections where the snapshot and log differ.
      3. Material Data Report expands eligible colorways into individual EAN/size rows.
      4. OOB does not limit the UA portfolio. It marks confirmed items and retains an
         individual UA EAN when that confirmed EAN is outside the active master scope.
      5. Centric Brands In-line files are a separate licensed-product stream and
         are appended directly from their own master data. MFO underwear is excluded.

    The result deliberately excludes standard-UA products that exist only in Material Data:
    these are not necessarily available for EMEA FW26 ordering. Centric In-line products
    are included without relying on the UA Line List or OOB.
    """
    raw_oob_rows = _read_sheet_from_xlsx(oob_bytes, "Sheet1")
    raw_material_rows = _read_sheet_from_xlsx(material_bytes, "Sheet 1")
    raw_line_rows = _read_sheet_from_xlsx(line_list_bytes, "Line List")
    raw_licensed_rows = _read_sheet_from_xlsx(line_list_bytes, "Licensed List")
    changes = _read_sheet_from_xlsx(changelog_bytes, "Adds-Drops")
    date_changes = _read_sheet_from_xlsx(changelog_bytes, "Date Changes")
    template_rows = _read_sheet_from_xlsx(template_bytes, "Sheet1")
    centric_underwear_in_line = _read_sheet_from_xlsx(centric_underwear_bytes, "In_line Underwear_Undershirts")
    centric_underwear_boys = _read_sheet_from_xlsx(centric_underwear_bytes, "Boys_Underwear")
    centric_outerwear = _read_sheet_from_xlsx(centric_outerwear_bytes, "Sheet1")
    centric_sportswear = _read_sheet_from_xlsx(centric_sportswear_bytes, "Sheet1")

    _require_headers(raw_oob_rows, OOB_HEADERS.values(), "OOB")
    _require_headers(raw_material_rows, MATERIAL_HEADERS.values(), "Material Data Report")

    # Some UA exports contain a trailing Excel "Grand Total" row. It is not a
    # product and must never turn into a fake EAN in the customer listing.
    oob_rows = [
        row for row in raw_oob_rows
        if not _is_summary_row(row, [OOB_HEADERS["ean"], OOB_HEADERS["article"]])
    ]
    material_rows = [
        row for row in raw_material_rows
        if not _is_summary_row(row, [MATERIAL_HEADERS["ean"], MATERIAL_HEADERS["article"]])
    ]
    line_rows = [
        row for row in raw_line_rows
        if not _is_summary_row(row, [LINE_HEADERS["colorway"]])
    ]
    licensed_rows = [
        row for row in raw_licensed_rows
        if not _is_summary_row(row, [LINE_HEADERS["colorway"]])
    ]
    _require_headers(line_rows, LINE_REQUIRED_HEADERS, "Line List")
    _require_headers(changes, ["Change Date", "Colorway Number", "Record Change"], "Change Log / Adds-Drops")
    _require_headers(
        date_changes,
        ["Change Date", "Colorway Number", "New Shipment Start Date.", "New Hard Launch Date"],
        "Change Log / Date Changes",
    )

    reference = _reference_indexes(template_rows)
    issues: list[dict[str, str]] = []
    counters = Counter()
    if LINE_HEADERS["hero_look_name"] not in line_rows[0]:
        issues.append({
            "Severity": "Review", "Type": "Line List Hero Look Name column missing", "EAN": "", "Article": "",
            "Detail": "The EMEA Line List does not contain the Hero Look Name column; all output GHL values default to NE.",
            "Recommended action": "Confirm whether Hero Look Name was removed or renamed before sending the customer listing.",
        })

    # Centric data is authoritative for the licensed product stream.  Underwear
    # deliberately excludes the MFO sheet; outerwear and sportswear use their In-line files.
    centric_sources = (
        (centric_underwear_in_line, "Centric Underwear In-line"),
        (centric_underwear_boys, "Centric Boys Underwear"),
        (centric_outerwear, "Centric Outerwear In-line"),
        (centric_sportswear, "Centric Sportswear In-line"),
    )
    centric_authoritative_articles = {
        norm(_centric_article_and_style(row)[0])
        for records, _ in centric_sources
        for row in records
        if norm(row.get(CENTRIC_HEADERS["range"])) == "IN-LINE" and _centric_article_and_style(row)[0]
    }
    centric_records = [
        item
        for records, source_name in centric_sources
        for item in _centric_rows(records, source_name, issues, counters)
    ]

    # Index all potential EAN master data. Material Data can contain products that
    # are technically in the season but are not part of the sellable EMEA line.
    material_by_ean: dict[str, list[dict[str, str]]] = defaultdict(list)
    material_by_colorway: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in material_rows:
        ean = norm(row.get(MATERIAL_HEADERS["ean"]))
        article = norm(row.get(MATERIAL_HEADERS["article"]))
        if ean:
            material_by_ean[ean].append(row)
        if article:
            material_by_colorway[article].append(row)

    # Current Line List is the base net assortment. Licensed List is equally valid
    # for licensed products (underwear, bottles, etc.).
    line_by_colorway: dict[str, dict[str, str]] = {}
    for source_name, records in (("Line List", line_rows), ("Licensed List", licensed_rows)):
        for row in records:
            colorway = norm(row.get(LINE_HEADERS["colorway"]))
            if not colorway:
                continue
            if colorway in line_by_colorway:
                counters["duplicate_line_colorways"] += 1
                issues.append({
                    "Severity": "Warning", "Type": "Duplicate Line List Colorway", "EAN": "", "Article": colorway,
                    "Detail": f"The colorway occurs more than once across Line List / Licensed List; the first row is used.",
                    "Recommended action": "Check whether the Line List export contains duplicate commercial records.",
                })
                continue
            line_by_colorway[colorway] = row

    latest_change = _latest_by_date(changes, "Colorway Number", "Change Date")
    latest_date_change = _latest_by_date(date_changes, "Colorway Number", "Change Date")
    line_colorways = set(line_by_colorway)
    oob_colorways = {norm(row.get(OOB_HEADERS["article"])) for row in oob_rows if norm(row.get(OOB_HEADERS["article"]))}

    drop_colorways = {
        colorway for colorway, change in latest_change.items()
        if norm(change.get("Record Change")) == "DROP"
    }
    add_colorways = {
        colorway for colorway, change in latest_change.items()
        if norm(change.get("Record Change")) == "ADD"
    }

    # A newer Change Log can correct a Line List snapshot. A DROP does not delete
    # an OOB-confirmed EAN; it only removes the colorway from the broad active
    # portfolio. The confirmed EAN itself is added back below as an OOB exception.
    dropped_from_current_line = line_colorways & drop_colorways
    active_colorways = line_colorways - dropped_from_current_line

    # A post-snapshot ADD can be included only if Material Data has actual EAN/size
    # records. No Material Data means there is nothing a customer could list yet.
    added_outside_current_line = add_colorways - line_colorways
    added_with_material = {
        colorway for colorway in added_outside_current_line
        if material_by_colorway.get(colorway)
    }
    active_colorways.update(added_with_material)
    added_without_material = added_outside_current_line - added_with_material
    for colorway in sorted(added_without_material):
        issues.append({
            "Severity": "Info", "Type": "Change Log ADD without Material Data", "EAN": "", "Article": colorway,
            "Detail": "Latest Change Log status is ADD, but the current Material Data Report contains no EAN/size rows for this colorway.",
            "Recommended action": "Do not list it yet. Include it automatically once the Material Data Report contains its EANs.",
        })

    for colorway in sorted(dropped_from_current_line):
        issues.append({
            "Severity": "Info", "Type": "Change Log DROP applied to Line List", "EAN": "", "Article": colorway,
            "Detail": "The colorway was removed from the broad active portfolio because its latest Change Log status is DROP.",
            "Recommended action": "No action unless an OOB-confirmed EAN for this colorway must remain listed.",
        })

    # Build the active portfolio from current active Line List colors, expanded by
    # all corresponding Master Data EANs. Use one final row per EAN.
    scope_records: list[dict[str, Any]] = []
    final_eans: set[str] = set()
    duplicate_active_master_eans: set[str] = set()
    for material in material_rows:
        article = norm(material.get(MATERIAL_HEADERS["article"]))
        if article not in active_colorways:
            continue
        if article in centric_authoritative_articles:
            counters["generic_ua_rows_replaced_by_centric"] += 1
            continue
        ean = norm(material.get(MATERIAL_HEADERS["ean"]))
        if not ean:
            counters["active_material_without_ean"] += 1
            issues.append({
                "Severity": "Warning", "Type": "Active Material row without EAN", "EAN": "", "Article": clean(material.get(MATERIAL_HEADERS["article"])),
                "Detail": "The colorway is active, but this Material Data row has no EAN and cannot become a valid customer listing row.",
                "Recommended action": "Request a refreshed Material Data Report before listing this size.",
            })
            continue
        if ean in final_eans:
            duplicate_active_master_eans.add(ean)
            continue
        final_eans.add(ean)
        source = "Active Line List" if article in line_colorways else "Change Log ADD"
        scope_records.append({
            "material": material,
            "oob": None,
            "source": source,
        })
    for ean in sorted(duplicate_active_master_eans):
        counters["duplicate_active_master_eans"] += 1
        issues.append({
            "Severity": "Warning", "Type": "Duplicate Active Master EAN", "EAN": ean, "Article": "",
            "Detail": "More than one active Material Data row has the same EAN; the first row was retained.",
            "Recommended action": "Check for duplicated EANs in the Material Data Report.",
        })

    # Index OOB by EAN. It is used for traceability of confirmed items and for
    # retaining exceptional confirmed EANs which are absent from the current active
    # Line List/Material scope.
    oob_by_ean: dict[str, dict[str, str]] = {}
    oob_without_ean: list[dict[str, str]] = []
    for oob in oob_rows:
        ean = norm(oob.get(OOB_HEADERS["ean"]))
        if not ean:
            oob_without_ean.append(oob)
            continue
        if ean in oob_by_ean:
            counters["duplicate_oob_eans"] += 1
            issues.append({
                "Severity": "Warning", "Type": "Duplicate OOB EAN", "EAN": clean(oob.get(OOB_HEADERS["ean"])), "Article": clean(oob.get(OOB_HEADERS["article"])),
                "Detail": "The OOB contains the same EAN more than once; the first row is used for traceability.",
                "Recommended action": "Check whether the OOB contains repeated line items.",
            })
            continue
        oob_by_ean[ean] = oob
    for oob in oob_without_ean:
        issues.append({
            "Severity": "Review", "Type": "OOB row without EAN", "EAN": "", "Article": clean(oob.get(OOB_HEADERS["article"])),
            "Detail": "This confirmed OOB row cannot be safely matched or retained without an EAN.",
            "Recommended action": "Correct the OOB row and rerun the listing.",
        })

    # Mark active-scope records ordered in OOB and retain each OOB EAN that would
    # otherwise be absent. This protects confirmed purchases without widening a
    # DROP colorway to all its Master Data sizes.
    oob_already_active = 0
    for record in scope_records:
        ean = norm(_scope_ean(record))
        oob = oob_by_ean.get(ean)
        if oob:
            record["oob"] = oob
            oob_already_active += 1
            material = record["material"]
            if (
                norm(oob.get(OOB_HEADERS["article"])) != norm(material.get(MATERIAL_HEADERS["article"]))
                or norm(oob.get(OOB_HEADERS["size"])) != norm(material.get(MATERIAL_HEADERS["size"]))
            ):
                counters["oob_master_mismatch"] += 1
                issues.append({
                    "Severity": "Review", "Type": "OOB / Material Data mismatch", "EAN": clean(oob.get(OOB_HEADERS["ean"])),
                    "Article": clean(oob.get(OOB_HEADERS["article"])),
                    "Detail": "The same EAN has different Article Generic or Size values in OOB and Material Data.",
                    "Recommended action": "Verify which source represents the currently sellable product code.",
                })

    oob_exceptions = 0
    for ean, oob in oob_by_ean.items():
        if ean in final_eans:
            continue
        material, ambiguous_material = _select_material(material_by_ean.get(ean, []), oob)
        if ambiguous_material:
            counters["ambiguous_material"] += 1
            issues.append({
                "Severity": "Warning", "Type": "Ambiguous Material Data", "EAN": clean(oob.get(OOB_HEADERS["ean"])), "Article": clean(oob.get(OOB_HEADERS["article"])),
                "Detail": "More than one Material Data Report row matched this OOB EAN with the same best score.",
                "Recommended action": "Review the source rows in the Material Data Report.",
            })
        if material is None:
            counters["missing_material"] += 1
            issues.append({
                "Severity": "Warning", "Type": "Missing Material Data", "EAN": clean(oob.get(OOB_HEADERS["ean"])), "Article": clean(oob.get(OOB_HEADERS["article"])),
                "Detail": "Confirmed OOB EAN is absent from the current Material Data Report.",
                "Recommended action": "Keep it only where delivery is certain; request the corresponding Master Data row from UA.",
            })
        scope_records.append({
            "material": material,
            "oob": oob,
            "source": "OOB confirmed exception",
        })
        final_eans.add(ean)
        oob_exceptions += 1

    # Add the separate Centric Brands In-line portfolio.  These rows never depend
    # on OOB, generic UA Material Data or the UA Licensed List.  If an EAN happens
    # to exist in the generic UA stream, Centric wins for that licensed product.
    existing_scope_by_ean = {norm(_scope_ean(record)): record for record in scope_records}
    for centric in centric_records:
        ean = norm(centric["ean"])
        old = existing_scope_by_ean.get(ean)
        if old is not None:
            old["discard"] = True
            issues.append({
                "Severity": "Info", "Type": "Centric data overrides generic UA row",
                "EAN": centric["ean"], "Article": centric["article"],
                "Detail": "The same EAN was present in the generic UA stream; Centric master data is authoritative for this licensed product.",
                "Recommended action": "No action. The final listing uses the Centric row.",
            })
        scope_records.append({"centric": centric, "material": None, "oob": None, "source": centric["source"]})
        existing_scope_by_ean[ean] = scope_records[-1]
        final_eans.add(ean)
    scope_records = [record for record in scope_records if not record.get("discard")]

    # Build fallback ranges from the whole final scope, not just from OOB sizes.
    style_sizes = _scope_style_sizes(scope_records)
    scope_records.sort(key=lambda record: (
        _scope_style(record),
        _scope_article(record),
        _size_order_key(_scope_size(record)),
        _scope_ean(record),
    ))

    output_rows: list[dict[str, str]] = []
    scope_audit_rows: list[dict[str, str]] = []
    for record in scope_records:
        material = record.get("material")
        centric = record.get("centric")
        oob = record.get("oob")
        source = clean(record.get("source"))
        article = _scope_article(record)
        ean = _scope_ean(record)
        size = _scope_size(record)
        line = line_by_colorway.get(norm(article))
        date_change = latest_date_change.get(norm(article))
        latest_add_drop = latest_change.get(norm(article), {})
        change_status = norm(latest_add_drop.get("Record Change"))

        # OOB confirmed EANs always survive a Drop warning. This is deliberately
        # per EAN, not per full colorway, so unconfirmed sizes are not resurrected.
        if oob is not None and change_status == "DROP":
            counters["oob_change_log_drop_conflicts"] += 1
            issues.append({
                "Severity": "Review", "Type": "OOB / Change Log DROP conflict", "EAN": ean, "Article": article,
                "Detail": (
                    "Latest Adds-Drops entry is DROP "
                    f"({clean(latest_add_drop.get('Change Date'))}), but this EAN is present in the current OOB. "
                    "The confirmed EAN was retained in the final listing."
                ),
                "Recommended action": "No listing action. Reconcile with UA only where source-data consistency is required.",
            })

        if line is None and centric is None:
            if source == "Change Log ADD":
                counters["add_without_line_row"] += 1
                issues.append({
                    "Severity": "Warning", "Type": "Change Log ADD missing Line List attributes", "EAN": ean, "Article": article,
                    "Detail": "The ADD was included from Master Data, but no current Line List row supplies commercial attributes such as dates and Size Scale.",
                    "Recommended action": "Refresh the Line List when available; Master Data fields were used in the meantime.",
                })
            elif source == "OOB confirmed exception":
                counters["missing_line"] += 1
                issues.append({
                    "Severity": "Warning", "Type": "Missing Line List Data", "EAN": ean, "Article": article,
                    "Detail": "Confirmed OOB EAN is outside the current Line List and Licensed List.",
                    "Recommended action": "Keep it as a confirmed exception; check whether it is a carryover, replacement, or discontinued SKU.",
                })

        style = clean(centric.get("style")) if centric else (clean(material.get(MATERIAL_HEADERS["style"])) if material else _style_from_article(article))
        division = clean(centric.get("division")) if centric else (clean(material.get(MATERIAL_HEADERS["division"])) if material else clean(oob.get(OOB_HEADERS["division"])) if oob else "")
        gender = clean(centric.get("gender")) if centric else (clean(material.get(MATERIAL_HEADERS["gender"])) if material else clean(oob.get(OOB_HEADERS["gender"])) if oob else "")
        size_group = clean(centric.get("size_group")) if centric else (clean(material.get(MATERIAL_HEADERS["size_group"])) if material else "")
        style_description = clean(centric.get("description")) if centric else (clean(material.get(MATERIAL_HEADERS["style_name"])) if material else clean(oob.get(OOB_HEADERS["description"])) if oob else "")

        ref_row = reference["exact"].get((norm(article), norm(size)))
        size_eur = clean(ref_row.get("Size EUR")) if ref_row else ""
        if not size_eur:
            candidates = reference["general_eu"].get((norm(size), norm(division), norm(gender)), set())
            if len(candidates) == 1:
                size_eur = next(iter(candidates))
                counters["size_eur_reference_fallback"] += 1
        if not size_eur and centric is not None:
            # Centric supplies age labels (e.g. 6-7YR), not a separate EU-size
            # conversion. Preserve the source label rather than inventing 122/128.
            size_eur = size
            counters["centric_size_eur_source_label"] += 1
        if not size_eur and norm(division) == "ACCESSORIES":
            size_eur = "UNI" if norm(size) in {"OSFA", "OSFM"} else size
        if not size_eur:
            counters["missing_size_eur"] += 1
            issues.append({
                "Severity": "Review", "Type": "Unresolved Size EUR", "EAN": ean, "Article": article,
                "Detail": f"No unambiguous Size EUR mapping for UA size '{size}', division '{division}', gender '{gender}'.",
                "Recommended action": "Add a mapping to the reference listing or apply a manual correction before customer export.",
            })

        size_scale = clean(line.get(LINE_HEADERS["size_scale"])) if line else ""
        if not size_scale:
            scales = reference["article_scale"].get(norm(article), set())
            if len(scales) == 1:
                size_scale = next(iter(scales))
        if not size_scale:
            scales = reference["style_scale"].get(norm(style), set())
            if len(scales) == 1:
                size_scale = next(iter(scales))
        if not size_scale:
            size_scale = _infer_size_scale(style_sizes.get(norm(style), []))
            if size_scale:
                counters["inferred_size_scale"] += 1
                issues.append({
                    "Severity": "Review", "Type": "Inferred Size Scale", "EAN": ean, "Article": article,
                    "Detail": f"No official Size Scale was available; generated from the final active listing sizes: '{size_scale}'.",
                    "Recommended action": "Confirm the range if a customer requires UA's official Size Scale text.",
                })

        ship_start = clean(line.get(LINE_HEADERS["ship_start"])) if line else ""
        launch_date = clean(line.get(LINE_HEADERS["hard_launch"])) if line else ""
        if date_change:
            new_ship_start = clean(date_change.get("New Shipment Start Date."))
            new_launch_date = clean(date_change.get("New Hard Launch Date"))
            if new_ship_start:
                ship_start = new_ship_start
                counters["date_changes_applied"] += 1
            if new_launch_date:
                launch_date = new_launch_date
                counters["date_changes_applied"] += 1

        line_copy = clean(line.get(LINE_HEADERS["catalog_copy"])) if line else ""
        volume, dimensions_in, dimensions_cm = _extract_volume_and_dimensions(line_copy)
        if not dimensions_cm:
            dimensions_cm = _material_dimension_string(material)

        row = {header: "" for header in OUTPUT_HEADERS}
        row.update({
            "Style": style,
            "Article": article,
            "SKU": f"{article}-{size}" if article and size else article,
            "usdis": f"{article}-{size}" if article and size else article,
            "Style Description": style_description,
            "Size UA": size,
            "Size US DIST.": size,
            "Size EUR": size_eur,
            "Size Scale": size_scale,
            "Selling Season": season,
            "EAN": ean,
            "Product Division": division,
            "Gender Description": gender,
            "Size Group Description": size_group,
            "End Use": clean(centric.get("end_use")) if centric else (clean(line.get(LINE_HEADERS["end_use"])) if line else (clean(material.get(MATERIAL_HEADERS["end_use"])) if material else clean(oob.get(OOB_HEADERS["end_use"])) if oob else "")),
            "Signature Collections": clean(line.get(LINE_HEADERS["signature"])) if line else "",
            # GHL is a commercial flag from the EMEA Line List.  Any non-empty
            # Hero Look Name marks the whole style/colorway as a Global Hero Look.
            # The current FW26 values are Q3/Q4 labels, but we intentionally use
            # non-empty as the rule to remain robust if UA changes the naming.
            "GHL": "ANO" if line and clean(line.get(LINE_HEADERS["hero_look_name"])) else "NE",
            "Story Tier": clean(line.get(LINE_HEADERS["storytier"])) if line else "",
            "Fit Type": clean(line.get(LINE_HEADERS["fit"])) if line else "",
            "Merch Department": clean(material.get(MATERIAL_HEADERS["merch_department"])) if material else "",
            "Class Description": clean(material.get(MATERIAL_HEADERS["merch_class"])) if material else "",
            "Sub Class Description": clean(material.get(MATERIAL_HEADERS["merch_subclass"])) if material else "",
            "Period": clean(line.get(LINE_HEADERS["period"])) if line else "",
            "Shipment Start Date": _date_string(ship_start),
            "Shipment End Date": _date_string(clean(line.get(LINE_HEADERS["ship_end"])) if line else ""),
            "Launch Date": _date_string(launch_date),
            "Color Group": clean(centric.get("color")) if centric else (clean(material.get(MATERIAL_HEADERS["color_group"])) if material else ""),
            "Primary Color": clean(centric.get("color")) if centric else (clean(material.get(MATERIAL_HEADERS["primary_color"])) if material else ""),
            "Secondary Color": clean(material.get(MATERIAL_HEADERS["secondary_color"])) if material else "",
            "Logo Colorway": clean(material.get(MATERIAL_HEADERS["logo_color"])) if material else "",
            "Product Ranking": clean(material.get(MATERIAL_HEADERS["ranking"])) if material else (clean(line.get(LINE_HEADERS["fashion_grade"])) if line else ""),
            "Gross Weight (kg)": clean(material.get(MATERIAL_HEADERS["weight"])) if material else "",
            "Article Lenght": clean(material.get(MATERIAL_HEADERS["length"])) if material else "",
            "Article Width": clean(material.get(MATERIAL_HEADERS["width"])) if material else "",
            "Article Height": clean(material.get(MATERIAL_HEADERS["height"])) if material else "",
            "Volume": volume,
            "Dimensions (inch)": dimensions_in,
            "Dimensions (cm)": dimensions_cm,
            "HTS Code": clean(centric.get("hts")) if centric else (clean(material.get(MATERIAL_HEADERS["hts"])) if material else ""),
            "COO COUNTRY": clean(centric.get("country")) if centric else (clean(material.get(MATERIAL_HEADERS["country"])) if material else ""),
            "FEDAS Code": clean(centric.get("fedas")) if centric else (clean(material.get(MATERIAL_HEADERS["fedas"])) if material else ""),
            "FEDAS Size Range": clean(centric.get("fedas_size")) if centric else (clean(material.get(MATERIAL_HEADERS["fedas_size"])) if material else ""),
            "Material": clean(centric.get("material")) if centric else (clean(material.get(MATERIAL_HEADERS["material"])) if material else ""),
            "Tech Platform": clean(line.get(LINE_HEADERS["tech"])) if line else (clean(material.get(MATERIAL_HEADERS["tech"])) if material else ""),
        })
        output_rows.append(row)
        scope_audit_rows.append({
            "EAN": ean,
            "Article": article,
            "Size UA": size,
            "Listing Scope": source,
            "In OOB": "Yes" if oob is not None else "No",
            "In Current Line List": "Yes" if line is not None else "No",
            "GHL": row["GHL"],
            "Hero Look Name": clean(line.get(LINE_HEADERS["hero_look_name"])) if line else "",
            "Latest Change Log": change_status,
            "Latest Change Date": clean(latest_add_drop.get("Change Date")),
            "Source Material Number": clean(centric.get("raw_material_number")) if centric else "",
        })

    summary: dict[str, Any] = {
        "Season": season,
        "OOB input rows": len(raw_oob_rows),
        "OOB product rows": len(oob_rows),
        "OOB summary rows ignored": len(raw_oob_rows) - len(oob_rows),
        "Current Line List colorways (incl. Licensed)": len(line_colorways),
        "Change Log DROP colorways removed from broad portfolio": len(dropped_from_current_line),
        "Change Log ADD colorways included outside current Line List": len(added_with_material),
        "Change Log ADD colorways without Material Data": len(added_without_material),
        "Active portfolio EANs from Line List / Master Data": sum(1 for record in scope_records if clean(record.get("source")) in {"Active Line List", "Change Log ADD"}),
        "OOB EANs already present in active portfolio": oob_already_active,
        "OOB confirmed exception EANs retained": oob_exceptions,
        "OOB / Change Log DROP conflict EANs retained": counters["oob_change_log_drop_conflicts"],
        "Centric Underwear In-line EANs": sum(1 for row in centric_records if row["source"] == "Centric Underwear In-line"),
        "Centric Boys Underwear EANs": sum(1 for row in centric_records if row["source"] == "Centric Boys Underwear"),
        "Centric Outerwear In-line EANs": sum(1 for row in centric_records if row["source"] == "Centric Outerwear In-line"),
        "Centric Sportswear In-line EANs": sum(1 for row in centric_records if row["source"] == "Centric Sportswear In-line"),
        "Centric MFO EANs included": 0,
        "Centric invalid / placeholder EAN rows excluded": counters["centric_invalid_ean"],
        "Centric missing COO rows": counters["centric_missing_coo"],
        "Generic UA Material Data rows superseded by Centric": counters["generic_ua_rows_replaced_by_centric"],
        "GHL = ANO rows": sum(1 for row in output_rows if row.get("GHL") == "ANO"),
        "GHL = NE rows": sum(1 for row in output_rows if row.get("GHL") == "NE"),
        "Final listing rows": len(output_rows),
        "Missing Material Data rows": counters["missing_material"],
        "Missing Line List rows": counters["missing_line"],
        "Change Log ADD rows missing Line List attributes": counters["add_without_line_row"],
        "Ambiguous Material Data rows": counters["ambiguous_material"],
        "OOB / Material Data mismatches": counters["oob_master_mismatch"],
        "Unresolved Size EUR rows": counters["missing_size_eur"],
        "Size EUR using reference fallback": counters["size_eur_reference_fallback"],
        "Inferred Size Scale rows": counters["inferred_size_scale"],
        "Date fields overwritten from Change Log": counters["date_changes_applied"],
    }
    return output_rows, summary, issues, scope_audit_rows

def _template_sheet_and_styles(template_bytes: bytes) -> tuple[str, dict[str, str], dict[str, str], str]:
    with zipfile.ZipFile(BytesIO(template_bytes)) as zf:
        paths = _sheet_paths(zf)
        if "Sheet1" not in paths:
            raise ValueError("Reference listing must contain a sheet called Sheet1.")
        root = ET.fromstring(zf.read(paths["Sheet1"]))
    header_styles: dict[str, str] = {}
    body_styles: dict[str, str] = {}
    header_row_attrs = ""
    header_row = root.find(".//m:sheetData/m:row[@r='1']", NS)
    if header_row is not None:
        safe_row_attributes = ("ht", "customHeight", "hidden", "collapsed", "outlineLevel", "s")
        header_row_attrs = " ".join(
            f'{key}="{escape(header_row.attrib[key])}"' for key in safe_row_attributes if key in header_row.attrib
        )
        for cell in header_row.findall("m:c", NS):
            ref = cell.attrib.get("r", "")
            col = "".join(char for char in ref if char.isalpha())
            if col:
                header_styles[col] = cell.attrib.get("s", "0")
    first_data_row = root.find(".//m:sheetData/m:row[@r='2']", NS)
    if first_data_row is not None:
        for cell in first_data_row.findall("m:c", NS):
            ref = cell.attrib.get("r", "")
            col = "".join(char for char in ref if char.isalpha())
            if col:
                body_styles[col] = cell.attrib.get("s", "0")
    return paths["Sheet1"], header_styles, body_styles, header_row_attrs


def _shared_strings_xml(strings: list[str], total_count: int) -> bytes:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<sst xmlns="{MAIN_NS}" count="{total_count}" uniqueCount="{len(strings)}">',
    ]
    for value in strings:
        safe = escape(value)
        # xml:space prevents Excel stripping spaces in product descriptions.
        parts.append(f'<si><t xml:space="preserve">{safe}</t></si>')
    parts.append("</sst>")
    return "".join(parts).encode("utf-8")


def _table_to_shared_strings(headers: list[str], rows: list[dict[str, str]]) -> tuple[dict[str, int], list[str], int]:
    table: OrderedDict[str, int] = OrderedDict()
    total = 0
    for value in headers:
        text = clean(value)
        if text not in table:
            table[text] = len(table)
        total += 1
    for row in rows:
        for header in headers:
            value = clean(row.get(header, ""))
            if not value:
                continue
            if value not in table:
                table[value] = len(table)
            total += 1
    return dict(table), list(table), total


def _listing_sheet_xml(
    headers: list[str],
    rows: list[dict[str, str]],
    value_index: dict[str, int],
    header_styles: dict[str, str],
    body_styles: dict[str, str],
    header_row_attrs: str,
) -> str:
    final_row = len(rows) + 1
    last_col = excel_col(len(headers))
    cells: list[str] = []

    def write_row(row_number: int, values: list[str]) -> None:
        extra_attrs = f" {header_row_attrs}" if row_number == 1 and header_row_attrs else ""
        parts = [f'<row r="{row_number}"{extra_attrs}>']
        for index, value in enumerate(values, start=1):
            text = clean(value)
            if not text:
                continue
            column = excel_col(index)
            style = (header_styles if row_number == 1 else body_styles).get(column, "0")
            parts.append(f'<c r="{column}{row_number}" s="{style}" t="s"><v>{value_index[text]}</v></c>')
        parts.append("</row>")
        cells.append("".join(parts))

    write_row(1, headers)
    for row_number, row in enumerate(rows, start=2):
        write_row(row_number, [row.get(header, "") for header in headers])
    sheet_data = "<sheetData>" + "".join(cells) + "</sheetData>"

    return sheet_data, f"A1:{last_col}{final_row}"


def listing_xlsx_from_template(template_bytes: bytes, rows: list[dict[str, str]]) -> bytes:
    """Create listing XLSX, preserving the source listing's sheet formatting."""
    sheet_path, header_styles, body_styles, header_row_attrs = _template_sheet_and_styles(template_bytes)
    value_index, strings, total = _table_to_shared_strings(OUTPUT_HEADERS, rows)
    new_sheet_data, final_range = _listing_sheet_xml(OUTPUT_HEADERS, rows, value_index, header_styles, body_styles, header_row_attrs)

    source = BytesIO(template_bytes)
    target = BytesIO()
    with zipfile.ZipFile(source) as input_zip, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as output_zip:
        for info in input_zip.infolist():
            data = input_zip.read(info.filename)
            if info.filename == sheet_path:
                xml = data.decode("utf-8")
                xml = re.sub(r"<dimension\s+ref=\"[^\"]+\"", f'<dimension ref="{final_range}"', xml, count=1)
                xml = re.sub(r"<sheetData>.*?</sheetData>", new_sheet_data, xml, flags=re.DOTALL, count=1)
                xml = re.sub(r"<autoFilter\s+ref=\"[^\"]+\"", f'<autoFilter ref="{final_range}"', xml, count=1)
                data = xml.encode("utf-8")
            elif info.filename == "xl/sharedStrings.xml":
                data = _shared_strings_xml(strings, total)
            output_zip.writestr(info, data)
    return target.getvalue()


def _minimal_content_types(sheet_count: int) -> str:
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, sheet_count + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        f'{overrides}</Types>'
    )


def _minimal_styles() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{MAIN_NS}"><fonts count="2">'
        '<font><sz val="10"/><name val="Arial"/></font>'
        '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
        '</fonts><fills count="3"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '</cellXfs></styleSheet>'
    )


def _write_minimal_xlsx(sheets: list[tuple[str, list[list[str]]]]) -> bytes:
    """Create a compact XLSX for the audit report without external packages."""
    all_strings: OrderedDict[str, int] = OrderedDict()
    count = 0
    for _, data in sheets:
        for row in data:
            for value in row:
                value = clean(value)
                if not value:
                    continue
                if value not in all_strings:
                    all_strings[value] = len(all_strings)
                count += 1

    def sheet_xml(data: list[list[str]]) -> str:
        row_xml: list[str] = []
        max_cols = max((len(row) for row in data), default=1)
        for row_index, row in enumerate(data, start=1):
            cells = [f'<row r="{row_index}">']
            for col_index, value in enumerate(row, start=1):
                text = clean(value)
                if not text:
                    continue
                style = 1 if row_index == 1 else 0
                cells.append(f'<c r="{excel_col(col_index)}{row_index}" s="{style}" t="s"><v>{all_strings[text]}</v></c>')
            cells.append("</row>")
            row_xml.append("".join(cells))
        last_row = max(1, len(data))
        last_col = excel_col(max_cols)
        widths = "".join(
            f'<col min="{i}" max="{i}" width="{min(42, max(12, max((len(clean(row[i-1])) if len(row) >= i else 0 for row in data), default=12) + 2))}" customWidth="1"/>'
            for i in range(1, max_cols + 1)
        )
        auto_filter = f'<autoFilter ref="A1:{last_col}{last_row}"/>' if len(data) > 1 else ""
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{MAIN_NS}"><dimension ref="A1:{last_col}{last_row}"/>'
            '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            f'<cols>{widths}</cols><sheetData>{"".join(row_xml)}</sheetData>{auto_filter}</worksheet>'
        )

    workbook_sheets = "".join(
        f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, (name, _) in enumerate(sheets, start=1)
    )
    workbook_rels = "".join(
        f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
        for index, _ in enumerate(sheets, start=1)
    )
    workbook_rels += '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    workbook_rels += '<Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'

    payload = BytesIO()
    with zipfile.ZipFile(payload, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        zf.writestr("[Content_Types].xml", _minimal_content_types(len(sheets)))
        zf.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        zf.writestr("xl/workbook.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{MAIN_NS}" xmlns:r="{DOCREL_NS}"><sheets>{workbook_sheets}</sheets></workbook>')
        zf.writestr("xl/_rels/workbook.xml.rels", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKGREL_NS}">{workbook_rels}</Relationships>')
        zf.writestr("xl/styles.xml", _minimal_styles())
        zf.writestr("xl/sharedStrings.xml", _shared_strings_xml(list(all_strings), count))
        for index, (_, data) in enumerate(sheets, start=1):
            zf.writestr(f"xl/worksheets/sheet{index}.xml", sheet_xml(data))
    return payload.getvalue()


def audit_xlsx(
    summary: dict[str, Any],
    issues: list[dict[str, str]],
    scope_rows: list[dict[str, str]],
) -> bytes:
    summary_rows = [["Metric", "Value"]] + [[key, str(value)] for key, value in summary.items()]
    issue_headers = ["Severity", "Type", "EAN", "Article", "Detail", "Recommended action"]
    issue_rows = [issue_headers] + [[issue.get(header, "") for header in issue_headers] for issue in issues]
    scope_headers = ["EAN", "Article", "Size UA", "Listing Scope", "In OOB", "In Current Line List", "GHL", "Hero Look Name", "Latest Change Log", "Latest Change Date", "Source Material Number"]
    scope_table = [scope_headers] + [[row.get(header, "") for header in scope_headers] for row in scope_rows]
    return _write_minimal_xlsx([("Summary", summary_rows), ("Exceptions", issue_rows), ("Scope", scope_table)])


def build_files(
    oob_bytes: bytes,
    material_bytes: bytes,
    line_list_bytes: bytes,
    changelog_bytes: bytes,
    template_bytes: bytes,
    centric_underwear_bytes: bytes,
    centric_outerwear_bytes: bytes,
    centric_sportswear_bytes: bytes,
    season: str,
) -> tuple[bytes, bytes, dict[str, Any], list[dict[str, str]]]:
    rows, summary, issues, scope_rows = build_listing(
        oob_bytes, material_bytes, line_list_bytes, changelog_bytes, template_bytes,
        centric_underwear_bytes, centric_outerwear_bytes, centric_sportswear_bytes, season
    )
    return (
        listing_xlsx_from_template(template_bytes, rows),
        audit_xlsx(summary, issues, scope_rows),
        summary,
        issues,
    )

def build_from_paths(
    oob_path: str | Path,
    material_path: str | Path,
    line_list_path: str | Path,
    changelog_path: str | Path,
    template_path: str | Path,
    centric_underwear_path: str | Path,
    centric_outerwear_path: str | Path,
    centric_sportswear_path: str | Path,
    season: str,
) -> tuple[bytes, bytes, dict[str, Any], list[dict[str, str]]]:
    return build_files(
        Path(oob_path).read_bytes(),
        Path(material_path).read_bytes(),
        Path(line_list_path).read_bytes(),
        Path(changelog_path).read_bytes(),
        Path(template_path).read_bytes(),
        Path(centric_underwear_path).read_bytes(),
        Path(centric_outerwear_path).read_bytes(),
        Path(centric_sportswear_path).read_bytes(),
        season,
    )
