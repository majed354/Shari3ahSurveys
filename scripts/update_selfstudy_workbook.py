#!/usr/bin/env python3

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from collections import defaultdict
from copy import deepcopy
from pathlib import Path
from typing import Dict, Iterable, List, Tuple
from xml.etree import ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"a": NS_MAIN, "pr": NS_REL}

ET.register_namespace("", NS_MAIN)

ROOT = Path(__file__).resolve().parents[1]
MASTER_PATH = ROOT / "استطلاعات_وتقييمات_كلية_الشريعة_1445_1446.xlsx"
SELFSTUDY_PATH = ROOT / "استطلاعات الدراسة الذاتية.xlsx"
BACKUP_PATH = ROOT / "استطلاعات الدراسة الذاتية.backup.xlsx"

SHEET_CONFIGS = {
    "بكالوريوس - فهرسة تفصيلية": {
        "path": "xl/worksheets/sheet1.xml",
        "key_headers": {
            "survey": "نوع الاستطلاع",
            "phrase": "نص العبارة",
            "program": "البرنامج",
            "degree": "الدرجة العلمية",
            "semester": "الفصل",
            "count": "عدد المقيمين",
            "avg": "متوسط التقييم",
            "male_count": "عدد المقيمين (ذكور)",
            "male_avg": "متوسط التقييم (ذكور)",
            "female_count": "عدد المقيمين (إناث)",
            "female_avg": "متوسط التقييم (إناث)",
        },
    },
    "ماجستير - فهرسة تفصيلية": {
        "path": "xl/worksheets/sheet3.xml",
        "key_headers": {
            "survey": "نوع الاستطلاع",
            "phrase": "نص العبارة",
            "program": "البرنامج",
            "degree": "الدرجة العلمية",
            "semester": "الفصل",
            "count": "عدد المقيمين",
            "avg": "متوسط التقييم",
            "male_count": "عدد المقيمين (ذكور)",
            "male_avg": "متوسط التقييم (ذكور)",
            "female_count": "عدد المقيمين (إناث)",
            "female_avg": "متوسط التقييم (إناث)",
        },
    },
    "عبارات غير مرتبطة": {
        "path": "xl/worksheets/sheet5.xml",
        "key_headers": {
            "survey": "نوع الاستطلاع",
            "phrase": "نص العبارة",
            "program": "البرنامج",
            "degree": "الدرجة العلمية",
            "semester": "الفصل",
            "count": "عدد المقيمين",
            "avg": "متوسط التقييم",
            "male_count": "عدد المقيمين (ذكور)",
            "male_avg": "متوسط التقييم (ذكور)",
            "female_count": "عدد المقيمين (إناث)",
            "female_avg": "متوسط التقييم (إناث)",
        },
    },
}


def clean_text(value: object) -> str:
    text = str(value or "")
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def numeric_int(value: object) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def numeric_float(value: object) -> float:
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return 0.0


def weighted_average(score: int, count: int) -> float:
    return round(score / count, 2) if count else 0.0


def column_to_index(column: str) -> int:
    value = 0
    for char in column:
        value = (value * 26) + (ord(char.upper()) - 64)
    return value


def index_to_column(index: int) -> str:
    letters: List[str] = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def split_cell_ref(ref: str) -> Tuple[str, str]:
    match = re.match(r"([A-Z]+)(\d+)", ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    return match.group(1), match.group(2)


def parse_range(ref: str) -> Tuple[str, int]:
    start, end = ref.split(":")
    end_col, end_row = split_cell_ref(end)
    return end_col, int(end_row)


def read_shared_strings(archive: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: List[str] = []
    for si in root.findall("a:si", NS):
        values.append("".join(node.text or "" for node in si.iterfind(".//a:t", NS)))
    return values


def cell_text(cell: ET.Element, shared_strings: List[str]) -> str:
    if cell is None:
        return ""
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return clean_text("".join(node.text or "" for node in cell.iterfind(".//a:t", NS)))
    value_node = cell.find("a:v", NS)
    if value_node is None:
        return ""
    raw = value_node.text or ""
    if cell_type == "s":
        raw = shared_strings[int(raw)]
    return clean_text(raw)


def build_master_index() -> Dict[Tuple[str, str, str, str], Dict[str, object]]:
    aggregate = defaultdict(
        lambda: {
            "total_count": 0,
            "total_score": 0,
            "male_count": 0,
            "male_score": 0,
            "female_count": 0,
            "female_score": 0,
            "semesters": set(),
        }
    )

    with zipfile.ZipFile(MASTER_PATH) as archive:
        shared_strings = read_shared_strings(archive)
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        rows = root.find("a:sheetData", NS).findall("a:row", NS)

        for row in rows[1:]:
            values: Dict[str, str] = {}
            for cell in row.findall("a:c", NS):
                ref = cell.get("r", "")
                column, _ = split_cell_ref(ref)
                values[column] = cell_text(cell, shared_strings)

            program = clean_text(values.get("F"))
            degree = clean_text(values.get("C"))
            survey = clean_text(values.get("G"))
            phrase = clean_text(values.get("J"))
            semester = clean_text(values.get("A"))
            gender = clean_text(values.get("B"))

            counts = [numeric_int(values.get(column)) for column in ("K", "L", "M", "N", "O")]
            response_count = sum(counts)
            if not (program and degree and survey and phrase and response_count):
                continue

            score_total = sum((index + 1) * count for index, count in enumerate(counts))
            item = aggregate[(program, degree, survey, phrase)]
            item["total_count"] += response_count
            item["total_score"] += score_total
            item["semesters"].add(semester)

            if gender == "ذكر":
                item["male_count"] += response_count
                item["male_score"] += score_total
            elif gender == "إناث":
                item["female_count"] += response_count
                item["female_score"] += score_total

    index: Dict[Tuple[str, str, str, str], Dict[str, object]] = {}
    for key, item in aggregate.items():
        index[key] = {
            "semester_text": "، ".join(sorted(item["semesters"])),
            "total_count": item["total_count"],
            "total_avg": weighted_average(item["total_score"], item["total_count"]),
            "male_count": item["male_count"],
            "male_avg": weighted_average(item["male_score"], item["male_count"]),
            "female_count": item["female_count"],
            "female_avg": weighted_average(item["female_score"], item["female_count"]),
        }
    return index


def create_inline_cell(ref: str, text: str, style: str | None) -> ET.Element:
    cell = ET.Element(f"{{{NS_MAIN}}}c", {"r": ref, "t": "inlineStr"})
    if style:
        cell.set("s", style)
    is_node = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
    text_node = ET.SubElement(is_node, f"{{{NS_MAIN}}}t")
    if text.startswith(" ") or text.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    return cell


def create_number_cell(ref: str, value: int | float, style: str | None) -> ET.Element:
    cell = ET.Element(f"{{{NS_MAIN}}}c", {"r": ref, "t": "n"})
    if style:
        cell.set("s", style)
    value_node = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
    if isinstance(value, int) or float(value).is_integer():
        value_node.text = str(int(round(float(value))))
    else:
        value_node.text = f"{float(value):.2f}"
    return cell


def upsert_inline_cell(row: ET.Element, column_index: int, text: str, style: str | None) -> None:
    target_ref = f"{index_to_column(column_index)}{row.get('r')}"
    for cell in row.findall("a:c", NS):
        if cell.get("r") == target_ref:
            cell.clear()
            cell.tag = f"{{{NS_MAIN}}}c"
            cell.set("r", target_ref)
            cell.set("t", "inlineStr")
            if style:
                cell.set("s", style)
            is_node = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
            text_node = ET.SubElement(is_node, f"{{{NS_MAIN}}}t")
            text_node.text = text
            return

    new_cell = create_inline_cell(target_ref, text, style)
    row.append(new_cell)
    reorder_row_cells(row)


def upsert_number_cell(row: ET.Element, column_index: int, value: int | float, style: str | None) -> None:
    target_ref = f"{index_to_column(column_index)}{row.get('r')}"
    for cell in row.findall("a:c", NS):
        if cell.get("r") == target_ref:
            cell.clear()
            cell.tag = f"{{{NS_MAIN}}}c"
            cell.set("r", target_ref)
            cell.set("t", "n")
            if style:
                cell.set("s", style)
            value_node = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
            if isinstance(value, int) or float(value).is_integer():
                value_node.text = str(int(round(float(value))))
            else:
                value_node.text = f"{float(value):.2f}"
            return

    new_cell = create_number_cell(target_ref, value, style)
    row.append(new_cell)
    reorder_row_cells(row)


def reorder_row_cells(row: ET.Element) -> None:
    cells = row.findall("a:c", NS)
    ordered = sorted(cells, key=lambda cell: column_to_index(split_cell_ref(cell.get("r"))[0]))
    for cell in cells:
        row.remove(cell)
    for cell in ordered:
        row.append(cell)


def insert_column(sheet_root: ET.Element, insert_at: int, header_style: str | None, data_style: str | None) -> None:
    sheet_data = sheet_root.find("a:sheetData", NS)
    rows = sheet_data.findall("a:row", NS)

    for row in rows:
        row_number = row.get("r")
        mapping: Dict[int, ET.Element] = {}
        for cell in row.findall("a:c", NS):
            column, _ = split_cell_ref(cell.get("r"))
            column_index = column_to_index(column)
            if column_index >= insert_at:
                column_index += 1
            new_ref = f"{index_to_column(column_index)}{row_number}"
            cell.set("r", new_ref)
            mapping[column_index] = cell

        for cell in row.findall("a:c", NS):
            row.remove(cell)

        style = header_style if row_number == "1" else data_style
        text = "الفصل" if row_number == "1" else ""
        mapping[insert_at] = create_inline_cell(f"{index_to_column(insert_at)}{row_number}", text, style)

        for column_index in sorted(mapping):
            row.append(mapping[column_index])

    cols = sheet_root.find("a:cols", NS)
    if cols is not None:
        templates = {int(col.get("min")): deepcopy(col) for col in cols.findall("a:col", NS)}
        for col in cols.findall("a:col", NS):
            cols.remove(col)

        template = deepcopy(templates.get(insert_at - 1) or next(iter(templates.values())))
        template.set("min", str(insert_at))
        template.set("max", str(insert_at))
        template.set("width", "16")
        template.set("customWidth", "1")

        rebuilt = {}
        for index, col in templates.items():
            target_index = index + 1 if index >= insert_at else index
            col.set("min", str(target_index))
            col.set("max", str(target_index))
            rebuilt[target_index] = col
        rebuilt[insert_at] = template

        for column_index in sorted(rebuilt):
            cols.append(rebuilt[column_index])

    dimension = sheet_root.find("a:dimension", NS)
    if dimension is not None and ":" in dimension.get("ref", ""):
        _, end_row = parse_range(dimension.get("ref"))
        new_last_col = index_to_column(insert_at + len(rows[0].findall("a:c", NS)) - insert_at)
        dimension.set("ref", f"A1:{new_last_col}{end_row}")

    auto_filter = sheet_root.find("a:autoFilter", NS)
    if auto_filter is not None and ":" in auto_filter.get("ref", ""):
        _, end_row = parse_range(auto_filter.get("ref"))
        new_last_col = index_to_column(len(rows[0].findall("a:c", NS)))
        auto_filter.set("ref", f"A1:{new_last_col}{end_row}")


def header_map(header_row: ET.Element) -> Dict[str, int]:
    mapping = {}
    for cell in header_row.findall("a:c", NS):
        column, _ = split_cell_ref(cell.get("r"))
        mapping[cell_text(cell, [])] = column_to_index(column)
    return mapping


def style_for_column(row: ET.Element, column_index: int) -> str | None:
    ref = f"{index_to_column(column_index)}{row.get('r')}"
    for cell in row.findall("a:c", NS):
        if cell.get("r") == ref:
            return cell.get("s")
    return None


def update_sheet(sheet_root: ET.Element, master_index: Dict[Tuple[str, str, str, str], Dict[str, object]], config: Dict[str, object]) -> Dict[str, int]:
    sheet_data = sheet_root.find("a:sheetData", NS)
    rows = sheet_data.findall("a:row", NS)
    header_row = rows[0]
    headers = header_map(header_row)
    degree_column = headers["الدرجة العلمية"]

    inserted = 0
    if "الفصل" not in headers:
        header_style = style_for_column(header_row, degree_column)
        data_style = style_for_column(rows[1], degree_column) if len(rows) > 1 else None
        insert_column(sheet_root, degree_column + 1, header_style, data_style)
        inserted = 1
        rows = sheet_root.find("a:sheetData", NS).findall("a:row", NS)
        header_row = rows[0]
        headers = header_map(header_row)

    semester_column = headers[config["key_headers"]["semester"]]
    survey_column = headers[config["key_headers"]["survey"]]
    phrase_column = headers[config["key_headers"]["phrase"]]
    program_column = headers[config["key_headers"]["program"]]
    degree_column = headers[config["key_headers"]["degree"]]

    count_column = headers[config["key_headers"]["count"]]
    avg_column = headers[config["key_headers"]["avg"]]
    male_count_column = headers[config["key_headers"]["male_count"]]
    male_avg_column = headers[config["key_headers"]["male_avg"]]
    female_count_column = headers[config["key_headers"]["female_count"]]
    female_avg_column = headers[config["key_headers"]["female_avg"]]

    corrected = 0
    missing = 0

    for row in rows[1:]:
        row_cells = {
            column_to_index(split_cell_ref(cell.get("r"))[0]): cell for cell in row.findall("a:c", NS)
        }
        survey = cell_text(row_cells.get(survey_column), [])
        phrase = cell_text(row_cells.get(phrase_column), [])
        program = cell_text(row_cells.get(program_column), [])
        degree = cell_text(row_cells.get(degree_column), [])

        key = (program, degree, survey, phrase)
        master = master_index.get(key)
        if master is None:
            missing += 1
            continue

        text_style = row_cells.get(degree_column).get("s") if row_cells.get(degree_column) is not None else None
        count_style = row_cells.get(count_column).get("s") if row_cells.get(count_column) is not None else None
        avg_style = row_cells.get(avg_column).get("s") if row_cells.get(avg_column) is not None else None

        current_semester = cell_text(row_cells.get(semester_column), [])
        current_total_count = numeric_int(cell_text(row_cells.get(count_column), []))
        current_total_avg = numeric_float(cell_text(row_cells.get(avg_column), []))
        current_male_count = numeric_int(cell_text(row_cells.get(male_count_column), []))
        current_male_avg = numeric_float(cell_text(row_cells.get(male_avg_column), []))
        current_female_count = numeric_int(cell_text(row_cells.get(female_count_column), []))
        current_female_avg = numeric_float(cell_text(row_cells.get(female_avg_column), []))

        needs_update = any(
            [
                current_semester != master["semester_text"],
                current_total_count != master["total_count"],
                abs(current_total_avg - master["total_avg"]) > 0.01,
                current_male_count != master["male_count"],
                abs(current_male_avg - master["male_avg"]) > 0.01,
                current_female_count != master["female_count"],
                abs(current_female_avg - master["female_avg"]) > 0.01,
            ]
        )

        upsert_inline_cell(row, semester_column, str(master["semester_text"]), text_style)
        upsert_number_cell(row, count_column, int(master["total_count"]), count_style)
        upsert_number_cell(row, avg_column, float(master["total_avg"]), avg_style)
        upsert_number_cell(row, male_count_column, int(master["male_count"]), count_style)
        upsert_number_cell(row, male_avg_column, float(master["male_avg"]), avg_style)
        upsert_number_cell(row, female_count_column, int(master["female_count"]), count_style)
        upsert_number_cell(row, female_avg_column, float(master["female_avg"]), avg_style)

        if needs_update:
            corrected += 1

    return {"inserted": inserted, "corrected": corrected, "missing": missing}


def main() -> None:
    master_index = build_master_index()

    shutil.copy2(SELFSTUDY_PATH, BACKUP_PATH)

    with zipfile.ZipFile(SELFSTUDY_PATH) as source:
        source_bytes = {name: source.read(name) for name in source.namelist()}

    results = {}
    for sheet_name, config in SHEET_CONFIGS.items():
        path = config["path"]
        root = ET.fromstring(source_bytes[path])
        results[sheet_name] = update_sheet(root, master_index, config)
        source_bytes[path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
        temp_path = Path(temp_file.name)

    try:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as output:
            for name, data in source_bytes.items():
                output.writestr(name, data)
        shutil.move(temp_path, SELFSTUDY_PATH)
    finally:
        if temp_path.exists():
            temp_path.unlink()

    print("تم تحديث ملف الدراسة الذاتية بنجاح.")
    print(f"نسخة احتياطية: {BACKUP_PATH.name}")
    for sheet_name, result in results.items():
        print(
            f"{sheet_name}: added_semester_column={result['inserted']} "
            f"corrected_rows={result['corrected']} missing_rows={result['missing']}"
        )


if __name__ == "__main__":
    main()
