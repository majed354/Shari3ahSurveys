#!/usr/bin/env python3

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Tuple
from xml.etree import ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"a": NS_MAIN, "pr": NS_REL}

ET.register_namespace("", NS_MAIN)

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_PATH = ROOT / "استطلاعات الدراسة الذاتية المحدث.xlsx"
BACKUP_PATH = ROOT / "استطلاعات الدراسة الذاتية المحدث.pre-unified-backup.xlsx"

PRIMARY_SHEET_NAME = "فهرسة تفصيلية موحدة"
LEGACY_BACHELOR_SHEET = "بكالوريوس - فهرسة تفصيلية"
LEGACY_MASTER_SHEET = "ماجستير - فهرسة تفصيلية"
ARCHIVE_SHEET_NAME = "أرشيف تفصيلي قديم"
SOURCE_DETAIL_SHEETS = (
    LEGACY_BACHELOR_SHEET,
    LEGACY_MASTER_SHEET,
    PRIMARY_SHEET_NAME,
    ARCHIVE_SHEET_NAME,
)

PG_TEXT_OVERRIDES = {
    "تتنوع استراتيجيات التعليم والتعلم وطرق التقييم في البرنامج بما يتناسب مع طبيعته ومستواه": {
        "criterion_code": "2-3-2",
        "standard": "التعليم والتعلم",
        "section": "جودة التدريس وتقييم الطلاب",
        "status": "معتمد",
        "note": "تم تصحيح الرقم وفق دليل الدراسات العليا.",
    },
    "ترتبط نواتج التعلم في المقررات مع نواتج التعلم في البرنامج (مصفوفة توزيع نواتج تعلم البرنامج على المقررات)": {
        "criterion_code": "3-2-2",
        "standard": "التعليم والتعلم",
        "section": "المنهج الدراسي",
        "status": "معتمد",
        "note": "تم تصحيح الرقم وفق دليل الدراسات العليا.",
    },
    "تُطبق آليات فعالة لتقويم كفاية وجودة الخدمات المقدمة للطلاب وقياس رضاهم عنها والاستفادة من النتائج في التحسين": {
        "criterion_code": "6-0-3",
        "standard": "الطلاب",
        "section": "",
        "status": "معتمد",
        "note": "تم تصحيح الرقم وفق دليل الدراسات العليا.",
    },
    "يتأكد البرنامج من تطبيق موحد للخطة الدراسية وتوصيف البرنامج والمقررات في أكثر من موقع": {
        "criterion_code": "4-2-2",
        "standard": "التعليم والتعلم",
        "section": "المنهج الدراسي",
        "status": "معتمد",
        "note": "تم تصحيح الرقم وفق دليل الدراسات العليا.",
    },
    "يُزوَّد الطلاب في بداية تدريس كل مقرر بمعلومات شاملة عنه: نواتج التعلم، واستراتيجيات التعليم والتعلم وطرق التقييم، ومواعيدها": {
        "criterion_code": "4-3-2",
        "standard": "التعليم والتعلم",
        "section": "جودة التدريس وتقييم الطلاب",
        "status": "معتمد",
        "note": "تم تصحيح الرقم وفق دليل الدراسات العليا.",
    },
    "يُقدَّم التدريب اللازم لهيئة التدريس على استراتيجيات التعليم والتعلم وطرق التقييم والاستخدام الفعال للتقنية": {
        "criterion_code": "3-3-2",
        "standard": "التعليم والتعلم",
        "section": "جودة التدريس وتقييم الطلاب",
        "status": "معتمد",
        "note": "تم تصحيح الرقم وفق دليل الدراسات العليا.",
    },
    "يطبق البرنامج آليات لدعم وتحفيز التميز في التدريس وتشجيع الإبداع والابتكار لدى هيئة التدريس": {
        "criterion_code": "5-3-2",
        "standard": "التعليم والتعلم",
        "section": "جودة التدريس وتقييم الطلاب",
        "status": "معتمد",
        "note": "تم تصحيح الرقم وفق دليل الدراسات العليا.",
    },
    "يطبق البرنامج إجراءات واضحة ومعلنة للتحقق من جودة طرق التقييم ومصداقيتها والتأكد من مستوى تحصيل الطلاب": {
        "criterion_code": "6-3-2",
        "standard": "التعليم والتعلم",
        "section": "جودة التدريس وتقييم الطلاب",
        "status": "معتمد",
        "note": "تم تصحيح الرقم وفق دليل الدراسات العليا.",
    },
}

PG_REVIEW_TEXTS = {
    "تتوفر لطلاب البرنامج أنشطة لا صفية في العديد من المجالات لتنمية قدراتهم ومهاراتهم": "لم يُعثر على رقم مطابق بوضوح في دليل الدراسات العليا، ويحتاج هذا الربط إلى مراجعة بشرية.",
    "يتحقق البرنامج من فعالية التدريب الميداني وجودة الإشراف عليه": "لم يُعثر على رقم مطابق بوضوح في دليل الدراسات العليا، ويحتاج هذا الربط إلى مراجعة بشرية.",
}

TARGET_HEADERS = [
    "المعيار",
    "القسم",
    "رقم المحك",
    "نص المحك",
    "جانب المحك المدعوم",
    "نوع الاستطلاع",
    "رقم العبارة",
    "نص العبارة",
    "البرنامج",
    "الدرجة العلمية",
    "الفصل الدراسي",
    "عدد المقيمين",
    "متوسط التقييم",
    "عدد المقيمين (ذكور)",
    "متوسط التقييم (ذكور)",
    "عدد المقيمين (إناث)",
    "متوسط التقييم (إناث)",
    "فئة الدرجة",
    "رقم المحك الأصلي",
    "حالة الربط",
    "ملاحظة الربط",
    "مصدر السجل",
]

TEXT_COLUMNS = {
    "المعيار",
    "القسم",
    "رقم المحك",
    "نص المحك",
    "جانب المحك المدعوم",
    "نوع الاستطلاع",
    "رقم العبارة",
    "نص العبارة",
    "البرنامج",
    "الدرجة العلمية",
    "الفصل الدراسي",
    "فئة الدرجة",
    "رقم المحك الأصلي",
    "حالة الربط",
    "ملاحظة الربط",
    "مصدر السجل",
}

INT_COLUMNS = {"عدد المقيمين", "عدد المقيمين (ذكور)", "عدد المقيمين (إناث)"}
FLOAT_COLUMNS = {"متوسط التقييم", "متوسط التقييم (ذكور)", "متوسط التقييم (إناث)"}


def clean_text(value: object) -> str:
    text = str(value or "")
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_criterion_code(value: object) -> str:
    text = clean_text(value)
    digits = re.findall(r"\d+", text)
    if len(digits) >= 3:
        return "-".join(digits[:3])
    return text


def degree_bucket(degree: str) -> str:
    degree = clean_text(degree)
    if degree in {"البكالوريوس", "بكالوريوس انتساب"}:
        return "بكالوريوس"
    if degree in {"الماجستير", "الدكتوراه"}:
        return "دراسات عليا"
    return degree


def numeric_int(value: object) -> int:
    try:
        return int(round(float(str(value or "0").replace(",", ""))))
    except (TypeError, ValueError):
        return 0


def numeric_float(value: object) -> float:
    try:
        return round(float(str(value or "0").replace(",", "")), 2)
    except (TypeError, ValueError):
        return 0.0


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


def read_shared_strings(archive: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: List[str] = []
    for si in root.findall("a:si", NS):
        values.append("".join(node.text or "" for node in si.iterfind(".//a:t", NS)))
    return values


def cell_text(cell: ET.Element | None, shared_strings: List[str]) -> str:
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


def load_sheet_paths(archive: zipfile.ZipFile) -> Dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.get("Id"): rel.get("Target") for rel in rels.findall("pr:Relationship", NS)}
    mapping: Dict[str, str] = {}
    for sheet in workbook.findall("a:sheets/a:sheet", NS):
        name = sheet.get("name") or ""
        rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rel_map.get(rid, "").lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        mapping[name] = target
    return mapping


def header_map(header_row: ET.Element, shared_strings: List[str]) -> Dict[str, int]:
    mapping = {}
    for cell in header_row.findall("a:c", NS):
        column, _ = split_cell_ref(cell.get("r"))
        mapping[cell_text(cell, shared_strings)] = column_to_index(column)
    return mapping


def style_for_column(row: ET.Element, column_index: int) -> str | None:
    ref = f"{index_to_column(column_index)}{row.get('r')}"
    for cell in row.findall("a:c", NS):
        if cell.get("r") == ref:
            return cell.get("s")
    return None


def iter_source_rows(
    archive: zipfile.ZipFile,
    shared_strings: List[str],
    sheet_paths: Dict[str, str],
) -> Iterable[Dict[str, object]]:
    required = {
        "المعيار",
        "القسم",
        "رقم المحك",
        "نص المحك",
        "جانب المحك المدعوم",
        "نوع الاستطلاع",
        "رقم العبارة",
        "نص العبارة",
        "البرنامج",
        "الدرجة العلمية",
        "الفصل الدراسي",
        "عدد المقيمين",
        "متوسط التقييم",
        "عدد المقيمين (ذكور)",
        "متوسط التقييم (ذكور)",
        "عدد المقيمين (إناث)",
        "متوسط التقييم (إناث)",
    }

    for sheet_name in SOURCE_DETAIL_SHEETS:
        path = sheet_paths.get(sheet_name)
        if not path:
            continue

        root = ET.fromstring(archive.read(path))
        rows = root.find("a:sheetData", NS).findall("a:row", NS)
        if not rows:
            continue

        headers = header_map(rows[0], shared_strings)
        if not required.issubset(headers):
            continue

        for row in rows[1:]:
            row_cells = {
                column_to_index(split_cell_ref(cell.get("r"))[0]): cell
                for cell in row.findall("a:c", NS)
            }

            def get(label: str) -> str:
                return cell_text(row_cells.get(headers[label]), shared_strings)

            yield {
                "sourceSheet": sheet_name,
                "standard": get("المعيار"),
                "section": get("القسم"),
                "criterionCode": normalize_criterion_code(get("رقم المحك")),
                "criterionText": get("نص المحك"),
                "supportedSide": get("جانب المحك المدعوم"),
                "surveyTitle": get("نوع الاستطلاع"),
                "itemNumber": get("رقم العبارة"),
                "itemText": get("نص العبارة"),
                "program": get("البرنامج"),
                "degree": get("الدرجة العلمية"),
                "semester": get("الفصل الدراسي"),
                "respondents": numeric_int(get("عدد المقيمين")),
                "average": numeric_float(get("متوسط التقييم")),
                "maleRespondents": numeric_int(get("عدد المقيمين (ذكور)")),
                "maleAverage": numeric_float(get("متوسط التقييم (ذكور)")),
                "femaleRespondents": numeric_int(get("عدد المقيمين (إناث)")),
                "femaleAverage": numeric_float(get("متوسط التقييم (إناث)")),
            }


def sort_key_for_criterion(code: str) -> Tuple[int, ...]:
    digits = [int(part) for part in re.findall(r"\d+", clean_text(code))]
    if not digits:
        return (999, 999, 999)
    while len(digits) < 3:
        digits.append(999)
    return tuple(digits[:3])


def resolve_mapping(row: Dict[str, object]) -> Dict[str, str]:
    degree = clean_text(row["degree"])
    text = clean_text(row["criterionText"])
    original_code = normalize_criterion_code(row["criterionCode"])
    standard = clean_text(row["standard"])
    section = clean_text(row["section"])

    if degree_bucket(degree) != "دراسات عليا":
        return {
            "criterionCode": original_code,
            "standard": standard,
            "section": section,
            "status": "معتمد",
            "note": "",
            "originalCode": original_code,
        }

    if text in PG_TEXT_OVERRIDES:
        override = PG_TEXT_OVERRIDES[text]
        return {
            "criterionCode": override["criterion_code"],
            "standard": override["standard"],
            "section": override["section"],
            "status": override["status"],
            "note": override["note"],
            "originalCode": original_code,
        }

    if text in PG_REVIEW_TEXTS:
        return {
            "criterionCode": "",
            "standard": standard,
            "section": section,
            "status": "يحتاج مراجعة",
            "note": PG_REVIEW_TEXTS[text],
            "originalCode": original_code,
        }

    return {
        "criterionCode": original_code,
        "standard": standard,
        "section": section,
        "status": "معتمد",
        "note": "",
        "originalCode": original_code,
    }


def build_unified_rows(
    archive: zipfile.ZipFile,
    shared_strings: List[str],
    sheet_paths: Dict[str, str],
) -> List[Dict[str, object]]:
    grouped: Dict[Tuple[str, ...], List[Dict[str, object]]] = defaultdict(list)

    for row in iter_source_rows(archive, shared_strings, sheet_paths):
        key = (
            clean_text(row["surveyTitle"]),
            clean_text(row["itemNumber"]),
            clean_text(row["itemText"]),
            clean_text(row["program"]),
            clean_text(row["degree"]),
            clean_text(row["semester"]),
            clean_text(row["criterionText"]),
            clean_text(row["supportedSide"]),
        )
        grouped[key].append(row)

    unified: List[Dict[str, object]] = []
    for rows in grouped.values():
        rows = sorted(rows, key=lambda item: item["sourceSheet"])
        base = rows[0]
        resolved = resolve_mapping(base)
        unified.append(
            {
                "المعيار": resolved["standard"],
                "القسم": resolved["section"],
                "رقم المحك": resolved["criterionCode"],
                "نص المحك": clean_text(base["criterionText"]),
                "جانب المحك المدعوم": clean_text(base["supportedSide"]),
                "نوع الاستطلاع": clean_text(base["surveyTitle"]),
                "رقم العبارة": clean_text(base["itemNumber"]),
                "نص العبارة": clean_text(base["itemText"]),
                "البرنامج": clean_text(base["program"]),
                "الدرجة العلمية": clean_text(base["degree"]),
                "الفصل الدراسي": clean_text(base["semester"]),
                "عدد المقيمين": int(base["respondents"]),
                "متوسط التقييم": float(base["average"]),
                "عدد المقيمين (ذكور)": int(base["maleRespondents"]),
                "متوسط التقييم (ذكور)": float(base["maleAverage"]),
                "عدد المقيمين (إناث)": int(base["femaleRespondents"]),
                "متوسط التقييم (إناث)": float(base["femaleAverage"]),
                "فئة الدرجة": degree_bucket(clean_text(base["degree"])),
                "رقم المحك الأصلي": resolved["originalCode"],
                "حالة الربط": resolved["status"],
                "ملاحظة الربط": resolved["note"],
                "مصدر السجل": " | ".join(sorted({clean_text(item["sourceSheet"]) for item in rows})),
            }
        )

    unified.sort(
        key=lambda row: (
            0 if row["فئة الدرجة"] == "بكالوريوس" else 1,
            clean_text(row["الدرجة العلمية"]),
            clean_text(row["البرنامج"]),
            clean_text(row["الفصل الدراسي"]),
            sort_key_for_criterion(clean_text(row["رقم المحك"])),
            clean_text(row["نص المحك"]),
            clean_text(row["نوع الاستطلاع"]),
            clean_text(row["رقم العبارة"]),
        )
    )
    return unified


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


def replace_primary_sheet_rows(
    sheet_root: ET.Element,
    unified_rows: List[Dict[str, object]],
) -> None:
    sheet_data = sheet_root.find("a:sheetData", NS)
    existing_rows = sheet_data.findall("a:row", NS)
    if not existing_rows:
        raise RuntimeError("Primary sheet has no rows to use as style template.")

    header_row = existing_rows[0]
    sample_row = existing_rows[1] if len(existing_rows) > 1 else existing_rows[0]

    header_style = style_for_column(header_row, 1)
    text_style = style_for_column(sample_row, 1) or style_for_column(sample_row, 2)
    int_style = style_for_column(sample_row, 12)
    float_style = style_for_column(sample_row, 13)

    for row in existing_rows:
        sheet_data.remove(row)

    new_rows: List[ET.Element] = []

    header = ET.Element(f"{{{NS_MAIN}}}row", {"r": "1"})
    for column_index, label in enumerate(TARGET_HEADERS, start=1):
        header.append(create_inline_cell(f"{index_to_column(column_index)}1", label, header_style))
    new_rows.append(header)

    for row_number, values in enumerate(unified_rows, start=2):
        row = ET.Element(f"{{{NS_MAIN}}}row", {"r": str(row_number)})
        for column_index, label in enumerate(TARGET_HEADERS, start=1):
            ref = f"{index_to_column(column_index)}{row_number}"
            value = values.get(label, "")
            if label in INT_COLUMNS:
                row.append(create_number_cell(ref, int(value), int_style))
            elif label in FLOAT_COLUMNS:
                row.append(create_number_cell(ref, float(value), float_style))
            else:
                row.append(create_inline_cell(ref, clean_text(value), text_style))
        new_rows.append(row)

    for row in new_rows:
        sheet_data.append(row)

    last_col = index_to_column(len(TARGET_HEADERS))
    last_row = len(unified_rows) + 1
    new_ref = f"A1:{last_col}{last_row}"

    dimension = sheet_root.find("a:dimension", NS)
    if dimension is None:
        dimension = ET.Element(f"{{{NS_MAIN}}}dimension")
        sheet_root.insert(0, dimension)
    dimension.set("ref", new_ref)

    auto_filter = sheet_root.find("a:autoFilter", NS)
    if auto_filter is None:
        auto_filter = ET.Element(f"{{{NS_MAIN}}}autoFilter")
        sheet_root.append(auto_filter)
    auto_filter.set("ref", new_ref)


def rename_and_hide_sheets(workbook_root: ET.Element) -> None:
    for sheet in workbook_root.findall("a:sheets/a:sheet", NS):
        name = sheet.get("name") or ""
        if name == LEGACY_BACHELOR_SHEET:
            sheet.set("name", PRIMARY_SHEET_NAME)
            if "state" in sheet.attrib:
                del sheet.attrib["state"]
        elif name == PRIMARY_SHEET_NAME:
            if "state" in sheet.attrib:
                del sheet.attrib["state"]
        elif name == LEGACY_MASTER_SHEET:
            sheet.set("name", ARCHIVE_SHEET_NAME)
            sheet.set("state", "hidden")
        elif name == ARCHIVE_SHEET_NAME:
            sheet.set("state", "hidden")


def update_defined_names(workbook_root: ET.Element, unified_row_count: int) -> None:
    defined_names = workbook_root.find("a:definedNames", NS)
    if defined_names is None:
        return

    primary_ref = f"'{PRIMARY_SHEET_NAME}'!$A$1:$V${unified_row_count + 1}"
    archive_ref = f"'{ARCHIVE_SHEET_NAME}'!$A$1:$P$7280"

    for defined in defined_names.findall("a:definedName", NS):
        text = defined.text or ""
        text = text.replace(f"'{LEGACY_BACHELOR_SHEET}'!$A$1:$P$7519", primary_ref)
        text = text.replace(f"'{PRIMARY_SHEET_NAME}'!$A$1:$P$7519", primary_ref)
        text = text.replace(f"'{PRIMARY_SHEET_NAME}'!$A$1:$V$7519", primary_ref)
        text = text.replace(f"'{LEGACY_MASTER_SHEET}'!$A$1:$P$7280", archive_ref)
        defined.text = text


def main() -> None:
    shutil.copy2(WORKBOOK_PATH, BACKUP_PATH)

    with zipfile.ZipFile(WORKBOOK_PATH) as source:
        source_bytes = {name: source.read(name) for name in source.namelist()}
        shared_strings = read_shared_strings(source)
        sheet_paths = load_sheet_paths(source)
        unified_rows = build_unified_rows(source, shared_strings, sheet_paths)

    workbook_root = ET.fromstring(source_bytes["xl/workbook.xml"])
    rename_and_hide_sheets(workbook_root)
    update_defined_names(workbook_root, len(unified_rows))
    source_bytes["xl/workbook.xml"] = ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True)

    primary_path = sheet_paths.get(LEGACY_BACHELOR_SHEET) or sheet_paths.get(PRIMARY_SHEET_NAME)
    if not primary_path:
        raise RuntimeError("Could not locate primary detailed sheet.")

    primary_root = ET.fromstring(source_bytes[primary_path])
    replace_primary_sheet_rows(primary_root, unified_rows)
    source_bytes[primary_path] = ET.tostring(primary_root, encoding="utf-8", xml_declaration=True)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
        temp_path = Path(temp_file.name)

    try:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as output:
            for name, data in source_bytes.items():
                output.writestr(name, data)
        shutil.move(temp_path, WORKBOOK_PATH)
    finally:
        if temp_path.exists():
            temp_path.unlink()

    corrected = sum(1 for row in unified_rows if clean_text(row["رقم المحك"]) and clean_text(row["رقم المحك"]) != clean_text(row["رقم المحك الأصلي"]))
    review = sum(1 for row in unified_rows if clean_text(row["حالة الربط"]) == "يحتاج مراجعة")
    print("تم توحيد ملف الدراسة الذاتية المحدث بنجاح.")
    print(f"النسخة الاحتياطية: {BACKUP_PATH.name}")
    print(f"السجلات الموحدة: {len(unified_rows)}")
    print(f"السجلات المصححة: {corrected}")
    print(f"السجلات التي تحتاج مراجعة: {review}")


if __name__ == "__main__":
    main()
