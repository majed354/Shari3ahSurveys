#!/usr/bin/env python3

from __future__ import annotations

import json
import re
import zipfile
from collections import OrderedDict
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Tuple
from xml.etree import ElementTree as ET

from build_survey_data import PROGRAM_ID_MAP, clean_text


NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_PATH = ROOT / "استطلاعات الدراسة الذاتية المحدث.xlsx"
OUTPUT_PATH = ROOT / "js" / "self-study-data.js"


def normalize_text(value: object) -> str:
    return clean_text(value)


def extract_year(value: str) -> str:
    text = normalize_text(value)
    match = re.search(r"(14\d{2})", text)
    return match.group(1) if match else ""


def read_shared_strings(archive: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: List[str] = []
    for si in root.findall("a:si", NS):
        values.append("".join(node.text or "" for node in si.iterfind(".//a:t", NS)))
    return values


def load_workbook(archive: zipfile.ZipFile, shared_strings: List[str]) -> Dict[str, List[Dict[str, str]]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.get("Id"): rel.get("Target") for rel in rels.findall("pr:Relationship", NS)}
    sheets: Dict[str, List[Dict[str, str]]] = {}

    for sheet in workbook.findall("a:sheets/a:sheet", NS):
        name = sheet.get("name") or ""
        rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rel_map.get(rid, "").lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        root = ET.fromstring(archive.read(target))
        rows = root.find("a:sheetData", NS).findall("a:row", NS)
        parsed_rows: List[Dict[str, str]] = []

        for row in rows:
            parsed: Dict[str, str] = {}
            for cell in row.findall("a:c", NS):
                ref = cell.get("r", "")
                column = "".join(ch for ch in ref if ch.isalpha())
                cell_type = cell.get("t")
                if cell_type == "inlineStr":
                    raw = "".join(node.text or "" for node in cell.iterfind(".//a:t", NS))
                else:
                    value_node = cell.find("a:v", NS)
                    if value_node is None:
                        continue
                    raw = value_node.text or ""
                    if cell_type == "s":
                        raw = shared_strings[int(raw)]
                parsed[column] = normalize_text(raw)
            parsed_rows.append(parsed)

        sheets[name] = parsed_rows

    return sheets


def iter_detailed_rows(sheets: Dict[str, List[Dict[str, str]]]) -> Iterable[Dict[str, str]]:
    seen: set[Tuple[str, ...]] = set()

    for sheet_name, rows in sheets.items():
        if "فهرسة تفصيلية" not in sheet_name:
            continue

        if not rows:
            continue

        header = rows[0]
        columns = {value: key for key, value in header.items()}
        required = [
            "رقم المحك",
            "نص المحك",
            "جانب المحك المدعوم",
            "نوع الاستطلاع",
            "نص العبارة",
            "البرنامج",
            "الدرجة العلمية",
            "الفصل الدراسي",
        ]
        if not all(label in columns for label in required):
            continue

        for row in rows[1:]:
            criterion_code = normalize_text(row.get(columns["رقم المحك"]))
            criterion_text = normalize_text(row.get(columns["نص المحك"]))
            supported_side = normalize_text(row.get(columns["جانب المحك المدعوم"]))
            survey_title = normalize_text(row.get(columns["نوع الاستطلاع"]))
            phrase_text = normalize_text(row.get(columns["نص العبارة"]))
            program_name = normalize_text(row.get(columns["البرنامج"]))
            degree = normalize_text(row.get(columns["الدرجة العلمية"]))
            semester_label = normalize_text(row.get(columns["الفصل الدراسي"]))
            year = extract_year(semester_label)
            standard_label = normalize_text(row.get(columns.get("المعيار", ""), ""))
            section_label = normalize_text(row.get(columns.get("القسم", ""), ""))

            program_id = PROGRAM_ID_MAP.get((program_name, degree))
            if not program_id or not year:
                continue

            dedupe_key = (
                program_id,
                year,
                survey_title,
                phrase_text,
                criterion_code,
                criterion_text,
                supported_side,
                standard_label,
                section_label,
            )
            if dedupe_key in seen:
                continue
            seen.add(dedupe_key)

            yield {
                "programId": program_id,
                "programName": program_name,
                "degree": degree,
                "year": year,
                "standard": standard_label,
                "section": section_label,
                "criterionCode": criterion_code,
                "criterionText": criterion_text,
                "supportedSide": supported_side,
                "surveyTitle": survey_title,
                "phraseText": phrase_text,
            }


def make_item_key(program_id: str, year: str, survey_title: str, phrase_text: str) -> str:
    return "||".join(
        [
            normalize_text(program_id),
            normalize_text(year),
            normalize_text(survey_title),
            normalize_text(phrase_text),
        ]
    )


def build_payload() -> Dict[str, object]:
    with zipfile.ZipFile(WORKBOOK_PATH) as archive:
        shared_strings = read_shared_strings(archive)
        sheets = load_workbook(archive, shared_strings)

    item_links: "OrderedDict[str, List[Dict[str, str]]]" = OrderedDict()
    unique_entries = 0

    for row in iter_detailed_rows(sheets):
        item_key = make_item_key(row["programId"], row["year"], row["surveyTitle"], row["phraseText"])
        entry = {
            "standard": row["standard"],
            "section": row["section"],
            "criterionCode": row["criterionCode"],
            "criterionText": row["criterionText"],
            "supportedSide": row["supportedSide"],
        }

        links = item_links.setdefault(item_key, [])
        if entry not in links:
            links.append(entry)
            unique_entries += 1

    for links in item_links.values():
        links.sort(
            key=lambda item: (
                item["criterionCode"],
                item["supportedSide"],
                item["criterionText"],
                item["standard"],
                item["section"],
            )
        )

    return {
        "sourceLabel": "استطلاعات الدراسة الذاتية",
        "sourceFile": WORKBOOK_PATH.name,
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "linkedItemCount": len(item_links),
        "linkEntryCount": unique_entries,
        "itemLinks": item_links,
    }


def main() -> None:
    payload = build_payload()
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(
        "window.SELF_STUDY_DATA = " + json.dumps(payload, ensure_ascii=False, indent=2) + ";\n",
        encoding="utf-8",
    )
    print(f"Generated {OUTPUT_PATH}")
    print(f"Linked items: {payload['linkedItemCount']}")
    print(f"Link entries: {payload['linkEntryCount']}")


if __name__ == "__main__":
    main()
