#!/usr/bin/env python3

from __future__ import annotations

import hashlib
import json
import re
import zipfile
from collections import OrderedDict
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Tuple
from xml.etree import ElementTree as ET


NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_PATH = ROOT / "استطلاعات_وتقييمات_كلية_الشريعة_1445_1446.xlsx"
OUTPUT_PATH = ROOT / "js" / "surveys-data.js"

PROGRAM_ID_MAP = {
    ("الأنظمة", "البكالوريوس"): "p01",
    ("الدراسات الإسلامية", "البكالوريوس"): "p02",
    ("الشريعة", "البكالوريوس"): "p03",
    ("القرآن وعلومه", "البكالوريوس"): "p04",
    ("القراءات", "البكالوريوس"): "p05",
    ("القانون", "الماجستير"): "p06",
    ("العقيدة", "الماجستير"): "p07",
    ("أصول الفقه", "الماجستير"): "p08",
    ("الفقه", "الماجستير"): "p09",
    ("الدراسات القرآنية المعاصرة", "الماجستير"): "p10",
    ("القراءات", "الماجستير"): "p11",
    ("أصول الفقه", "الدكتوراه"): "p12",
    ("الفقه", "الدكتوراه"): "p13",
    ("الدراسات القرآنية", "الدكتوراه"): "p14",
    ("القراءات", "الدكتوراه"): "p15",
}

SURVEY_SECTION_MAP = {
    "تقييم الطلبة للمقرر والمحاضر": "learning",
    "استبانة تقويم المقرر": "learning",
    "آراء الطلبة في التعلم الالكتروني": "learning",
    "أراء الطلبة في التعلم الالكتروني": "learning",
    "استبانة التعليم الالكتروني": "learning",
    "استبانة تقويم برنامج": "management",
    "رضا الطلبة عن الجامعة وخدماتها - المكتبة": "students",
    "رضا الطلبة عن الجامعة وخدماتها - الفصول الدراسية والمرافق": "students",
    "رضا الطلبة عن الجامعة وخدماتها - القبول والتسجيل والإرشاد الجامعي": "students",
    "رضا الطلبة عن الجامعة وخدماتها - شؤون الطلبة وهيئة التدريس بالجامعة": "students",
    "رضا الطلبة عن الجامعة وخدماتها - الموارد المؤسسية بالجامعة والشراكة المجتمعية": "management",
    "رضا الطلبة عن الجامعة وخدماتها - الرسالة والحوكمة": "management",
    "استبانة تقويم خبرة الطالب/ الطالبة": "students",
    "خدمات عمادة القبول والتسجيل": "students",
    "استبيان لقياس رضا المستفيدين عن أداء الأقسام الأكاديمية في الجامعة": "management",
    "استبيان لقياس رضا المستفيدين عن أداء الإجراءات بمختلف إدارات الجامعة": "management",
    "استطلاع رأي أصحاب المصلحة في رؤية ورسالة وقيم واهداف جامعة الطائف الاستراتيجية": "management",
    "رضا الطلبة ذوي الاحتياجات الخاصة عن الجامعة وخدماتها": "students",
    "رضا الطلبة عن لجنة حماية حقوق الطلبة": "students",
    "استبانة تقويم الطلبة للخبرة الميدانية": "market",
}


def clean_text(value: object) -> str:
    text = str(value or "")
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s*-\s*", " - ", text)
    return text.strip()


def normalize_gender(value: str) -> str:
    value = clean_text(value)
    mapping = {
        "ذكر": "ذكر",
        "ذكور": "ذكر",
        "male": "ذكر",
        "m": "ذكر",
        "أنثى": "إناث",
        "إناث": "إناث",
        "اناث": "إناث",
        "female": "إناث",
        "f": "إناث",
    }
    return mapping.get(value, value)


def read_shared_strings(archive: zipfile.ZipFile) -> List[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: List[str] = []
    for si in root.findall("a:si", NS):
        values.append("".join(node.text or "" for node in si.iterfind(".//a:t", NS)))
    return values


def iter_sheet_rows(archive: zipfile.ZipFile, shared_strings: List[str]) -> Iterable[Dict[str, str]]:
    root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    rows = root.find("a:sheetData", NS).findall("a:row", NS)

    for row in rows[1:]:
        values: Dict[str, str] = {}
        for cell in row.findall("a:c", NS):
            ref = cell.get("r", "")
            column = "".join(ch for ch in ref if ch.isalpha())
            node = cell.find("a:v", NS)
            if node is None:
                continue
            raw = node.text or ""
            if cell.get("t") == "s":
                raw = shared_strings[int(raw)]
            values[column] = clean_text(raw)
        if values:
            yield values


def year_from_semester_code(code: str) -> str:
    code = clean_text(code)
    if len(code) >= 2 and code[:2].isdigit():
        return f"14{code[:2]}"
    return code


def numeric(value: str) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def make_id(*parts: str) -> str:
    raw = "||".join(clean_text(part) for part in parts if clean_text(part))
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]


def section_for_survey(title: str) -> str:
    title = clean_text(title)
    if title in SURVEY_SECTION_MAP:
        return SURVEY_SECTION_MAP[title]
    if "المقرر" in title or "التعلم الالكتروني" in title or "التعليم الالكتروني" in title:
        return "learning"
    if "المكتبة" in title or "القبول والتسجيل" in title or "الطلبة" in title:
        return "students"
    if "التدريب الميداني" in title:
        return "market"
    return "management"


def item_sort_key(item: Dict[str, object]) -> Tuple[int, str]:
    number_text = clean_text(item.get("number"))
    digits = re.sub(r"[^\d]", "", number_text)
    if digits.isdigit():
        return (int(digits), clean_text(item.get("label")))
    return (999999, clean_text(item.get("label")))


def build_payload() -> Dict[str, object]:
    datasets: "OrderedDict[str, Dict[str, object]]" = OrderedDict()
    available_program_years: Dict[str, set] = {}
    available_genders = set()
    skipped_programs = set()

    with zipfile.ZipFile(WORKBOOK_PATH) as archive:
        shared_strings = read_shared_strings(archive)

        for row in iter_sheet_rows(archive, shared_strings):
            program_name = clean_text(row.get("F"))
            degree = clean_text(row.get("C"))
            program_id = PROGRAM_ID_MAP.get((program_name, degree))
            if not program_id:
                skipped_programs.add((program_name, degree))
                continue

            year = year_from_semester_code(row.get("A", ""))
            survey_title = clean_text(row.get("G"))
            topic_label = clean_text(row.get("H")) or survey_title
            item_number = clean_text(row.get("I"))
            item_label = clean_text(row.get("J")) or topic_label
            gender = normalize_gender(row.get("B"))

            counts = [numeric(row.get(column, "0")) for column in ("K", "L", "M", "N", "O")]
            response_sum = sum(counts)
            score_total = sum((index + 1) * count for index, count in enumerate(counts))
            if not response_sum:
                continue

            dataset_key = f"{program_id}::{year}"
            dataset = datasets.setdefault(
                dataset_key,
                {
                    "mode": "survey-items",
                    "source": (
                        f"استيراد تلقائي من ملف {WORKBOOK_PATH.name} "
                        "المستخرج من منصة ذكاء الأعمال. جميع العناصر الحالية من استطلاعات الطلاب."
                    ),
                    "_surveys": OrderedDict(),
                },
            )

            survey = dataset["_surveys"].setdefault(
                survey_title,
                {
                    "id": make_id(program_id, year, survey_title),
                    "title": survey_title,
                    "stakeholder": "students",
                    "sectionId": section_for_survey(survey_title),
                    "_topics": OrderedDict(),
                },
            )

            topic = survey["_topics"].setdefault(
                topic_label,
                {
                    "label": topic_label,
                    "_items": OrderedDict(),
                },
            )

            item_key = (item_number, item_label)
            item = topic["_items"].setdefault(
                item_key,
                {
                    "number": item_number,
                    "label": item_label,
                    "_genders": OrderedDict(),
                },
            )

            gender_entry = item["_genders"].setdefault(
                gender,
                {
                    "gender": gender,
                    "responses": 0,
                    "scoreTotal": 0,
                },
            )
            gender_entry["responses"] += response_sum
            gender_entry["scoreTotal"] += score_total

            available_program_years.setdefault(program_id, set()).add(year)
            if gender:
                available_genders.add(gender)

    extracted_data: "OrderedDict[str, Dict[str, object]]" = OrderedDict()
    survey_count = 0
    item_record_count = 0

    for dataset_key, dataset in datasets.items():
        surveys = []
        for survey in dataset["_surveys"].values():
            topics = []
            for topic in survey["_topics"].values():
                items = []
                for item in topic["_items"].values():
                    genders = list(item["_genders"].values())
                    item_record_count += len(genders)
                    items.append(
                        {
                            "number": item["number"],
                            "label": item["label"],
                            "genders": genders,
                        }
                    )
                items.sort(key=item_sort_key)
                topics.append(
                    {
                        "label": topic["label"],
                        "items": items,
                    }
                )

            surveys.append(
                {
                    "id": survey["id"],
                    "title": survey["title"],
                    "stakeholder": survey["stakeholder"],
                    "sectionId": survey["sectionId"],
                    "topics": topics,
                }
            )

        extracted_data[dataset_key] = {
            "mode": dataset["mode"],
            "source": dataset["source"],
            "surveys": surveys,
        }
        survey_count += len(surveys)

    payload = {
        "sourceLabel": "منصة ذكاء الأعمال",
        "sourceFile": WORKBOOK_PATH.name,
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "datasetCount": len(extracted_data),
        "surveyCount": survey_count,
        "itemRecordCount": item_record_count,
        "availableGenders": sorted(available_genders),
        "availableProgramYears": {
            program_id: sorted(years, reverse=True)
            for program_id, years in sorted(available_program_years.items())
        },
        "skippedPrograms": [
            {"program": program, "degree": degree}
            for program, degree in sorted(skipped_programs)
        ],
        "extractedData": extracted_data,
    }
    return payload


def main() -> None:
    payload = build_payload()
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(
        "window.SURVEYS_DATA = " + json.dumps(payload, ensure_ascii=False, indent=2) + ";\n",
        encoding="utf-8",
    )
    print(f"Generated {OUTPUT_PATH}")
    print(f"Datasets: {payload['datasetCount']}")
    print(f"Surveys: {payload['surveyCount']}")
    print(f"Item records: {payload['itemRecordCount']}")
    if payload["skippedPrograms"]:
        print("Skipped programs:")
        for item in payload["skippedPrograms"]:
            print(f" - {item['program']} | {item['degree']}")


if __name__ == "__main__":
    main()
