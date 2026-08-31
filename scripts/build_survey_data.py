#!/usr/bin/env python3

from __future__ import annotations

import hashlib
import json
import re
import zipfile
from collections import OrderedDict, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple
from xml.etree import ElementTree as ET


NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
ROOT = Path(__file__).resolve().parents[1]
LEGACY_WORKBOOK_PATH = ROOT / "استطلاعات_وتقييمات_كلية_الشريعة_1445_1446.xlsx"
SUPPLEMENTAL_DATA_DIR = ROOT / "data" / "bi-1447"
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

PROGRAM_DEFINITIONS = {
    "p01": {"name": "الأنظمة", "degree": "البكالوريوس", "dept": "الأنظمة"},
    "p02": {"name": "الدراسات الإسلامية", "degree": "البكالوريوس", "dept": "الدراسات الإسلامية"},
    "p03": {"name": "الشريعة", "degree": "البكالوريوس", "dept": "الشريعة"},
    "p04": {"name": "القرآن وعلومه", "degree": "البكالوريوس", "dept": "القراءات"},
    "p05": {"name": "القراءات", "degree": "البكالوريوس", "dept": "القراءات"},
    "p06": {"name": "القانون", "degree": "الماجستير", "dept": "الأنظمة"},
    "p07": {"name": "العقيدة", "degree": "الماجستير", "dept": "الدراسات الإسلامية"},
    "p08": {"name": "أصول الفقه", "degree": "الماجستير", "dept": "الشريعة"},
    "p09": {"name": "الفقه", "degree": "الماجستير", "dept": "الشريعة"},
    "p10": {"name": "الدراسات القرآنية المعاصرة", "degree": "الماجستير", "dept": "القراءات"},
    "p11": {"name": "القراءات", "degree": "الماجستير", "dept": "القراءات"},
    "p12": {"name": "أصول الفقه", "degree": "الدكتوراه", "dept": "الشريعة"},
    "p13": {"name": "الفقه", "degree": "الدكتوراه", "dept": "الشريعة"},
    "p14": {"name": "الدراسات القرآنية", "degree": "الدكتوراه", "dept": "القراءات"},
    "p15": {"name": "القراءات", "degree": "الدكتوراه", "dept": "القراءات"},
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

STUDENT_1447_SOURCES = [
    {"path": SUPPLEMENTAL_DATA_DIR / "تقييمات الطلاب 471.xlsx", "semester": "471"},
    {"path": SUPPLEMENTAL_DATA_DIR / "تقييمات الطلاب 472.xlsx", "semester": "472"},
]

FACULTY_1447_SOURCES = [
    {"path": SUPPLEMENTAL_DATA_DIR / "استبيانات الكادر - أنظمة.xlsx", "department": "الأنظمة"},
    {"path": SUPPLEMENTAL_DATA_DIR / "استبيانات الكادر - قراءات.xlsx", "department": "القراءات"},
    {"path": SUPPLEMENTAL_DATA_DIR / "الكادر - ثقافة.xlsx", "department": "الدراسات الإسلامية"},
    {"path": SUPPLEMENTAL_DATA_DIR / "الكادر - شريعة.xlsx", "department": "الشريعة"},
]

DEPARTMENT_PROGRAMS = {
    "الأنظمة": ["p01", "p06"],
    "الدراسات الإسلامية": ["p02", "p07"],
    "الشريعة": ["p03", "p08", "p09", "p12", "p13"],
    "القراءات": ["p04", "p05", "p10", "p11", "p14", "p15"],
}

SKIPPED_STUDENT_PROGRAM_LABELS = {
    "السنة الأولى للشريعة والأنظمة",
    "الشريعة والدراسات الإسلامية",
    "الشريعة والدراسات الاسلامية",
    "-",
}

STUDENT_PROGRAM_ALIASES = {
    "الثقافة الإسلامية": ["p02"],
    "الدراسات الاسلامية - شعبة القرآن الكريم وعلومه": ["p04"],
}

E_LEARNING_PREFERRED_PHRASES = {
    "أتاحت لي الأنشطة التعليمية وأدوات المقرر فرصة للتفاعل مع المحتوى",
    "ارتبطت أهداف الوحدات الدراسية بالأهداف العامة للمقرر",
    "بشكل عام أنا راضي عن هذا المقرر",
    "تتوفر أكثر من وسيلة للتواصل مع عضو هيئة التدريس",
    "تلقيت تغذية راجعة عن أدائي في بعض المهام من قبل عضو هيئة التدريس",
    "وضح عضو هيئة التدريس آلية تصحيح الاختبارات والأنشطة والتكليفات",
    "تم عرض خطة المقرر بالتفصيل في بداية الفصل الدراسي",
    "تم عرض الدروس بأشكال وأساليب مختلفة بالتعلم الالكتروني",
    "اشعر بالرضا عن شرح عضو هيئة التدريس في المقرر من خلال التعلم الالكتروني",
    "شعر بالرضا عن شرح عضو هيئة التدريس في المقرر من خلال التعلم الالكتروني",
}

LEGACY_SOURCE_NOTE = (
    f"استيراد مباشر من ملف {LEGACY_WORKBOOK_PATH.name} المستخرج من منصة ذكاء الأعمال. "
    "يتضمن بيانات الطلاب مع عدد المقيمين."
)

STUDENT_1447_SOURCE_NOTE = (
    "استيراد مفسر من ملفات استطلاعات الطلاب 1447هـ بمنصة ذكاء الأعمال. "
    "هذا المصدر لا يتضمن الدرجة العلمية ولا عدد المقيمين؛ لذا حُفظ المتوسط كما هو، "
    "ووُزعت السجلات ذات اسم البرنامج المشترك على البرامج الرسمية المطابقة للاسم نفسه."
)

FACULTY_1447_SOURCE_NOTE = (
    "استيراد مفسر من ملفات استطلاعات الكادر الأكاديمي 1447هـ بمنصة ذكاء الأعمال. "
    "هذا المصدر قسميّ النطاق ولا يتضمن اسم البرنامج ولا عدد المقيمين؛ "
    "لذا أُلحق بكل برامج القسم مع الحفاظ على الجهة = أعضاء هيئة التدريس."
)

NAME_ONLY_PROGRAMS: Dict[str, List[str]] = defaultdict(list)
for (program_name, _degree), program_id in PROGRAM_ID_MAP.items():
    NAME_ONLY_PROGRAMS[program_name].append(program_id)


def clean_text(value: object) -> str:
    text = str(value or "")
    text = text.replace("\u00a0", " ").replace("\u200f", "").replace("\u200e", "")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s*-\s*", " - ", text)
    return text.strip()


def normalize_match_text(value: object) -> str:
    text = clean_text(value)
    replacements = {
        "؟": "",
        "?": "",
        "٫": ".",
        "،": "،",
        "أراء": "آراء",
        "الاكاديمي": "الأكاديمي",
        "الارشاد": "الإرشاد",
        "ادارة": "إدارة",
        "اهداف": "أهداف",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" .؛:،-")


def has_outer_question_wrap(value: object) -> bool:
    text = clean_text(value)
    return bool(text) and (
        text.startswith("?")
        or text.endswith("?")
        or text.startswith("؟")
        or text.endswith("؟")
    )


def sanitize_display_phrase(value: object) -> str:
    text = clean_text(value)
    while text.startswith("?") or text.startswith("؟"):
        text = text[1:].lstrip()
    while text.endswith("?") or text.endswith("؟"):
        text = text[:-1].rstrip()
    return text


def normalize_gender(value: str) -> str:
    value = normalize_match_text(value)
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
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: List[str] = []
    for si in root.findall("a:si", NS):
        values.append("".join(node.text or "" for node in si.iterfind(".//a:t", NS)))
    return values


def iter_sheet_rows(archive: zipfile.ZipFile, worksheet_path: str, shared_strings: List[str]) -> Iterable[Dict[str, str]]:
    root = ET.fromstring(archive.read(worksheet_path))
    rows = root.find("a:sheetData", NS).findall("a:row", NS)

    for row in rows:
        values: Dict[str, str] = {}
        for cell in row.findall("a:c", NS):
            ref = cell.get("r", "")
            column = "".join(ch for ch in ref if ch.isalpha())
            if cell.get("t") == "inlineStr":
                raw = "".join(node.text or "" for node in cell.iterfind(".//a:t", NS))
            else:
                node = cell.find("a:v", NS)
                if node is None:
                    continue
                raw = node.text or ""
                if cell.get("t") == "s":
                    raw = shared_strings[int(raw)]
            values[column] = clean_text(raw)
        if values:
            yield values


def read_first_sheet_rows(path: Path) -> List[Dict[str, str]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = read_shared_strings(archive)
        return list(iter_sheet_rows(archive, "xl/worksheets/sheet1.xml", shared_strings))


def year_from_semester_code(code: str) -> str:
    code = clean_text(code)
    if len(code) >= 2 and code[:2].isdigit():
        return f"14{code[:2]}"
    return code


def numeric(value: str) -> int:
    try:
        return int(float(clean_text(value).replace("٫", ".")))
    except (TypeError, ValueError):
        return 0


def decimal_number(value: str) -> float | None:
    text = clean_text(value).replace("٫", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def make_id(*parts: str) -> str:
    raw = "||".join(clean_text(part) for part in parts if clean_text(part))
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]


def round_score(value: float | None) -> float | None:
    if value is None:
        return None
    return round(float(value), 4)


def section_for_survey(title: str) -> str:
    title = clean_text(title)
    if title in SURVEY_SECTION_MAP:
        return SURVEY_SECTION_MAP[title]
    if "المقرر" in title or "التعلم الالكتروني" in title or "التعليم الالكتروني" in title:
        return "learning"
    if "المكتبة" in title or "القبول والتسجيل" in title or "الطلبة" in title:
        return "students"
    if "التدريب الميداني" in title or "ريادة الأعمال" in title:
        return "market"
    if "هيئة التدريس" in title or "الموارد البشرية" in title:
        return "faculty"
    return "management"


def item_sort_key(item: Dict[str, object]) -> Tuple[int, str]:
    number_text = clean_text(item.get("number"))
    digits = re.sub(r"[^\d]", "", number_text)
    if digits.isdigit():
        return (int(digits), clean_text(item.get("label")))
    return (999999, clean_text(item.get("label")))


def split_numbered_phrase(phrase_text: str) -> Tuple[str, str]:
    text = clean_text(phrase_text)
    match = re.match(r"^(\d+)\s*[-–]\s*(.+)$", text)
    if match:
        return match.group(1), match.group(2).strip()
    return "", text


def parse_numbered_phrase(phrase_text: str) -> Tuple[str, str]:
    item_number, item_label = split_numbered_phrase(phrase_text)
    return item_number, sanitize_display_phrase(item_label)


def sorted_program_ids(program_ids: Sequence[str]) -> List[str]:
    return sorted(set(program_ids), key=lambda value: int(value.replace("p", "")))


def resolve_student_program_ids(program_label: str) -> List[str]:
    label = clean_text(program_label)
    if not label or label in SKIPPED_STUDENT_PROGRAM_LABELS:
        return []
    if label in STUDENT_PROGRAM_ALIASES:
        return sorted_program_ids(STUDENT_PROGRAM_ALIASES[label])
    if label in NAME_ONLY_PROGRAMS:
        return sorted_program_ids(NAME_ONLY_PROGRAMS[label])
    return []


def build_legacy_phrase_catalog(rows: Sequence[Dict[str, str]]) -> Dict[str, List[Dict[str, str]]]:
    catalog: Dict[str, List[Dict[str, str]]] = defaultdict(list)
    for row in rows:
        survey_title = clean_text(row.get("G"))
        topic_label = clean_text(row.get("H")) or survey_title
        item_number = clean_text(row.get("I"))
        phrase_text = clean_text(row.get("J")) or topic_label
        if not phrase_text:
            continue
        entry = {
            "surveyTitle": survey_title,
            "topicLabel": topic_label,
            "sectionId": section_for_survey(survey_title),
            "itemNumber": item_number,
            "itemLabel": phrase_text,
        }
        key = normalize_match_text(phrase_text)
        if entry not in catalog[key]:
            catalog[key].append(entry)
    return catalog


def choose_legacy_match(phrase_text: str, matches: Sequence[Dict[str, str]]) -> Dict[str, str] | None:
    if not matches:
        return None

    unique_matches: List[Dict[str, str]] = []
    for match in matches:
        if match not in unique_matches:
            unique_matches.append(match)

    if len(unique_matches) == 1:
        return unique_matches[0]

    normalized_phrase = normalize_match_text(phrase_text)
    wrapped_phrase = has_outer_question_wrap(phrase_text)
    e_learning_matches = [
        match for match in unique_matches
        if normalize_match_text(match["surveyTitle"]) in {
            "آراء الطلبة في التعلم الالكتروني",
            "أراء الطلبة في التعلم الالكتروني",
        }
    ]

    if wrapped_phrase and e_learning_matches and (
        "التعلم الالكتروني" in normalized_phrase
        or normalized_phrase in {normalize_match_text(item) for item in E_LEARNING_PREFERRED_PHRASES}
    ):
        return e_learning_matches[0]

    exact_course_matches = [
        match for match in unique_matches
        if normalize_match_text(match["surveyTitle"]) == normalize_match_text("تقييم الطلبة للمقرر والمحاضر")
    ]

    if wrapped_phrase and e_learning_matches:
        return e_learning_matches[0]

    if exact_course_matches and (
        normalized_phrase.startswith("شعر بالرضا عن شرح عضو هيئة التدريس")
        or "التعلم الالكتروني" in normalized_phrase
        or normalized_phrase in {normalize_match_text(item) for item in E_LEARNING_PREFERRED_PHRASES}
    ):
        return exact_course_matches[0]

    if e_learning_matches:
        return e_learning_matches[0]

    return unique_matches[0]


def topic_for_elearning_phrase(phrase_text: str) -> str:
    normalized_phrase = normalize_match_text(phrase_text)
    rules = [
        (
            (
                "خطة المقرر",
                "الخدمات الأكاديمية",
                "الخدمات الاكاديمية",
                "آلية تصحيح",
                "وضح عضو هيئة التدريس",
                "تم عرض الدروس",
            ),
            "أولاً: تنظيم المقرر وتخطيطه:",
        ),
        (
            (
                "المشاركة أثناء المحاضرة",
                "التواصل مع عضو هيئة التدريس",
                "وسيلة للتواصل مع عضو هيئة التدريس",
                "تغذية راجعة عن أدائي",
                "مناقشة زملائي",
                "الدعم الفني",
                "أشعر بالمتعة",
            ),
            "ثانياً التواصل والتفاعل:",
        ),
        (
            (
                "الأنشطة التعليمية",
                "مهارتي في استخدام التكنولوجيا",
                "التدريب والإرشاد",
                "تم تنفيذ وعرض المقرر",
                "متابعة ما تغيبت",
                "تنوعت الوسائط",
                "التعلم المستمر",
            ),
            "ثالثاً: عناصر تصميم المقرر ووسائطه المتعددة:",
        ),
        (
            (
                "أهداف الوحدات الدراسية",
                "التكليفات",
                "تغذية راجعة فورية بعد الاختبار",
                "يوفر التعلم الالكتروني وقت الطالب",
            ),
            "رابعاً: المواد التعليمية والتقييمات:",
        ),
        (
            (
                "بشكل عام أنا راضي",
                "شعر بالرضا عن شرح عضو هيئة التدريس",
                "اشعر بالرضا عن شرح عضو هيئة التدريس",
            ),
            "خامساً: الرضا العام عن المقرر وعضو هيئة التدريس:",
        ),
    ]

    for keywords, topic in rules:
        if any(keyword in normalized_phrase for keyword in keywords):
            return topic

    return "ثالثاً: عناصر تصميم المقرر ووسائطه المتعددة:"


def classify_student_phrase(phrase_text: str, legacy_catalog: Dict[str, List[Dict[str, str]]]) -> Dict[str, str]:
    item_number, item_label = parse_numbered_phrase(phrase_text)
    _raw_item_number, raw_item_label = split_numbered_phrase(phrase_text)
    normalized_label = normalize_match_text(item_label)

    direct_matches = legacy_catalog.get(normalized_label, [])
    if direct_matches:
        match = choose_legacy_match(raw_item_label, direct_matches)
        if match:
            return {
                "surveyTitle": match["surveyTitle"],
                "topicLabel": match["topicLabel"],
                "sectionId": match["sectionId"],
                "itemNumber": item_number or match["itemNumber"],
                "itemLabel": match["itemLabel"],
            }

    if "التعلم الالكتروني" in normalized_label or normalized_label in {
        normalize_match_text("شعر بالرضا عن شرح عضو هيئة التدريس في المقرر من خلال التعلم الالكتروني"),
        normalize_match_text("اشعر بالرضا عن شرح عضو هيئة التدريس في المقرر من خلال التعلم الالكتروني"),
    }:
        return {
            "surveyTitle": "أراء الطلبة في التعلم الالكتروني",
            "topicLabel": topic_for_elearning_phrase(item_label),
            "sectionId": "learning",
            "itemNumber": item_number,
            "itemLabel": item_label,
        }

    if (
        "ريادة الأعمال" in normalized_label
        or "ريادية" in normalized_label
        or "رواد الأعمال" in normalized_label
        or "مستقبلي الوظيفي" in normalized_label
    ):
        return {
            "surveyTitle": "استطلاع ريادة الأعمال والمهارات المستقبلية",
            "topicLabel": "ريادة الأعمال والمهارات المستقبلية",
            "sectionId": "market",
            "itemNumber": item_number,
            "itemLabel": item_label,
        }

    if (
        "ذوي الإعاقة" in normalized_label
        or "ذوي الاعاقة" in normalized_label
        or "إعاقتي" in normalized_label
        or "طلب بدل إعاقة" in normalized_label
        or "الطلبة ذوي الإعاقة" in normalized_label
    ):
        return {
            "surveyTitle": "رضا الطلبة ذوي الاحتياجات الخاصة عن الجامعة وخدماتها",
            "topicLabel": "ذوو الاحتياجات الخاصة",
            "sectionId": "students",
            "itemNumber": item_number,
            "itemLabel": item_label,
        }

    if "لجنة حماية حقوق الطلبة" in normalized_label or "منصة حماية حقوق الطلبة" in normalized_label:
        return {
            "surveyTitle": "رضا الطلبة عن لجنة حماية حقوق الطلبة",
            "topicLabel": "لجنة حماية حقوق الطلبة",
            "sectionId": "students",
            "itemNumber": item_number,
            "itemLabel": item_label,
        }

    if normalized_label.startswith("ما مدى رضاك عن") or "الخدمات الإدارية" in normalized_label or "التعاملات الإدارية" in normalized_label:
        return {
            "surveyTitle": "استبيان لقياس رضا المستفيدين عن أداء الإجراءات بمختلف إدارات الجامعة",
            "topicLabel": "الإجراءات الإدارية",
            "sectionId": "management",
            "itemNumber": item_number,
            "itemLabel": item_label,
        }

    if (
        "أعضاء هيئة التدريس" in normalized_label
        or "الكادر الإداري" in normalized_label
        or "الفنيين" in normalized_label
        or "تعيين أعضاء هيئة التدريس" in normalized_label
        or "الموارد البشرية" in normalized_label
    ):
        return {
            "surveyTitle": "استطلاع الموارد البشرية والدعم الأكاديمي",
            "topicLabel": "هيئة التدريس والموارد البشرية",
            "sectionId": "faculty",
            "itemNumber": item_number,
            "itemLabel": item_label,
        }

    if (
        "طرق التدريس" in normalized_label
        or "معايير التقييم" in normalized_label
        or "محتوى المقررات" in normalized_label
        or "المقررات الدراسية" in normalized_label
        or "التغذية الراجعة" in normalized_label
        or "مصادر تعلم حديثة" in normalized_label
        or "تنوع أساليب التعلم" in normalized_label
        or "التفاعل والمشاركة" in normalized_label
        or "دراستي" in normalized_label
        or "أتعلم كيف أعمل" in normalized_label
    ):
        return {
            "surveyTitle": "استطلاع الخبرة التعليمية في البرنامج",
            "topicLabel": "الخبرة التعليمية والمقررات",
            "sectionId": "learning",
            "itemNumber": item_number,
            "itemLabel": item_label,
        }

    return {
        "surveyTitle": "استطلاع الخبرة التعليمية في البرنامج",
        "topicLabel": "بنود متنوعة",
        "sectionId": "learning",
        "itemNumber": item_number,
        "itemLabel": item_label,
    }


def classify_faculty_phrase(phrase_text: str, department_label: str) -> Dict[str, str]:
    normalized_phrase = normalize_match_text(phrase_text)

    def result(title_suffix: str, topic_label: str, section_id: str) -> Dict[str, str]:
        return {
            "surveyTitle": f"استطلاع الكادر الأكاديمي - {department_label} - {title_suffix}",
            "topicLabel": topic_label,
            "sectionId": section_id,
            "itemNumber": "",
            "itemLabel": clean_text(phrase_text),
        }

    if any(keyword in normalized_phrase for keyword in ("البلاك بورد", "التعلم الإلكتروني", "التعلم الالكتروني", "منصة", "الخدمات الإلكترونية", "الخدمات الالكترونية", "النظام", "دعم فني", "المحتوى التدريبي", "مهارات")):
        return result("التعليم الإلكتروني", "التعليم الإلكتروني والخدمات التقنية", "learning")

    if any(keyword in normalized_phrase for keyword in ("المكتبة", "المكتبة الرقمية", "الاستعارة", "قواعد المعلومات", "الدوريات", "المجلات", "الكتب", "المراجع", "الحاسب", "التصوير")):
        return result("الموارد التعليمية", "المكتبة ومصادر التعلم", "learning")

    if any(keyword in normalized_phrase for keyword in ("الخدمات الإدارية", "الخدمات الادارية", "الإدارة", "الادارة", "الموظفين", "الشفافية", "التعاملات الإدارية", "التعاملات الادارية")):
        return result("الخدمات الإدارية", "الإجراءات والخدمات الإدارية", "management")

    if any(keyword in normalized_phrase for keyword in ("العبء التدريسي", "الأعباء التدريسية", "هيئة التدريس", "التنمية المهنية", "إعادة توزيع الأعباء", "ذوي المؤهلات", "ذوي المؤهلات والخبرات", "توزيع المهام", "توزيع العبء", "النصاب")):
        return result("هيئة التدريس", "هيئة التدريس والموارد البشرية", "faculty")

    if any(keyword in normalized_phrase for keyword in ("مخرجات تعلم الطلاب", "المستفيدين من خارج الجامعة", "البرامج الأكاديمية", "البرامج الاكاديمية", "المقررات الدراسية", "تقويم وتعديل البرامج", "ضمان جودة البرامج", "توقعات المستفيدين", "سوق العمل", "مراجع علمية حديثة")):
        if "المستفيدين من خارج الجامعة" in normalized_phrase or "سوق العمل" in normalized_phrase:
            return result("جودة البرنامج وسوق العمل", "جودة البرنامج وسوق العمل", "market")
        return result("جودة البرنامج", "جودة البرنامج والمقررات", "learning")

    if any(keyword in normalized_phrase for keyword in ("القيادات", "الرؤية", "الرسالة", "الهيكل التنظيمي", "النزاهة", "العدالة", "المساواة", "الشكاوي", "الشكاوى", "التظلم", "اتخاذ القرار", "إتخاذ القرار", "أخلاقيات العمل", "السلوك الوظيفي", "الأمان العلمية", "المليكية الفكرية", "المناخ التنظيمي", "اللوائح", "الأنظمة الجامعية", "الجامعة")):
        return result("الحوكمة والقيادة", "الحوكمة والقيادة والإدارة", "management")

    return result("بنود متنوعة", "بنود متنوعة", "management")


def create_dataset_store() -> Tuple["OrderedDict[str, Dict[str, object]]", Dict[str, set], set, set]:
    return OrderedDict(), defaultdict(set), set(), set()


def ensure_dataset(datasets: "OrderedDict[str, Dict[str, object]]", dataset_key: str) -> Dict[str, object]:
    dataset = datasets.setdefault(
        dataset_key,
        {
            "mode": "survey-items",
            "_sources": set(),
            "_surveys": OrderedDict(),
        },
    )
    return dataset


def add_item_measurement(
    datasets: "OrderedDict[str, Dict[str, object]]",
    available_program_years: Dict[str, set],
    available_genders: set,
    *,
    program_ids: Sequence[str],
    year: str,
    survey_title: str,
    stakeholder: str,
    section_id: str,
    topic_label: str,
    item_number: str,
    item_label: str,
    gender: str,
    source_note: str,
    responses: int | None = None,
    score_total: int | None = None,
    average: float | None = None,
) -> None:
    for program_id in sorted_program_ids(program_ids):
        dataset_key = f"{program_id}::{year}"
        dataset = ensure_dataset(datasets, dataset_key)
        dataset["_sources"].add(source_note)

        survey_key = (stakeholder, section_id, survey_title)
        survey = dataset["_surveys"].setdefault(
            survey_key,
            {
                "id": make_id(program_id, year, stakeholder, section_id, survey_title),
                "title": survey_title,
                "stakeholder": stakeholder,
                "sectionId": section_id,
                "_topics": OrderedDict(),
            },
        )

        topic = survey["_topics"].setdefault(
            topic_label or survey_title,
            {
                "label": topic_label or survey_title,
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

        gender_key = gender or ""
        gender_entry = item["_genders"].setdefault(
            gender_key,
            {
                "gender": gender,
                "responses": None,
                "scoreTotal": None,
                "_averageTotal": 0.0,
                "_averageSamples": 0,
            },
        )

        if responses is not None and score_total is not None and responses > 0:
            gender_entry["responses"] = (gender_entry["responses"] or 0) + responses
            gender_entry["scoreTotal"] = (gender_entry["scoreTotal"] or 0) + score_total
        elif average is not None:
            gender_entry["_averageTotal"] += float(average)
            gender_entry["_averageSamples"] += 1

        available_program_years[program_id].add(year)
        if gender:
            available_genders.add(gender)


def import_legacy_student_surveys(
    legacy_rows: Sequence[Dict[str, str]],
    datasets: "OrderedDict[str, Dict[str, object]]",
    available_program_years: Dict[str, set],
    available_genders: set,
    skipped_programs: set,
) -> None:
    for row in legacy_rows:
        program_name = clean_text(row.get("F"))
        degree = clean_text(row.get("C"))
        program_id = PROGRAM_ID_MAP.get((program_name, degree))
        if not program_id:
            skipped_programs.add((program_name, degree, "legacy"))
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

        add_item_measurement(
            datasets,
            available_program_years,
            available_genders,
            program_ids=[program_id],
            year=year,
            survey_title=survey_title,
            stakeholder="students",
            section_id=section_for_survey(survey_title),
            topic_label=topic_label,
            item_number=item_number,
            item_label=item_label,
            gender=gender,
            source_note=LEGACY_SOURCE_NOTE,
            responses=response_sum,
            score_total=score_total,
        )


def import_student_1447_surveys(
    datasets: "OrderedDict[str, Dict[str, object]]",
    available_program_years: Dict[str, set],
    available_genders: set,
    skipped_programs: set,
    legacy_catalog: Dict[str, List[Dict[str, str]]],
) -> None:
    for source in STUDENT_1447_SOURCES:
        path = source["path"]
        if not path.exists():
            continue

        rows = read_first_sheet_rows(path)
        if not rows:
            continue

        header = {label: column for column, label in rows[0].items()}
        program_col = header.get("التخصص")
        gender_col = header.get("الجنس")
        average_col = header.get("المتوسط")
        phrase_col = header.get("الاسئلة")
        if not all([program_col, gender_col, average_col, phrase_col]):
            continue

        for row in rows[1:]:
            program_label = clean_text(row.get(program_col))
            program_ids = resolve_student_program_ids(program_label)
            if not program_ids:
                skipped_programs.add((program_label, "غير محدد", path.name))
                continue

            phrase_text = clean_text(row.get(phrase_col))
            average = decimal_number(row.get(average_col))
            if not phrase_text or average is None:
                continue

            classification = classify_student_phrase(phrase_text, legacy_catalog)
            add_item_measurement(
                datasets,
                available_program_years,
                available_genders,
                program_ids=program_ids,
                year="1447",
                survey_title=classification["surveyTitle"],
                stakeholder="students",
                section_id=classification["sectionId"],
                topic_label=classification["topicLabel"],
                item_number=classification["itemNumber"],
                item_label=classification["itemLabel"],
                gender=normalize_gender(row.get(gender_col)),
                source_note=f"{STUDENT_1447_SOURCE_NOTE} المصدر الفرعي: {path.name}.",
                average=average,
            )


def import_faculty_1447_surveys(
    datasets: "OrderedDict[str, Dict[str, object]]",
    available_program_years: Dict[str, set],
    available_genders: set,
) -> None:
    for source in FACULTY_1447_SOURCES:
        path = source["path"]
        department_label = source["department"]
        if not path.exists():
            continue

        program_ids = DEPARTMENT_PROGRAMS.get(department_label, [])
        if not program_ids:
            continue

        rows = read_first_sheet_rows(path)
        if not rows:
            continue

        header = {label: column for column, label in rows[0].items()}
        gender_col = header.get("الجنس")
        average_col = header.get("المتوسط")
        phrase_col = header.get("الاسئلة")
        if not all([gender_col, average_col, phrase_col]):
            continue

        for row in rows[1:]:
            phrase_text = clean_text(row.get(phrase_col))
            average = decimal_number(row.get(average_col))
            if not phrase_text or average is None:
                continue

            classification = classify_faculty_phrase(phrase_text, department_label)
            add_item_measurement(
                datasets,
                available_program_years,
                available_genders,
                program_ids=program_ids,
                year="1447",
                survey_title=classification["surveyTitle"],
                stakeholder="faculty",
                section_id=classification["sectionId"],
                topic_label=classification["topicLabel"],
                item_number=classification["itemNumber"],
                item_label=classification["itemLabel"],
                gender=normalize_gender(row.get(gender_col)),
                source_note=f"{FACULTY_1447_SOURCE_NOTE} المصدر الفرعي: {path.name}.",
                average=average,
            )


def finalize_payload(
    datasets: "OrderedDict[str, Dict[str, object]]",
    available_program_years: Dict[str, set],
    available_genders: set,
    skipped_programs: set,
) -> Dict[str, object]:
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
                    genders = []
                    for gender_entry in item["_genders"].values():
                        responses = gender_entry.get("responses")
                        score_total = gender_entry.get("scoreTotal")
                        average_samples = int(gender_entry.get("_averageSamples") or 0)
                        average_total = float(gender_entry.get("_averageTotal") or 0)

                        final_entry = {"gender": gender_entry["gender"]}
                        if responses and score_total is not None:
                            final_entry["responses"] = int(responses)
                            final_entry["scoreTotal"] = int(score_total)
                        elif average_samples:
                            final_entry["average"] = round_score(average_total / average_samples)
                            final_entry["measurementCount"] = average_samples
                        else:
                            continue

                        genders.append(final_entry)

                    if not genders:
                        continue

                    item_record_count += len(genders)
                    items.append(
                        {
                            "number": item["number"],
                            "label": item["label"],
                            "genders": genders,
                        }
                    )

                if not items:
                    continue

                items.sort(key=item_sort_key)
                topics.append({"label": topic["label"], "items": items})

            if not topics:
                continue

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
            "source": " | ".join(sorted(dataset["_sources"])),
            "surveys": surveys,
        }
        survey_count += len(surveys)

    return {
        "sourceLabel": "منصة ذكاء الأعمال",
        "sourceFile": LEGACY_WORKBOOK_PATH.name,
        "sourceFiles": [
            LEGACY_WORKBOOK_PATH.name,
            *[item["path"].name for item in STUDENT_1447_SOURCES if item["path"].exists()],
            *[item["path"].name for item in FACULTY_1447_SOURCES if item["path"].exists()],
        ],
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
            {"program": program, "degree": degree, "source": source}
            for program, degree, source in sorted(skipped_programs)
        ],
        "notes": [
            "بيانات 1445-1446 تحتفظ بعدد المقيمين الأصلي من منصة ذكاء الأعمال.",
            "بيانات 1447 العامة والبيانات القسمية للكادر الأكاديمي لا تتضمن عدد المقيمين، لذلك يعرض الموقع المتوسط فقط ويجعل عدد المقيمين غير متاح عند الحاجة.",
        ],
        "extractedData": extracted_data,
    }


def build_payload() -> Dict[str, object]:
    datasets, available_program_years, available_genders, skipped_programs = create_dataset_store()

    legacy_rows = read_first_sheet_rows(LEGACY_WORKBOOK_PATH)
    if legacy_rows:
        data_rows = legacy_rows[1:]
        import_legacy_student_surveys(
            data_rows,
            datasets,
            available_program_years,
            available_genders,
            skipped_programs,
        )
        legacy_catalog = build_legacy_phrase_catalog(data_rows)
    else:
        legacy_catalog = {}

    import_student_1447_surveys(
        datasets,
        available_program_years,
        available_genders,
        skipped_programs,
        legacy_catalog,
    )
    import_faculty_1447_surveys(
        datasets,
        available_program_years,
        available_genders,
    )

    return finalize_payload(datasets, available_program_years, available_genders, skipped_programs)


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
        print("Skipped program labels:")
        for item in payload["skippedPrograms"]:
            print(f" - {item['program']} | {item['degree']} | {item['source']}")


if __name__ == "__main__":
    main()
