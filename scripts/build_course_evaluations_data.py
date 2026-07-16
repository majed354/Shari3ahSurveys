#!/usr/bin/env python3

from __future__ import annotations

import base64
import csv
import json
import os
import re
import unicodedata
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from statistics import median
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = ROOT / "js" / "course-evaluations-data.js"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
PBKDF2_ITERATIONS = 600_000
AAD = b"shari3ah-course-evaluations-v1"

PROGRAM_NAMES = [
    "بكالوريوس القرآن وعلومه",
    "ماجستير الدراسات القرآنية المعاصرة",
    "دكتوراه الدراسات القرآنية",
    "بكالوريوس القراءات",
    "ماجستير القراءات",
    "دكتوراه القراءات",
    "بكالوريوس الدراسات الإسلامية",
    "بكالوريوس الشريعة",
    "بكالوريوس الأنظمة",
    "ماجستير أصول الفقه",
    "ماجستير الفقه",
    "ماجستير القانون",
    "ماجستير العقيدة",
    "دكتوراه الفقه",
    "دكتوراه أصول الفقه",
    "دبلومات وانتساب",
]

PROGRAM_BY_DEGREE_SPECIALTY = {
    ("البكالوريوس", "القرآن وعلومه"): PROGRAM_NAMES[0],
    ("الماجستير", "الدراسات القرآنية المعاصرة"): PROGRAM_NAMES[1],
    ("الدكتوراه", "الدراسات القرآنية"): PROGRAM_NAMES[2],
    ("البكالوريوس", "القراءات"): PROGRAM_NAMES[3],
    ("الماجستير", "القراءات"): PROGRAM_NAMES[4],
    ("الدكتوراه", "القراءات"): PROGRAM_NAMES[5],
    ("البكالوريوس", "الدراسات الإسلامية"): PROGRAM_NAMES[6],
    ("البكالوريوس", "الشريعة"): PROGRAM_NAMES[7],
    ("البكالوريوس", "الأنظمة"): PROGRAM_NAMES[8],
    ("الماجستير", "أصول الفقه"): PROGRAM_NAMES[9],
    ("الماجستير", "الفقه"): PROGRAM_NAMES[10],
    ("الماجستير", "القانون"): PROGRAM_NAMES[11],
    ("الماجستير", "العقيدة"): PROGRAM_NAMES[12],
    ("الدكتوراه", "الفقه"): PROGRAM_NAMES[13],
    ("الدكتوراه", "أصول الفقه"): PROGRAM_NAMES[14],
}

DIMENSIONS = [
    ("plan", "وضوح الخطة والتعليمات"),
    ("strategies", "استراتيجيات التدريس"),
    ("interaction", "التواصل والتفاعل"),
    ("assessment", "التقويم والتغذية الراجعة"),
    ("technology", "التقنية والتعلم الإلكتروني"),
    ("alignment", "ارتباط الأنشطة بالأهداف"),
    ("satisfaction", "الرضا الكلي"),
]


def clean_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\u00a0", " ")).strip()


@lru_cache(maxsize=None)
def normalize_arabic(value: object) -> str:
    text = unicodedata.normalize("NFKC", clean_text(value)).lower()
    text = "".join(char for char in unicodedata.normalize("NFD", text) if unicodedata.category(char) != "Mn")
    replacements = {"أ": "ا", "إ": "ا", "آ": "ا", "ؤ": "و", "ئ": "ي", "ء": "", "ة": "ه", "ى": "ي"}
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"[^\w\u0600-\u06ff()]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def numeric_id(value: object) -> str:
    text = clean_text(value)
    try:
        number = float(text)
        return str(int(number)) if number.is_integer() else text
    except (TypeError, ValueError):
        return text


def col_index(ref: str) -> int:
    match = re.match(r"[A-Z]+", ref)
    letters = match.group(0) if match else "A"
    value = 0
    for char in letters:
        value = value * 26 + ord(char) - 64
    return value - 1


def load_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in si.iter(f"{NS}t")) for si in root.findall(f"{NS}si")]


def row_values(row: ET.Element, shared: list[str]) -> list[object]:
    values: dict[int, object] = {}
    for cell in row.findall(f"{NS}c"):
        index = col_index(cell.get("r", "A1"))
        cell_type = cell.get("t")
        raw: object = ""
        if cell_type == "inlineStr":
            raw = "".join(node.text or "" for node in cell.iter(f"{NS}t"))
        else:
            value_node = cell.find(f"{NS}v")
            raw = value_node.text if value_node is not None and value_node.text is not None else ""
            if cell_type == "s" and raw != "":
                raw = shared[int(raw)]
            elif cell_type in {None, "n"} and raw != "":
                try:
                    number = float(raw)
                    raw = int(number) if number.is_integer() else number
                except ValueError:
                    pass
        values[index] = raw
    width = max(values, default=-1) + 1
    return [values.get(index, "") for index in range(width)]


def find_source_files() -> tuple[Path, Path, Path]:
    workbooks = [path for path in ROOT.rglob("*.xlsx") if path.name.startswith("استطلاعات الطلاب حول كافة المقررات")]
    if not workbooks:
        raise FileNotFoundError("لم يُعثر على مصنف تقييمات الطلاب.")
    source = sorted(workbooks, key=lambda path: ("كواد تقييم الطلاب" not in normalize_arabic(path.parent.name), len(str(path))))[0]
    text_files = list(ROOT.rglob("*.txt"))
    materials = next(
        (path for path in text_files if normalize_arabic(path.stem) == normalize_arabic("مواد البرنامج")),
        None,
    )
    faculty = next(
        (path for path in text_files if normalize_arabic(path.stem) == normalize_arabic("أسماء الأعضاء")),
        None,
    )
    if materials is None or faculty is None:
        raise FileNotFoundError("لم يُعثر على أكواد تصنيف المقررات أو مطابقة الأعضاء.")
    return source, materials, faculty


def parse_course_catalog(path: Path) -> dict[str, dict[str, set[str]]]:
    text = path.read_text(encoding="utf-8-sig")
    catalog: dict[str, dict[str, set[str]]] = {name: defaultdict(set) for name in PROGRAM_NAMES}
    pattern = re.compile(r"GetP(\d+)\s*=\s*\"([^\"]+)\"")
    for match in pattern.finditer(text):
        program_index = int(match.group(1)) - 1
        if not 0 <= program_index < len(PROGRAM_NAMES):
            continue
        value = clean_text(match.group(2))
        if "|" not in value:
            continue
        code, course_name = [clean_text(part) for part in value.split("|", 1)]
        if code and course_name and code != "???":
            catalog[PROGRAM_NAMES[program_index]][normalize_arabic(course_name)].add(code)
    return catalog


def parse_faculty_names(path: Path) -> dict[str, str]:
    text = path.read_text(encoding="utf-8-sig")
    mapping: dict[str, str] = {}
    for match in re.finditer(r'arr\(\d+\)\s*=\s*\"([^\"]+)\|([^|\"]+)\|([^\"]+)\"', text):
        name, faculty_id = clean_text(match.group(1)), numeric_id(match.group(2))
        if name and faculty_id:
            mapping[faculty_id] = name
    return mapping


def find_faculty_csv() -> Path:
    configured = clean_text(os.environ.get("COURSE_EVAL_FACULTY_CSV", ""))
    candidates = [
        Path(configured).expanduser() if configured else None,
        ROOT.parent / "faculty-activities-main" / "data" / "faculty.csv",
    ]
    for candidate in candidates:
        if candidate is not None and candidate.is_file():
            return candidate
    raise FileNotFoundError(
        "لم يُعثر على faculty.csv. يمكن تحديده في COURSE_EVAL_FACULTY_CSV."
    )


@lru_cache(maxsize=None)
def normalize_person_name(value: object) -> str:
    tokens = normalize_arabic(value).split()
    title_tokens = {"ا", "د", "اد", "استاذ", "الاستاذ", "دكتور", "الدكتور"}
    while tokens and tokens[0] in title_tokens:
        tokens.pop(0)
    text = " ".join(tokens)
    compounds = {
        "عبد الله": "عبدالله",
        "عبد الرحمن": "عبدالرحمن",
        "عبد العزيز": "عبدالعزيز",
        "عبد الحميد": "عبدالحميد",
        "عبد الواحد": "عبدالواحد",
        "عبد الغفور": "عبدالغفور",
    }
    for separated, joined in compounds.items():
        text = text.replace(separated, joined)
    return re.sub(r"\s+", " ", text).strip()


def parse_faculty_departments(path: Path) -> dict[str, object]:
    by_id_year: dict[tuple[str, str], set[str]] = defaultdict(set)
    by_name_year: dict[tuple[str, str], set[str]] = defaultdict(set)
    latest_by_id: dict[str, tuple[int, set[str]]] = {}
    latest_by_name: dict[str, tuple[int, set[str]]] = {}
    departments: set[str] = set()
    row_count = 0

    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        for row in csv.DictReader(stream):
            faculty_id = numeric_id(row.get("id", ""))
            year = clean_text(row.get("year", ""))
            name_key = normalize_person_name(row.get("name", ""))
            department = clean_text(row.get("department", ""))
            if not department:
                continue
            try:
                year_number = int(float(year))
            except (TypeError, ValueError):
                continue
            row_count += 1
            departments.add(department)
            if faculty_id:
                by_id_year[(faculty_id, year)].add(department)
                previous = latest_by_id.get(faculty_id)
                if previous is None or year_number > previous[0]:
                    latest_by_id[faculty_id] = (year_number, {department})
                elif year_number == previous[0]:
                    previous[1].add(department)
            if name_key:
                by_name_year[(name_key, year)].add(department)
                previous = latest_by_name.get(name_key)
                if previous is None or year_number > previous[0]:
                    latest_by_name[name_key] = (year_number, {department})
                elif year_number == previous[0]:
                    previous[1].add(department)

    return {
        "byIdYear": by_id_year,
        "byNameYear": by_name_year,
        "latestById": latest_by_id,
        "latestByName": latest_by_name,
        "departments": departments,
        "rows": row_count,
    }


def unique_department(values: set[str] | None) -> str | None:
    return next(iter(values)) if values and len(values) == 1 else None


def resolve_faculty_department(
    profiles: dict[str, object],
    faculty_id: str,
    faculty_name: str,
    year: str,
    *,
    latest: bool = False,
) -> tuple[str, str]:
    name_key = normalize_person_name(faculty_name)
    if latest:
        if faculty_id:
            latest_id = profiles["latestById"].get(faculty_id)
            department = unique_department(latest_id[1]) if latest_id else None
            if department:
                return department, "idLatest"
        if name_key:
            latest_name = profiles["latestByName"].get(name_key)
            department = unique_department(latest_name[1]) if latest_name else None
            if department:
                return department, "nameLatest"
        return "غير محدد", "unmatched"

    if faculty_id:
        department = unique_department(profiles["byIdYear"].get((faculty_id, year)))
        if department:
            return department, "idYear"
        latest_id = profiles["latestById"].get(faculty_id)
        department = unique_department(latest_id[1]) if latest_id else None
        if department:
            return department, "idLatest"
    if name_key:
        department = unique_department(profiles["byNameYear"].get((name_key, year)))
        if department:
            return department, "nameYear"
        latest_name = profiles["latestByName"].get(name_key)
        department = unique_department(latest_name[1]) if latest_name else None
        if department:
            return department, "nameLatest"
    return "غير محدد", "unmatched"


def classify_dimension(statement: str) -> str | None:
    text = normalize_arabic(statement)

    plan_terms = (
        "عرض خطه المقرر", "اليه تصحيح", "الخدمات الاكاديميه", "الخطوط الاساسيه",
        "متطلبات النجاح", "مصادر مساعدتي", "ملتزما باعطاء المقرر بشكل كامل",
        "المصادر التي احتجتها في هذا المقرر متوافره",
    )
    assessment_terms = (
        "تغذيه راجعه", "درجات الواجبات", "تصحيح واجباتي", "اختباراتي عادلا",
    )
    interaction_terms = (
        "فرص متكافيه", "التواصل مع عضو", "وسيله للتواصل", "مناقشه زملائي",
        "مهتما بمدي تقدمي", "موجودا للمساعده", "القاء الاسيله", "الاتصال بفاعليه",
    )
    technology_terms = (
        "الدعم الفني", "استخدام التكنولوجيا", "التدريب والارشاد", "متابعه ما تغيبت",
        "يوفر التعلم الالكتروني وقت", "استخدام فعال للتقنيه",
    )
    strategy_terms = (
        "عرض الدروس باشكال", "اسلوب شيق", "تنوعت الوسايط", "متحمسا لما يقوم بتدريسه",
        "المام كامل بمحتوي", "حديثا ومفيدا",
    )
    alignment_terms = (
        "الانشطه التعليميه وادوات المقرر", "ارتبطت اهداف", "الانشطه والتكليفات",
        "متسقه مع الخطوط", "تطوير معرفتي ومهاراتي", "كميه العمل", "الصله بين هذا المقرر",
        "التفكير وحل المشكلات", "العمل علي شكل فريق", "مهم وسيفيدني مستقبلا",
    )
    satisfaction_terms = (
        "راضي عن", "اشعر بالرضا", "شعر بالرضا عن", "اشعر بالمتعه", "شجعني التعلم", "تقديم افضل ما عندي",
    )

    for dimension, terms in (
        ("plan", plan_terms),
        ("assessment", assessment_terms),
        ("interaction", interaction_terms),
        ("technology", technology_terms),
        ("strategies", strategy_terms),
        ("alignment", alignment_terms),
        ("satisfaction", satisfaction_terms),
    ):
        if any(normalize_arabic(term) in text for term in terms):
            return dimension
    return None


def is_research_course(course_name: str) -> bool:
    text = normalize_arabic(course_name)
    return text in {"الرساله", "بحث التخرج"} or "مشروع بحث" in text or "مشروع تخرج بحثي" in text


def program_name(degree: str, specialty: str) -> str:
    return PROGRAM_BY_DEGREE_SPECIALTY.get((degree, specialty), f"{degree} {specialty}".strip())


def evidence_band(responses: int) -> str:
    if responses < 5:
        return "فردي/غير كافٍ"
    if responses < 10:
        return "محدود"
    if responses < 20:
        return "متوسط"
    return "قوي"


def score_grade(score: float) -> str:
    if score >= 4.5:
        return "ممتاز"
    if score >= 4.0:
        return "جيد جداً"
    if score >= 3.5:
        return "جيد"
    return "مقبول"


def recommendation(score: float, dimension_scores: dict[str, float | None]) -> str:
    available = [(key, value) for key, value in dimension_scores.items() if value is not None]
    weakest = min(available, key=lambda item: item[1]) if available else None
    label_lookup = dict(DIMENSIONS)
    if score >= 4.5:
        return "المحافظة على التميز وتوثيق الممارسة ونشرها."
    if score >= 4.0:
        return f"استمرار التحسين مع متابعة محور {label_lookup.get(weakest[0], 'الأداء الأقل')} بشكل دوري." if weakest else "استمرار التحسين والمتابعة الدورية."
    if score >= 3.5:
        return f"إعداد إجراء تحسيني مركز لمحور {label_lookup.get(weakest[0], 'الأداء الأقل')}." if weakest else "إعداد إجراء تحسيني ومتابعته في الفصل التالي."
    return "إعداد خطة تصحيحية موثقة، ومتابعة التنفيذ والقياس في الفصل التالي."


def parse_sections(
    source: Path,
    faculty_names: dict[str, str],
    faculty_profiles: dict[str, object],
) -> tuple[list[dict[str, object]], dict[str, int]]:
    sheets = {"1447": "xl/worksheets/sheet1.xml", "1446": "xl/worksheets/sheet2.xml", "1445": "xl/worksheets/sheet3.xml"}
    sections: dict[tuple[object, ...], dict[str, object]] = {}
    quality = Counter()
    unclassified_statements = Counter()

    with zipfile.ZipFile(source) as archive:
        shared = load_shared_strings(archive)
        for year, xml_path in sheets.items():
            row_number = 0
            with archive.open(xml_path) as stream:
                for _, elem in ET.iterparse(stream, events=("end",)):
                    if elem.tag != f"{NS}row":
                        continue
                    values = row_values(elem, shared)
                    elem.clear()
                    row_number += 1
                    if row_number == 1:
                        continue
                    values.extend([""] * (19 - len(values)))
                    semester, degree, campus, college, department, specialty = [clean_text(value) for value in values[0:6]]
                    instructor_id = numeric_id(values[6])
                    raw_instructor_name = clean_text(values[7])
                    canonical_name = faculty_names.get(instructor_id, raw_instructor_name)
                    course, course_type, section = [clean_text(value) for value in values[8:11]]
                    topic, item_number, statement = values[11:14]

                    if not course:
                        quality["missingCourseRows"] += 1
                        continue
                    counts: list[int] = []
                    for value in values[14:19]:
                        try:
                            counts.append(int(float(value or 0)))
                        except (TypeError, ValueError):
                            counts.append(0)
                    response_count = sum(counts)
                    if response_count <= 0:
                        quality["zeroResponseItems"] += 1
                        continue
                    weighted_score = sum((index + 1) * count for index, count in enumerate(counts))
                    dimension = classify_dimension(clean_text(statement))
                    if dimension is None:
                        quality["unclassifiedItems"] += 1
                        unclassified_statements[clean_text(statement)] += 1

                    instructor_key = (
                        f"id:{instructor_id}"
                        if instructor_id
                        else f"name:{normalize_arabic(raw_instructor_name)}"
                    )
                    key = (
                        year, semester, degree, campus, college, department, specialty,
                        instructor_key, canonical_name, course, course_type, section,
                    )
                    if key not in sections:
                        faculty_department, faculty_department_source = resolve_faculty_department(
                            faculty_profiles,
                            instructor_id,
                            canonical_name,
                            year,
                        )
                        faculty_department_latest, faculty_department_latest_source = resolve_faculty_department(
                            faculty_profiles,
                            instructor_id,
                            canonical_name,
                            year,
                            latest=True,
                        )
                        sections[key] = {
                            "year": year,
                            "semester": semester,
                            "degree": degree,
                            "campus": campus,
                            "college": college,
                            "department": department,
                            "specialty": specialty,
                            "program": program_name(degree, specialty),
                            "instructorKey": instructor_key,
                            "instructorName": canonical_name,
                            "facultyDepartment": faculty_department,
                            "facultyDepartmentSource": faculty_department_source,
                            "facultyDepartmentLatest": faculty_department_latest,
                            "facultyDepartmentLatestSource": faculty_department_latest_source,
                            "course": course,
                            "courseType": course_type,
                            "section": section,
                            "responseCounts": [],
                            "scoreNumerator": 0,
                            "scoreDenominator": 0,
                            "dimensions": {key: [0, 0] for key, _ in DIMENSIONS},
                            "items": 0,
                        }
                    aggregate = sections[key]
                    aggregate["responseCounts"].append(response_count)
                    aggregate["scoreNumerator"] += weighted_score
                    aggregate["scoreDenominator"] += response_count
                    aggregate["items"] += 1
                    if dimension:
                        aggregate["dimensions"][dimension][0] += weighted_score
                        aggregate["dimensions"][dimension][1] += response_count

    finalized: list[dict[str, object]] = []
    for aggregate in sections.values():
        positive_counts = [count for count in aggregate.pop("responseCounts") if count > 0]
        responses = int(median(positive_counts)) if positive_counts else 0
        denominator = aggregate.pop("scoreDenominator")
        numerator = aggregate.pop("scoreNumerator")
        score = round(numerator / denominator, 4) if denominator else 0.0
        dimension_scores = {
            key: round(values[0] / values[1], 4) if values[1] else None
            for key, values in aggregate.pop("dimensions").items()
        }
        finalized.append({
            **aggregate,
            "responses": responses,
            "score": score,
            "dimensionScores": dimension_scores,
            "researchCourse": is_research_course(str(aggregate["course"])),
        })

    bachelor_groups: dict[tuple[str, ...], list[dict[str, object]]] = defaultdict(list)
    for record in finalized:
        if str(record["degree"]) != "البكالوريوس":
            continue
        group_key = (
            str(record["year"]),
            str(record["degree"]),
            str(record["program"]),
            str(record["instructorKey"]),
            normalize_arabic(record["course"]),
        )
        bachelor_groups[group_key].append(record)

    bachelor_small_sections = [
        record
        for records in bachelor_groups.values()
        for record in records
        if int(record["responses"]) <= 2
    ]
    quality["bachelorSmallSectionsBeforeGrouping"] = len(bachelor_small_sections)
    quality["bachelorSmallStudentsBeforeGrouping"] = sum(
        int(record["responses"]) for record in bachelor_small_sections
    )

    excluded_record_ids: set[int] = set()
    for group_records in bachelor_groups.values():
        total_responses = sum(int(record["responses"]) for record in group_records)
        small_sections = [record for record in group_records if int(record["responses"]) <= 2]
        if total_responses <= 2:
            excluded_record_ids.update(id(record) for record in group_records)
            quality["excludedBachelorSmallGroups"] += 1
            quality["excludedBachelorSmallSections"] += len(group_records)
            quality["excludedBachelorSmallStudents"] += total_responses
        else:
            quality["recoveredBachelorSmallSections"] += len(small_sections)
            quality["recoveredBachelorSmallStudents"] += sum(
                int(record["responses"]) for record in small_sections
            )

    quality["sectionsBeforeBachelorGroupRule"] = len(finalized)
    finalized = [record for record in finalized if id(record) not in excluded_record_ids]
    quality["sections"] = len(finalized)
    quality["facultyCatalogMatches"] = len(faculty_names)
    department_sources = Counter(str(record["facultyDepartmentSource"]) for record in finalized)
    for source_name, count in department_sources.items():
        quality[f"facultyDepartment{source_name[0].upper()}{source_name[1:]}Sections"] = count
    quality["facultyDepartmentMatchedSections"] = sum(
        count for source_name, count in department_sources.items() if source_name != "unmatched"
    )
    quality["facultyDepartmentUnmatchedSections"] = department_sources["unmatched"]
    if os.environ.get("COURSE_EVAL_DEBUG"):
        print("أكثر العبارات غير المصنفة:")
        for statement, count in unclassified_statements.most_common(40):
            print(f"{count:>6} | {statement}")
    return finalized, dict(quality)


def match_course_code(program: str, course: str, catalog: dict[str, dict[str, set[str]]]) -> tuple[str, str]:
    normalized = normalize_arabic(course)
    candidates = set(catalog.get(program, {}).get(normalized, set()))
    if len(candidates) == 1:
        return next(iter(candidates)), "matched"
    if len(candidates) > 1:
        return "", "ambiguous"
    global_candidates: set[str] = set()
    for program_catalog in catalog.values():
        global_candidates.update(program_catalog.get(normalized, set()))
    if len(global_candidates) == 1:
        return next(iter(global_candidates)), "cross-program"
    if len(global_candidates) > 1:
        return "", "ambiguous"
    return "", "unmatched"


def build_course_records(sections: list[dict[str, object]], catalog: dict[str, dict[str, set[str]]]) -> tuple[list[dict[str, object]], dict[str, int]]:
    groups: dict[tuple[str, ...], dict[str, object]] = {}
    code_statuses = Counter()
    for section in sections:
        for semester_key in (str(section["semester"]), "all"):
            key = (
                str(section["year"]),
                semester_key,
                str(section["degree"]),
                str(section["program"]),
                str(section["course"]),
            )
            if key not in groups:
                groups[key] = {
                    "year": section["year"],
                    "semester": semester_key,
                    "degree": section["degree"],
                    "program": section["program"],
                    "courseName": section["course"],
                    "campuses": set(),
                    "respondents": 0,
                    "sections": 0,
                    "scoreNumerator": 0.0,
                    "scoreDenominator": 0.0,
                    "dimensions": {dimension_key: [0.0, 0.0] for dimension_key, _ in DIMENSIONS},
                    "researchCourse": False,
                }
            group = groups[key]
            responses = int(section["responses"])
            group["campuses"].add(section["campus"])
            group["respondents"] += responses
            group["sections"] += 1
            group["scoreNumerator"] += float(section["score"]) * responses
            group["scoreDenominator"] += responses
            group["researchCourse"] = bool(group["researchCourse"] or section["researchCourse"])
            for dimension, value in section["dimensionScores"].items():
                if value is not None:
                    group["dimensions"][dimension][0] += float(value) * responses
                    group["dimensions"][dimension][1] += responses

    records: list[dict[str, object]] = []
    for group in groups.values():
        denominator = group.pop("scoreDenominator")
        numerator = group.pop("scoreNumerator")
        dimensions = group.pop("dimensions")
        score = round(numerator / denominator, 2) if denominator else 0.0
        dimension_scores = {
            key: round(values[0] / values[1], 2) if values[1] else None
            for key, values in dimensions.items()
        }
        code, code_status = match_course_code(str(group["program"]), str(group["courseName"]), catalog)
        code_statuses[code_status] += 1
        respondents = int(group["respondents"])
        records.append({
            **group,
            "campuses": sorted(group["campuses"]),
            "courseCode": code,
            "codeStatus": code_status,
            "score": score,
            "evidence": evidence_band(respondents),
            "recommendation": recommendation(score, dimension_scores),
        })
    def semester_sort_key(value: object) -> tuple[int, str]:
        semester = str(value)
        if semester == "all":
            return (99_999, semester)
        try:
            return (int(semester), semester)
        except ValueError:
            return (99_998, semester)

    records.sort(
        key=lambda row: (
            str(row["year"]),
            semester_sort_key(row["semester"]),
            str(row["program"]),
            normalize_arabic(row["courseName"]),
        )
    )
    counters = Counter()
    for record in records:
        counter_key = f"{record['year']}||{record['semester']}"
        counters[counter_key] += 1
        record["reportId"] = f"CE-{record['year']}-{record['semester']}-{counters[counter_key]:04d}"
    return records, dict(code_statuses)


def build_faculty_records(sections: list[dict[str, object]]) -> list[dict[str, object]]:
    groups: dict[tuple[str, str, str, str], dict[str, object]] = {}
    for section in sections:
        for year_key in (str(section["year"]), "all"):
            key = (year_key, str(section["degree"]), str(section["program"]), str(section["instructorKey"]))
            department_key = "facultyDepartmentLatest" if year_key == "all" else "facultyDepartment"
            faculty_department = str(section.get(department_key) or "غير محدد")
            if key not in groups:
                groups[key] = {
                    "year": year_key,
                    "degree": section["degree"],
                    "program": section["program"],
                    "instructorName": section["instructorName"],
                    "facultyDepartment": faculty_department,
                    "sectionsData": [],
                }
            elif groups[key]["facultyDepartment"] == "غير محدد" and faculty_department != "غير محدد":
                groups[key]["facultyDepartment"] = faculty_department
            groups[key]["sectionsData"].append(section)

    records: list[dict[str, object]] = []
    for group in groups.values():
        all_section_rows = group.pop("sectionsData")
        is_bachelor = str(group["degree"]) == "البكالوريوس"
        eligible_section_rows = [
            section
            for section in all_section_rows
            if not is_bachelor or int(section["responses"]) >= 5
        ]
        excluded_bachelor_rows = [
            section
            for section in all_section_rows
            if is_bachelor and int(section["responses"]) < 5
        ]

        total_responses = sum(int(section["responses"]) for section in eligible_section_rows)
        weighted_score = (
            sum(float(section["score"]) * int(section["responses"]) for section in eligible_section_rows)
            / total_responses
            if total_responses else 0.0
        )
        dimension_sums = {dimension: 0.0 for dimension, _ in DIMENSIONS}
        dimension_weights = {dimension: 0.0 for dimension, _ in DIMENSIONS}
        for section in eligible_section_rows:
            responses = int(section["responses"])
            for dimension, value in section["dimensionScores"].items():
                if value is None:
                    continue
                dimension_sums[dimension] += float(value) * responses
                dimension_weights[dimension] += responses

        weighted_dimensions = {
            dimension: round(dimension_sums[dimension] / dimension_weights[dimension], 2)
            if dimension_weights[dimension] else None
            for dimension, _ in DIMENSIONS
        }
        if is_bachelor:
            eligible = total_responses >= 20
            if eligible:
                status = "مؤهل للمقارنة"
            elif not eligible_section_rows:
                status = "لا توجد شعبة بخمس استجابات"
            else:
                status = "أقل من 20 استجابة"
        else:
            eligible = total_responses >= 4
            status = "مؤهل ضمن الدراسات العليا" if eligible else "أقل من 4 استجابات"

        records.append({
            **group,
            "scoreWeighted": round(weighted_score, 2),
            "respondents": total_responses,
            "sections": len(eligible_section_rows),
            "courses": len({normalize_arabic(section["course"]) for section in eligible_section_rows}),
            "excludedBachelorSections": len(excluded_bachelor_rows),
            "researchSectionsIncluded": sum(
                1 for section in eligible_section_rows if bool(section["researchCourse"])
            ),
            "minSectionRespondents": min(
                (int(section["responses"]) for section in eligible_section_rows),
                default=0,
            ),
            "evidence": evidence_band(total_responses),
            "grade": score_grade(weighted_score) if total_responses else "غير محتسب",
            "eligible": eligible,
            "status": status,
            "dimensionScores": weighted_dimensions,
        })
    records.sort(key=lambda row: (str(row["year"]), str(row["program"]), not bool(row["eligible"]), -float(row["scoreWeighted"])))
    return records


def build_public_excellence_leaders(
    faculty_records: list[dict[str, object]],
) -> list[dict[str, object]]:
    groups: dict[tuple[str, str, str], list[dict[str, object]]] = defaultdict(list)
    for record in faculty_records:
        if not bool(record["eligible"]):
            continue
        key = (str(record["year"]), str(record["degree"]), str(record["program"]))
        groups[key].append(record)

    public_records: list[dict[str, object]] = []
    for (year, degree, program), records in groups.items():
        selection_departments = ["all", *sorted({str(record["facultyDepartment"]) for record in records})]
        for selection_department in selection_departments:
            candidates = [
                record
                for record in records
                if selection_department == "all"
                or str(record["facultyDepartment"]) == selection_department
            ]
            candidates.sort(
                key=lambda record: (
                    -float(record["scoreWeighted"]),
                    normalize_person_name(record["instructorName"]),
                )
            )
            last_score: float | None = None
            displayed_rank = 0
            for position, record in enumerate(candidates[:5], start=1):
                current_score = float(record["scoreWeighted"])
                if last_score is None or current_score != last_score:
                    displayed_rank = position
                last_score = current_score
                public_records.append({
                    "year": year,
                    "degree": degree,
                    "program": program,
                    "selectionDepartment": selection_department,
                    "facultyDepartment": record["facultyDepartment"],
                    "rank": displayed_rank,
                    "instructorName": record["instructorName"],
                    "scoreWeighted": record["scoreWeighted"],
                    "grade": record["grade"],
                    "respondents": record["respondents"],
                    "sections": record["sections"],
                    "courses": record["courses"],
                })
    public_records.sort(
        key=lambda record: (
            str(record["year"]),
            str(record["program"]),
            str(record["degree"]),
            str(record["selectionDepartment"]),
            int(record["rank"]),
            normalize_person_name(record["instructorName"]),
        )
    )
    return public_records


def encrypt_private_payload(payload: dict[str, object], pin: str) -> dict[str, object]:
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    salt = os.urandom(16)
    iv = os.urandom(12)
    key = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=PBKDF2_ITERATIONS,
    ).derive(pin.encode("utf-8"))
    plaintext = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    ciphertext = AESGCM(key).encrypt(iv, plaintext, AAD)
    return {
        "algorithm": "AES-GCM",
        "kdf": "PBKDF2-SHA256",
        "iterations": PBKDF2_ITERATIONS,
        "salt": base64.b64encode(salt).decode("ascii"),
        "iv": base64.b64encode(iv).decode("ascii"),
        "aad": AAD.decode("ascii"),
        "ciphertext": base64.b64encode(ciphertext).decode("ascii"),
    }


def validate_outputs(
    course_records: list[dict[str, object]],
    faculty_records: list[dict[str, object]],
    public_leaders: list[dict[str, object]],
) -> None:
    report_ids = [str(record["reportId"]) for record in course_records]
    if len(report_ids) != len(set(report_ids)):
        raise ValueError("معرّفات التقرير السنوي غير فريدة.")

    for record in course_records:
        if any(
            private_key in record
            for private_key in ("instructorName", "instructorKey", "facultyDepartment")
        ):
            raise ValueError("تسرّب حقل خاص بالأعضاء إلى البيانات العامة.")
        if not 1 <= float(record["score"]) <= 5:
            raise ValueError("درجة مقرر خارج النطاق 1-5.")

    for record in faculty_records:
        if not clean_text(record["instructorName"]):
            raise ValueError("سجل عضو بلا اسم.")
        if not clean_text(record["facultyDepartment"]):
            raise ValueError("سجل عضو بلا قسم أو حالة قسم.")
        if bool(record["eligible"]):
            if str(record["degree"]) == "البكالوريوس":
                valid = int(record["respondents"]) >= 20 and int(record["minSectionRespondents"]) >= 5
            else:
                valid = int(record["respondents"]) >= 4
            if not valid:
                raise ValueError("سجل مؤهل لا يحقق حد الأدلة المعلن.")
        if int(record["respondents"]) > 0 and not 1 <= float(record["scoreWeighted"]) <= 5:
            raise ValueError("متوسط موزون لعضو خارج النطاق 1-5.")

    public_group_counts = Counter(
        (
            str(record["year"]),
            str(record["degree"]),
            str(record["program"]),
            str(record["selectionDepartment"]),
        )
        for record in public_leaders
    )
    if any(count > 5 for count in public_group_counts.values()):
        raise ValueError("تجاوزت قائمة عامة الحد الأقصى المسموح وهو خمسة أعضاء.")
    allowed_public_fields = {
        "year", "degree", "program", "selectionDepartment", "facultyDepartment",
        "rank", "instructorName", "scoreWeighted", "grade", "respondents", "sections", "courses",
    }
    private_eligible_keys = {
        (
            str(record["year"]),
            str(record["degree"]),
            str(record["program"]),
            str(record["facultyDepartment"]),
            str(record["instructorName"]),
        )
        for record in faculty_records
        if bool(record["eligible"])
    }
    for record in public_leaders:
        if set(record) != allowed_public_fields:
            raise ValueError("تتضمن قائمة الخمسة العامة تفاصيل زائدة غير مصرح بها.")
        key = (
            str(record["year"]),
            str(record["degree"]),
            str(record["program"]),
            str(record["facultyDepartment"]),
            str(record["instructorName"]),
        )
        if key not in private_eligible_keys:
            raise ValueError("تتضمن قائمة الخمسة العامة عضواً غير مؤهل.")


def main() -> None:
    pin = os.environ.get("COURSE_EVAL_ADMIN_PIN", "").strip()
    if not pin:
        raise RuntimeError("عيّن COURSE_EVAL_ADMIN_PIN عند توليد بيانات تقرير التميز.")

    source, materials_path, faculty_path = find_source_files()
    faculty_csv_path = find_faculty_csv()
    catalog = parse_course_catalog(materials_path)
    faculty_names = parse_faculty_names(faculty_path)
    faculty_profiles = parse_faculty_departments(faculty_csv_path)
    sections, quality = parse_sections(source, faculty_names, faculty_profiles)
    course_records, code_statuses = build_course_records(sections, catalog)
    faculty_records = build_faculty_records(sections)
    public_leaders = build_public_excellence_leaders(faculty_records)
    validate_outputs(course_records, faculty_records, public_leaders)

    section_bands = Counter(evidence_band(int(section["responses"])) for section in sections)
    degrees = sorted({str(section["degree"]) for section in sections})
    programs = sorted({str(section["program"]) for section in sections})
    years = sorted({str(section["year"]) for section in sections}, reverse=True)
    semesters_by_year = {
        year: sorted(
            {str(section["semester"]) for section in sections if str(section["year"]) == year},
            key=lambda semester: int(semester) if str(semester).isdigit() else str(semester),
        )
        for year in years
    }
    program_department_counts: dict[str, Counter[str]] = defaultdict(Counter)
    for section in sections:
        department = clean_text(section.get("department"))
        if department:
            program_department_counts[str(section["program"])][department] += 1
    program_default_departments = {
        program: counts.most_common(1)[0][0]
        for program, counts in program_department_counts.items()
        if counts
    }
    available_departments = set(faculty_profiles["departments"])
    available_departments.update(
        clean_text(record.get("facultyDepartment")) for record in faculty_records
    )
    available_departments.discard("")
    preferred_department_order = ["القراءات", "الشريعة", "الأنظمة", "الثقافة الإسلامية"]
    departments = [
        department for department in preferred_department_order if department in available_departments
    ]
    departments.extend(sorted(available_departments - set(departments) - {"غير محدد"}))
    if "غير محدد" in available_departments:
        departments.append("غير محدد")

    private_payload = {
        "facultyRecords": faculty_records,
        "dimensions": [{"id": key, "label": label} for key, label in DIMENSIONS],
        "methodology": {
            "scoreMethod": "المتوسط الموزون بعدد الاستجابات: مجموع (درجة الشعبة × عدد مستجيبيها) ÷ مجموع المستجيبين",
            "bachelorSectionMinimum": "لا تدخل شعبة البكالوريوس في تقرير التميز إذا كان عدد مستجيبيها أقل من 5",
            "bachelorEligibility": "20 استجابة محتسبة للعضو، سواء من مقرر واحد أو مقررات متعددة",
            "postgraduateEligibility": "4 استجابات محتسبة للعضو، مع استخدام المتوسط الموزون نفسه",
            "researchRule": "تدخل الرسالة والمشروع البحثي ضمن المتوسط الموزون ودرجة التميز، وتخضع لحدود البكالوريوس أو الدراسات العليا بحسب الدرجة",
            "departmentRule": "يُحدد قسم العضو من سجل أعضاء هيئة التدريس بحسب سنة التقييم، وتستخدم أحدث عضوية متاحة عند تجميع السنوات",
        },
    }

    public_payload = {
        "meta": {
            "sourceLabel": "استطلاعات الطلاب التفصيلية لمقررات كلية الشريعة والأنظمة 1445-1447هـ",
            "sourceModified": datetime.fromtimestamp(source.stat().st_mtime).astimezone().isoformat(timespec="seconds"),
            "years": years,
            "semestersByYear": semesters_by_year,
            "degrees": degrees,
            "programs": programs,
            "departments": departments,
            "programDefaultDepartments": program_default_departments,
            "dimensions": [{"id": key, "label": label} for key, label in DIMENSIONS],
            "methodology": {
                "scale": "1-5",
                "bachelorGroupingRule": "تُجمع شعب المقرر للعضو نفسه داخل السنة والبرنامج، ثم تُستبعد المجموعة فقط إذا كان مجموع طلابها طالبًا واحدًا أو طالبين",
                "excellenceScoreMethod": "المتوسط الموزون بعدد الاستجابات، مع حد 5 لكل شعبة بكالوريوس و20 للعضو، وحد 4 للعضو في الدراسات العليا، وتشمل الرسائل والمشروعات البحثية",
                "departmentFilterRule": "القسم الافتراضي في تقرير التميز هو القسم الذي يتبعه البرنامج، ويمكن اختيار كل الأقسام أو قسم آخر؛ ويُعرض غير المطابقين تحت غير محدد",
                "publicLeadersRule": "تظهر للعامة بيانات الخمسة الأعلى فقط في كل برنامج ودرجة ونطاق قسم مختار، بينما تبقى القائمة الكاملة والتفاصيل الإضافية مشفرة بكلمة مرور المدير",
                "evidenceBands": [
                    {"label": "فردي/غير كافٍ", "range": "1-4"},
                    {"label": "محدود", "range": "5-9"},
                    {"label": "متوسط", "range": "10-19"},
                    {"label": "قوي", "range": "20 فأكثر"},
                ],
                "participationNote": "نسبة المشاركة في جدول التقرير السنوي 100٪؛ لأن السجل يمثل جميع الطلاب الذين درسوا المقرر في السنة المختارة بعد تجميع شعب العضو وتطبيق حد الاستبعاد على المجموعة.",
            },
            "dataQuality": {
                **quality,
                "courseRecords": len(course_records),
                "facultyRecordsEncrypted": len(faculty_records),
                "publicExcellenceLeaderRecords": len(public_leaders),
                "facultyDepartmentCsvRows": int(faculty_profiles["rows"]),
                "codeStatuses": code_statuses,
                "sectionEvidenceBands": dict(section_bands),
                "researchSections": sum(1 for section in sections if section["researchCourse"]),
                "excellenceResearchSectionsIncluded": sum(
                    1
                    for section in sections
                    if bool(section["researchCourse"])
                    and (
                        str(section["degree"]) != "البكالوريوس"
                        or int(section["responses"]) >= 5
                    )
                ),
                "excellenceBachelorSectionsBelowFive": sum(
                    1
                    for section in sections
                    if str(section["degree"]) == "البكالوريوس"
                    and int(section["responses"]) < 5
                ),
                "excellenceBachelorStudentsBelowFive": sum(
                    int(section["responses"])
                    for section in sections
                    if str(section["degree"]) == "البكالوريوس"
                    and int(section["responses"]) < 5
                ),
            },
        },
        "courseRecords": course_records,
        "publicExcellenceLeaders": public_leaders,
        "encryptedFaculty": encrypt_private_payload(private_payload, pin),
    }

    OUTPUT_PATH.write_text(
        "window.COURSE_EVALUATIONS_DATA = "
        + json.dumps(public_payload, ensure_ascii=False, separators=(",", ":"))
        + ";\n",
        encoding="utf-8",
    )
    print(json.dumps({
        "output": str(OUTPUT_PATH),
        "bytes": OUTPUT_PATH.stat().st_size,
        "courseRecords": len(course_records),
        "facultyRecordsEncrypted": len(faculty_records),
        "publicExcellenceLeaderRecords": len(public_leaders),
        "codeStatuses": code_statuses,
        "quality": quality,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
