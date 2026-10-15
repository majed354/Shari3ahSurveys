#!/usr/bin/env python3
"""Build the anonymous v2 API from the original workbook (Python 3.10+, stdlib).

No lecturer attributes leave this importer. The private evaluation key is used
only to detect overlapping source reports, never to infer programme or gender.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
import unicodedata
from xml.etree import ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SOURCE_NAME = "استطلاعات الطلاب حول كافة المقررات في برامج كلية الشريعة 1445-1447.xlsx"
DEFAULT_SOURCE = ROOT / "أكواد تقييم الطلاب" / SOURCE_NAME
DEFAULT_OUTPUT = ROOT / "data/course-evaluations/v2/student-course-evaluations.json"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
DOC_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
HEADERS = ["الفصل الدراسي", "الدرجة العلمية", "المقر", "الكلية", "القسم", "التخصص",
           "رقم المحاضر", "اسم المحاضر", "المقرر", "نوع المقرر", "الشعبة", "الموضوع",
           "رقم العبارة", "وصف العبارة", "غير موافق بشدة", "غير موافق", "محايد", "موافق", "موافق بشدة"]
# IDs match the existing survey API. Degree + specialty identify the programme;
# the source department is validated, never substituted for the specialty.
PROGRAMS = [
    ("p01", "البكالوريوس", "الأنظمة", "الأنظمة"),
    ("p02", "البكالوريوس", "الدراسات الإسلامية", "الثقافة الإسلامية"),
    ("p03", "البكالوريوس", "الشريعة", "الشريعة"),
    ("p04", "البكالوريوس", "القرآن وعلومه", "القراءات"),
    ("p05", "البكالوريوس", "القراءات", "القراءات"),
    ("p06", "الماجستير", "القانون", "الأنظمة"),
    ("p07", "الماجستير", "العقيدة", "الثقافة الإسلامية"),
    ("p08", "الماجستير", "أصول الفقه", "الشريعة"),
    ("p09", "الماجستير", "الفقه", "الشريعة"),
    ("p10", "الماجستير", "الدراسات القرآنية المعاصرة", "القراءات"),
    ("p11", "الماجستير", "القراءات", "القراءات"),
    ("p12", "الدكتوراه", "أصول الفقه", "الشريعة"),
    ("p13", "الدكتوراه", "الفقه", "الشريعة"),
    ("p14", "الدكتوراه", "الدراسات القرآنية", "القراءات"),
    ("p15", "الدكتوراه", "القراءات", "القراءات"),
]
PROGRAM_MAP = {(degree, specialty): (pid, department) for pid, degree, specialty, department in PROGRAMS}
CAMPUSES = {"حوية": "male", "حوية طالبات": "female"}
INSTRUMENT_TOPICS = {
    "legacy": {"أسئلة خاصة بـبداية المقرر", "أسئلة خاصة بما حدث خلال المقرر", "تقويم المقرر", "التقويم العام"},
    "electronic_learning": {"أولاً: تنظيم المقرر وتخطيطه:", "ثانياً التواصل والتفاعل:",
        "ثالثاً: عناصر تصميم المقرر ووسائطه المتعددة:", "رابعاً: المواد التعليمية والتقييمات:",
        "خامساً: الرضا العام عن المقرر وعضو هيئة التدريس:"},
}


def instrument(topic):
    for name, topics in INSTRUMENT_TOPICS.items():
        if topic in topics:
            return name
    raise ValueError(f"Unrecognised survey topic {topic!r}; register the source instrument explicitly")


def known_counts(item):
    return item["knownResponseCounts"] if item["responseCounts"] is None else item["responseCounts"]


def text(value):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(value if value is not None else ""))).strip()


def integer(value, *, nullable=False):
    value = text(value).translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    if not value and nullable:
        return None
    try:
        n = Decimal(value)
        if not n.is_finite() or n < 0 or n != int(n):
            raise ValueError
        return int(n)
    except (InvalidOperation, ValueError, OverflowError):
        raise ValueError(f"Expected a nonnegative integer, got {value!r}") from None


def stable_id(parts):
    return hashlib.sha256(json.dumps(parts, ensure_ascii=False, separators=(",", ":")).encode()).hexdigest()[:24]


def add_counts(target, values):
    for i, value in enumerate(values):
        if value is not None:
            target[i] += value


def read_xlsx(source):
    """Resolve sheet names through workbook relationships; stream actual rows."""
    with zipfile.ZipFile(source) as archive:
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            strings = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared = ["".join(t.text or "" for t in si.iter(NS + "t")) for si in strings]
        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        paths = {r.get("Id"): r.get("Target") for r in relations.findall(REL_NS + "Relationship")}
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = workbook.find(NS + "sheets")
        names = [text(s.get("name")) for s in sheets]
        if Counter(names) != Counter(["1445", "1446", "1447"]):
            raise ValueError(f"Expected exactly the 1445, 1446, 1447 sheets; got {names}")
        for sheet in sorted(sheets, key=lambda s: text(s.get("name"))):
            year = text(sheet.get("name"))
            target = paths[sheet.get(DOC_REL_NS + "id")]
            path = target.lstrip("/") if target.startswith("/") else str(PurePosixPath("xl") / target)
            with archive.open(path) as stream:
                root = None
                header_seen = False
                for event, elem in ET.iterparse(stream, events=("start", "end")):
                    if root is None:
                        root = elem
                    if event != "end" or elem.tag != NS + "row":
                        continue
                    number = int(elem.get("r"))
                    values = [None] * 19
                    for cell in elem.findall(NS + "c"):
                        letters = re.match(r"[A-Z]+", cell.get("r")).group()
                        col = 0
                        for letter in letters:
                            col = col * 26 + ord(letter) - 64
                        raw = cell.findtext(NS + "v")
                        kind = cell.get("t")
                        if kind == "s":
                            raw = shared[int(raw)]
                        elif kind == "inlineStr":
                            raw = "".join(t.text or "" for t in cell.iter(NS + "t"))
                        if cell.find(NS + "f") is not None:
                            raise ValueError(f"{year}:{number}: formulas are not supported source values")
                        if col > 19 and text(raw):
                            raise ValueError(f"{year}:{number}: unexpected extra source column")
                        if col <= 19:
                            values[col - 1] = raw
                    elem.clear()
                    root.clear()
                    if not header_seen:
                        if [text(v) for v in values] != HEADERS:
                            raise ValueError(f"{year}:{number}: source headers changed")
                        header_seen = True
                    elif any(text(v) for v in values):
                        yield year, number, values
                if not header_seen:
                    raise ValueError(f"{year}: empty worksheet")


def parse_rows(rows):
    evaluations = {}
    templates = defaultdict(dict)
    controls = {}
    unassigned = []
    lecturer_names = set()
    for year, row_number, raw in rows:
        v = [text(x) for x in raw]
        if len(v) != 19:
            raise ValueError(f"{year}:{row_number}: expected 19 source columns")
        semester = str(integer(v[0]))
        if not re.fullmatch(re.escape(year[-2:]) + r"[123]", semester):
            raise ValueError(f"{year}:{row_number}: semester does not match sheet year")
        identity = (v[1], v[5])
        if identity not in PROGRAM_MAP or v[4] != PROGRAM_MAP[identity][1]:
            raise ValueError(f"{year}:{row_number}: unrecognised degree/specialty/department identity")
        pid = PROGRAM_MAP[identity][0]
        if v[2] not in CAMPUSES:
            raise ValueError(f"{year}:{row_number}: unrecognised campus {v[2]!r}")
        gender = CAMPUSES[v[2]]
        counts = [integer(x, nullable=True) for x in v[14:19]]
        partition = (year, semester, pid, gender)
        control = controls.setdefault(partition, {"sourceRows": 0, "sourceCounts": [0]*5,
            "acceptedRows": 0, "acceptedCounts": [0]*5, "duplicateRowsRemoved": 0,
            "duplicateCountsRemoved": [0]*5, "quarantinedRows": 0, "quarantinedCounts": [0]*5,
            "unassignedRows": 0, "unassignedCounts": [0]*5, "missingCountCells": 0})
        control["sourceRows"] += 1
        control["missingCountCells"] += counts.count(None)
        add_counts(control["sourceCounts"], counts)
        if v[7]:
            lecturer_names.add(v[7])
        if not v[8]:
            control["unassignedRows"] += 1
            add_counts(control["unassignedCounts"], counts)
            unassigned.append({"academicYear": year, "sourceRow": row_number, "semester": semester,
                               "programId": pid, "gender": gender, "reason": "missing_course_name"})
            continue
        if not v[10] or not v[9] or not v[3] or not (v[6] or v[7]):
            raise ValueError(f"{year}:{row_number}: incomplete evaluation identity")
        item_number = integer(v[12])
        if item_number < 1 or not v[11]:
            raise ValueError(f"{year}:{row_number}: incomplete item identity")
        item_key = (v[11], item_number)
        instrument(v[11])
        statement = v[13] or None
        template = templates[(year, semester)]
        if item_key in template and statement and template[item_key] and statement != template[item_key]:
            raise ValueError(f"{year}:{row_number}: conflicting item wording within a period")
        if statement or item_key not in template:
            template[item_key] = statement
        # Both staff columns participate privately in duplicate detection. They
        # do not participate in public course IDs, programme or gender inference.
        evaluation_key = (year, semester, *v[1:11])
        evaluation = evaluations.setdefault(evaluation_key, {
            "partition": partition, "courseName": v[8], "campus": v[2], "college": v[3],
            "courseType": v[9], "section": v[10], "items": defaultdict(list)})
        evaluation["items"][item_key].append({"row": row_number, "counts": counts, "statement": statement})
    return evaluations, templates, controls, unassigned, lecturer_names


def metrics(counts, reasons):
    reasons = sorted(set(reasons))
    responses = sum(counts)
    numerator = sum((i+1)*n for i, n in enumerate(counts))
    return {"status": "partial" if reasons else ("complete" if responses else "no_responses"),
            **({"missingReasons": reasons, "knownResponseCounts": counts} if reasons else {}),
            "responseCounts": None if reasons else counts, "responses": None if reasons else responses,
            "scoreNumerator": None if reasons else numerator,
            "mean": round(numerator/responses, 6) if responses and not reasons else None}


def aggregate(evaluations, templates, controls):
    groups = {}
    issues = []
    for evaluation in evaluations.values():
        year, semester, pid, gender = evaluation["partition"]
        course = evaluation["courseName"]
        group_key = (year, semester, pid, course, gender)
        group = groups.setdefault(group_key, {"items": {}, "sections": set(), "colleges": set(),
            "courseTypes": set(), "evaluationCount": 0, "quarantinedEvaluationCount": 0})
        group["sections"].add((evaluation["campus"], evaluation["courseType"], evaluation["section"]))
        group["colleges"].add(evaluation["college"])
        group["courseTypes"].add(evaluation["courseType"])
        group["evaluationCount"] += 1
        # If any repeated item disagrees, the WHOLE overlapping evaluation is
        # ambiguous, including items whose counts happen to agree by chance.
        conflicting = any(len({(tuple(r["counts"]), r["statement"]) for r in records}) > 1
                          for records in evaluation["items"].values())
        control = controls[evaluation["partition"]]
        if conflicting:
            group["quarantinedEvaluationCount"] += 1
            issues.append({"academicYear": year, "semester": semester, "programId": pid,
                "courseName": course, "gender": gender, "reason": "conflicting_repeated_evaluation",
                "sourceRows": sorted(r["row"] for records in evaluation["items"].values() for r in records)})
        instruments = {instrument(k[0]) for k in evaluation["items"]}
        expected = {k: v for k, v in templates[(year, semester)].items() if instrument(k[0]) in instruments}
        for item_key, template_statement in expected.items():
            item = group["items"].setdefault(item_key, {"counts": [0]*5, "reasons": set(),
                "statements": set(), "missingStatement": False, "sourceRows": 0,
                "acceptedRows": 0, "contributingEvaluations": 0})
            records = evaluation["items"].get(item_key, [])
            item["sourceRows"] += len(records)
            if not records:
                item["reasons"].add("missing_item_in_evaluation")
                continue
            for record in records:
                if record["statement"] is None:
                    item["missingStatement"] = True
                else:
                    item["statements"].add(record["statement"])
            if conflicting:
                item["reasons"].add("ambiguous_source_evaluation")
                for record in records:
                    control["quarantinedRows"] += 1
                    add_counts(control["quarantinedCounts"], record["counts"])
                continue
            first, *duplicates = records
            for duplicate in duplicates:
                control["duplicateRowsRemoved"] += 1
                add_counts(control["duplicateCountsRemoved"], duplicate["counts"])
            control["acceptedRows"] += 1
            add_counts(control["acceptedCounts"], first["counts"])
            add_counts(item["counts"], first["counts"])
            item["acceptedRows"] += 1
            if None in first["counts"]:
                item["reasons"].add("missing_response_count")
            else:
                item["contributingEvaluations"] += 1
        missing = [key for key in expected if key not in evaluation["items"]]
        if missing:
            issues.append({"academicYear": year, "semester": semester, "programId": pid,
                "courseName": course, "gender": gender, "reason": "missing_items_in_evaluation",
                "missingItems": [{"topic": k[0], "number": k[1]} for k in sorted(missing)]})
    return groups, issues


def cohort(group, template, period):
    items = []
    for (topic, number), item in sorted(group["items"].items()):
        statement = None if item["missingStatement"] else next(iter(item["statements"]), None)
        statement_source = "source_rows" if statement else "missing_in_source"
        if not item["sourceRows"]:
            statement = template[(topic, number)]
            statement_source = "period_template"
        items.append({"itemId": stable_id((*period, topic, number)),
            "statementSource": statement_source,
            **metrics(item["counts"], item["reasons"]), "sourceRows": item["sourceRows"],
            "acceptedRows": item["acceptedRows"], "contributingEvaluations": item["contributingEvaluations"]})
    combined = [sum(known_counts(i)[j] for i in items) for j in range(5)]
    reasons = [r for i in items for r in i.get("missingReasons", [])]
    summary = metrics(combined, reasons)
    return {"sectionCount": len(group["sections"]), "sectionEvaluationCount": group["evaluationCount"],
        "quarantinedEvaluationCount": group["quarantinedEvaluationCount"],
        "respondents": None, "respondentCountStatus": "unique_students_not_identifiable",
        "allItemResponseSummary": summary, "items": items}


def combine_groups(groups):
    total = {"items": {}, "sections": set(), "colleges": set(), "courseTypes": set(),
             "evaluationCount": 0, "quarantinedEvaluationCount": 0}
    for group in groups:
        for field in ("sections", "colleges", "courseTypes"):
            total[field].update(group[field])
        for field in ("evaluationCount", "quarantinedEvaluationCount"):
            total[field] += group[field]
        for key, item in group["items"].items():
            target = total["items"].setdefault(key, {"counts": [0]*5, "reasons": set(),
                "statements": set(), "missingStatement": False, "sourceRows": 0,
                "acceptedRows": 0, "contributingEvaluations": 0})
            add_counts(target["counts"], item["counts"])
            target["reasons"].update(item["reasons"])
            target["statements"].update(item["statements"])
            target["missingStatement"] |= item["missingStatement"]
            for field in ("sourceRows", "acceptedRows", "contributingEvaluations"):
                target[field] += item[field]
    return total


def build_payload(rows, source_metadata):
    evaluations, templates, controls, unassigned, names = parse_rows(rows)
    groups, issues = aggregate(evaluations, templates, controls)
    course_keys = sorted({key[:4] for key in groups})
    records = []
    for key in course_keys:
        year, semester, pid, course = key
        available = [gender for gender in ("male", "female") if (*key, gender) in groups]
        available_groups = [groups[(*key, gender)] for gender in available]
        total_group = combine_groups(available_groups)
        cohorts = {gender: cohort(groups[(*key, gender)], templates[(year, semester)], (year, semester))
                   if gender in available else None for gender in ("male", "female")}
        cohorts["total"] = {**cohort(total_group, templates[(year, semester)], (year, semester)),
            "source": "sum_of_available_cohort_item_counts", "includedCohorts": available,
            "missingCohorts": [g for g in ("male", "female") if g not in available],
            "coverage": "both_cohorts" if len(available) == 2 else "available_cohort_only"}
        cohort_item_ids = {g: {i["itemId"] for i in cohorts[g]["items"]} if cohorts[g] else set()
                           for g in ("male", "female")}
        for item in cohorts["total"]["items"]:
            item["includedCohorts"] = [g for g in ("male", "female") if item["itemId"] in cohort_item_ids[g]]
        records.append({"id": stable_id(key), "academicYear": year, "semester": semester,
            "semesterNumber": int(semester[-1]), "programId": pid, "courseName": course,
            "courseCode": None, "courseCodeStatus": "not_in_source", "sourceSheet": year,
            "colleges": sorted(total_group["colleges"]), "courseTypes": sorted(total_group["courseTypes"]),
            "cohorts": cohorts})
    published = defaultdict(lambda: [0]*5)
    published_rows = Counter()
    for record in records:
        for gender in ("male", "female"):
            c = record["cohorts"][gender]
            if c is not None:
                partition = (record["academicYear"], record["semester"], record["programId"], gender)
                for item in c["items"]:
                    add_counts(published[partition], known_counts(item))
                    published_rows[partition] += item["acceptedRows"]
    reconciliations = []
    for partition, control in sorted(controls.items()):
        if control["acceptedCounts"] != published[partition] or control["acceptedRows"] != published_rows[partition]:
            raise ValueError("Source-to-published partition reconciliation failed")
        accounted = [0]*5
        for field in ("acceptedCounts", "duplicateCountsRemoved", "quarantinedCounts", "unassignedCounts"):
            add_counts(accounted, control[field])
        if accounted != control["sourceCounts"] or control["sourceRows"] != sum(control[f] for f in
                ("acceptedRows", "duplicateRowsRemoved", "quarantinedRows", "unassignedRows")):
            raise ValueError("Source accounting failed")
        year, semester, pid, gender = partition
        reconciliations.append({"academicYear": year, "semester": semester, "programId": pid,
            "gender": gender, **control, "publishedKnownCounts": published[partition], "passed": True})
    totals = {field: sum(c[field] for c in controls.values()) for field in
              ("sourceRows", "acceptedRows", "duplicateRowsRemoved", "quarantinedRows", "unassignedRows", "missingCountCells")}
    for field in ("sourceCounts", "acceptedCounts", "duplicateCountsRemoved", "quarantinedCounts", "unassignedCounts"):
        totals[field] = [sum(c[field][i] for c in controls.values()) for i in range(5)]
    payload = {"schemaVersion": "course-student-evaluations-v2", "source": source_metadata,
        "scale": [{"score": i+1, "label": label} for i, label in enumerate(HEADERS[14:19])],
        "programs": [{"id": pid, "degree": degree, "specialty": specialty, "sourceDepartment": dept}
                     for pid, degree, specialty, dept in PROGRAMS],
        "itemDefinitions": {stable_id((year, semester, topic, number)):
            {"academicYear": year, "semester": semester, "instrument": instrument(topic),
             "topic": topic, "number": number, "statement": statement}
            for (year, semester), template in sorted(templates.items())
            for (topic, number), statement in sorted(template.items())},
        "methodology": {
            "genderMapping": CAMPUSES,
            "itemMean": "sum(score * response_count) / sum(response_count); sum counts across evaluations first",
            "total": "sum available male/female item counts, then recompute mean; no imputation",
            "summary": "all-item response-weighted mean, not a learning-domain or satisfaction-only measure",
            "studentCount": "not identifiable; respondents is null; responses are answers to ONE item",
            "courseIdentity": "exact source course name (NFC and whitespace only) within degree/specialty/year/semester; no guessed code",
            "duplicatePolicy": "remove identical repeated items only if the evaluation has no conflicting repeated items; quarantine entire evaluation on any conflict",
            "missingPolicy": "missing cohorts stay null; partial item counts/means stay null; knownResponseCounts retain accepted observed counts",
            "itemCoverage": "expected items are source topic/number identities for the observed instrument(s) in each year/semester; missing items flag partial coverage; different instruments remain distinct",
            "privacy": "only course aggregates; no lecturer names, numbers, individual results or private identity hashes"},
        "quality": {**totals, "courseSemesterRecords": len(records),
            "repeatedItemRows": sum(len(rows)-1 for e in evaluations.values() for rows in e["items"].values()),
            "quarantinedEvaluations": sum(g["quarantinedEvaluationCount"] for g in groups.values()),
            "recordIdsUnique": len({r["id"] for r in records}) == len(records),
            "reconciliationPassed": True, "reconciliations": reconciliations,
            "issues": sorted(issues + unassigned, key=lambda x: json.dumps(x, ensure_ascii=False, sort_keys=True))},
        "records": records}
    # Check all public string values once (including labels, metadata and issue
    # text), rather than repeatedly scanning the large numeric JSON download.
    public_strings = set()
    def collect_strings(value):
        if isinstance(value, str):
            public_strings.add(value)
        elif isinstance(value, dict):
            for child in value.values():
                collect_strings(child)
        elif isinstance(value, list):
            for child in value:
                collect_strings(child)
    collect_strings(payload)
    serialized = "\n".join(sorted(public_strings))
    for name in names:
        if name in serialized:
            raise ValueError("Lecturer name leaked into public output")
    if not payload["quality"]["recordIdsUnique"]:
        raise ValueError("Duplicate public record IDs")
    return payload


def serialize(payload):
    # Compact JSON limits download size; deterministic bytes permit --check.
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"), allow_nan=False) + "\n"


def build(source):
    source = Path(source)
    digest = hashlib.sha256()
    with source.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024*1024), b""):
            digest.update(chunk)
    metadata = {"filename": source.name, "sha256": digest.hexdigest(), "bytes": source.stat().st_size,
                "sheets": ["1445", "1446", "1447"], "format": "original_xlsx"}
    return build_payload(read_xlsx(source), metadata)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--check", action="store_true", help="Verify committed JSON matches source without rewriting it")
    args = parser.parse_args()
    payload = build(args.source)
    result = serialize(payload)
    if args.check:
        if not args.output.is_file() or args.output.read_text(encoding="utf-8") != result:
            raise SystemExit("Published JSON differs from source; rebuild required")
    else:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        temporary = args.output.with_suffix(".json.tmp")
        temporary.write_text(result, encoding="utf-8")
        temporary.replace(args.output)
        if args.output.resolve() == DEFAULT_OUTPUT.resolve():
            # Every published refresh reclassifies items through item-classification.json;
            # a genuinely new statement stops here with the nearest classified item suggested.
            import build_course_domain_scores
            domains = build_course_domain_scores.build_payload(
                payload, json.loads(build_course_domain_scores.CLASSIFICATION.read_text()),
                build_course_domain_scores.CLASSIFICATION.read_bytes())
            build_course_domain_scores.OUTPUT.write_text(build_course_domain_scores.serialize(domains))
    print(json.dumps({"output": str(args.output), **{k: v for k, v in payload["quality"].items()
          if k not in ("issues", "reconciliations")}}, ensure_ascii=False))


if __name__ == "__main__":
    main()
