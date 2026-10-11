"""Build lightweight, equal-course item statistics from the public v2 API."""
import argparse
from collections import defaultdict
import hashlib
import json
import math
from pathlib import Path
import re
import unicodedata

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "data/course-evaluations/v2/student-course-evaluations.json"
OUTPUT = ROOT / "data/course-evaluations/item-quality.json"


def normalized(text):
    # Keep wording, Arabic letters and instrument distinctions intact.
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(text or ""))).strip()


def build_payload(source):
    if source.get("schemaVersion") != "course-student-evaluations-v2":
        raise ValueError("Expected course-student-evaluations-v2")
    definitions = {}
    item_ids = {}
    for source_id, definition in source["itemDefinitions"].items():
        if not definition.get("statement"):
            item_ids[source_id] = None
            continue
        identity = [definition["instrument"], normalized(definition["topic"]),
                    definition["number"], normalized(definition["statement"])]
        item_id = hashlib.sha256(json.dumps(identity, ensure_ascii=False).encode()).hexdigest()[:24]
        item_ids[source_id] = item_id
        definitions[item_id] = dict(zip(("instrument", "topic", "number", "statement"), identity))

    # Pool periods within the same course BEFORE averaging across courses.
    courses = {}
    for record in source["records"]:
        for cohort_name in ("total", "male", "female"):
            cohort = record["cohorts"].get(cohort_name)
            if cohort is None:
                continue
            for semester in (record["semester"], "all"):
                key = (record["academicYear"], semester, record["programId"], cohort_name,
                       normalized(record["courseName"]))
                course = courses.setdefault(key, {"items": {}, "unnamed": 0})
                for item in cohort["items"]:
                    item_id = item_ids[item["itemId"]]
                    if item_id is None:
                        course["unnamed"] += 1
                        continue
                    value = course["items"].setdefault(item_id, {"numerator": 0, "responses": 0, "partial": False})
                    if item["status"] == "partial" or item.get("statementSource") == "missing_in_source":
                        value["partial"] = True
                    elif item["status"] == "complete" and item.get("responses", 0) > 0:
                        value["numerator"] += item["scoreNumerator"]
                        value["responses"] += item["responses"]

    groups = {}
    for key, course in sorted(courses.items()):
        group_key = key[:4]
        group = groups.setdefault(group_key, {"courseCount": 0, "unnamedItemObservations": 0, "items": {}})
        group["courseCount"] += 1
        group["unnamedItemObservations"] += course["unnamed"]
        for item_id, value in course["items"].items():
            stats = group["items"].setdefault(item_id, {"means": [], "responseCount": 0, "excludedCourseCount": 0})
            if value["partial"] or not value["responses"]:
                stats["excludedCourseCount"] += 1
            else:
                stats["means"].append(value["numerator"] / value["responses"])
                stats["responseCount"] += value["responses"]

    records = []
    for (year, semester, program_id, cohort), group in sorted(groups.items()):
        items = {}
        for item_id, stats in sorted(group["items"].items()):
            items[item_id] = {"courseMeanSum": math.fsum(stats["means"]), "courseCount": len(stats["means"]),
                              "responseCount": stats["responseCount"], "excludedCourseCount": stats["excludedCourseCount"]}
        records.append({"year": year, "semester": semester, "programId": program_id, "cohort": cohort,
                        **group, "items": items})
    degree_names = {"البكالوريوس": "بكالوريوس", "الماجستير": "ماجستير", "الدكتوراه": "دكتوراه"}
    return {"schemaVersion": "course-item-quality-v1", "source": "course-student-evaluations-v2",
            "weighting": "equal courses; pool item responses within a course across selected semesters first",
            "programs": [{**p, "name": f"{degree_names[p['degree']]} {p['specialty']}"} for p in source["programs"]],
            "years": sorted({r["academicYear"] for r in source["records"]}, reverse=True),
            "itemDefinitions": definitions, "records": records}


def serialize(payload):
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"), allow_nan=False) + "\n"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=SOURCE)
    parser.add_argument("--output", type=Path, default=OUTPUT)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    output = serialize(build_payload(json.loads(args.source.read_text())))
    if args.check:
        if not args.output.exists() or args.output.read_text() != output:
            raise SystemExit("Item quality data is out of date")
        print("Item quality data matches public v2 source")
    else:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output)
        print(f"Wrote {args.output} ({len(output.encode()):,} bytes)")


if __name__ == "__main__":
    main()
