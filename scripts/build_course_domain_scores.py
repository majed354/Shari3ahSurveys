"""Build per-course learning-domain and quality scores from the public v2 API.

Every item is classified only through data/course-evaluations/item-classification.json
(semantic map, not keywords). An item absent from the map stops the build.
"""
import argparse
import hashlib
import json
from pathlib import Path
import re
import unicodedata

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "data/course-evaluations/v2/student-course-evaluations.json"
CLASSIFICATION = ROOT / "data/course-evaluations/item-classification.json"
OUTPUT = ROOT / "data/course-evaluations/course-domain-scores.json"
# overall_satisfaction is the "overall" subset of quality; it is scored separately.
SCORED = ("knowledge", "skills", "values", "overall_satisfaction", "quality", "other", "review")
COHORTS = ("total", "male", "female")


def normalized(text):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(text or ""))).strip()


def load_map(classification):
    if classification.get("schemaVersion") != "course-item-classification-v1":
        raise ValueError("Expected course-item-classification-v1")
    allowed = {c["id"] for c in classification["categories"]}
    by_id, by_statement = {}, {}
    for item in classification["items"]:
        if item["category"] not in allowed:
            raise ValueError(f"Unknown category: {item['category']}")
        if item["overall"] and item["category"] != "quality":
            raise ValueError("overall=true is only valid for quality items")
        if (item["category"] == "review") != (item["status"] == "needs_review"):
            raise ValueError("review category and needs_review status must agree")
        key = normalized(item["statement"])
        if key in by_statement:
            raise ValueError(f"Duplicate statement in classification: {key}")
        by_statement[key] = item
        for item_id in item["itemIds"]:
            by_id[item_id] = item
    return by_id, by_statement


def categories_of(item):
    if item["overall"]:
        return ("quality", "overall_satisfaction")
    return (item["category"],)


def build_payload(source, classification):
    if source.get("schemaVersion") != "course-student-evaluations-v2":
        raise ValueError("Expected course-student-evaluations-v2")
    by_id, by_statement = load_map(classification)
    item_category = {}
    for item_id, definition in source["itemDefinitions"].items():
        statement = definition.get("statement")
        if not statement:
            item_category[item_id] = None
            continue
        mapped = by_id.get(item_id) or by_statement.get(normalized(statement))
        if mapped is None or normalized(mapped["statement"]) != normalized(statement):
            raise ValueError(f"Item not classified: {item_id} {statement}")
        item_category[item_id] = categories_of(mapped)

    records = []
    for record in source["records"]:
        cohorts = {}
        for cohort_name in COHORTS:
            cohort = record["cohorts"].get(cohort_name)
            if cohort is None:
                continue
            # [scoreNumerator, responses, itemCount, nonCompleteItems]
            totals = {}
            unnamed = 0
            for item in cohort["items"]:
                categories = item_category[item["itemId"]]
                if categories is None or item.get("statementSource") == "missing_in_source":
                    unnamed += 1
                    continue
                for category in categories:
                    value = totals.setdefault(category, [0, 0, 0, 0])
                    value[0] += item.get("scoreNumerator") or 0
                    value[1] += item.get("responses") or 0
                    value[2] += 1
                    value[3] += item["status"] != "complete"
            cohorts[cohort_name] = {"categories": {c: totals[c] for c in SCORED if c in totals},
                                    **({"unnamedItems": unnamed} if unnamed else {})}
        records.append({"year": record["academicYear"], "semester": record["semester"],
                        "programId": record["programId"], "courseName": record["courseName"],
                        "recordId": record["id"], "cohorts": cohorts})
    records.sort(key=lambda r: (r["year"], r["semester"], r["programId"], r["courseName"]))
    degree_names = {"البكالوريوس": "بكالوريوس", "الماجستير": "ماجستير", "الدكتوراه": "دكتوراه"}
    labels = {c["id"]: c["label"] for c in classification["categories"]}
    return {"schemaVersion": "course-domain-scores-v1", "source": "course-student-evaluations-v2",
            "classificationSha256": hashlib.sha256(CLASSIFICATION.read_bytes()).hexdigest()
            if CLASSIFICATION.exists() else None,
            "tuple": ["scoreNumerator", "responses", "itemCount", "nonCompleteItems"],
            "categories": [{"id": c, "label": "الرضا العام عن المقرر" if c == "overall_satisfaction" else labels[c]}
                           for c in SCORED],
            "scoring": classification["scoring"],
            "programs": [{**p, "name": f"{degree_names[p['degree']]} {p['specialty']}"} for p in source["programs"]],
            "years": sorted({r["academicYear"] for r in source["records"]}, reverse=True),
            "records": records}


def category_score(values):
    """Score of one category after summing tuples over the selected periods."""
    if not values:
        return None
    numerator, responses, items, non_complete = values
    return numerator / responses if items and not non_complete and responses > 0 else None


def serialize(payload):
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"), allow_nan=False) + "\n"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=SOURCE)
    parser.add_argument("--classification", type=Path, default=CLASSIFICATION)
    parser.add_argument("--output", type=Path, default=OUTPUT)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    output = serialize(build_payload(json.loads(args.source.read_text()), json.loads(args.classification.read_text())))
    if args.check:
        if not args.output.exists() or args.output.read_text() != output:
            raise SystemExit("Course domain scores are out of date")
        print("Course domain scores match public v2 source and classification")
    else:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output)
        print(f"Wrote {args.output} ({len(output.encode()):,} bytes)")


if __name__ == "__main__":
    main()
