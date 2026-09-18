#!/usr/bin/env python3
"""Publish anonymous course aggregates from the university HTML/XLS reports.

Raw files stay outside the repository. Identity is department/course/semester,
not a guessed programme, campus or student gender. Uses only the standard library.
"""
from __future__ import annotations

import argparse
import hashlib
from html.parser import HTMLParser
from html import unescape
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "data/course-evaluations/1447/course-item-scores.json"


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


class RowParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.cells = []
        self.cell = None
        self.span = 1

    def handle_starttag(self, tag, attrs):
        if tag.lower() in ("td", "th"):
            if self.cell is not None:
                raise ValueError("Nested cells in source")
            self.span = max(1, int(dict(attrs).get("colspan", "1")))
            self.cell = []

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)

    def handle_endtag(self, tag):
        if tag.lower() in ("td", "th") and self.cell is not None:
            self.cells.append((self.span, clean("".join(self.cell))))
            self.cell = None


def rows(text):
    for row_number, match in enumerate(re.finditer(r"<tr\b[^>]*>.*?</tr\s*>", text, re.I | re.S), 1):
        parser = RowParser()
        parser.feed(match.group())
        position = 0
        grid = {}
        for span, value in parser.cells:
            if value:
                grid[position] = value
            position += span
        yield row_number, grid


def number(value):
    text = str(value or "").translate(str.maketrans("٠١٢٣٤٥٦٧٨٩٫", "0123456789."))
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        return None
    return float(text)


def topic_order(topic):
    for index, prefix in enumerate(("أولاً", "ثانياً", "ثالثاً", "رابعاً", "خامساً", "سادساً")):
        if topic.startswith(prefix):
            return index
    return 99


def parse_report(path):
    path = Path(path)
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp1256")
    if "<html" not in text[:1000].lower():
        raise ValueError(f"{path.name}: expected university HTML/XLS export")
    period_text = clean(unescape(re.sub(r"<[^>]+>", " ", text[:20000])))
    period = re.search(r"الفصل الدراسي (الأول|الثاني) للعام الجامعي\s*(14\d{2})", period_text)
    if not period:
        raise ValueError(f"{path.name}: explicit academic period is missing")
    year = period.group(2)
    semester = year[-2:] + ("1" if period.group(1) == "الأول" else "2")
    reports = []
    department = course = teacher = topic = ""
    enrolled = None
    items = []
    pending = None
    layout_verified = False

    def finish_item():
        nonlocal pending
        if pending is not None:
            items.append(pending)
            pending = None

    def finish_report():
        nonlocal items, topic
        finish_item()
        if items:
            if not department or not course or not teacher:
                raise ValueError(f"{path.name}: incomplete report identity")
            unique = {}
            for item in items:
                key = (item["topic"], item["number"])
                if key in unique:
                    # A repeated page is allowed only when its values agree.
                    old = unique[key]
                    if (old["counts"], old["label"]) != (item["counts"], item["label"]):
                        raise ValueError(f"{path.name}: contradictory repeated item at row {item['sourceRow']}")
                    continue
                unique[key] = item
            reports.append({"department": department, "courseName": course,
                            "enrolled": enrolled, "items": list(unique.values()),
                            "_teacher": teacher})
        items = []
        topic = ""

    for row_number, grid in rows(text):
        values = list(grid.values())
        joined = " ".join(values)
        if "بنود التقييم" in values:
            if grid.get(38) != "المتوسطالحسابي" or grid.get(36) != "ن":
                raise ValueError(f"{path.name}: unsupported numeric layout at row {row_number}")
            layout_verified = True
        if "القسم :" in values:
            next_department = values[values.index("القسم :") + 1]
            next_enrolled = number(values[values.index("عدد المسجلين :") + 1])
            if next_department != department or next_enrolled != enrolled:
                finish_report()
            department, enrolled = next_department, next_enrolled
        if "المقرر :" in values:
            next_course = values[values.index("المقرر :") + 1]
            next_teacher = values[1]
            if next_course != course or next_teacher != teacher:
                finish_report()
            course, teacher = next_course, next_teacher
        if re.match(r"^(أولاً|ثانياً|ثالثاً|رابعاً|خامساً|سادساً)", joined):
            if joined.startswith("أولاً") and (items or pending is not None):
                finish_report()
            finish_item()
            topic = joined
            continue
        item_number = number(grid.get(4))
        if topic and item_number is not None and item_number.is_integer() and not all(p in grid for p in (13, 19, 26, 29, 34, 36)):
            raise ValueError(f"{path.name}: incomplete response row at {row_number}")
        if item_number is not None and item_number.is_integer() and all(p in grid for p in (13, 19, 26, 29, 34, 36)):
            finish_item()
            if not layout_verified or not topic:
                raise ValueError(f"{path.name}: item precedes layout/topic at row {row_number}")
            counts = [number(grid[p]) for p in (13, 19, 26, 29, 34)]
            count = number(grid[36])
            if any(n is None or not n.is_integer() for n in counts) or count is None or sum(counts) != count:
                raise ValueError(f"{path.name}: response counts disagree at row {row_number}")
            exact = sum((i + 1) * n for i, n in enumerate(counts)) / count if count else None
            printed = number(grid.get(38))
            percentage = number(grid.get(39))
            if exact is not None and printed is not None and abs(exact - printed) > .011:
                raise ValueError(f"{path.name}: printed mean disagrees at row {row_number}")
            # Published percentage is often calculated from the rounded mean.
            if exact is not None and percentage is not None and abs(exact * 20 - percentage) > .23:
                raise ValueError(f"{path.name}: printed percentage disagrees at row {row_number}")
            label = grid.get(6, "")
            pending = {"topic": topic, "number": int(item_number), "label": label,
                       "counts": [int(n) for n in counts], "respondents": int(count),
                       "sourceRow": row_number, "printedMean": printed}
        elif pending is not None and not pending["label"]:
            # Some labels occupy a separate physical row in this export.
            labels = [v for p, v in grid.items() if 5 <= p <= 11 and re.search(r"[\u0621-\u064a]", v)]
            if labels and not any(p >= 13 for p in grid):
                pending["label"] = clean(" ".join(labels))
                pending["labelSourceRow"] = row_number
        if "المجموع" in values:
            finish_item()
    finish_report()
    if not reports:
        raise ValueError(f"{path.name}: no course reports found")
    return {"id": f"survey-{semester}", "filename": path.name, "academicYear": year,
            "semester": semester, "sha256": hashlib.sha256(raw).hexdigest(),
            "bytes": len(raw)}, reports


def aggregate(sources):
    records = {}
    for source, reports in sources:
        for report in reports:
            key = (source["academicYear"], source["semester"], report["department"], report["courseName"])
            group = records.setdefault(key, {"sourceId": source["id"], "reports": [], "items": {}})
            group["reports"].append(report)
            for item in report["items"]:
                item_key = (item["topic"], item["number"], item["label"])
                total = group["items"].setdefault(item_key, {"counts": [0] * 5, "reportCount": 0})
                total["counts"] = [a + b for a, b in zip(total["counts"], item["counts"])]
                total["reportCount"] += 1
    output = []
    for key, group in sorted(records.items()):
        year, semester, department, course = key
        items = []
        for (topic, num, label), item in sorted(group["items"].items(), key=lambda pair: (topic_order(pair[0][0]), pair[0][1], pair[0][2])):
            counts = item["counts"]
            n = sum(counts)
            items.append({"topic": topic, "number": num, "statement": label or None,
                          "statementStatus": "present" if label else "missing_in_source",
                          "responseCounts": counts, "responses": n,
                          "scoreNumerator": sum((i + 1) * v for i, v in enumerate(counts)),
                          "mean": round(sum((i + 1) * v for i, v in enumerate(counts)) / n, 6) if n else None,
                          "reportCount": item["reportCount"]})
        sizes = [{i["respondents"] for i in report["items"] if i["respondents"]} for report in group["reports"]]
        respondents = sum(next(iter(s)) for s in sizes) if all(len(s) == 1 for s in sizes) else None
        output.append({"id": hashlib.sha256("|".join(key).encode()).hexdigest()[:20],
                       "academicYear": year, "semester": semester, "department": department,
                       "courseName": course, "courseCode": None, "program": None, "gender": None,
                       "scope": "department_course", "sourceId": group["sourceId"],
                       "reportCount": len(group["reports"]), "respondents": respondents,
                       "respondentCountStatus": "consistent_within_reports" if respondents is not None else "varies_by_item",
                       "items": items})
    return output


def build(paths, destination):
    parsed = [parse_report(p) for p in paths]
    ids = [s[0]["id"] for s in parsed]
    if len(ids) != len(set(ids)):
        raise ValueError("Duplicate semester source; refusing double counting")
    records = aggregate(parsed)
    source_items = sum(len(r["items"]) for _, reports in parsed for r in reports)
    source_responses = sum(sum(i["counts"]) for _, reports in parsed for r in reports for i in r["items"])
    published_responses = sum(i["responses"] for r in records for i in r["items"])
    assert source_responses == published_responses
    result = {"schemaVersion": "course-item-evaluations-v1", "sources": [s for s, _ in parsed],
              "scope": "department_course", "scale": {"min": 1, "max": 5},
              "method": "تجميع أعداد الاستجابات لكل عبارة داخل المقرر والقسم والفصل؛ المتوسط = مجموع (الدرجة × عدد الإجابات) ÷ عدد الإجابات. أعداد الإجابات عبر العبارات ليست أعداد طلاب فريدين.",
              "identityNotes": "المصدر لا يحدد رمز المقرر أو البرنامج أو الشطر؛ تبقى هذه الحقول فارغة ولا تُستنتج من اسم الأستاذ أو القسم. لا يُستبدل به تلقائيًا قياس برنامج أو شطر محدد.",
              "privacy": "تجميع على مستوى المقرر، بلا أسماء أعضاء هيئة التدريس أو معرفاتهم أو نتائجهم الفردية.",
              "quality": {"sourceReports": sum(len(reports) for _, reports in parsed), "sourceItems": source_items,
                          "courseSemesterRecords": len(records), "publishedItems": sum(len(r["items"]) for r in records),
                          "missingStatementItems": sum(i["statement"] is None for r in records for i in r["items"]),
                          "sourceResponseTotal": source_responses, "publishedResponseTotal": published_responses},
              "records": records}
    serialized = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    for _, reports in parsed:
        for r in reports:
            if r["_teacher"] in serialized:
                raise ValueError("Instructor name appeared in public output")
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(serialized, encoding="utf-8")
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("sources", nargs="+", type=Path)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    result = build(args.sources, args.output)
    print(json.dumps({"output": str(args.output), **result["quality"]}, ensure_ascii=False))
