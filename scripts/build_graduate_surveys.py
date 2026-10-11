#!/usr/bin/env python3
"""Publish aggregated external alumni questions, without individual submissions."""
from __future__ import annotations

import json
from collections import Counter, OrderedDict

try:
    from scripts.build_survey_data import (
        ROOT, GRADUATE_SAMPLE_SOURCES, PROGRAM_DEFINITIONS, clean_text,
        read_first_sheet_rows, resolve_graduate_sample_program_id,
        parse_graduate_year, parse_range_midpoint,
    )
except ModuleNotFoundError:
    from build_survey_data import (
        ROOT, GRADUATE_SAMPLE_SOURCES, PROGRAM_DEFINITIONS, clean_text,
        read_first_sheet_rows, resolve_graduate_sample_program_id,
        parse_graduate_year, parse_range_midpoint,
    )

DESCRIPTION = (
    "أُجريت هذه الاستطلاعات عبر شيت مستقل لجمع استجابات الخريجين، ثم أُدرجت نتائجها هنا. "
    "تعبّر النتائج عن الخريجين المشاركين في الاستطلاع."
)


def question_definition(label, number):
    if "جودة التعلم" in label:
        axis, kind, scale = "جودة التعلم في البرنامج", "rating", 5
    elif "الإشراف" in label:
        axis, kind, scale = "الإشراف العلمي", "rating", 5
    elif "الخدمات" in label:
        axis, kind, scale = "الخدمات المقدمة", "rating", 5
    elif "الاختبارات" in label:
        axis, kind, scale = "الاختبارات الوطنية والمهنية", "range", 100
    elif "ما درجة تقييم" in label:
        axis, kind, scale = "الأداء الوظيفي", "range", 5
    elif "وضعك الحالي" in label or "مجال عملك" in label:
        axis, kind, scale = "الوضع بعد التخرج", "category", None
    else:
        axis, kind, scale = "المقترحات والتعليقات", "text", None
    return {"id": f"q{number}", "number": number, "label": label,
            "axis": axis, "kind": kind, "scale": scale}


def build_external_payload():
    questionnaires, groups = [], []
    for source in GRADUATE_SAMPLE_SOURCES:
        rows = read_first_sheet_rows(source["path"])
        if not rows:
            continue
        headers = rows[0]
        program_column = next(c for c, label in headers.items() if label == "ما اسم البرنامج")
        year_column = next(c for c, label in headers.items() if label == "سنة التخرج من البرنامج الأكاديمي")
        columns = [(c, label) for c, label in headers.items()
                   if c not in (program_column, year_column) and label != "طابع زمني"]
        questions = [question_definition(label, i + 1) for i, (_, label) in enumerate(columns)]
        questionnaires.append({"id": source["track"], "label": "البكالوريوس" if source["track"] == "bachelor" else "الدراسات العليا",
                               "sourceFile": source["path"].name, "questions": questions})
        buckets = OrderedDict()
        for row in rows[1:]:
            program = resolve_graduate_sample_program_id(row.get(program_column), source["track"])
            year = parse_graduate_year(row.get(year_column))
            if not program or not year:
                raise ValueError("Unmapped graduate program/year; do not silently omit a response")
            buckets.setdefault((program, year), []).append(row)
        for (program, year), responses in buckets.items():
            items = []
            for (column, _), question in zip(columns, questions):
                values = [clean_text(row.get(column)) for row in responses if clean_text(row.get(column))]
                counts = Counter(values)
                item = {"questionId": question["id"], "responseCount": len(values)}
                if question["kind"] in ("rating", "range"):
                    numbers = [parse_range_midpoint(v) for v in values]
                    if any(v is None or not 0 <= v <= question["scale"] for v in numbers):
                        raise ValueError(f"Invalid scale in {source['path'].name}, {question['id']}")
                    item["scoreTotal"] = round(sum(numbers), 8)
                # Free comments are counted, never published verbatim. Fields of
                # employment are grouped verbatim without linking to a submission.
                item["categories"] = [] if question["kind"] == "text" else [
                    {"label": label, "count": count} for label, count in sorted(counts.items())]
                items.append(item)
            groups.append({"questionnaireId": source["track"], "programId": program,
                           "year": year, "respondents": len(responses), "items": items})
    program_ids = {g["programId"] for g in groups}
    return {"title": "استطلاعات الخريجين الخارجية", "description": DESCRIPTION,
            "axisNote": "جُمعت الأسئلة تحت محاور لتسهيل العرض؛ الشيت الأصلي لا يتضمن عناوين محاور.",
            "questionnaires": questionnaires, "groups": groups,
            "years": sorted({g["year"] for g in groups}, reverse=True),
            "programs": [{"id": pid, **definition} for pid, definition in PROGRAM_DEFINITIONS.items() if pid in program_ids]}


def main():
    payload = build_external_payload()
    target = ROOT / "js" / "graduate-surveys-data.js"
    target.write_text("window.GRADUATE_SURVEYS_DATA = " + json.dumps(payload, ensure_ascii=False, indent=2) + ";\n", encoding="utf-8")
    print(f"Generated {target.name}: {sum(g['respondents'] for g in payload['groups'])} responses, {len(payload['groups'])} program/year groups")


if __name__ == "__main__":
    main()
