import json
from pathlib import Path
import unittest

from scripts import build_item_quality_data as api


def fixture():
    return {"schemaVersion": "course-student-evaluations-v2",
            "programs": [{"id": "p01", "degree": "البكالوريوس", "specialty": "الشريعة"}],
            "itemDefinitions": {
                "first": {"instrument": "electronic_learning", "topic": "أولاً", "number": 1, "statement": "بند ثابت"},
                "second": {"instrument": "electronic_learning", "topic": "أولاً", "number": 1, "statement": "بند ثابت"},
            }, "records": []}


def record(course, semester="471", numerator=5, responses=1, status="complete", source="source_rows", item_id="first"):
    item = {"itemId": item_id, "status": status, "statementSource": source,
            "scoreNumerator": numerator, "responses": responses}
    return {"academicYear": "14" + semester[:2], "semester": semester, "programId": "p01", "courseName": course,
            "cohorts": {"total": {"items": [item]}, "male": {"items": [item]}, "female": None}}


class ItemQualityTests(unittest.TestCase):
    def stats(self, payload, semester="all", year="1447", cohort="total"):
        group = next(r for r in payload["records"] if r["semester"] == semester and r["year"] == year and r["cohort"] == cohort)
        return group, next(iter(group["items"].values()))

    def test_equal_courses_not_equal_responses_and_annual_course_not_counted_twice(self):
        data = fixture()
        data["records"] = [record("أ", numerator=5), record("أ", "472", numerator=27, responses=9, item_id="second"),
                           record("ب", numerator=200, responses=100)]
        payload = api.build_payload(data)
        group, stats = self.stats(payload)
        self.assertEqual(group["courseCount"], 2)
        # Course A: (5 + 27)/(1 + 9) = 3.2; B = 2; their mean = 2.6.
        self.assertAlmostEqual(stats["courseMeanSum"] / stats["courseCount"], 2.6)
        self.assertEqual(stats["responseCount"], 110)
        self.assertEqual(self.stats(payload, "471")[1]["courseMeanSum"] / 2, 3.5)
        self.assertFalse(any(r["cohort"] == "female" for r in payload["records"]))

    def test_partial_or_missing_statement_excludes_course_item_without_zero_imputation(self):
        for field in ("partial", "missing_in_source"):
            data = fixture()
            data["records"] = [record("أ"), record("أ", "472", status="partial" if field == "partial" else "complete",
                source="missing_in_source" if field == "missing_in_source" else "source_rows"), record("ب", numerator=4)]
            _, stats = self.stats(api.build_payload(data))
            self.assertEqual(stats["courseMeanSum"], 4)
            self.assertEqual(stats["courseCount"], 1)
            self.assertEqual(stats["excludedCourseCount"], 1)

    def test_same_number_is_not_identity_when_wording_or_instrument_differs(self):
        for field, value in (("statement", "بند مختلف"), ("instrument", "legacy")):
            data = fixture()
            data["itemDefinitions"]["second"][field] = value
            data["records"] = [record("أ"), record("أ", "472", item_id="second")]
            payload = api.build_payload(data)
            self.assertEqual(len(payload["itemDefinitions"]), 2)
            self.assertEqual(len(self.stats(payload)[0]["items"]), 2)

    def test_null_statement_is_never_invented_and_zero_answers_are_not_zero_score(self):
        data = fixture()
        data["itemDefinitions"]["second"]["statement"] = None
        data["records"] = [record("أ", numerator=0, responses=0, status="no_responses"), record("ب", item_id="second")]
        group, stats = self.stats(api.build_payload(data))
        self.assertEqual(group["unnamedItemObservations"], 1)
        self.assertEqual(stats["courseCount"], 0)
        self.assertEqual(stats["excludedCourseCount"], 1)

    def test_published_summary_matches_public_source(self):
        source = json.loads(api.SOURCE.read_text())
        self.assertEqual(api.OUTPUT.read_text(), api.serialize(api.build_payload(source)))
        self.assertLess(api.OUTPUT.stat().st_size, 2_000_000)


if __name__ == "__main__":
    unittest.main()
