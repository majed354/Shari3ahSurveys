import hashlib
import json
from pathlib import Path
import tempfile
import unittest
from xml.etree import ElementTree as ET
import zipfile

from scripts import build_student_course_evaluations as api
from scripts.build_survey_data import PROGRAM_ID_MAP

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "tests/fixtures/student-course-samples-1447.json"


def row(number=2, counts=None, campus="حوية", semester="471", course="مقرر تجريبي",
        specialty="القرآن وعلومه", degree="البكالوريوس", department="القراءات",
        staff_id="99", staff_name="محاضر خاص للاختبار", section="1", item=1,
        topic="أولاً: تنظيم المقرر وتخطيطه:", statement="عبارة تجريبية"):
    year = "14" + semester[:2]
    return year, number, [semester, degree, campus, "كلية الشريعة والأنظمة", department,
        specialty, staff_id, staff_name, course, "نظري", section, topic, item, statement,
        *(counts if counts is not None else [0, 0, 0, 0, 1])]


def payload(rows):
    return api.build_payload(rows, {"filename": "test.xlsx", "sha256": "test"})


class ImportTests(unittest.TestCase):
    def test_sums_bins_before_weighting_and_uses_campus(self):
        d = payload([row(counts=[1, 0, 0, 0, 0]),
            row(3, counts=[0, 0, 0, 0, 9], section="2"),
            row(4, counts=[0, 0, 2, 0, 0], campus="حوية طالبات", staff_name="محاضر ذكر تجريبي")])
        r = d["records"][0]
        self.assertEqual(r["cohorts"]["male"]["items"][0]["responseCounts"], [1, 0, 0, 0, 9])
        self.assertEqual(r["cohorts"]["male"]["items"][0]["mean"], 4.6)
        self.assertEqual(r["cohorts"]["female"]["items"][0]["mean"], 3)
        self.assertAlmostEqual(r["cohorts"]["total"]["items"][0]["mean"], 52/12, places=5)
        self.assertEqual(r["cohorts"]["male"]["sectionCount"], 2)
        self.assertNotIn("محاضر ذكر تجريبي", api.serialize(d))

    def test_absent_cohort_and_zero_responses_remain_distinct(self):
        d = payload([row(counts=[0]*5, campus="حوية طالبات")])
        r = d["records"][0]
        self.assertIsNone(r["cohorts"]["male"])
        self.assertEqual(r["cohorts"]["female"]["items"][0]["responses"], 0)
        self.assertIsNone(r["cohorts"]["female"]["items"][0]["mean"])
        self.assertEqual(r["cohorts"]["total"]["missingCohorts"], ["male"])
        self.assertEqual(r["cohorts"]["total"]["coverage"], "available_cohort_only")
        self.assertIsNone(r["cohorts"]["total"]["respondents"])

    def test_identical_duplicates_are_removed(self):
        d = payload([row(), row(3)])
        self.assertEqual(d["quality"]["sourceCounts"], [0, 0, 0, 0, 2])
        self.assertEqual(d["quality"]["acceptedCounts"], [0, 0, 0, 0, 1])
        self.assertEqual(d["quality"]["duplicateRowsRemoved"], 1)
        self.assertEqual(d["records"][0]["cohorts"]["male"]["sectionEvaluationCount"], 1)

    def test_conflicting_duplicate_quarantines_whole_evaluation(self):
        rows = [row(), row(3, counts=[1, 0, 0, 0, 0]), row(4, item=2),
                row(5, item=2), row(6, section="2", item=1), row(7, section="2", item=2)]
        d = payload(rows)
        q = d["quality"]
        self.assertEqual((q["quarantinedRows"], q["acceptedRows"], q["duplicateRowsRemoved"]), (4, 2, 0))
        self.assertEqual(q["quarantinedEvaluations"], 1)
        for c in ("male", "total"):
            for item in d["records"][0]["cohorts"][c]["items"]:
                self.assertIsNone(item["mean"])
                self.assertIsNone(item["responseCounts"])
                self.assertEqual(item["knownResponseCounts"], [0, 0, 0, 0, 1])
        self.assertTrue(all(x["passed"] for x in q["reconciliations"]))

    def test_missing_count_and_missing_item_never_become_zero(self):
        d = payload([row(counts=[None, 0, 0, 0, 1]), row(3, section="2"), row(4, section="2", item=2)])
        items = d["records"][0]["cohorts"]["male"]["items"]
        self.assertEqual(d["quality"]["missingCountCells"], 1)
        self.assertEqual({r for i in items for r in i.get("missingReasons", [])},
                         {"missing_response_count", "missing_item_in_evaluation"})
        self.assertTrue(all(i["mean"] is None for i in items))
        self.assertIsNone(d["records"][0]["cohorts"]["male"]["allItemResponseSummary"]["mean"])

    def test_missing_statement_does_not_invent_text(self):
        d = payload([row(statement="")])
        self.assertIsNone(next(iter(d["itemDefinitions"].values()))["statement"])
        self.assertEqual(d["records"][0]["cohorts"]["male"]["items"][0]["statementSource"], "missing_in_source")

    def test_two_instruments_do_not_create_false_missing_items(self):
        d = payload([row(semester="452", topic="التقويم العام"),
                     row(3, semester="452", section="2"), row(4, semester="452", section="2", item=2)])
        items = d["records"][0]["cohorts"]["male"]["items"]
        self.assertTrue(all(i["status"] == "complete" for i in items))
        self.assertEqual({i["instrument"] for i in d["itemDefinitions"].values()}, {"legacy", "electronic_learning"})

    def test_programme_and_course_scope_are_kept_separate(self):
        d = payload([row(course="التفسير الموضوعي"),
            row(3, course="التفسير الموضوعي", degree="الماجستير", specialty="الدراسات القرآنية المعاصرة"),
            row(4, course="تلاوة وحفظ (6)"), row(5, course="تلاوة وحفظ (7)")])
        self.assertEqual(len(d["records"]), 4)
        self.assertEqual({r["programId"] for r in d["records"]}, {"p04", "p10"})
        self.assertTrue(all(r["courseCode"] is None for r in d["records"]))

    def test_invalid_identities_and_counts_refuse_publication(self):
        invalid = [row(department="الشريعة"), row(specialty="غير معروف"), row(campus="غير معروف"),
                   ("1446", 2, row()[2]), row(counts=[0, 0, -1, 0, 1]), row(counts=[0, 0, .5, 0, 1])]
        for r in invalid:
            with self.subTest(row_number=r[1]), self.assertRaises(ValueError):
                payload([r])

    def test_missing_course_is_accounted_without_guessed_identity(self):
        d = payload([row(course="")])
        self.assertEqual(d["records"], [])
        self.assertEqual(d["quality"]["unassignedRows"], 1)
        self.assertEqual(d["quality"]["unassignedCounts"], [0, 0, 0, 0, 1])

    def test_output_is_deterministic_under_row_order_changes(self):
        rows = [row(section="1"), row(3, section="2", campus="حوية طالبات"), row(4, item=2)]
        self.assertEqual(api.serialize(payload(rows)), api.serialize(payload(list(reversed(rows)))))

    def test_private_attributes_never_appear_in_public_ids_or_fields(self):
        d = payload([row()])
        changed = payload([row(staff_id="123456789", staff_name="اسم آخر خاص")])
        self.assertEqual(d["records"], changed["records"])
        self.assertNotIn("123456789", api.serialize(changed))


def make_workbook(path, bad_header=False):
    """Tiny real XLSX with reordered sheets, arbitrary relationship filenames."""
    workbook = ET.Element(api.NS + "workbook")
    sheets = ET.SubElement(workbook, api.NS + "sheets")
    relations = ET.Element(api.REL_NS + "Relationships")
    with zipfile.ZipFile(path, "w") as z:
        for i, year in enumerate(("1446", "1445", "1447")):
            rid = f"r{i}"
            ET.SubElement(sheets, api.NS + "sheet", {"name": year, "sheetId": str(i+1), api.DOC_REL_NS + "id": rid})
            target = f"worksheets/arbitrary-{i}.xml"
            ET.SubElement(relations, api.REL_NS + "Relationship", {"Id": rid, "Target": target})
            sheet = ET.Element(api.NS + "worksheet")
            data = ET.SubElement(sheet, api.NS + "sheetData")
            headers = list(api.HEADERS)
            if bad_header:
                headers[2] = "مقر مختلف"
            for number, values in ((1, headers), (2, row(semester=year[-2:]+"1")[2])):
                xmlrow = ET.SubElement(data, api.NS + "row", {"r": str(number)})
                for col, value in enumerate(values):
                    cell = ET.SubElement(xmlrow, api.NS + "c", {"r": chr(65+col)+str(number), "t": "inlineStr"})
                    ET.SubElement(ET.SubElement(cell, api.NS + "is"), api.NS + "t").text = str(value)
            z.writestr("xl/"+target, ET.tostring(sheet))
        z.writestr("xl/workbook.xml", ET.tostring(workbook))
        z.writestr("xl/_rels/workbook.xml.rels", ET.tostring(relations))


class WorkbookTests(unittest.TestCase):
    def test_worksheet_names_and_relationships_are_resolved(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "source.xlsx"
            make_workbook(p)
            d = api.build(p)
            self.assertEqual({r["academicYear"] for r in d["records"]}, {"1445", "1446", "1447"})
            self.assertEqual(d["source"]["sha256"], hashlib.sha256(p.read_bytes()).hexdigest())
            make_workbook(p, bad_header=True)
            with self.assertRaisesRegex(ValueError, "headers changed"):
                api.build(p)


class PublishedDataTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.d = json.loads(api.DEFAULT_OUTPUT.read_text(encoding="utf-8"))
        cls.samples = json.loads(FIXTURES.read_text(encoding="utf-8"))

    def test_all_counts_reconcile_by_year_semester_programme_and_cohort(self):
        d = self.d
        self.assertEqual(len({r["id"] for r in d["records"]}), len(d["records"]))
        self.assertEqual({r["academicYear"] for r in d["records"]}, {"1445", "1446", "1447"})
        published = defaultdict_counts()
        for r in d["records"]:
            self.assertIsNone(r["courseCode"])
            for gender in ("male", "female"):
                c = r["cohorts"][gender]
                if c:
                    self.assertIsNone(c["respondents"])
                    for i in c["items"]:
                        key = (r["academicYear"], r["semester"], r["programId"], gender)
                        api.add_counts(published[key], api.known_counts(i))
        for x in d["quality"]["reconciliations"]:
            key = tuple(x[k] for k in ("academicYear", "semester", "programId", "gender"))
            self.assertEqual(x["acceptedCounts"], published[key])
            accounted = [0]*5
            for k in ("acceptedCounts", "duplicateCountsRemoved", "quarantinedCounts", "unassignedCounts"):
                api.add_counts(accounted, x[k])
            self.assertEqual(accounted, x["sourceCounts"])
            self.assertEqual(x["sourceRows"], sum(x[k] for k in
                ("acceptedRows", "duplicateRowsRemoved", "quarantinedRows", "unassignedRows")))

    def test_every_item_mean_and_combined_cohort_is_correct(self):
        definitions = self.d["itemDefinitions"]
        for r in self.d["records"]:
            maps = {g: {i["itemId"]: i for i in c["items"]} if c else {}
                    for g, c in r["cohorts"].items()}
            for g, items in maps.items():
                for key, i in items.items():
                    definition = definitions[key]
                    self.assertEqual((definition["academicYear"], definition["semester"]), (r["academicYear"], r["semester"]))
                    if i["status"] == "partial":
                        self.assertTrue(i["missingReasons"])
                        self.assertTrue(all(i[k] is None for k in ("responseCounts", "mean", "responses", "scoreNumerator")))
                    else:
                        counts = i["responseCounts"]
                        n = sum(counts)
                        num = sum((j+1)*v for j, v in enumerate(counts))
                        self.assertEqual((i["responses"], i["scoreNumerator"]), (n, num))
                        self.assertEqual(i["mean"], round(num/n, 6) if n else None)
                    if g == "total":
                        counts = [0]*5
                        included = []
                        for gender in ("male", "female"):
                            if key in maps[gender]:
                                api.add_counts(counts, api.known_counts(maps[gender][key]))
                                included.append(gender)
                        self.assertEqual(counts, api.known_counts(i))
                        self.assertEqual(included, i["includedCohorts"])

    def test_original_1447_samples_against_independent_excel_extraction(self):
        self.assertEqual(self.samples["sourceSha256"], self.d["source"]["sha256"])
        pids = {(p["degree"], p["specialty"]): p["id"] for p in self.d["programs"]}
        records = {(r["semester"], r["programId"], r["courseName"]): r for r in self.d["records"] if r["academicYear"] == "1447"}
        for sample in self.samples["samples"]:
            key = (sample["semester"], pids[(sample["degree"], sample["specialty"])], sample["courseName"])
            r = records[key]
            for gender, expected in sample["cohorts"].items():
                with self.subTest(key=key, gender=gender):
                    c = r["cohorts"][gender]
                    if expected is None:
                        self.assertIsNone(c)
                        continue
                    summary = c["allItemResponseSummary"]
                    self.assertEqual(summary["responseCounts"], expected["counts"])
                    self.assertEqual(summary["responses"], expected["responses"])
                    self.assertEqual(summary["scoreNumerator"], expected["numerator"])
                    self.assertEqual(summary["mean"], expected["mean"])
                    if "firstItemCounts" in expected:
                        item = next(i for i in c["items"] if self.d["itemDefinitions"][i["itemId"]]["topic"] == "أولاً: تنظيم المقرر وتخطيطه:"
                                    and self.d["itemDefinitions"][i["itemId"]]["number"] == 1)
                        self.assertEqual(item["responseCounts"], expected["firstItemCounts"])

    def test_existing_programme_ids_and_public_field_allowlist(self):
        for p in self.d["programs"]:
            self.assertEqual(PROGRAM_ID_MAP[(p["specialty"], p["degree"])], p["id"])
        allowed = {"id", "academicYear", "semester", "semesterNumber", "programId", "courseName", "courseCode",
                   "courseCodeStatus", "sourceSheet", "colleges", "courseTypes", "cohorts"}
        for r in self.d["records"]:
            self.assertEqual(set(r), allowed)
        serialized = api.serialize(self.d)
        for field in ('"instructorName"', '"instructorId"', '"lecturer"', '"staff_id"', '"اسم المحاضر"', '"رقم المحاضر"'):
            self.assertNotIn(field, serialized)

    @unittest.skipUnless(api.DEFAULT_SOURCE.exists(), "Private original workbook is not present in CI")
    def test_rebuild_matches_committed_bytes(self):
        self.assertEqual(api.serialize(api.build(api.DEFAULT_SOURCE)), api.DEFAULT_OUTPUT.read_text(encoding="utf-8"))


def defaultdict_counts():
    from collections import defaultdict
    return defaultdict(lambda: [0]*5)


if __name__ == "__main__":
    unittest.main()
