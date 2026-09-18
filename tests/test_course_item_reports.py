import importlib.util
import json
from pathlib import Path
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("import_items", ROOT / "scripts/import_course_item_reports.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


def row(values):
    return "<tr>" + "".join(f"<td>{values.get(i, '')}</td>" for i in range(46)) + "</tr>"


def report(counts, teacher="عضو اختبار", split=False, missing=False, bad_mean=False):
    n = sum(counts)
    mean = sum((i + 1) * c for i, c in enumerate(counts)) / n
    label = "تحسنت مهارتي في استخدام التكنولوجيا"
    identity = row({0: "القسم :", 1: "القراءات", 2: "عدد المسجلين :", 3: str(n)}) + row({0: "أستاذ المقرر :", 1: teacher, 2: "المقرر :", 3: "التفسير (1)"})
    numeric = {4: "1", 13: str(counts[0]), 19: str(counts[1]), 26: str(counts[2]), 29: str(counts[3]), 34: str(counts[4]), 36: str(n), 38: str(2 if bad_mean else round(mean, 2)), 39: str(round(mean * 20, 1))}
    if not split and not missing:
        numeric[6] = label
    content = identity + row({0: "أولاً: تنظيم المقرر"}) + row({6: "بنود التقييم", 36: "ن", 38: "المتوسطالحسابي"}) + row(numeric)
    if split:
        content += row({5: label})
    # Repeated page header must not create a second report.
    return content + identity + row({0: "المجموع"})


def document(content, semester="الأول"):
    return f"<html><table>{row({0: f'الفصل الدراسي {semester} للعام الجامعي 1447'})}{content}</table></html>".encode("cp1256")


class CourseItemsTest(unittest.TestCase):
    def build(self, body):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        source = Path(tmp.name) / "source.xls"
        source.write_bytes(document(body))
        destination = Path(tmp.name) / "result.json"
        return module.build([source], destination), source, destination

    def test_split_label_period_and_repeated_page_header(self):
        data, _, _ = self.build(report([0, 0, 0, 1, 2], split=True))
        record = data["records"][0]
        self.assertEqual((record["academicYear"], record["semester"], record["reportCount"]), ("1447", "471", 1))
        self.assertEqual(record["items"][0]["statement"], "تحسنت مهارتي في استخدام التكنولوجيا")
        self.assertEqual(record["items"][0]["responseCounts"], [0, 0, 0, 1, 2])
        self.assertIsNone(record["program"])
        self.assertIsNone(record["gender"])

    def test_aggregation_is_weighted_and_has_no_staff_fields(self):
        data, _, destination = self.build(report([0, 0, 0, 1, 2], teacher="أستاذ تجريبي أول") + report([0, 0, 1, 0, 0], teacher="أستاذ تجريبي ثان"))
        record = data["records"][0]
        self.assertEqual(record["reportCount"], 2)
        self.assertEqual(record["respondents"], 4)
        self.assertEqual(record["items"][0]["mean"], 4.25)
        self.assertEqual(data["quality"]["sourceResponseTotal"], data["quality"]["publishedResponseTotal"])
        self.assertNotIn("أستاذ تجريبي", destination.read_text())
        self.assertNotIn("_teacher", destination.read_text())

    def test_unknown_statement_keeps_counts_and_is_not_invented(self):
        data, _, _ = self.build(report([0, 0, 0, 1, 2], missing=True))
        item = data["records"][0]["items"][0]
        self.assertIsNone(item["statement"])
        self.assertEqual(item["statementStatus"], "missing_in_source")
        self.assertEqual(item["responses"], 3)

    def test_inconsistent_mean_refuses_publication(self):
        with self.assertRaisesRegex(ValueError, "printed mean disagrees"):
            self.build(report([0, 0, 0, 1, 2], bad_mean=True))

    def test_duplicate_semester_refuses_double_counting(self):
        _, source, destination = self.build(report([0, 0, 0, 1, 2]))
        before = destination.read_bytes()
        with self.assertRaisesRegex(ValueError, "Duplicate semester"):
            module.build([source, source], destination)
        self.assertEqual(destination.read_bytes(), before)

    def test_published_data_reconciles_counts_and_means(self):
        data = json.loads((ROOT / "data/course-evaluations/1447/course-item-scores.json").read_text())
        self.assertEqual({r["semester"] for r in data["records"]}, {"471", "472"})
        self.assertEqual(len({r["id"] for r in data["records"]}), len(data["records"]))
        total = 0
        for record in data["records"]:
            self.assertTrue(all(record[k] is None for k in ("program", "gender", "courseCode")))
            for item in record["items"]:
                counts = item["responseCounts"]
                self.assertEqual(sum(counts), item["responses"])
                self.assertEqual(sum((i + 1) * n for i, n in enumerate(counts)), item["scoreNumerator"])
                if item["responses"]:
                    self.assertAlmostEqual(item["mean"], item["scoreNumerator"] / item["responses"], places=5)
                total += item["responses"]
        self.assertEqual(total, data["quality"]["sourceResponseTotal"])
        self.assertEqual(total, data["quality"]["publishedResponseTotal"])


if __name__ == "__main__":
    unittest.main()
