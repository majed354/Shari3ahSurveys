import copy
import json
import unittest

from scripts import build_course_domain_scores as api

LEARNING = {"knowledge", "skills", "values"}


def classification():
    return json.loads(api.CLASSIFICATION.read_text())


def fixture():
    definitions = {
        "s": {"academicYear": "1447", "semester": "471", "statement": "تحسنت مهارتي في استخدام التكنولوجيا والاتصالات من خلال التعلم الالكتروني."},
        "o": {"academicYear": "1447", "semester": "471", "statement": "بشكل عام أنا راضي عن هذا المقرر."},
        "p": {"academicYear": "1447", "semester": "471", "statement": "تم عرض خطة المقرر بالتفصيل في بداية الفصل الدراسي."},
        "m": {"academicYear": "1447", "semester": "471", "statement": None},
    }

    def item(item_id, numerator, responses, status="complete", source="source_rows"):
        return {"itemId": item_id, "scoreNumerator": numerator, "responses": responses, "status": status,
                "statementSource": source}

    record = {"id": "r1", "academicYear": "1447", "semester": "471", "programId": "p10", "courseName": "مقرر",
              "cohorts": {"total": {"items": [item("s", 8, 2), item("o", 9, 2), item("p", 6, 2),
                                              item("m", 10, 2, source="missing_in_source")]},
                          "male": None, "female": None}}
    return {"schemaVersion": "course-student-evaluations-v2", "itemDefinitions": definitions, "records": [record],
            "programs": [{"id": "p10", "degree": "الماجستير", "specialty": "الدراسات القرآنية المعاصرة"}]}


class ClassificationMapTests(unittest.TestCase):
    def test_every_published_item_is_classified_once(self):
        source = json.loads(api.SOURCE.read_text())
        by_id, by_statement = api.load_map(classification())
        named = {k for k, v in source["itemDefinitions"].items() if v.get("statement")}
        self.assertEqual(named, set(by_id))
        self.assertEqual(len(by_statement), len(classification()["items"]))

    def test_every_instrument_measures_all_three_domains_and_overall(self):
        # User decision: no learning domain may be empty in any instrument.
        items = classification()["items"]
        for instrument in ("legacy", "electronic_learning"):
            present = {i["category"] for i in items if i["instrument"] == instrument}
            self.assertTrue(LEARNING <= present, instrument)
            self.assertTrue(any(i["overall"] for i in items if i["instrument"] == instrument), instrument)

    def test_direct_and_proxy_learning_items_are_explicit(self):
        items = classification()["items"]
        direct = {i["statement"] for i in items if i.get("measurementBasis") == "direct_gain"}
        self.assertEqual(direct, {
            "ساعدني هذا المقرر على تحسين قدرتي على التفكير وحل المشكلات بدلاً من حفظ المعلومات فقط.",
            "ساعدني هذا المقرر على تحسين قدرتي على الاتصال بفاعلية.",
            "تحسنت مهارتي في استخدام التكنولوجيا والاتصالات من خلال التعلم الالكتروني.",
            "ساعدني هذا المقرر على تحسين مهاراتي في العمل على شكل فريق.",
        })
        proxy = {i["statement"]: i["category"] for i in items if i.get("measurementBasis") == "proxy"}
        self.assertEqual(proxy, {
            "ساعدت الأشياء التي طلبت مني في هذا المقرر (النشطة الصفية، المعامل، وهكذا) في تطوير معرفتي ومهاراتي التي يهدف المقرر لتعليمها.": "knowledge",
            "ما تعلمته في هذا المقرر مهم وسيفيدني مستقبلاً.": "knowledge",
            "أتاحت لي الأنشطة التعليمية وأدوات المقرر فرصة للتفاعل مع المحتوى.": "knowledge",
            "شجعني التعلم الالكتروني على التعلم المستمر.": "values",
            "يتاح لي في التعلم الالكتروني مناقشة زملائي فيما يرتبط بالمقرر وأنشطته.": "values",
        })
        # Teaching, services and satisfaction never become learning items, even as proxies.
        for statement in ("اشعر بالمتعة في الدراسة من خلال التعلم الالكتروني.",
                          "لدى عضو هيئة التدريس الذي يقوم بتقديم هذا المقرر إلمام كامل بمحتوى المقرر.",
                          "تم عرض خطة المقرر بالتفصيل في بداية الفصل الدراسي.",
                          "يتوفر لنا الدعم الفني في التعلم الالكتروني."):
            self.assertNotIn(next(i for i in items if i["statement"] == statement)["category"], LEARNING)

    def test_overall_satisfaction_and_review_items_are_explicit(self):
        items = classification()["items"]
        self.assertEqual({i["statement"] for i in items if i["overall"]},
                         {"أشعر بالرضا بشكل عام عن مستوى جودة هذا المقرر.", "بشكل عام أنا راضي عن هذا المقرر."})
        # Every former review item is decided; unsuitable items stay unclassified.
        self.assertFalse([i for i in items if i["category"] == "review"])
        self.assertEqual(len([i for i in items if "decision" in i]), 3)
        for item in items:
            if item["category"] == "review":
                self.assertEqual(item["status"], "needs_review")
                self.assertIn(item["proposedCategory"], {c["id"] for c in classification()["categories"]})


class DomainScoreTests(unittest.TestCase):
    def test_scores_use_response_counts_and_skip_missing_statements(self):
        payload = api.build_payload(fixture(), classification())
        categories = payload["records"][0]["cohorts"]["total"]["categories"]
        self.assertEqual(categories["skills"], [8, 2, 1, 0, 0])
        self.assertEqual(categories["overall_satisfaction"], [9, 2, 1, 0, 0])
        self.assertEqual(categories["quality"], [15, 4, 2, 0, 0])  # overall item is also a quality item
        self.assertNotIn("values", categories)
        self.assertEqual(payload["records"][0]["cohorts"]["total"]["unnamedItems"], 1)
        self.assertEqual(api.category_score(categories["skills"]), 4)
        self.assertIsNone(api.category_score(categories.get("knowledge")))
        self.assertNotIn("male", payload["records"][0]["cohorts"])

    def test_partial_item_uses_accepted_counts_and_is_flagged(self):
        data = fixture()
        data["records"][0]["cohorts"]["total"]["items"][0].update(
            status="partial", scoreNumerator=None, responses=None, knownResponseCounts=[0, 0, 1, 2, 0])
        categories = api.build_payload(data, classification())["records"][0]["cohorts"]["total"]["categories"]
        self.assertEqual(categories["skills"], [11, 3, 1, 1, 0])
        self.assertAlmostEqual(api.category_score(categories["skills"]), 11 / 3)
        self.assertTrue(api.is_partial(categories["skills"]))
        self.assertFalse(api.is_partial(categories["overall_satisfaction"]))

    def test_only_a_fully_quarantined_category_has_no_score(self):
        data = fixture()
        data["records"][0]["cohorts"]["total"]["items"][0].update(
            status="partial", scoreNumerator=None, responses=None, knownResponseCounts=[0, 0, 0, 0, 0])
        categories = api.build_payload(data, classification())["records"][0]["cohorts"]["total"]["categories"]
        self.assertEqual(categories["skills"], [0, 0, 1, 1, 1])
        self.assertIsNone(api.category_score(categories["skills"]))

    def test_unclassified_statement_stops_the_build(self):
        data = fixture()
        data["itemDefinitions"]["x"] = {"statement": "بند جديد لم يصنف"}
        with self.assertRaises(ValueError):
            api.build_payload(data, classification())

    def test_invalid_map_is_rejected(self):
        bad = copy.deepcopy(classification())
        bad["items"][0]["overall"] = True
        bad["items"][0]["category"] = "skills"
        with self.assertRaises(ValueError):
            api.load_map(bad)

    def test_published_scores_match_public_source(self):
        source = json.loads(api.SOURCE.read_text())
        raw = api.CLASSIFICATION.read_bytes()
        self.assertEqual(api.OUTPUT.read_text(), api.serialize(api.build_payload(source, json.loads(raw), raw)))
        self.assertLess(api.OUTPUT.stat().st_size, 3_000_000)


if __name__ == "__main__":
    unittest.main()
