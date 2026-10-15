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
        by_id, by_statement, _ = api.load_map(classification())
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
        self.assertNotIn("proxyCategories", payload["records"][0]["cohorts"]["total"])  # skills item is direct

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
        with self.assertRaisesRegex(ValueError, "بنود جديدة تحتاج تصنيفًا"):
            api.build_payload(data, classification())

    def cases(self):
        return json.loads((api.ROOT / "tests/fixtures/item-matching-cases.json").read_text())

    def test_new_meaning_is_never_matched(self):
        # Independent adversarial cases: negations, المقرر/البرنامج, الطالب/المدرس, other skills or services.
        _, _, keyed = api.load_map(classification())
        for case in self.cases()["newMeaning"]:
            item, method, nearest = api.resolve(case["statement"], keyed, api.equivalents_of(classification()))
            self.assertIsNone(item, case["statement"])
            self.assertEqual(method, "unmatched")

    def test_wording_variants_never_reach_a_wrong_item(self):
        _, _, keyed = api.load_map(classification())
        matched = 0
        for case in self.cases()["variants"]:
            item = api.resolve(case["statement"], keyed, api.equivalents_of(classification()))[0]
            if item is not None:
                self.assertEqual(item["statement"], case["item"], case["statement"])
                matched += 1
        self.assertGreaterEqual(matched, 50)  # the rest stop for a human alias decision

    def test_spelling_variants_match_and_letter_swaps_or_numbers_do_not(self):
        _, _, keyed = api.load_map(classification())
        for variant, original in [
                ("بشكل عام أنا راضٍ عن هذا المقرر", "بشكل عام أنا راضي عن هذا المقرر."),
                ("أشعر بالرضا عن شرح عضو هيئة التدريس في المقرر من خلال التعلم الإلكتروني.", "شعر بالرضا عن شرح عضو هيئة التدريس في المقرر من خلال التعلم الالكتروني."),
                ("تحسنت مهاراتي في استخدام التكنولوجيا والاتصالات من خلال التعلم الإلكتروني", "تحسنت مهارتي في استخدام التكنولوجيا والاتصالات من خلال التعلم الالكتروني."),
                ("أتاحت لي الأنشطة التعليمية و أدوات المقرر فرصة للتفاعل مع المحتوى", "أتاحت لي الأنشطة التعليمية وأدوات المقرر فرصة للتفاعل مع المحتوى."),
                ("شجعني التعلّم الإلكتروني على التعلم المستمر", "شجعني التعلم الالكتروني على التعلم المستمر.")]:
            self.assertEqual(api.resolve(variant, keyed, api.equivalents_of(classification()))[0]["statement"], original, variant)
        for text in self.cases()["edge"]:
            item = api.resolve(text, keyed)[0]
            if item is not None:  # only separator differences such as «بشكل_عام» are equal after normalisation
                self.assertEqual(api.match_key(text), api.match_key(item["statement"]), text)
        for a, b in [("اختبار", "اختيار"), ("غير", "غيرة"), ("مهم", "مهمل"), ("الجيد", "الجديد"), ("بعد", "بعدم"), ("الصفيه", "الصيفيه")]:
            self.assertFalse(api.same_word(a, b, api.equivalents_of(classification())), (a, b))

    def test_human_alias_makes_a_rephrasing_match_exactly(self):
        data = copy.deepcopy(classification())
        rephrased = "ساعدني هذا المقرر على تحسين قدرتي على التواصل بفاعلية."
        _, _, keyed = api.load_map(data)
        self.assertIsNone(api.resolve(rephrased, keyed)[0])
        next(i for i in data["items"] if i["statement"] == "ساعدني هذا المقرر على تحسين قدرتي على الاتصال بفاعلية.")["aliases"] = [rephrased]
        _, _, keyed = api.load_map(data)
        self.assertEqual(api.resolve(rephrased, keyed)[0]["category"], "skills")

    def test_variant_in_source_is_scored_and_reported(self):
        data = fixture()
        data["itemDefinitions"]["s"]["statement"] = "تحسنت مهاراتي في استخدام التكنولوجيا والاتصالات من خلال التعلم الإلكتروني"
        payload = api.build_payload(data, classification())
        self.assertEqual(payload["records"][0]["cohorts"]["total"]["categories"]["skills"], [8, 2, 1, 0, 0])
        self.assertEqual([m["itemId"] for m in payload["itemMatches"]], ["s"])

    def test_no_two_classified_items_are_variants_of_each_other(self):
        items = classification()["items"]
        for n, a in enumerate(items):
            for b in items[n + 1:]:
                self.assertFalse(api.is_variant(api.tokens(a["statement"]), api.tokens(b["statement"])), (a["statement"], b["statement"]))

    def test_dashboard_course_code_is_published_per_semester_then_year(self):
        data = fixture()
        codes = {("1447", "471", "ماجستير الدراسات القرآنية المعاصرة", api.match_key("مقرر")): ("2002851-4", "matched")}
        record = api.build_payload(data, classification(), site_codes=codes)["records"][0]
        self.assertEqual((record["courseCode"], record["courseCodeSource"]), ("2002851-4", "dashboard_matched"))
        codes = {("1447", "all", "ماجستير الدراسات القرآنية المعاصرة", api.match_key("مقرر")): ("2002851-4", "cross-program")}
        self.assertEqual(api.build_payload(data, classification(), site_codes=codes)["records"][0]["courseCode"], "2002851-4")
        self.assertNotIn("courseCode", api.build_payload(data, classification(), site_codes={})["records"][0])

    def test_approved_manual_link_overrides_and_proposed_is_ignored(self):
        data = fixture()
        codes = {("1447", "471", "ماجستير الدراسات القرآنية المعاصرة", api.match_key("مقرر")): ("2002851-4", "matched")}
        record = api.build_payload(data, classification(), site_codes=codes, manual_links={("p10", "مقرر"): "2002999-2"})["records"][0]
        self.assertEqual((record["courseCode"], record["courseCodeSource"]), ("2002999-2", "approved_link"))
        links = json.loads(api.COURSE_LINKS.read_text())
        self.assertTrue(all(link["status"] in ("approved", "proposed") for link in links["links"]))
        approved = {(l["programId"], n) for l in links["links"] if l["status"] == "approved" for n in l["surveyNames"]}
        self.assertEqual(set(api.load_manual_links()), {(p, api.normalized(n)) for p, n in approved})

    def test_published_master_courses_carry_their_codes(self):
        published = json.loads(api.OUTPUT.read_text())
        rows = {r["courseName"]: r.get("courseCode") for r in published["records"] if r["programId"] == "p10" and r["year"] == "1447"}
        self.assertEqual(rows["علم القراءات دراية"], "2002854-2")
        self.assertEqual(rows["مشروع بحثي"], "2002888-4")
        self.assertTrue(all(rows.values()))

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
