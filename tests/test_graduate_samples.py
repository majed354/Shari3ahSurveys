#!/usr/bin/env python3

import unittest

from scripts.build_survey_data import PROGRAM_DEFINITIONS, build_payload


class GraduateSampleImportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.payload = build_payload()
        cls.samples = cls.payload["graduateSampleKpis"]

    def test_all_program_year_groups_are_imported(self):
        self.assertEqual(len(self.samples), 25)
        bachelor = [
            entry
            for entry in self.samples.values()
            if PROGRAM_DEFINITIONS[entry["programId"]]["degree"] == "البكالوريوس"
        ]
        postgraduate = [entry for entry in self.samples.values() if entry not in bachelor]
        self.assertEqual(len(bachelor), 10)
        self.assertEqual(len(postgraduate), 15)

    def test_only_complementary_kpis_are_exposed(self):
        bachelor_allowed = {"performance_rate", "employment_rate", "eval_employers"}
        postgraduate_allowed = {"eval_supervision", "eval_services", "eval_employers"}

        for entry in self.samples.values():
            degree = PROGRAM_DEFINITIONS[entry["programId"]]["degree"]
            allowed = bachelor_allowed if degree == "البكالوريوس" else postgraduate_allowed
            self.assertTrue(set(entry["metrics"]).issubset(allowed))
            self.assertNotIn("eval_experience", entry["metrics"])
            self.assertNotIn("eval_courses", entry["metrics"])

    def test_every_graduate_result_is_marked_as_sample(self):
        for entry in self.samples.values():
            self.assertEqual(entry["sourceType"], "sample")
            self.assertEqual(entry["sourceLabel"], "استطلاع عينة")
            for metric in entry["metrics"].values():
                self.assertEqual(metric["sourceType"], "sample")
                self.assertEqual(metric["sourceLabel"], "استطلاع عينة")
                self.assertGreater(metric["sampleCount"], 0)

        sample_surveys = [
            survey
            for dataset in self.payload["extractedData"].values()
            for survey in dataset["surveys"]
            if survey.get("surveyType") == "sample"
        ]
        self.assertEqual(len(sample_surveys), 49)
        self.assertTrue(all("استطلاع عينة" in survey["title"] for survey in sample_surveys))

    def test_existing_surveys_are_not_relabelled(self):
        original_surveys = [
            survey
            for dataset in self.payload["extractedData"].values()
            for survey in dataset["surveys"]
            if survey.get("surveyType") != "sample"
        ]
        self.assertTrue(original_surveys)
        self.assertTrue(all("استطلاع شامل" not in survey["title"] for survey in original_surveys))

    def test_known_bachelor_result(self):
        metrics = self.samples["p01::1445"]["metrics"]
        self.assertAlmostEqual(metrics["performance_rate"]["value"], 81.1667, places=4)
        self.assertEqual(metrics["performance_rate"]["sampleCount"], 6)
        self.assertAlmostEqual(metrics["employment_rate"]["value"], 42.8571, places=4)
        self.assertEqual(metrics["employment_rate"]["positiveCount"], 3)
        self.assertEqual(metrics["employment_rate"]["sampleCount"], 7)


if __name__ == "__main__":
    unittest.main()
