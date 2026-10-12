#!/usr/bin/env python3

import unittest

from scripts.build_survey_data import (
    align_legacy_rows,
    create_dataset_store,
    import_legacy_student_surveys,
    import_student_1447_surveys,
)

HEADER = [
    "الفصل الدراسي", "الجنس", "الدرجة العلمية", "المقر", "الكلية", "البرنامج", "التقييم",
    "وصف الموضوع", "رقم العبارة", "وصف العبارة",
    "لا أوافق بشدة", "لا أوافق", "أوافق الى حد ما", "أوافق", "أوافق بشدة",
]


def sheet(header, *rows):
    letters = [chr(ord("A") + index) for index in range(len(header))]
    return [dict(zip(letters, header))] + [dict(zip(letters, row)) for row in rows]


def program_row(program, degree, semester="471", counts=(0, 0, 2, 4, 4)):
    return [
        semester, "ذكر", degree, "حوية", "كلية الشريعة والأنظمة", program, "استبانة تقويم برنامج",
        "التقويم العام", "1", "أشعر بالرضا بشكل عام عن مستوى جودة خبرتي التعليمية في الجامعة",
        *[str(count) for count in counts],
    ]


class AlignLegacyRowsTests(unittest.TestCase):
    def test_columns_are_matched_by_name_not_position(self):
        reordered = list(reversed(HEADER))
        row = list(reversed(program_row("القراءات", "البكالوريوس")))
        aligned = align_legacy_rows(sheet(reordered, row))
        self.assertEqual(aligned[0]["F"], "القراءات")
        self.assertEqual(aligned[0]["A"], "471")
        self.assertEqual(aligned[0]["O"], "4")

    def test_alternative_header_names_are_accepted(self):
        header = [name.replace("البرنامج", "التخصص").replace("لا أوافق بشدة", "غير موافق بشدة") for name in HEADER]
        self.assertEqual(align_legacy_rows(sheet(header, program_row("الشريعة", "البكالوريوس")))[0]["F"], "الشريعة")

    def test_unknown_layout_is_rejected(self):
        self.assertIsNone(align_legacy_rows(sheet(["الجنس", "التخصص", "المتوسط", "الاسئلة"], ["ذكر", "القراءات", "4", "س"])))


class ProgramLevelPrecedenceTests(unittest.TestCase):
    def build(self, program_rows):
        datasets, years, genders, skipped = create_dataset_store()
        coverage = import_legacy_student_surveys(align_legacy_rows(sheet(HEADER, *program_rows)), datasets, years, genders, skipped)
        import_student_1447_surveys(datasets, years, genders, skipped, {}, coverage)
        return datasets, coverage

    @staticmethod
    def measurement_kinds(dataset):
        kinds = set()
        for survey in dataset["_surveys"].values():
            for topic in survey["_topics"].values():
                for item in topic["_items"].values():
                    for entry in item["_genders"].values():
                        kinds.add("responses" if entry["responses"] else "average")
        return kinds

    def test_program_level_1447_export_replaces_college_values_for_that_programme_only(self):
        datasets, coverage = self.build([program_row("القراءات", "البكالوريوس")])
        self.assertIn(("p05", "1447"), coverage)
        self.assertEqual(self.measurement_kinds(datasets["p05::1447"]), {"responses"})
        self.assertEqual(self.measurement_kinds(datasets["p04::1447"]), {"average"})

    def test_without_program_level_files_college_data_is_kept_for_every_programme(self):
        datasets, coverage = self.build([])
        self.assertEqual(coverage, set())
        for program_id in ("p01", "p05", "p15"):
            self.assertEqual(self.measurement_kinds(datasets[f"{program_id}::1447"]), {"average"})


if __name__ == "__main__":
    unittest.main()
