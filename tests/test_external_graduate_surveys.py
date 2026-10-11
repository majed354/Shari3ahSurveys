import json
import unittest
from scripts.build_graduate_surveys import build_external_payload, ROOT
from scripts.build_survey_data import GRADUATE_SAMPLE_SOURCES, read_first_sheet_rows


class ExternalGraduateSurveysTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data = build_external_payload()

    def test_every_submission_and_original_question_is_represented(self):
        self.assertEqual(sum(g['respondents'] for g in self.data['groups']), 141)
        self.assertEqual(len(self.data['groups']), 25)
        for source, questionnaire in zip(GRADUATE_SAMPLE_SOURCES, self.data['questionnaires']):
            rows = read_first_sheet_rows(source['path'])
            labels = [v for v in rows[0].values() if v not in ('طابع زمني', 'ما اسم البرنامج', 'سنة التخرج من البرنامج الأكاديمي')]
            self.assertEqual([q['label'] for q in questionnaire['questions']], labels)
            self.assertEqual(sum(g['respondents'] for g in self.data['groups'] if g['questionnaireId'] == questionnaire['id']), len(rows) - 1)

    def test_question_totals_match_source_and_empty_answers_are_excluded(self):
        for source, questionnaire in zip(GRADUATE_SAMPLE_SOURCES, self.data['questionnaires']):
            rows = read_first_sheet_rows(source['path'])
            for q in questionnaire['questions']:
                col = next(c for c, label in rows[0].items() if label == q['label'])
                expected = sum(bool(row.get(col)) for row in rows[1:])
                items = [i for g in self.data['groups'] if g['questionnaireId'] == questionnaire['id'] for i in g['items'] if i['questionId'] == q['id']]
                self.assertEqual(sum(i['responseCount'] for i in items), expected)
                if q['kind'] != 'text':
                    self.assertEqual(sum(c['count'] for i in items for c in i['categories']), expected)
        bachelor = [g for g in self.data['groups'] if g['questionnaireId'] == 'bachelor']
        self.assertEqual(sum(next(i['scoreTotal'] for i in g['items'] if i['questionId'] == 'q1') for g in bachelor), 249)

    def test_no_submission_timestamps_or_free_comments_are_published(self):
        encoded = json.dumps(self.data, ensure_ascii=False)
        self.assertNotIn('طابع زمني', encoded)
        for questionnaire in self.data['questionnaires']:
            text_ids = {q['id'] for q in questionnaire['questions'] if q['kind'] == 'text'}
            for group in self.data['groups']:
                if group['questionnaireId'] == questionnaire['id']:
                    self.assertTrue(all(not i['categories'] for i in group['items'] if i['questionId'] in text_ids))

    def test_published_aggregate_matches_workbooks(self):
        source = (ROOT / 'js/graduate-surveys-data.js').read_text()
        self.assertEqual(json.loads(source.split(' = ', 1)[1].rstrip(';\n')), self.data)
