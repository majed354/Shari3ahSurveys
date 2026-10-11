const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const analysis = require('../js/graduate-surveys-analysis.js');
const context = { window: {} };
vm.runInNewContext(fs.readFileSync(require.resolve('../js/graduate-surveys-data.js'), 'utf8'), context);
const data = JSON.parse(JSON.stringify(context.window.GRADUATE_SURVEYS_DATA));

test('all 141 responses and both questionnaires, without combining different questions', () => {
    const report = analysis.buildReport(data, {year: 'all', degree: 'all', program: 'all'});
    assert.equal(report.respondentCount, 141);
    assert.equal(report.rows.length, 13);
    assert.equal(report.sections[0].respondentCount, 62);
    assert.equal(report.sections[1].respondentCount, 79);
    assert.equal(report.sections[0].rows[0].average, 249 / 62);
    const testScores = report.sections[0].rows.find(r => r.scale === 100);
    assert.equal(testScores.responseCount, 33);
    assert.equal(testScores.missingCount, 29);
    assert.ok(Math.abs(testScores.categories.reduce((n, c) => n + c.percentage, 0) - 100) < 1e-9);
});
test('program, degree, graduation year, axis and text filters use matching source groups', () => {
    const report = analysis.buildReport(data, {degree: 'البكالوريوس', program: 'p01', year: '1445'});
    assert.equal(report.respondentCount, 7);
    assert.equal(report.sections.length, 1);
    assert.equal(report.rows.length, 6);
    const axis = analysis.buildReport(data, {axis: 'جودة التعلم في البرنامج', query: 'جودة التعلم'});
    assert.equal(axis.rows.length, 2);
    assert.equal(axis.respondentCount, 141);
    assert.equal(analysis.buildReport(data, {year: '1447'}).rows.length, 0);
    assert.equal(analysis.buildReport(data, {degree: 'الدكتوراه', program: 'p01'}).respondentCount, 0);
    assert.equal(analysis.buildReport(data, {query: 'عبارة غير موجودة'}).rows.length, 0);
});
test('exports preserve typed scores, scale, denominators and filtered distribution', () => {
    const report = analysis.buildReport(data, {degree: 'البكالوريوس', program: 'p01', year: '1445'});
    const rows = analysis.exportRows(report, 'الأنظمة 1445');
    assert.ok(rows.every(r => r['النطاق'] === 'الأنظمة 1445' && r['الاستطلاع'] === 'البكالوريوس'));
    assert.equal(typeof rows[0]['متوسط البند'], 'number');
    assert.equal(rows[0]['المقياس'], 5);
    assert.equal(rows[0]['المشاركون في النطاق'], 7);
    assert.ok(rows.some(r => r['ملاحظة'].includes('منتصف')));
    assert.ok(rows.some(r => r['الإجابة'].includes('العدد فقط')));
});
