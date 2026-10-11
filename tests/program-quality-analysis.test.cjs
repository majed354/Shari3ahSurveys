const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const { prepareData, buildReport, exportRows } = require("../js/program-quality-analysis.js");
const programs = [{ id: "p01", degree: "بكالوريوس", label: "الشريعة" }, { id: "p02", degree: "الماجستير", label: "الفقه" }];
const entry = (gender, responses, scoreTotal) => ({ gender, responses, scoreTotal });
const item = (number, label, genders) => ({ number, label, genders });
const survey = (items, extras = {}) => ({ title: "استبانة تقويم برنامج", stakeholder: "students", sectionId: "management", topics: [{ label: "الدعم", items }], ...extras });
const dataset = (...surveys) => ({ surveys });
const scope = { year: "1446", program: "p01", degree: "all", cohort: "total" };
function fixture() {
    return { extractedData: {
        "p01::1446": dataset(survey([
            item("1", "وضوح الخطة", [entry("ذكر", 10, 50), entry("إناث", 30, 90)]),
            item("2", "الدعم الأكاديمي", [entry("ذكر", 10, 10), entry("إناث", 30, 90)]),
        ])),
        "p01::1445": dataset(survey([
            item("1", "وضوح الخطة", [entry("ذكر", 10, 40), entry("إناث", 30, 60)]),
            item("2", "الدعم الأكاديمي", [entry("ذكر", 10, 20), entry("إناث", 30, 120)]),
        ])),
        "p02::1446": dataset(survey([item("1", "وضوح الخطة", [entry("ذكر", 20, 100)])])),
    } };
}
const report = (payload, filters = {}) => buildReport(prepareData(payload, programs), { ...scope, ...filters });

test("programme item means pool actual responses and baseline averages all current items", () => {
    const result = report(fixture());
    assert.equal(result.rows[0].current, 3.5); // 140/40, not (5+3)/2.
    assert.equal(result.rows[0].responseCount, 40);
    assert.equal(result.rows[0].previous, 2.5);
    assert.equal(result.rows[0].annualChange, 1);
    assert.equal(result.overall, 3);
    assert.equal(result.rows[0].overallDifference, 0.5);
    assert.equal(result.comparableItemCount, 2);
});
test("cohort comparison keeps separate baselines, triplets, scope and visible export order", () => {
    const data = prepareData(fixture(), programs), filters = { ...scope, compareCohorts: true, itemSort: "lowest" };
    const result = buildReport(data, filters);
    assert.deepEqual(result.cohortReports.map((row) => row.overall), [3, 3, 3]);
    assert.deepEqual(result.rows.map((row) => [row.number, row.cohort]), [["2", "male"], ["2", "female"], ["2", "total"], ["1", "male"], ["1", "female"], ["1", "total"]]);
    assert.deepEqual(result.rows.slice(3).map((row) => row.overallDifference), [2, 0, 0.5]);
    assert.deepEqual(exportRows(result, filters, data).slice(0, 3).map((row) => row["الفئة"]), ["الطلاب", "الطالبات", "الإجمالي المتاح"]);
    const searched = report(fixture(), { compareCohorts: true, query: "وضوح" });
    assert.equal(searched.visibleItemCount, 1);
    assert.deepEqual(searched.cohortReports, result.cohortReports);
});
test("programme and degree filtering applies equally to current and previous years", () => {
    assert.equal(report(fixture(), { program: "all" }).rows[0].current, 4);
    const filtered = report(fixture(), { program: "all", degree: "الماجستير" });
    assert.equal(filtered.rows[0].current, 5);
    assert.equal(filtered.rows[0].previous, null);
    assert.equal(filtered.programCount, 1);
    assert.equal(report(fixture(), { degree: "الماجستير" }).rows.length, 0);
});
test("ready averages never use measurementCount as respondents or weighting", () => {
    const payload = { extractedData: { "p01::1447": dataset(survey([item("1", "وضوح", [
        { gender: "ذكر", average: 5, measurementCount: 99 }, { gender: "إناث", average: 1, measurementCount: 1 },
    ])], { title: "استطلاع الخبرة التعليمية في البرنامج", sectionId: "learning" })) } };
    const result = report(payload, { year: "1447", compareCohorts: true });
    assert.deepEqual(result.rows.map((row) => row.current), [5, 1, 3]);
    assert.equal(result.rows[2].responseCount, null);
    assert.equal(result.hasUnknownCounts, true);
    assert.equal(exportRows(result, scope, prepareData(payload, programs))[2]["إجابات السنة الحالية"], null);
});
test("missing current cohort keeps matching prior value and null differences", () => {
    const payload = fixture();
    payload.extractedData["p01::1446"].surveys[0].topics[0].items[1].genders = [entry("ذكر", 10, 10)];
    const result = report(payload, { compareCohorts: true });
    const missing = result.rows.find((row) => row.number === "2" && row.cohort === "female");
    assert.equal(missing.current, null);
    assert.equal(missing.previous, 4);
    assert.equal(missing.annualChange, null);
    assert.equal(missing.overallDifference, null);
});
test("instrument, original topic, item number and wording must all match for annual comparisons", () => {
    for (const mutate of [
        (s) => { s.title = "استطلاع الخبرة التعليمية في البرنامج"; },
        (s) => { s.topics[0].label = "محور مختلف"; },
        (s) => { s.topics[0].items[0].number = "9"; },
        (s) => { s.topics[0].items[0].label = "صياغة مختلفة"; },
    ]) {
        const payload = fixture();
        mutate(payload.extractedData["p01::1445"].surveys[0]);
        assert.equal(report(payload).rows[0].previous, null);
    }
    assert.equal(report(fixture(), { year: "1445" }).rows[0].previous, null);
});
test("identical duplicate aliases are ignored, conflicting duplicates remain unavailable", () => {
    const payload = fixture(), genders = payload.extractedData["p01::1446"].surveys[0].topics[0].items[0].genders;
    genders.push(entry("ذكر", 10, 50));
    assert.equal(report(payload).rows[0].responseCount, 40);
    genders.push(entry("ذكر", 10, 40), entry("ذكر", 10, 50));
    const result = report(payload, { compareCohorts: true });
    assert.equal(result.rows[0].current, null);
    assert.equal(result.rows[2].current, 3);
    assert.equal(result.hasExcluded, true);
});
test("KPI survey lookup excludes faculty and unrelated samples; original source stays unchanged", () => {
    const payload = fixture(), original = payload.extractedData["p01::1446"].surveys[0];
    payload.extractedData["p01::1446"].surveys = [survey([], { title: "استطلاع الخريجين" }), survey([], { stakeholder: "faculty" }), survey([], { title: "استطلاع الخبرة التعليمية في البرنامج" }), original];
    const before = JSON.stringify(payload), result = report(payload);
    assert.equal(result.rows[0].surveyTitle, "استبانة تقويم برنامج");
    assert.equal(result.rows[0].sectionId, "management");
    assert.equal(result.rows[0].topic, "الدعم");
    assert.equal(JSON.stringify(payload), before);
});
test("published survey payload resolves both models with original classifications and known counts", () => {
    const context = { window: {} };
    vm.runInNewContext(fs.readFileSync(require.resolve("../js/surveys-data.js"), "utf8"), context);
    const data = prepareData(context.window.SURVEYS_DATA, [{ id: "p04", degree: "بكالوريوس", label: "القرآن وعلومه" }]);
    const recent = buildReport(data, { year: "1447", program: "p04", compareCohorts: true });
    assert.equal(recent.itemCount, 22);
    assert.equal(recent.comparableItemCount, 0);
    assert.equal(recent.rows[0].sectionId, "learning");
    assert.equal(recent.rows[2].current, (3.6606 + 3.501) / 2);
    assert.equal(recent.rows[2].responseCount, null);
    const legacy = buildReport(data, { year: "1446", program: "p04", compareCohorts: true });
    assert.equal(legacy.itemCount, 23);
    assert.equal(legacy.comparableItemCount, 23);
    assert.equal(legacy.rows[0].sectionId, "management");
    assert.equal(legacy.rows[2].responseCount, 297);
    assert.equal(legacy.rows[2].current, Number(((717 + 269) / (223 + 74)).toFixed(6)));
});
