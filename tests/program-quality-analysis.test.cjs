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

test("programme item means and baseline pool actual responses according to KPI", () => {
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
test("ready averages follow KPI measurement weights without treating them as respondents", () => {
    const payload = { extractedData: { "p01::1447": dataset(survey([item("1", "وضوح", [
        { gender: "ذكر", average: 5, measurementCount: 99 }, { gender: "إناث", average: 1, measurementCount: 1 },
    ])], { title: "استطلاع الخبرة التعليمية في البرنامج", sectionId: "learning" })) } };
    const result = report(payload, { year: "1447", compareCohorts: true });
    assert.deepEqual(result.rows.map((row) => row.current), [5, 1, 4.96]);
    assert.equal(result.overall, 4.96);
    assert.equal(result.rows[2].aggregationWeight, 100);
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
    payload.extractedData["p01::1446"].surveys = [survey([], { title: "استطلاع الخريجين" }), survey([], { stakeholder: "faculty" }), original, survey([], { title: "استطلاع الخبرة التعليمية في البرنامج" })];
    const before = JSON.stringify(payload), result = report(payload);
    assert.equal(result.rows[0].surveyTitle, "استبانة تقويم برنامج");
    assert.equal(result.rows[0].sectionId, "management");
    assert.equal(result.rows[0].topic, "الدعم");
    assert.equal(JSON.stringify(payload), before);
});

// Independent compatibility oracle for aggregateProgramExperienceMetricFromSurvey
// from the published KPI_TaifShare3h/main/js/app.js, checked on 2026-10-11.
function publishedKpiValue(survey) {
    let numerator = 0, denominator = 0;
    for (const topic of survey.topics || []) for (const item of topic.items || []) for (const entry of item.genders || []) {
        const responses = Number(entry.responses || 0), total = Number(entry.scoreTotal || 0);
        if (Number.isFinite(responses) && responses > 0) {
            numerator += Number.isFinite(total) ? total : 0; denominator += responses;
        } else {
            const average = Number(entry.average), weight = Number(entry.measurementCount || 0);
            if (Number.isFinite(average) && Number.isFinite(weight) && weight > 0) { numerator += average * weight; denominator += weight; }
        }
    }
    return denominator ? Math.round(numerator / denominator * 100) / 100 : null;
}
test("overall pools unequal item counts instead of averaging their means", () => {
    const payload = { extractedData: { "p01::1446": dataset(survey([
        item("1", "الأول", [entry("ذكر", 10, 50)]), item("2", "الثاني", [entry("ذكر", 90, 90)]),
    ])) } };
    const result = report(payload);
    assert.equal(result.overall, 1.4); // 140/100; mean of item averages would be 3.
    assert.deepEqual(result.rows.map((row) => row.overallDifference), [3.6, -0.4]);
    assert.equal(result.overall, publishedKpiValue(payload.extractedData["p01::1446"].surveys[0]));
});
test("KPI weights of ready averages also determine overall and filtered cohort references", () => {
    const payload = { extractedData: { "p01::1447": dataset(survey([
        item("1", "الأول", [{ gender: "ذكر", average: 5, measurementCount: 1 }, { gender: "إناث", average: 1, measurementCount: 3 }]),
        item("2", "الثاني", [{ gender: "ذكر", average: 3, measurementCount: 2 }]),
    ], { title: "استطلاع الخبرة التعليمية في البرنامج", sectionId: "learning" })) } };
    const result = report(payload, { year: "1447", compareCohorts: true });
    assert.equal(result.overall, 2.333333);
    assert.equal(Math.round(result.overall * 100) / 100, publishedKpiValue(payload.extractedData["p01::1447"].surveys[0]));
    assert.deepEqual(result.cohortReports.map((row) => row.overall), [3.666667, 1, 2.333333]);
    assert.equal(result.rows[2].current, 2);
    assert.equal(result.rows[2].responseCount, null);
});
test("all published programme-year totals match the KPI formula and each cohort uses the same formula", () => {
    const context = { window: {} };
    vm.runInNewContext(fs.readFileSync(require.resolve("../js/surveys-data.js"), "utf8"), context);
    const payload = context.window.SURVEYS_DATA;
    const allPrograms = Array.from({ length: 15 }, (_, i) => ({ id: `p${String(i + 1).padStart(2, "0")}`, degree: "all", label: String(i + 1) }));
    const data = prepareData(payload, allPrograms);
    let compared = 0;
    for (const [key, dataset] of Object.entries(payload.extractedData)) {
        const [program, year] = key.split("::");
        const source = dataset.surveys?.find((s) => ["استبانة تقويم برنامج", "استطلاع الخبرة التعليمية في البرنامج"].includes(s.title));
        if (!source || !allPrograms.some((p) => p.id === program)) continue;
        const result = buildReport(data, { program, year, compareCohorts: true });
        for (const cohort of ["male", "female", "total"]) {
            const filtered = { ...source, topics: source.topics.map((topic) => ({ ...topic, items: topic.items.map((item) => ({ ...item,
                genders: item.genders.filter((g) => cohort === "total" || (cohort === "male" ? g.gender === "ذكر" : g.gender === "إناث")),
            })) })) };
            const actual = result.cohortReports.find((r) => r.cohort === cohort).overall;
            assert.equal(actual == null ? null : Math.round(actual * 100) / 100, publishedKpiValue(filtered), `${key}:${cohort}`);
        }
        compared++;
    }
    assert.equal(compared, 35);
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
