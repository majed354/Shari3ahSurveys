const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const { RULES, resolveAll } = require("../js/program-experience-kpi.js");
const { loadSurveysData, render, OUTPUT_PATH } = require("../scripts/build_program_experience_kpi.cjs");

const entry = (gender, responses, scoreTotal) => ({ gender, responses, scoreTotal });
const ready = (gender, average, measurementCount = 1) => ({ gender, average, measurementCount });
const item = (label, genders, number = "") => ({ number, label, genders });
const topic = (label, items) => ({ label, items });
const survey = (title, topics, extras = {}) => ({ title, stakeholder: "students", topics, ...extras });
const OVERALL = "أشعر بالرضا بشكل عام عن مستوى جودة خبرتي التعليمية في الجامعة";
const print = RULES.programEvaluation.fingerprint;

function programEvaluation(title = "استبانة تقويم برنامج") {
    return survey(title, [
        topic("المساعدة والدعم", print.slice(0, 13).map((label) => item(label, [entry("ذكر", 10, 30), entry("إناث", 10, 30)]))),
        topic("تقويم التعليم", [...print.slice(13, 21), `${OVERALL}.`].map((label) => item(label, [entry("ذكر", 10, 20)]))),
        topic("التقويم العام", [item(OVERALL, [entry("ذكر", 10, 40), entry("إناث", 30, 90)])]),
    ]);
}
const resolve = (surveys, key = "p01::1446") => resolveAll({ extractedData: { [key]: { surveys } } })[key];

test("tier 1 uses the explicit overall item, pooled by responses, not the item mean", () => {
    const record = resolve([programEvaluation()]);
    assert.equal(record.tier, 1);
    assert.equal(record.value, 3.25); // (40 + 90) / (10 + 30)
    assert.equal(record.respondents, 40);
    assert.equal(record.item, OVERALL);
    assert.notEqual(record.itemsMean, record.value);
});

test("a renamed survey is still recognised from its item fingerprint", () => {
    const record = resolve([programEvaluation("استبيان الطلبة الختامي")]);
    assert.equal(record.tier, 1);
    assert.equal(record.identifiedBy, "fingerprint");
});

test("an unknown title without the fingerprint is not treated as the programme survey", () => {
    const other = survey("استبانة المكتبة", [topic("عام", [item("الرضا بشكل عام عن البرنامج", [entry("ذكر", 5, 25)])])]);
    assert.equal(resolve([other]), undefined);
});

test("tier 2 averages the survey's own items and drops unclassified topics and foreign items", () => {
    const record = resolve([survey("استطلاع الخبرة التعليمية في البرنامج", [
        topic("الخبرة التعليمية والمقررات", [item("طرق التدريس تساعدني على الفهم", [ready("ذكر", 4), ready("إناث", 3)])]),
        topic("بنود متنوعة", [item("التوزيع على أماكن التدريب يتم وفق أسس موضوعية", [ready("إناث", 1)])]),
    ])], "p01::1447");
    assert.equal(record.tier, 2);
    assert.equal(record.value, 3.5);
    assert.equal(record.respondents, null);
    assert.equal(record.weighting, "averages");
});

test("measurementCount of ready averages counts source aliases, so it never weights the mean", () => {
    const record = resolve([survey("استطلاع الخبرة التعليمية في البرنامج", [
        topic("الخبرة التعليمية والمقررات", [item("طرق التدريس تساعدني على الفهم", [ready("ذكر", 4, 20), ready("إناث", 2, 2)])]),
    ])], "p01::1447");
    assert.equal(record.value, 3);
});

test("tier 3 falls back to the learning topic of the student experience survey", () => {
    const record = resolve([survey("استبانة تقويم خبرة الطالب/ الطالبة", [
        topic("التعــلم والتـعليم", [item("يحفزني البرنامج على المزيد من التعلم", [entry("ذكر", 4, 16)])]),
        topic("التــقــويم العــــام", [item("أشعر بالرضا بشكل عام عن حياتي الطلابية", [entry("ذكر", 4, 4)])]),
    ])]);
    assert.equal(record.tier, 3);
    assert.equal(record.value, 4);
});

test("invalid or out-of-range cells are ignored", () => {
    const pes = programEvaluation();
    pes.topics[2].items[0].genders.push(entry("ذكر", 2, 50), ready("إناث", 9, 3), entry("غير محدد", 5, 25));
    assert.equal(resolve([pes]).value, 3.25);
});

test("identical values across programmes are flagged as shared, never hidden", () => {
    const records = resolveAll({ extractedData: Object.fromEntries(["p01", "p02", "p03"].map((id) => [`${id}::1447`, { surveys: [programEvaluation()] }])) });
    assert.deepEqual(Object.values(records).map((record) => [record.scope, record.value]), [["shared", 3.25], ["shared", 3.25], ["shared", 3.25]]);
});

test("program-level data in the same year is not marked as shared", () => {
    const varied = programEvaluation();
    varied.topics[2].items[0].genders = [entry("ذكر", 10, 50)];
    const records = resolveAll({ extractedData: {
        "p01::1446": { surveys: [programEvaluation()] }, "p02::1446": { surveys: [programEvaluation()] }, "p03::1446": { surveys: [varied] },
    } });
    assert.ok(Object.values(records).every((record) => record.scope === "program"));
});

test("published KPI file is in sync with js/surveys-data.js", () => {
    assert.equal(fs.readFileSync(OUTPUT_PATH, "utf8"), render(loadSurveysData()));
});
