const { test } = require("node:test");
const assert = require("node:assert/strict");
const { buildReport, exportRows } = require("../js/item-quality-analysis.js");

const stats = (sum, count = 1) => ({ courseMeanSum: sum, courseCount: count, responseCount: count * 10, excludedCourseCount: 0 });
function fixture() {
    return { years: ["1447", "1446"], programs: [
        { id: "p01", name: "بكالوريوس الشريعة", degree: "البكالوريوس" },
        { id: "p02", name: "ماجستير الفقه", degree: "الماجستير" },
    ], itemDefinitions: {
        a: { instrument: "electronic_learning", topic: "أولاً", number: 1, statement: "وضوح الخطة" },
        b: { instrument: "electronic_learning", topic: "أولاً", number: 2, statement: "التغذية الراجعة" },
        old: { instrument: "legacy", topic: "أولاً", number: 1, statement: "وضوح الخطة القديمة" },
    }, records: [
        { year: "1447", semester: "all", programId: "p01", cohort: "total", courseCount: 2, items: { a: stats(9, 2), b: stats(6, 2) } },
        { year: "1446", semester: "all", programId: "p01", cohort: "total", courseCount: 1, items: { a: stats(4), b: stats(3.5), old: stats(5) } },
        { year: "1447", semester: "471", programId: "p01", cohort: "total", courseCount: 1, items: { a: stats(5) } },
        { year: "1446", semester: "461", programId: "p01", cohort: "total", courseCount: 1, items: { a: stats(3) } },
        { year: "1447", semester: "all", programId: "p01", cohort: "female", courseCount: 1, items: { a: stats(2) } },
        { year: "1447", semester: "all", programId: "p02", cohort: "total", courseCount: 1, items: { a: stats(3) } },
    ] };
}
const scope = { year: "1447", semester: "all", degree: "البكالوريوس", program: "all", cohort: "total" };

test("mean of course means and baseline mean of items; annual and baseline differences stay distinct", () => {
    const report = buildReport(fixture(), scope);
    assert.equal(report.overall, 3.75);
    assert.equal(report.rows[0].current, 4.5);
    assert.equal(report.rows[0].previous, 4);
    assert.equal(report.rows[0].annualChange, 0.5);
    assert.equal(report.rows[0].overallDifference, 0.75);
    assert.equal(report.rows[1].annualChange, -0.5);
    assert.equal(report.rows[1].overallDifference, -0.75);
    assert.equal(report.rows.length, 2); // Previous-only instrument does not become a current item.
});

test("semester comparison uses the same semester of the previous year", () => {
    const report = buildReport(fixture(), { ...scope, semester: "471" });
    assert.equal(report.previousSemester, "461");
    assert.equal(report.rows[0].annualChange, 2);
});

test("program, degree and cohort filters apply to both years without duplication", () => {
    assert.equal(buildReport(fixture(), { ...scope, cohort: "female" }).rows[0].current, 2);
    assert.equal(buildReport(fixture(), { ...scope, program: "ماجستير الفقه" }).rows.length, 0);
    const all = buildReport(fixture(), { ...scope, degree: "all" });
    assert.equal(all.rows[0].current, 4); // (9 + 3) / (2 + 1), not mean of program means.
    assert.equal(all.courseCount, 3);
});

test("text search and sorting preserve the full scope baseline; exports follow visible order", () => {
    const report = buildReport(fixture(), { ...scope, query: "التغذية", itemSort: "lowest" });
    assert.equal(report.rows.length, 1);
    assert.equal(report.overall, 3.75);
    assert.equal(report.rows[0].overallDifference, -0.75);
    assert.equal(exportRows(report, scope)[0]["متوسط السنة الحالية"], 3);
    assert.equal(buildReport(fixture(), { ...scope, itemSort: "lowest" }).rows[0].id, "b");
});

test("missing previous wording or incomplete current values remain null, including in exports", () => {
    const data = fixture();
    delete data.records[1].items.a;
    data.records[0].items.b = { ...stats(0, 0), excludedCourseCount: 2 };
    const report = buildReport(data, { ...scope, itemSort: "lowest" });
    assert.equal(report.rows[0].previous, null);
    assert.equal(report.rows[0].annualChange, null);
    assert.equal(report.rows[1].current, null);
    assert.equal(report.rows[1].overallDifference, null);
    assert.equal(report.excludedCourseItems, 2);
    assert.equal(exportRows(report, scope)[0]["متوسط السنة الماضية"], null);
});

test("earliest year and empty scopes do not substitute other years or invent zero", () => {
    const report = buildReport(fixture(), { ...scope, year: "1445" });
    assert.equal(report.overall, null);
    assert.equal(report.previousYear, "1444");
    assert.deepEqual(report.rows, []);
});
