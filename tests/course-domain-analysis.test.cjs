const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { buildRows, categoryScore } = require("../js/course-domain-analysis.js");

function fixture() {
    return {
        programs: [{ id: "p10", name: "ماجستير الدراسات القرآنية المعاصرة", degree: "الماجستير" },
            { id: "p04", name: "بكالوريوس القرآن وعلومه", degree: "البكالوريوس" }],
        records: [
            { year: "1447", semester: "471", programId: "p10", courseName: "أصول التفسير",
                cohorts: { total: { categories: { skills: [8, 2, 1, 0, 0], overall_satisfaction: [9, 2, 1, 0, 0], unclassified: [6, 2, 1, 0, 0] } } } },
            { year: "1447", semester: "472", programId: "p10", courseName: "أصول  التفسير",
                cohorts: { total: { categories: { skills: [12, 3, 1, 0, 0], overall_satisfaction: [10, 3, 1, 0, 0] } } } },
            { year: "1447", semester: "471", programId: "p04", courseName: "تلاوة",
                cohorts: { total: { categories: { skills: [6, 2, 1, 1, 0] } }, male: { categories: { skills: [4, 1, 1, 0, 0] } }, female: { categories: { skills: [0, 0, 1, 1, 1] } } } },
        ],
    };
}

test("semesters are pooled by response counts before dividing", () => {
    const rows = buildRows(fixture(), { year: "1447", semester: "all", program: "ماجستير الدراسات القرآنية المعاصرة" });
    assert.equal(rows.length, 1);
    assert.equal(rows[0].scores.skills, 20 / 5);
    assert.equal(rows[0].scores.overall_satisfaction, 19 / 5);
    assert.equal(rows[0].scores.knowledge, null); // no knowledge item: not measured, never zero
    assert.equal(rows[0].excludedItems, 1);
    assert.deepEqual(rows[0].semesters, ["471", "472"]);
});

test("a partial item is scored from accepted counts and flagged; cohorts stay separate", () => {
    const total = buildRows(fixture(), { year: "1447", semester: "471", degree: "البكالوريوس" });
    assert.equal(total[0].scores.skills, 3);
    assert.equal(total[0].partial.skills, true);
    const male = buildRows(fixture(), { year: "1447", semester: "471", degree: "البكالوريوس", cohort: "male" });
    assert.equal(male[0].scores.skills, 4);
    assert.equal(male[0].partial.skills, false);
    const female = buildRows(fixture(), { year: "1447", semester: "471", degree: "البكالوريوس", cohort: "female" });
    assert.equal(female[0].scores.skills, null); // only a quarantined report: no accepted answer
    assert.equal(categoryScore(undefined), null);
});

test("every published 1447 and 1446 course has all three domains and overall satisfaction", () => {
    const data = JSON.parse(fs.readFileSync(path.join(__dirname, "../data/course-evaluations/course-domain-scores.json"), "utf8"));
    for (const year of ["1447", "1446"]) {
        const rows = buildRows(data, { year, semester: "all" });
        rows.forEach((row) => ["knowledge", "skills", "values", "overall_satisfaction"].forEach((id) => {
            assert.notEqual(row.scores[id], null, `${year} ${row.program} ${row.courseName} ${id}`);
        }));
    }
});

test("published scores load and every row has bounded scores", () => {
    const data = JSON.parse(fs.readFileSync(path.join(__dirname, "../data/course-evaluations/course-domain-scores.json"), "utf8"));
    assert.equal(data.schemaVersion, "course-domain-scores-v1");
    const rows = buildRows(data, { year: "1447", semester: "all" });
    assert.ok(rows.length > 100);
    rows.forEach((row) => Object.values(row.scores).forEach((value) => {
        if (value != null) assert.ok(value >= 1 && value <= 5);
    }));
});
