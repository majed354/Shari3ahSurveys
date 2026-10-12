#!/usr/bin/env node
/*
 * يولّد data/kpi/program-experience.json من js/surveys-data.js.
 * الملف الناتج هو ما تقرؤه لوحة مؤشرات الأداء (KPI_TaifShare3h) رقمًا جاهزًا.
 *
 * الاستعمال:
 *   node scripts/build_program_experience_kpi.cjs          # يكتب الملف
 *   node scripts/build_program_experience_kpi.cjs --check  # يتحقق أن الملف محدَّث
 */
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { buildKpiFile } = require("../js/program-experience-kpi.js");

const ROOT = path.resolve(__dirname, "..");
const SOURCE_PATH = path.join(ROOT, "js", "surveys-data.js");
const OUTPUT_PATH = path.join(ROOT, "data", "kpi", "program-experience.json");

function loadSurveysData(sourcePath = SOURCE_PATH) {
    const sandbox = { window: {} };
    vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), sandbox, { filename: sourcePath });
    return sandbox.window.SURVEYS_DATA;
}

function render(payload) {
    return JSON.stringify(buildKpiFile(payload), null, 2) + "\n";
}

function main() {
    const content = render(loadSurveysData());
    if (process.argv.includes("--check")) {
        const current = fs.existsSync(OUTPUT_PATH) ? fs.readFileSync(OUTPUT_PATH, "utf8") : "";
        if (current !== content) {
            console.error(`${path.relative(ROOT, OUTPUT_PATH)} غير محدَّث؛ شغّل: node scripts/build_program_experience_kpi.cjs`);
            process.exit(1);
        }
        console.log(`${path.relative(ROOT, OUTPUT_PATH)} محدَّث.`);
        return;
    }
    fs.mkdirSync(path.dirname(OUTPUT_PATH), { recursive: true });
    fs.writeFileSync(OUTPUT_PATH, content, "utf8");
    const records = Object.values(JSON.parse(content).records);
    const tiers = records.reduce((acc, record) => ({ ...acc, [record.tier]: (acc[record.tier] || 0) + 1 }), {});
    console.log(`Generated ${path.relative(ROOT, OUTPUT_PATH)}: ${records.length} records, tiers ${JSON.stringify(tiers)}`);
}

if (require.main === module) main();

module.exports = { loadSurveysData, render, OUTPUT_PATH };
