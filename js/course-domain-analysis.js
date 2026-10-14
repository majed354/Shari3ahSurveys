(function (root, factory) {
    "use strict";
    const api = factory();
    if (typeof module === "object" && module.exports) module.exports = api;
    else root.CourseDomainAnalysis = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
    "use strict";
    // Columns shown in the indirect-assessment table; "review" and "other" stay out of outcome averages.
    const COLUMNS = ["overall_satisfaction", "knowledge", "skills", "values", "quality", "other"];
    const normalize = (text) => String(text || "").normalize("NFC").replace(/\s+/g, " ").trim();

    // Same rule as the automation pipeline: a category is scored only when every item is complete.
    function categoryScore(values) {
        if (!values) return null;
        const [numerator, responses, items, nonComplete] = values;
        return items && !nonComplete && responses > 0 ? numerator / responses : null;
    }

    function addTuple(target, values) {
        if (!values) return target;
        const base = target || [0, 0, 0, 0];
        return base.map((value, index) => value + values[index]);
    }

    function buildRows(data, filters, codeLookup = () => "") {
        const programs = new Map(data.programs.map((program) => [program.id, program]));
        const query = normalize(filters.query).toLocaleLowerCase("ar");
        const cohort = filters.cohort || "total";
        const groups = new Map();
        data.records.forEach((record) => {
            const program = programs.get(record.programId);
            if (!program) return;
            if (filters.year && filters.year !== "all" && record.year !== filters.year) return;
            if (filters.semester && filters.semester !== "all" && record.semester !== filters.semester) return;
            if (filters.degree && filters.degree !== "all" && program.degree !== filters.degree) return;
            if (filters.program && filters.program !== "all" && program.name !== filters.program) return;
            const values = record.cohorts[cohort];
            if (!values) return;
            const key = [record.year, record.programId, normalize(record.courseName)].join("||");
            if (!groups.has(key)) {
                groups.set(key, { year: record.year, program: program.name, degree: program.degree,
                    courseName: record.courseName, semesters: [], categories: {} });
            }
            const group = groups.get(key);
            group.semesters.push(record.semester);
            // Same review items recur each semester; report distinct items, not occurrences.
            group.reviewItems = Math.max(group.reviewItems || 0, values.categories.review?.[2] || 0);
            Object.entries(values.categories).forEach(([category, tuple]) => {
                group.categories[category] = addTuple(group.categories[category], tuple);
            });
        });
        return [...groups.values()].map((group) => {
            const courseCode = codeLookup(group) || "";
            const scores = Object.fromEntries(COLUMNS.map((id) => [id, categoryScore(group.categories[id])]));
            const responses = Math.max(0, ...Object.values(group.categories).map((tuple) => tuple[2] ? tuple[1] / tuple[2] : 0));
            return { ...group, semesters: group.semesters.sort(), courseCode, scores,
                itemResponsesPerItem: Math.round(responses), reviewItems: group.reviewItems || 0 };
        }).filter((row) => !query || [row.courseName, row.courseCode, row.program].join(" ").toLocaleLowerCase("ar").includes(query))
            .sort((a, b) => a.program.localeCompare(b.program, "ar") || a.courseName.localeCompare(b.courseName, "ar"));
    }

    return { COLUMNS, buildRows, categoryScore };
});
