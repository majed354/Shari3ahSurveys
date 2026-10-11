(function (root, factory) {
    "use strict";
    const api = factory();
    if (typeof module === "object" && module.exports) module.exports = api;
    else root.ItemQualityAnalysis = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
    "use strict";
    const normalize = (text) => String(text || "").normalize("NFC").replace(/\s+/g, " ").trim();
    const searchable = (text) => normalize(text).replace(/[\u064b-\u065fـ]/g, "").replace(/[أإآ]/g, "ا");
    const mean = (values) => values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
    const rounded = (value) => value == null ? null : Number(value.toFixed(6));

    function collect(data, filters, year, semester) {
        const programs = new Map(data.programs.map((program) => [program.id, program]));
        const items = new Map();
        let courseCount = 0;
        let unnamed = 0;
        data.records.forEach((record) => {
            const program = programs.get(record.programId);
            if (!program || record.year !== year || record.semester !== semester || record.cohort !== (filters.cohort || "total")) return;
            if (filters.program && filters.program !== "all" && program.name !== filters.program) return;
            if (filters.degree && filters.degree !== "all" && program.degree !== filters.degree) return;
            courseCount += record.courseCount;
            unnamed += record.unnamedItemObservations || 0;
            Object.entries(record.items).forEach(([id, value]) => {
                if (!items.has(id)) items.set(id, { sum: 0, courseCount: 0, responseCount: 0, excludedCourseCount: 0 });
                const entry = items.get(id);
                entry.sum += value.courseMeanSum;
                entry.courseCount += value.courseCount;
                entry.responseCount += value.responseCount;
                entry.excludedCourseCount += value.excludedCourseCount;
            });
        });
        items.forEach((entry) => { entry.average = entry.courseCount ? entry.sum / entry.courseCount : null; });
        return { items, courseCount, unnamed };
    }

    function topicOrder(topic) {
        const prefixes = ["أولاً", "ثانياً", "ثالثاً", "رابعاً", "خامساً"];
        const index = prefixes.findIndex((prefix) => searchable(topic).startsWith(searchable(prefix)));
        if (index >= 0) return index;
        return ["أسئلة خاصة بـبداية المقرر", "أسئلة خاصة بما حدث خلال المقرر", "تقويم المقرر", "التقويم العام"].indexOf(topic);
    }

    function buildReport(data, filters = {}) {
        const year = String(filters.year === "all" || !filters.year ? data.years[0] : filters.year);
        const previousYear = String(Number(year) - 1);
        const semester = filters.semester || "all";
        const previousSemester = semester === "all" ? "all" : String(Number(semester) - 10);
        const current = collect(data, filters, year, semester);
        const previous = collect(data, filters, previousYear, previousSemester);
        const overall = mean([...current.items.values()].map((item) => item.average).filter((value) => value != null));
        let rows = [...current.items.entries()].map(([id, item]) => {
            const definition = data.itemDefinitions[id];
            const prior = previous.items.get(id);
            return { id, ...definition, current: rounded(item.average), previous: rounded(prior?.average ?? null),
                annualChange: rounded(item.average != null && prior?.average != null ? item.average - prior.average : null),
                overallDifference: rounded(item.average != null && overall != null ? item.average - overall : null),
                courseCount: item.courseCount, previousCourseCount: prior?.courseCount || 0,
                responseCount: item.responseCount, previousResponseCount: prior?.responseCount || 0,
                excludedCourseCount: item.excludedCourseCount, previousExcludedCourseCount: prior?.excludedCourseCount || 0 };
        });
        const itemCount = rows.filter((row) => row.current != null).length;
        const comparableItemCount = rows.filter((row) => row.annualChange != null).length;
        const excludedCourseItems = rows.reduce((sum, row) => sum + row.excludedCourseCount, 0);
        const query = searchable(filters.query);
        if (query) rows = rows.filter((row) => searchable(`${row.statement} ${row.topic}`).includes(query));
        const sort = filters.itemSort || "source";
        const field = sort === "change" ? "annualChange" : sort === "difference" ? "overallDifference" : "current";
        rows.sort((a, b) => {
            if (sort !== "source") {
                if (a[field] == null && b[field] != null) return 1;
                if (a[field] != null && b[field] == null) return -1;
                const diff = (a[field] || 0) - (b[field] || 0);
                if (diff) return sort === "highest" ? -diff : diff;
            }
            return a.instrument.localeCompare(b.instrument) || topicOrder(a.topic) - topicOrder(b.topic)
                || a.topic.localeCompare(b.topic, "ar") || Number(a.number) - Number(b.number);
        });
        return { year, previousYear, semester, previousSemester, overall: rounded(overall), courseCount: current.courseCount,
            previousCourseCount: previous.courseCount, itemCount, comparableItemCount, excludedCourseItems,
            unnamedItemObservations: current.unnamed, rows };
    }

    function exportRows(report, scope) {
        return report.rows.map((row) => ({
            "البرنامج": scope.program === "all" ? "كل البرامج" : scope.program,
            "الدرجة": scope.degree === "all" ? "كل الدرجات" : scope.degree,
            "الشطر": { total: "الإجمالي المتاح", male: "الطلاب", female: "الطالبات" }[scope.cohort || "total"],
            "السنة الماضية": report.previousYear, "السنة الحالية": report.year,
            "الفصل الحالي": report.semester === "all" ? "كل الفصول" : report.semester,
            "الفصل السابق": report.previousSemester === "all" ? "كل الفصول" : report.previousSemester,
            "النموذج": row.instrument === "legacy" ? "تقويم المقرر" : "التعلم الإلكتروني",
            "المحور": row.topic, "رقم البند داخل المحور": row.number, "البند": row.statement,
            "متوسط السنة الماضية": row.previous, "متوسط السنة الحالية": row.current,
            "التغير السنوي": row.annualChange, "الفرق عن المتوسط العام الحالي": row.overallDifference,
            "المتوسط العام الحالي": report.overall, "مقررات البند السنة الماضية": row.previousCourseCount,
            "مقررات البند السنة الحالية": row.courseCount, "إجابات البند السنة الماضية": row.previousResponseCount,
            "إجابات البند السنة الحالية": row.responseCount,
            "مقررات مستبعدة من البند السنة الماضية": row.previousExcludedCourseCount,
            "مقررات مستبعدة من البند السنة الحالية": row.excludedCourseCount,
            "طريقة الحساب": "أوزان متساوية للمقررات؛ المتوسط العام من جميع البنود في النطاق قبل البحث النصي",
        }));
    }
    return { buildReport, exportRows };
});
