(function (root, factory) {
    "use strict";
    const api = factory();
    if (typeof module === "object" && module.exports) module.exports = api;
    else root.ProgramQualityAnalysis = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
    "use strict";
    const clean = (value) => String(value ?? "").normalize("NFC").replace(/\s+/g, " ").trim();
    const searchable = (value) => clean(value).replace(/[\u064b-\u065fـ]/g, "").replace(/[أإآ]/g, "ا");
    const numeric = (value) => value == null || clean(value) === "" || !Number.isFinite(Number(value)) ? null : Number(value);
    const rounded = (value) => value == null ? null : Number(value.toFixed(6));
    const mean = (values) => values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
    const cohortLabel = (cohort) => ({ male: "الطلاب", female: "الطالبات", total: "الإجمالي المتاح" })[cohort];
    const sectionLabel = (id) => ({ management: "إدارة البرنامج وضمان الجودة", learning: "التعليم والتعلم" })[id] || id;

    function instrument(title) {
        const value = searchable(title);
        if (value === searchable("استبانة تقويم برنامج")) return "program_evaluation";
        if (value === searchable("استطلاع الخبرة التعليمية في البرنامج")) return "program_experience";
        return null;
    }

    function gender(value) {
        const name = searchable(value);
        if (["ذكر", "ذكور", "طالب", "طلاب", "male"].includes(name)) return "male";
        if (["اناث", "انثى", "طالبة", "طالبات", "female"].includes(name)) return "female";
        return null;
    }

    function measurement(entry) {
        const count = numeric(entry.responses), total = numeric(entry.scoreTotal), average = numeric(entry.average);
        if (count > 0 && total != null && total >= count && total <= count * 5) return { mode: "responses", count, total, value: total / count };
        if (average != null && average >= 1 && average <= 5) return { mode: "averages", count: null, total: null, value: average };
        return null;
    }

    function prepareData(payload, programs) {
        const definitions = new Map(), cells = new Map(), years = new Set();
        for (const [datasetKey, dataset] of Object.entries(payload.extractedData || {})) {
            const [programId, year] = datasetKey.split("::");
            if (!programs.some((program) => program.id === programId)) continue;
            const surveys = (dataset.surveys || []).filter((survey) => instrument(survey.title) && (!survey.stakeholder || survey.stakeholder === "students"));
            // Same title resolution used by KPI_TaifShare3h; prefer the original
            // programme instrument when both instruments exist in one dataset.
            const survey = surveys.find((entry) => instrument(entry.title) === "program_evaluation") || surveys[0];
            if (!survey) continue;
            years.add(year);
            (survey.topics || []).forEach((topic, topicIndex) => (topic.items || []).forEach((item, itemIndex) => {
                if (!clean(item.label)) return;
                const id = JSON.stringify([instrument(survey.title), clean(topic.label), clean(item.number), clean(item.label)]);
                if (!definitions.has(id)) definitions.set(id, { id, instrument: instrument(survey.title), surveyTitle: survey.title,
                    sectionId: survey.sectionId, topic: topic.label, topicOrder: topicIndex, itemOrder: itemIndex,
                    number: item.number || "", statement: item.label });
                (item.genders || []).forEach((entry) => {
                    const cohort = gender(entry.gender);
                    if (!cohort) return;
                    const semester = clean(entry.semester || item.semester || topic.semester || survey.semester || "annual");
                    const key = JSON.stringify([programId, year, id, cohort, semester]);
                    const value = measurement(entry);
                    const cell = { programId, year, id, cohort, semester, value, conflict: false };
                    if (!cells.has(key)) cells.set(key, cell);
                    else if (JSON.stringify(cells.get(key).value) !== JSON.stringify(value)) {
                        cells.get(key).value = null;
                        cells.get(key).conflict = true;
                    }
                });
            }));
        }
        return { programs: programs.filter((program) => program.id !== "all"), years: [...years].sort((a, b) => Number(b) - Number(a)),
            definitions, cells: [...cells.values()] };
    }

    function aggregate(cells) {
        const valid = cells.filter((cell) => cell.value);
        const weighted = valid.length > 0 && valid.every((cell) => cell.value.mode === "responses");
        const responses = weighted ? valid.reduce((sum, cell) => sum + cell.value.count, 0) : null;
        const average = weighted ? valid.reduce((sum, cell) => sum + cell.value.total, 0) / responses : mean(valid.map((cell) => cell.value.value));
        return { average, responseCount: responses, programCount: new Set(valid.map((cell) => cell.programId)).size,
            method: !valid.length ? "unavailable" : weighted ? "responses" : "averages", excluded: cells.length - valid.length };
    }

    function collect(data, filters, year, cohort) {
        const programs = new Map(data.programs.map((program) => [program.id, program]));
        const grouped = new Map();
        const cells = data.cells.filter((cell) => {
            const program = programs.get(cell.programId);
            return cell.year === year && (cohort === "total" || cell.cohort === cohort)
                && (!filters.program || filters.program === "all" || cell.programId === filters.program)
                && (!filters.degree || filters.degree === "all" || program.degree === filters.degree);
        });
        cells.forEach((cell) => {
            if (!grouped.has(cell.id)) grouped.set(cell.id, []);
            grouped.get(cell.id).push(cell);
        });
        return { items: new Map([...grouped].map(([id, entries]) => [id, aggregate(entries)])),
            programCount: new Set(cells.filter((cell) => cell.value).map((cell) => cell.programId)).size };
    }

    function buildReport(data, filters = {}) {
        const year = String(filters.year || data.years[0]), previousYear = String(Number(year) - 1);
        const cohorts = filters.compareCohorts ? ["male", "female", "total"] : [filters.cohort || "total"];
        const scopes = cohorts.map((cohort) => {
            const current = collect(data, filters, year, cohort), previous = collect(data, filters, previousYear, cohort);
            return { cohort, current, previous, overall: mean([...current.items.values()].map((item) => item.average).filter((value) => value != null)) };
        });
        const ids = [...new Set(scopes.flatMap((scope) => [...scope.current.items.keys()]))];
        const empty = { average: null, responseCount: null, programCount: 0, method: "unavailable", excluded: 0 };
        const rowsById = new Map(ids.map((id) => [id, scopes.map((scope) => {
            const current = scope.current.items.get(id) || empty, previous = scope.previous.items.get(id) || empty;
            return { ...data.definitions.get(id), cohort: scope.cohort, overall: rounded(scope.overall),
                current: rounded(current.average), previous: rounded(previous.average),
                annualChange: current.average == null || previous.average == null ? null : rounded(current.average - previous.average),
                overallDifference: current.average == null || scope.overall == null ? null : rounded(current.average - scope.overall),
                responseCount: current.responseCount, previousResponseCount: previous.responseCount,
                programCount: current.programCount, previousProgramCount: previous.programCount,
                method: current.method, previousMethod: previous.method, excluded: current.excluded + previous.excluded };
        })]));
        const field = filters.itemSort === "change" ? "annualChange" : filters.itemSort === "difference" ? "overallDifference" : "current";
        const reference = (id) => rowsById.get(id).at(-1);
        const query = searchable(filters.query);
        const visible = ids.filter((id) => !query || searchable(`${reference(id).statement} ${reference(id).topic} ${sectionLabel(reference(id).sectionId)}`).includes(query));
        visible.sort((a, b) => {
            const first = reference(a), second = reference(b);
            if (filters.itemSort && filters.itemSort !== "source") {
                if (first[field] == null && second[field] != null) return 1;
                if (first[field] != null && second[field] == null) return -1;
                const difference = (first[field] || 0) - (second[field] || 0);
                if (difference) return filters.itemSort === "highest" ? -difference : difference;
            }
            return first.instrument.localeCompare(second.instrument) || first.topicOrder - second.topicOrder || first.itemOrder - second.itemOrder || first.statement.localeCompare(second.statement, "ar");
        });
        const cohortReports = scopes.map((scope) => ({ cohort: scope.cohort, overall: rounded(scope.overall),
            programCount: scope.current.programCount, itemCount: [...scope.current.items.values()].filter((item) => item.average != null).length }));
        const referenceScope = scopes.at(-1);
        return { year, previousYear, comparison: Boolean(filters.compareCohorts), cohortReports,
            overall: rounded(referenceScope.overall), programCount: referenceScope.current.programCount,
            itemCount: ids.filter((id) => rowsById.get(id).some((row) => row.current != null)).length,
            comparableItemCount: ids.filter((id) => rowsById.get(id).some((row) => row.annualChange != null)).length,
            visibleItemCount: visible.length, rows: visible.flatMap((id) => rowsById.get(id)),
            hasUnknownCounts: [...rowsById.values()].flat().some((row) => row.method === "averages" || row.previousMethod === "averages"),
            hasExcluded: [...rowsById.values()].flat().some((row) => row.excluded > 0) };
    }

    function exportRows(report, filters, data) {
        const methodLabel = (method) => ({ responses: "موزون بعدد الإجابات", averages: "متوسط المتوسطات المتاحة؛ العدد غير متاح", unavailable: "غير متاح" })[method];
        return report.rows.map((row) => ({
            "البرنامج": filters.program === "all" ? "كل البرامج" : data.programs.find((program) => program.id === filters.program)?.label,
            "الدرجة": filters.degree === "all" ? "كل الدرجات" : filters.degree,
            "الفئة": cohortLabel(row.cohort), "التصنيف الأصلي": sectionLabel(row.sectionId),
            "الاستطلاع": row.surveyTitle, "المحور الأصلي": row.topic, "رقم البند": row.number, "البند": row.statement,
            "السنة الماضية": report.previousYear, "السنة الحالية": report.year,
            "متوسط السنة الماضية": row.previous, "متوسط السنة الحالية": row.current,
            "التغير السنوي": row.annualChange, "الفرق عن المتوسط العام للفئة": row.overallDifference,
            "المتوسط العام الحالي للفئة": row.overall, "إجابات السنة الماضية": row.previousResponseCount, "إجابات السنة الحالية": row.responseCount,
            "برامج السنة الماضية": row.previousProgramCount, "برامج السنة الحالية": row.programCount,
            "طريقة حساب السنة الماضية": methodLabel(row.previousMethod), "طريقة حساب السنة الحالية": methodLabel(row.method),
        }));
    }
    return { prepareData, buildReport, exportRows, sectionLabel, instrument };
});
