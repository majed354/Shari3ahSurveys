(function (root) {
    "use strict";
    function buildReport(data, filters = {}) {
        const programs = new Map(data.programs.map(p => [p.id, p]));
        const groups = data.groups.filter(g => (filters.year === "all" || !filters.year || g.year === filters.year)
            && (filters.program === "all" || !filters.program || g.programId === filters.program)
            && (filters.degree === "all" || !filters.degree || programs.get(g.programId)?.degree === filters.degree));
        const query = String(filters.query || "").trim();
        const sections = data.questionnaires.map(questionnaire => {
            const selected = groups.filter(g => g.questionnaireId === questionnaire.id);
            if (!selected.length) return null;
            const respondentCount = selected.reduce((sum, g) => sum + g.respondents, 0);
            const rows = questionnaire.questions.map(question => {
                const entries = selected.flatMap(g => g.items.filter(i => i.questionId === question.id));
                const responseCount = entries.reduce((sum, i) => sum + i.responseCount, 0);
                const categories = new Map();
                entries.forEach(i => i.categories.forEach(c => categories.set(c.label, (categories.get(c.label) || 0) + c.count)));
                return { ...question, questionnaire: questionnaire.label, respondentCount, responseCount,
                    missingCount: respondentCount - responseCount,
                    average: question.scale && responseCount ? entries.reduce((sum, i) => sum + i.scoreTotal, 0) / responseCount : null,
                    categories: [...categories].map(([label, count]) => ({ label, count, percentage: responseCount ? count / responseCount * 100 : null })).sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, "ar")) };
            }).filter(row => (!filters.axis || filters.axis === "all" || row.axis === filters.axis)
                && (!query || `${row.label} ${row.axis}`.includes(query)));
            return { id: questionnaire.id, label: questionnaire.label, sourceFile: questionnaire.sourceFile, respondentCount, rows };
        }).filter(Boolean);
        return { sections, rows: sections.flatMap(s => s.rows),
            respondentCount: groups.reduce((sum, g) => sum + g.respondents, 0),
            programCount: new Set(groups.map(g => g.programId)).size };
    }

    function exportRows(report, scope) {
        return report.rows.flatMap(row => (row.categories.length ? row.categories : [{ label: row.kind === "text" ? "إجابات نصية مفتوحة (العدد فقط)" : "لا توجد إجابات", count: row.responseCount, percentage: null }]).map(category => ({
            "النطاق": scope, "الاستطلاع": row.questionnaire, "المحور": row.axis, "رقم البند": row.number, "البند": row.label,
            "متوسط البند": row.average == null ? "" : Number(row.average.toFixed(4)), "المقياس": row.scale || "",
            "المشاركون في النطاق": row.respondentCount, "عدد إجابات البند": row.responseCount, "دون إجابة": row.missingCount,
            "الإجابة": category.label, "العدد": category.count, "النسبة من إجابات البند (%)": category.percentage == null ? "" : Number(category.percentage.toFixed(2)),
            "ملاحظة": row.kind === "range" ? "المتوسط تقديري باستخدام منتصف كل فئة" : row.kind === "text" ? "لا تُنشر نصوص التعليقات الفردية" : "",
        })));
    }
    const api = { buildReport, exportRows };
    root.GraduateSurveysAnalysis = api;
    if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
