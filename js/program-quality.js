(function () {
    "use strict";
    const analysis = window.ProgramQualityAnalysis;
    const view = window.SurveyItemReportView;
    const state = { year: "", degree: "all", program: "all", cohort: "total", compareCohorts: false, query: "", itemSort: "source" };
    let data, root, searchTimer;
    const escape = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" })[char]);
    const number = (value) => Number(value).toLocaleString("ar-SA");
    const score = (value) => value == null ? "غير متاح" : value.toFixed(2);
    const metadata = (row) => [analysis.sectionLabel(row.sectionId), row.topic, row.number ? `البند ${row.number}` : ""].filter(Boolean).join(" · ");
    const countLabel = (row, previous) => {
        const value = previous ? row.previous : row.current;
        const count = previous ? row.previousResponseCount : row.responseCount;
        return value == null ? "لا توجد نتيجة مطابقة" : count == null ? "عدد الإجابات غير متاح" : `${number(count)} إجابة`;
    };
    const filename = (report) => `متوسط_جودة_بنود_البرنامج_${report.year}${report.comparison ? "_مقارنة_الشطرين" : ""}`;
    function scopeLabel(report) {
        return [state.program === "all" ? "كل البرامج" : data.programs.find((program) => program.id === state.program)?.label,
            state.degree === "all" ? "كل الدرجات" : state.degree,
            report.comparison ? "مقارنة الطلاب والطالبات والإجمالي" : view.cohortLabel(state.cohort),
            state.query ? `البحث: ${state.query}` : ""].filter(Boolean).join(" · ");
    }
    const options = (values, selected) => values.map(([value, label]) => `<option value="${escape(value)}"${value === selected ? " selected" : ""}>${escape(label)}</option>`).join("");

    function createShell() {
        root.innerHTML = `<article class="card ce-filter-card"><div class="ce-filters-grid pq-filters-grid">
            <div class="field-group"><label for="pqYear">السنة</label><select id="pqYear"></select></div>
            <div class="field-group"><label for="pqDegree">الدرجة</label><select id="pqDegree"></select></div>
            <div class="field-group"><label for="pqProgram">البرنامج</label><select id="pqProgram"></select></div>
            <div class="field-group"><label for="pqCohort">الشطر</label><select id="pqCohort"></select></div>
            <div class="field-group pq-comparison-field"><label class="ce-item-compare" for="pqCompareCohorts"><input id="pqCompareCohorts" type="checkbox">عرض مقارنة الطلاب والطالبات</label></div>
            <div class="field-group ce-search-field"><label for="pqQuery">بحث في البنود والمحاور</label><input id="pqQuery" type="search" placeholder="نص البند أو المحور" autocomplete="off"></div>
            <button class="btn btn-ghost ce-reset" type="button" data-pq-action="reset">إعادة الضبط</button>
        </div></article><div id="pqNotice" class="ce-notice" role="status"></div><div id="pqContent"></div>`;
        root.addEventListener("change", (event) => {
            const key = { pqYear: "year", pqDegree: "degree", pqProgram: "program", pqCohort: "cohort", pqItemSort: "itemSort" }[event.target.id];
            if (event.target.id === "pqCompareCohorts") state.compareCohorts = event.target.checked;
            else if (key) state[key] = event.target.value;
            else return;
            if (key === "degree" && !data.programs.some((program) => program.id === state.program && (state.degree === "all" || program.degree === state.degree))) state.program = "all";
            renderAll();
        });
        root.addEventListener("input", (event) => {
            if (event.target.id !== "pqQuery") return;
            state.query = event.target.value;
            clearTimeout(searchTimer);
            searchTimer = setTimeout(renderContent, 180);
        });
        root.addEventListener("click", (event) => {
            const button = event.target.closest("[data-pq-action]");
            if (!button || button.disabled) return;
            const action = button.dataset.pqAction;
            if (action === "reset") {
                clearTimeout(searchTimer);
                Object.assign(state, { year: data.years[0], degree: "all", program: "all", cohort: "total", compareCohorts: false, query: "", itemSort: "source" });
                renderAll();
                return;
            }
            const report = analysis.buildReport(data, state);
            if (!report.rows.length) return;
            if (action === "xlsx") view.workbook(analysis.exportRows(report, state, data), "جودة بنود البرنامج", `${filename(report)}.xlsx`);
            if (action === "csv") exportCsv(report);
            if (action === "pdf") view.pdf({ report, scope: { ...state }, button, title: "متوسط جودة بنود البرنامج",
                weighting: report.hasUnknownCounts ? "المتوسطات المتاحة؛ أعداد بعض الإجابات غير متاحة" : "متوسطات موزونة بعدد الإجابات",
                metadata, countLabel: (row, previous) => {
                    const value = previous ? row.previous : row.current;
                    const count = previous ? row.previousResponseCount : row.responseCount;
                    return value == null ? "لا توجد نتيجة" : count == null ? "العدد غير متاح" : `${number(count)} إجابة`;
                }, scopeLabel: scopeLabel(report), footer: "استطلاع جودة البرنامج", filename: `${filename(report)}.pdf` });
        });
    }

    function renderAll() {
        const select = (id, values, selected) => { root.querySelector(`#${id}`).innerHTML = options(values, selected); };
        select("pqYear", data.years.map((year) => [year, `${year}هـ`]), state.year);
        select("pqDegree", [["all", "كل الدرجات"], ...[...new Set(data.programs.map((program) => program.degree))].map((degree) => [degree, degree])], state.degree);
        select("pqProgram", [["all", "كل البرامج"], ...data.programs.filter((program) => state.degree === "all" || program.degree === state.degree).map((program) => [program.id, program.label])], state.program);
        select("pqCohort", [["total", "الإجمالي المتاح"], ["male", "الطلاب"], ["female", "الطالبات"]], state.cohort);
        root.querySelector("#pqCohort").disabled = state.compareCohorts;
        root.querySelector("#pqCompareCohorts").checked = state.compareCohorts;
        root.querySelector("#pqQuery").value = state.query;
        renderContent();
    }

    function renderContent() {
        const report = analysis.buildReport(data, state);
        const disabled = report.rows.length ? "" : " disabled";
        const sources = [...new Set(report.rows.map((row) => row.surveyTitle))];
        const summary = report.comparison ? report.cohortReports.map((cohort) => ({
            label: `المتوسط العام للبنود ${report.year}هـ · ${view.cohortLabel(cohort.cohort)}`,
            value: score(cohort.overall), note: `من 5 · ${number(cohort.programCount)} برنامجًا · ${number(cohort.itemCount)} بندًا ذا نتيجة`,
        })) : [
            { label: `المتوسط العام للبنود ${report.year}هـ`, value: score(report.overall), note: "متوسط جميع البنود في النطاق · من 5" },
            { label: "البرامج في النطاق", value: number(report.programCount), note: "البرامج ذات النتائج المتاحة" },
            { label: "البنود ذات النتائج", value: number(report.itemCount), note: `${number(report.visibleItemCount)} بندًا ظاهرًا بعد البحث` },
            { label: "البنود القابلة للمقارنة", value: number(report.comparableItemCount), note: `بين ${report.previousYear} و${report.year}هـ` },
        ];
        root.querySelector("#pqContent").innerHTML = `${view.kpis(summary)}<article class="card">
            <div class="card-head ce-item-head"><div><h3>متوسط جودة بنود البرنامج</h3><p class="meta-text">${escape(scopeLabel(report))}</p></div>
            <div class="action-group"><button class="btn" type="button" data-pq-action="xlsx"${disabled}>Excel</button><button class="btn" type="button" data-pq-action="csv"${disabled}>CSV</button><button class="btn btn-primary" type="button" data-pq-action="pdf"${disabled}>تحميل PDF</button></div></div>
            <div class="ce-item-toolbar"><label for="pqItemSort">ترتيب البنود</label><select id="pqItemSort">${options([
                ["source", "ترتيب الاستبانة"], ["lowest", "الأقل في السنة الحالية"], ["highest", "الأعلى في السنة الحالية"],
                ["change", "الأقل تحسنًا عن السنة الماضية"], ["difference", "الأدنى مقارنة بالمتوسط العام"],
            ], state.itemSort)}</select></div>
            <p class="meta-text ce-item-caption">الفروق بالنقاط على مقياس 5. ${report.comparison ? `${number(report.visibleItemCount)} بندًا ظاهرًا، ولكل بند ثلاث فئات. يُقارن كل صف بالمتوسط العام للفئة نفسها، والترتيب العددي حسب الإجمالي. ` : ""}المتوسط العام لجميع البنود في النطاق المختار، ولا يتغير بالبحث النصي.</p>
            ${sources.length ? `<p class="meta-text">مصدر النتائج: ${escape(sources.join(" · "))}</p>` : ""}
            <div class="table-wrap">${view.table(report, report.rows, { metadata, countLabel })}</div></article>`;
        const notes = ["تُعرض البنود بترتيبها ومحاورها الأصلية. تُقارن السنة الحالية بالسنة الماضية مباشرة، عند تطابق نموذج البند ومحوره ورقمه ونصه؛ وإلا تظهر «غير متاح»."];
        if (report.hasUnknownCounts) notes.push("عند غياب أعداد الإجابات، يُحسب المتوسط من متوسطات الفئات المتاحة في البرامج المختارة بأوزان متساوية؛ ولا تُفترض أعداد للمشاركين.");
        else notes.push("يُحسب متوسط البند من مجموع درجات الإجابات مقسومًا على عددها في البرامج والفئة المختارة؛ والإجمالي يجمع إجابات الطلاب والطالبات المتاحة.");
        if (report.hasExcluded) notes.push("لم تُحتسب النتائج المفقودة أو المتعارضة.");
        root.querySelector("#pqNotice").textContent = notes.join(" ");
    }

    function exportCsv(report) {
        const rows = analysis.exportRows(report, state, data), headers = Object.keys(rows[0]);
        const cell = (value) => `"${String(value ?? "").replace(/"/g, '""')}"`;
        const lines = [headers.map(cell).join(","), ...rows.map((row) => headers.map((key) => cell(row[key])).join(","))];
        const url = URL.createObjectURL(new Blob(["\ufeff", lines.join("\r\n")], { type: "text/csv;charset=utf-8" }));
        const link = document.createElement("a");
        link.href = url; link.download = `${filename(report)}.csv`;
        document.body.appendChild(link); link.click(); link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 0);
    }

    window.ProgramQualityApp = { render({ programs }) {
        if (!analysis || !view) return;
        if (!data) {
            data = analysis.prepareData(window.SURVEYS_DATA || {}, programs);
            state.year = data.years[0] || "";
            root = document.getElementById("programQualityApp");
            if (!root) return;
            createShell();
        }
        renderAll();
    } };
})();
