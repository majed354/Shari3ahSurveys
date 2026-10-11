(function () {
    "use strict";
    const data = window.GRADUATE_SURVEYS_DATA;
    const analysis = window.GraduateSurveysAnalysis;
    const state = { year: "all", degree: "all", program: "all", axis: "all", query: "" };
    let root, timer;
    const escape = value => String(value ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" })[c]);
    const number = value => Number(value).toLocaleString("ar-SA", { maximumFractionDigits: 2 });
    const options = (values, selected) => values.map(([value, label]) => `<option value="${escape(value)}"${value === selected ? " selected" : ""}>${escape(label)}</option>`).join("");
    const programLabel = p => `${p.degree} — ${p.name}`;
    const scope = () => [state.year === "all" ? "كل سنوات التخرج" : `سنة التخرج ${state.year}هـ`, state.degree === "all" ? "كل الدرجات" : state.degree,
        state.program === "all" ? "كل البرامج" : programLabel(data.programs.find(p => p.id === state.program)),
        state.axis === "all" ? "كل المحاور" : state.axis, state.query ? `البحث: ${state.query}` : ""].filter(Boolean).join(" · ");
    const filename = ext => `استطلاعات_الخريجين_الخارجية_${state.year === "all" ? "كل_السنوات" : state.year}.${ext}`;
    const result = row => row.average != null ? `${number(row.average)} من ${number(row.scale)}${row.kind === "range" ? " (تقديري)" : ""}`
        : row.kind === "text" ? `${number(row.responseCount)} إجابة نصية` : row.responseCount ? "توزيع الإجابات" : "لا توجد إجابات";

    function createShell() {
        root.innerHTML = `<article class="card gs-intro"><p>${escape(data.description)}</p><p class="meta-text">${escape(data.axisNote)}</p>
            <p class="meta-text">التصفية بحسب سنة التخرج. متوسطات الفئات الرقمية تقديرية باستخدام منتصف كل فئة. تُحسب نسب الخيارات من إجابات البند، و«دون إجابة» تشمل الأسئلة غير المنطبقة. تُعرض التعليقات المفتوحة بعددها دون نشر نصوص فردية.</p></article>
            <article class="card"><div class="ce-filters-grid gs-filters-grid">
            <div class="field-group"><label for="gsYear">سنة التخرج</label><select id="gsYear"></select></div>
            <div class="field-group"><label for="gsDegree">الدرجة</label><select id="gsDegree"></select></div>
            <div class="field-group"><label for="gsProgram">البرنامج</label><select id="gsProgram"></select></div>
            <div class="field-group"><label for="gsAxis">المحور</label><select id="gsAxis"></select></div>
            <div class="field-group ce-search-field"><label for="gsQuery">بحث في البنود والمحاور</label><input id="gsQuery" type="search" placeholder="نص البند أو المحور"></div>
            <button type="button" class="btn btn-ghost" data-gs-action="reset">إعادة الضبط</button></div></article>
            <div class="card gs-toolbar"><p id="gsScope" class="meta-text"></p><div class="action-group">
            <button type="button" class="btn" data-gs-action="xlsx">Excel</button><button type="button" class="btn" data-gs-action="csv">CSV</button><button type="button" class="btn btn-primary" data-gs-action="pdf">تحميل PDF</button></div></div>
            <p id="gsStatus" class="meta-text" role="status"></p><div id="gsContent"></div>`;
        root.addEventListener("change", event => {
            const key = { gsYear: "year", gsDegree: "degree", gsProgram: "program", gsAxis: "axis" }[event.target.id];
            if (!key) return;
            state[key] = event.target.value;
            renderControls(); renderContent();
        });
        root.addEventListener("input", event => {
            if (event.target.id !== "gsQuery") return;
            state.query = event.target.value; clearTimeout(timer); timer = setTimeout(renderContent, 150);
        });
        root.addEventListener("click", async event => {
            const button = event.target.closest("[data-gs-action]");
            if (!button || button.disabled) return;
            const action = button.dataset.gsAction;
            if (action === "reset") { clearTimeout(timer); Object.assign(state, { year: "all", degree: "all", program: "all", axis: "all", query: "" }); renderControls(); renderContent(); return; }
            const report = analysis.buildReport(data, state);
            if (!report.rows.length) return;
            const rows = analysis.exportRows(report, scope());
            if (action === "xlsx") window.SurveyItemReportView.workbook(rows, "استطلاعات الخريجين الخارجية", filename("xlsx"));
            if (action === "csv") exportCsv(rows);
            if (action === "pdf") await exportPdf(report, button);
        });
    }

    function renderControls() {
        const select = (id, values, selected) => { root.querySelector(`#${id}`).innerHTML = options(values, selected); };
        select("gsYear", [["all", "كل سنوات التخرج"], ...data.years.map(y => [y, `${y}هـ`])], state.year);
        select("gsDegree", [["all", "كل الدرجات"], ...[...new Set(data.programs.map(p => p.degree))].map(d => [d, d])], state.degree);
        const available = data.programs.filter(p => (state.degree === "all" || p.degree === state.degree)
            && data.groups.some(g => g.programId === p.id && (state.year === "all" || g.year === state.year)));
        if (state.program !== "all" && !available.some(p => p.id === state.program)) state.program = "all";
        select("gsProgram", [["all", "كل البرامج"], ...available.map(p => [p.id, programLabel(p)])], state.program);
        const report = analysis.buildReport(data, { ...state, axis: "all", query: "" });
        const axes = [...new Set(report.rows.map(r => r.axis))];
        if (state.axis !== "all" && !axes.includes(state.axis)) state.axis = "all";
        select("gsAxis", [["all", "كل المحاور"], ...axes.map(a => [a, a])], state.axis);
        root.querySelector("#gsQuery").value = state.query;
    }

    function rowMarkup(row) {
        const distribution = row.categories.length ? `<details><summary>تفاصيل الإجابات (${number(row.categories.length)})</summary><ul class="gs-distribution">${row.categories.map(c => `<li><span>${escape(c.label)}</span><strong>${number(c.count)} (${number(c.percentage)}٪)</strong></li>`).join("")}</ul></details>`
            : row.kind === "text" ? "إجابات نصية مفتوحة؛ يُعرض عددها فقط" : "لا توجد إجابات";
        return `<tr><td>${number(row.number)}</td><th scope="row">${escape(row.label)}</th><td><strong>${result(row)}</strong></td><td>${number(row.responseCount)}<small class="gs-count-note">دون إجابة: ${number(row.missingCount)}</small></td><td>${distribution}</td></tr>`;
    }

    function renderContent() {
        const report = analysis.buildReport(data, state);
        root.querySelector("#gsScope").textContent = scope();
        root.querySelectorAll("[data-gs-action]").forEach(b => { if (b.dataset.gsAction !== "reset") b.disabled = !report.rows.length; });
        root.querySelector("#gsStatus").textContent = `${number(report.respondentCount)} استجابة في النطاق · ${number(report.rows.length)} بندًا ظاهرًا`;
        const kpis = window.SurveyItemReportView.kpis([
            { label: "الاستجابات في النطاق", value: number(report.respondentCount), note: "عدد الاستجابات المسجلة في الشيتين" },
            { label: "البرامج ذات الاستجابات", value: number(report.programCount), note: "حسب الدرجة وسنة التخرج المختارتين" },
            { label: "البنود الظاهرة", value: number(report.rows.length), note: "بحسب المحور والبحث" },
        ]);
        root.querySelector("#gsContent").innerHTML = kpis + (report.rows.length ? report.sections.filter(s => s.rows.length).map(section => `<article class="card gs-questionnaire"><div class="card-head"><div><h3>استطلاع خريجي ${escape(section.label)}</h3><p class="meta-text">${number(section.respondentCount)} استجابة · المصدر: ${escape(section.sourceFile)}</p></div><span class="status-pill">استطلاع عينة</span></div>
            ${[...new Set(section.rows.map(r => r.axis))].map(axis => `<section class="gs-axis"><h4>${escape(axis)}</h4><div class="table-wrap"><table class="data-table gs-table"><thead><tr><th>البند</th><th>نص السؤال</th><th>النتيجة</th><th>عدد الإجابات</th><th>توزيع الإجابات</th></tr></thead><tbody>${section.rows.filter(r => r.axis === axis).map(rowMarkup).join("")}</tbody></table></div></section>`).join("")}</article>`).join("") : `<article class="card empty-state">لا توجد بنود مطابقة للفلاتر والبحث. غيّر الاختيارات أو أعد الضبط.</article>`);
    }

    function exportCsv(rows) {
        const cell = value => { let text = String(value ?? ""); if (typeof value === "string" && /^[=+@-]/.test(text)) text = `'${text}`; return `"${text.replace(/"/g, '""')}"`; };
        const csv = [Object.keys(rows[0]), ...rows.map(Object.values)].map(row => row.map(cell).join(",")).join("\r\n");
        const url = URL.createObjectURL(new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8" }));
        const link = document.createElement("a"); link.href = url; link.download = filename("csv"); link.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    async function exportPdf(report, button) {
        if (!window.jspdf?.jsPDF) { root.querySelector("#gsStatus").textContent = "تعذر تحميل أداة PDF. أعد تحميل الصفحة ثم حاول مجددًا."; return; }
        button.disabled = true; button.textContent = "جارٍ التوليد…";
        try {
            await document.fonts.ready;
            const pdf = new window.jspdf.jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
            let canvas, ctx, y, page = 0;
            const draw = (line, size = 23, bold = false, color = "#163e39") => {
                ctx.font = `${bold ? 700 : 400} ${size}px Tajawal, sans-serif`; ctx.fillStyle = color; ctx.textAlign = "right"; ctx.direction = "rtl"; ctx.fillText(line, 740, y); y += size + 12;
            };
            const wrap = (text, size = 23) => {
                ctx.font = `400 ${size}px Tajawal, sans-serif`; const lines = []; let line = "";
                String(text).split(/\s+/).forEach(word => { const next = line ? `${line} ${word}` : word; if (line && ctx.measureText(next).width > 680) { lines.push(line); line = word; } else line = next; });
                if (line) lines.push(line); return lines;
            };
            const start = () => { canvas = document.createElement("canvas"); canvas.width = 1600; canvas.height = 2260; ctx = canvas.getContext("2d"); ctx.scale(2, 2); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, 800, 1130); y = 65; draw(data.title, 30, true); wrap(scope(), 19).forEach(line => draw(line, 19)); y += 15; };
            const finish = () => { ctx.fillStyle = "#607d76"; ctx.font = "18px Tajawal, sans-serif"; ctx.fillText(`استطلاع عينة · ${++page}`, 740, 1095); if (page > 1) pdf.addPage(); pdf.addImage(canvas.toDataURL("image/jpeg", 0.93), "JPEG", 0, 0, 210, 297); };
            const lines = (text, size = 23, bold = false) => wrap(text, size).forEach(line => { if (y + size + 12 > 1050) { finish(); start(); } draw(line, size, bold); });
            start(); lines(data.description, 19); lines(data.axisNote, 19); lines("سنة التخرج هي السنة المسجلة في الشيت. متوسط الفئات تقديري باستخدام منتصفها. النسب من إجابات البند. التعليقات المفتوحة معروضة بالعدد فقط.", 19); y += 15;
            report.sections.filter(s => s.rows.length).forEach(section => {
                if (y > 850) { finish(); start(); }
                lines(`استطلاع خريجي ${section.label} · ${number(section.respondentCount)} استجابة`, 26, true);
                lines(`المصدر: ${section.sourceFile}`, 18);
                [...new Set(section.rows.map(row => row.axis))].forEach(axis => {
                    if (y > 850) { finish(); start(); }
                    y += 12; lines(axis, 25, true);
                    section.rows.filter(row => row.axis === axis).forEach(row => {
                        if (y > 900) { finish(); start(); }
                        y += 8; lines(`${number(row.number)}. ${row.label}`, 23, true);
                        lines(`${result(row)} · ${number(row.responseCount)} إجابة · دون إجابة: ${number(row.missingCount)}`, 21);
                        row.categories.forEach(c => lines(`${c.label}: ${number(c.count)} (${number(c.percentage)}٪)`, 20));
                    });
                });
            });
            finish(); pdf.save(filename("pdf"));
            root.querySelector("#gsStatus").textContent = "تم إعداد ملف PDF وفق الفلاتر الحالية.";
        } catch (error) { console.error(error); root.querySelector("#gsStatus").textContent = "تعذر إعداد PDF. حاول مرة أخرى."; }
        finally { button.disabled = false; button.textContent = "تحميل PDF"; }
    }

    window.GraduateSurveysApp = { render() { if (!data || !analysis) return; if (!root) { root = document.getElementById("graduateSurveysApp"); if (!root) return; createShell(); } renderControls(); renderContent(); } };
})();
