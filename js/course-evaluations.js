(function () {
    "use strict";

    const packageData = window.COURSE_EVALUATIONS_DATA || {};
    const meta = packageData.meta || {};
    const courseRecords = Array.isArray(packageData.courseRecords) ? packageData.courseRecords : [];
    const publicExcellenceLeaders = Array.isArray(packageData.publicExcellenceLeaders) ? packageData.publicExcellenceLeaders : [];
    const outcomeCategories = Array.isArray(meta.outcomeCategories) ? meta.outcomeCategories : [];
    const PAGE_SIZE = 75;
    const LOCK_AFTER_MS = 15 * 60 * 1000;

    const state = {
        initialized: false,
        tab: "annual",
        year: (meta.years || [])[0] || "all",
        semester: "all",
        degree: "all",
        program: "all",
        department: "all",
        departmentAuto: true,
        query: "",
        page: 1,
        privatePayload: null,
        showDescriptive: false,
        lockTimer: null,
    };

    const $ = (selector, root = document) => root.querySelector(selector);

    function escapeHtml(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function normalizeDigits(value) {
        return String(value || "")
            .replace(/[٠-٩]/g, (digit) => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
            .replace(/[۰-۹]/g, (digit) => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
            .trim();
    }

    function score(value) {
        return Number.isFinite(Number(value)) ? Number(value).toFixed(2) : "—";
    }

    function scoreTone(value) {
        const number = Number(value);
        if (number >= 4.5) return "excellent";
        if (number >= 4) return "good";
        if (number >= 3.5) return "fair";
        return "low";
    }

    function scorePill(value) {
        return `<span class="ce-score ce-score--${scoreTone(value)}">${score(value)}</span>`;
    }

    function optionList(values, selected, allLabel = "الكل") {
        const options = [`<option value="all"${selected === "all" ? " selected" : ""}>${escapeHtml(allLabel)}</option>`];
        values.forEach((value) => {
            options.push(`<option value="${escapeHtml(value)}"${String(selected) === String(value) ? " selected" : ""}>${escapeHtml(value)}</option>`);
        });
        return options.join("");
    }

    function defaultDepartmentForProgram(program) {
        if (!program || program === "all") return "all";
        return meta.programDefaultDepartments?.[program] || "all";
    }

    function availableSemesters() {
        if (state.year === "all") return [];
        const values = meta.semestersByYear?.[state.year];
        return Array.isArray(values) ? values : [];
    }

    function semesterFieldVisible() {
        return state.tab === "annual" || state.tab === "indirect";
    }

    function ensureSemesterSelection() {
        const semesters = availableSemesters();
        if (!semesters.includes(state.semester)) state.semester = "all";
    }

    function activeCourseSemester() {
        if (state.year === "all") return "all";
        return state.semester === "all" ? "all" : state.semester;
    }

    function coursePeriodLabel() {
        if (state.year === "all") return "كل السنوات";
        if (state.semester === "all") return `${state.year}هـ - كل الفصول`;
        return `${state.year}هـ - الفصل ${state.semester}`;
    }

    function createShell(root) {
        root.innerHTML = `
            <div class="ce-tabs" role="tablist" aria-label="تقارير تقييم المقررات">
                <button class="ce-tab is-active" type="button" data-ce-tab="annual">التقرير السنوي</button>
                <button class="ce-tab" type="button" data-ce-tab="indirect">التقييم غير المباشر</button>
                <button class="ce-tab" type="button" data-ce-tab="excellence">التميز في التدريس</button>
                <button class="ce-tab" type="button" data-ce-tab="methodology">المنهجية والعدالة</button>
            </div>

            <article id="ceFilters" class="card ce-filter-card">
                <div class="ce-filters-grid">
                    <div class="field-group">
                        <label for="ceYear">السنة</label>
                        <select id="ceYear">${optionList(meta.years || [], state.year, "كل السنوات")}</select>
                    </div>
                    <div class="field-group" data-ce-semester-field>
                        <label for="ceSemester">الفصل</label>
                        <select id="ceSemester">${optionList(availableSemesters(), state.semester, "كل الفصول")}</select>
                    </div>
                    <div class="field-group">
                        <label for="ceDegree">الدرجة</label>
                        <select id="ceDegree">${optionList(meta.degrees || [], state.degree)}</select>
                    </div>
                    <div class="field-group">
                        <label for="ceProgram">البرنامج</label>
                        <select id="ceProgram">${optionList(meta.programs || [], state.program)}</select>
                    </div>
                    <div class="field-group hidden" data-ce-department-field>
                        <label for="ceDepartment">قسم عضو هيئة التدريس</label>
                        <select id="ceDepartment">${optionList(meta.departments || [], state.department, "كل الأقسام")}</select>
                    </div>
                    <div class="field-group ce-search-field">
                        <label for="ceQuery">بحث</label>
                        <input id="ceQuery" type="search" placeholder="اسم المقرر أو رمزه" autocomplete="off">
                    </div>
                    <button class="btn btn-ghost ce-reset" type="button" data-ce-action="reset">إعادة الضبط</button>
                </div>
            </article>

            <div id="ceNotice" class="ce-notice" role="status"></div>
            <div id="ceContent"></div>
        `;
        bindEvents(root);
        state.initialized = true;
    }

    function bindEvents(root) {
        root.addEventListener("click", async (event) => {
            const tab = event.target.closest("[data-ce-tab]");
            if (tab) {
                state.tab = tab.dataset.ceTab;
                if (state.tab === "excellence" && state.departmentAuto) {
                    state.department = defaultDepartmentForProgram(state.program);
                }
                state.page = 1;
                renderAll();
                return;
            }

            const action = event.target.closest("[data-ce-action]")?.dataset.ceAction;
            if (!action) return;
            if (action === "reset") resetFilters();
            if (action === "prev" && state.page > 1) { state.page -= 1; renderContent(); }
            if (action === "next") { state.page += 1; renderContent(); }
            if (action === "unlock") await unlockExcellence();
            if (action === "lock") lockExcellence("تم إغلاق القائمة الكاملة والتفاصيل الإدارية.");
            if (action === "export-annual") exportAnnual();
            if (action === "export-indirect") exportIndirect();
            if (action === "export-excellence") exportExcellence();
            if (action === "csv-public-excellence") exportPublicExcellenceCsv();
            if (action === "pdf-public-excellence") await exportPublicExcellencePdf();
            if (action === "pdf-excellence") await exportExcellencePdf();
        });

        root.addEventListener("change", (event) => {
            if (event.target.id === "ceProgram") {
                state.program = event.target.value;
                state.department = defaultDepartmentForProgram(state.program);
                state.departmentAuto = true;
                state.page = 1;
                renderAll();
                return;
            }
            if (event.target.id === "ceYear") {
                state.year = event.target.value;
                state.semester = "all";
                state.page = 1;
                renderAll();
                return;
            }
            if (event.target.id === "ceSemester") {
                state.semester = event.target.value;
                state.page = 1;
                renderContent();
                return;
            }
            if (event.target.id === "ceDepartment") {
                state.department = event.target.value;
                state.departmentAuto = false;
                state.page = 1;
                renderContent();
                return;
            }
            const map = { ceDegree: "degree" };
            if (map[event.target.id]) {
                state[map[event.target.id]] = event.target.value;
                state.page = 1;
                renderContent();
            }
            if (event.target.id === "ceShowDescriptive") {
                state.showDescriptive = event.target.checked;
                state.page = 1;
                renderContent();
            }
        });

        let searchTimer;
        root.addEventListener("input", (event) => {
            if (event.target.id !== "ceQuery") return;
            clearTimeout(searchTimer);
            searchTimer = setTimeout(() => {
                state.query = event.target.value.trim();
                state.page = 1;
                renderContent();
            }, 180);
        });

        root.addEventListener("keydown", async (event) => {
            if (event.target.id === "ceAdminPin" && event.key === "Enter") {
                event.preventDefault();
                await unlockExcellence();
            }
        });
    }

    function resetFilters() {
        state.year = (meta.years || [])[0] || "all";
        state.semester = "all";
        state.degree = "all";
        state.program = "all";
        state.department = "all";
        state.departmentAuto = true;
        state.query = "";
        state.page = 1;
        renderAll();
    }

    function renderAll() {
        const root = $("#courseEvaluationsApp");
        if (!root) return;
        root.querySelectorAll("[data-ce-tab]").forEach((button) => {
            const active = button.dataset.ceTab === state.tab;
            button.classList.toggle("is-active", active);
            button.setAttribute("aria-selected", active ? "true" : "false");
        });
        ensureSemesterSelection();
        $("#ceYear", root).innerHTML = optionList(meta.years || [], state.year, "كل السنوات");
        const semesterField = $("[data-ce-semester-field]", root);
        const semesterSelect = $("#ceSemester", root);
        const semesterOptions = availableSemesters();
        semesterSelect.innerHTML = optionList(semesterOptions, state.semester, "كل الفصول");
        semesterSelect.disabled = !semesterOptions.length;
        if (semesterField) semesterField.classList.toggle("hidden", !semesterFieldVisible());
        $("#ceDegree", root).innerHTML = optionList(meta.degrees || [], state.degree);
        $("#ceProgram", root).innerHTML = optionList(meta.programs || [], state.program);
        $("#ceDepartment", root).innerHTML = optionList(meta.departments || [], state.department, "كل الأقسام");
        const departmentField = $("[data-ce-department-field]", root);
        if (departmentField) departmentField.classList.toggle("hidden", state.tab !== "excellence");
        const queryInput = $("#ceQuery", root);
        queryInput.value = state.query;
        queryInput.placeholder = state.tab === "excellence" ? "اسم عضو هيئة التدريس" : "اسم المقرر أو رمزه";
        const filters = $("#ceFilters", root);
        if (filters) filters.classList.toggle("hidden", state.tab === "methodology");
        renderContent();
    }

    function filterCourseRows() {
        const query = state.query.toLocaleLowerCase("ar");
        const semester = activeCourseSemester();
        return courseRecords.filter((row) => {
            if (state.year !== "all" && String(row.year) !== state.year) return false;
            if (String(row.semester || "all") !== semester) return false;
            if (state.degree !== "all" && String(row.degree) !== state.degree) return false;
            if (state.program !== "all" && String(row.program) !== state.program) return false;
            if (query && ![row.courseName, row.courseCode, row.program].join(" ").toLocaleLowerCase("ar").includes(query)) return false;
            return true;
        });
    }

    function filterFacultyRows() {
        if (!state.privatePayload) return [];
        const query = state.query.toLocaleLowerCase("ar");
        return state.privatePayload.facultyRecords.filter((row) => {
            const targetYear = state.year === "all" ? "all" : state.year;
            if (String(row.year) !== targetYear) return false;
            if (state.degree !== "all" && String(row.degree) !== state.degree) return false;
            if (state.program !== "all" && String(row.program) !== state.program) return false;
            if (state.department !== "all" && String(row.facultyDepartment) !== state.department) return false;
            if (!state.showDescriptive && !row.eligible) return false;
            if (query && ![row.instructorName, row.program, row.facultyDepartment].join(" ").toLocaleLowerCase("ar").includes(query)) return false;
            return true;
        });
    }

    function filterPublicLeaderRows() {
        const query = state.query.toLocaleLowerCase("ar");
        const targetYear = state.year === "all" ? "all" : state.year;
        return publicExcellenceLeaders.filter((row) => {
            if (String(row.year) !== targetYear) return false;
            if (state.degree !== "all" && String(row.degree) !== state.degree) return false;
            if (state.program !== "all" && String(row.program) !== state.program) return false;
            if (String(row.selectionDepartment) !== state.department) return false;
            if (query && ![row.instructorName, row.program, row.facultyDepartment].join(" ").toLocaleLowerCase("ar").includes(query)) return false;
            return true;
        });
    }

    function renderContent() {
        const content = $("#ceContent");
        if (!content) return;
        if (state.tab === "annual") renderAnnual(content);
        if (state.tab === "indirect") renderIndirect(content);
        if (state.tab === "excellence") renderExcellence(content);
        if (state.tab === "methodology") renderMethodology(content);
    }

    function kpis(items) {
        return `<div class="kpi-grid ce-kpis">
            ${items.map((item) => `<article class="card kpi-card"><span class="kpi-label">${escapeHtml(item.label)}</span><strong class="kpi-value">${escapeHtml(item.value)}</strong><small>${escapeHtml(item.note || "")}</small></article>`).join("")}
        </div>`;
    }

    function pageRows(rows) {
        const pages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
        state.page = Math.min(state.page, pages);
        const start = (state.page - 1) * PAGE_SIZE;
        return { rows: rows.slice(start, start + PAGE_SIZE), pages, start };
    }

    function pagination(total, pages) {
        if (pages <= 1) return "";
        return `<div class="ce-pagination">
            <button class="btn" type="button" data-ce-action="prev"${state.page === 1 ? " disabled" : ""}>السابق</button>
            <span>صفحة ${state.page} من ${pages} · ${total.toLocaleString("ar-SA")} سجل</span>
            <button class="btn" type="button" data-ce-action="next"${state.page === pages ? " disabled" : ""}>التالي</button>
        </div>`;
    }

    function renderAnnual(content) {
        const rows = filterCourseRows();
        const average = rows.length ? rows.reduce((sum, row) => sum + Number(row.score), 0) / rows.length : 0;
        const respondents = rows.reduce((sum, row) => sum + Number(row.respondents || 0), 0);
        const page = pageRows(rows);
        content.innerHTML = `
            ${kpis([
                { label: "المقررات", value: rows.length.toLocaleString("ar-SA"), note: "بحسب الفلاتر الحالية" },
                { label: "الطلاب المقيمون", value: respondents.toLocaleString("ar-SA"), note: "بعد جمع شعب المقرر في السنة" },
                { label: "متوسط التقييم", value: score(average), note: "من 5" },
                { label: "ترميز المقررات", value: rows.filter((row) => row.courseCode).length.toLocaleString("ar-SA"), note: "المطابق بثقة" },
            ])}
            <article class="card">
                <div class="card-head"><div><h3>جدول التقييم للتقرير السنوي</h3><p class="meta-text">تُجمع جميع شعب المقرر في السنة المختارة، ونسبة المشاركة 100٪.</p></div><button class="btn btn-primary" type="button" data-ce-action="export-annual">تصدير Excel</button></div>
                <div class="table-wrap"><table class="data-table ce-wide-table"><thead><tr>
                    <th>رمز المقرر</th><th>اسم المقرر</th><th>عدد الطلاب الذين قيموا المقرر</th><th>نسبة المشاركين</th><th>نتيجة التقييم</th><th>التوصيات التطويرية</th>
                </tr></thead><tbody>
                    ${page.rows.map((row) => `<tr><td>${escapeHtml(row.courseCode || "—")}</td><td><strong>${escapeHtml(row.courseName)}</strong></td><td>${Number(row.respondents).toLocaleString("ar-SA")}</td><td>100٪</td><td>${scorePill(row.score)}</td><td>${escapeHtml(row.recommendation)}</td></tr>`).join("") || `<tr><td colspan="6"><div class="empty-state">لا توجد بيانات مطابقة.</div></td></tr>`}
                </tbody></table></div>
                ${pagination(rows.length, page.pages)}
            </article>`;
        setNotice(`الفترة الحالية: ${coursePeriodLabel()}. تُجمع شعب المقرر للعضو في السنة والبرنامج أولًا، ثم تُستبعد المجموعة إذا بقي مجموعها طالبًا أو طالبين.`, "info");
    }

    function renderIndirect(content) {
        const rows = filterCourseRows();
        if (!outcomeCategories.length) {
            content.innerHTML = `
                <article class="card ce-lock-card">
                    <div class="ce-lock-icon" aria-hidden="true">⚠️</div>
                    <div><h3>يلزم استكمال خريطة تصنيف بنود التقييم</h3><p>الملف التفصيلي يحتوي البنود ودرجاتها، والكود المرفق يقرأ تصنيفًا جاهزًا باسم «جودة المقرر/معرفي/مهاري/قيمي» من عمود غير موجود في ملف المصدر. لذلك أوقفت الجدول القديم ذي الجوانب السبعة حتى لا يعرض تصنيفًا غير مطابق.</p></div>
                </article>`;
            setNotice("يلزم ملف أو كود يربط رقم البند أو نصه بتصنيفه قبل حساب المتوسطات المعرفية والمهارية والقيمية بدقة.", "info");
            return;
        }
        const page = pageRows(rows);
        content.innerHTML = `
            <article class="card">
                <div class="card-head"><div><h3>التقييم غير المباشر للمقررات</h3><p class="meta-text">متوسطات التصنيف المعرفي والمهاري والقيمي وما بقي بلا تصنيف، بعد جمع شعب المقرر في السنة.</p></div><button class="btn btn-primary" type="button" data-ce-action="export-indirect">تصدير Excel</button></div>
                <div class="table-wrap"><table class="data-table ce-wide-table"><thead><tr>
                    <th>رمز المقرر</th><th>اسم المقرر</th><th>عدد الطلاب</th><th>جودة المقرر</th>${outcomeCategories.map((item) => `<th>${escapeHtml(item.label)}</th>`).join("")}
                </tr></thead><tbody>
                    ${page.rows.map((row) => `<tr><td>${escapeHtml(row.courseCode || "—")}</td><td><strong>${escapeHtml(row.courseName)}</strong></td><td>${Number(row.respondents).toLocaleString("ar-SA")}</td><td>${scorePill(row.score)}</td>${outcomeCategories.map((item) => `<td>${row.outcomeScores?.[item.id] == null ? "—" : scorePill(row.outcomeScores[item.id])}</td>`).join("")}</tr>`).join("") || `<tr><td colspan="${4 + outcomeCategories.length}"><div class="empty-state">لا توجد بيانات مطابقة.</div></td></tr>`}
                </tbody></table></div>
                ${pagination(rows.length, page.pages)}
            </article>`;
        setNotice("الشرطة (—) تعني عدم وجود بنود من التصنيف المحدد في نموذج السنة المختارة.", "info");
    }

    function rankFaculty(rows) {
        const grouped = new Map();
        rows.forEach((row) => {
            const key = `${row.degree}||${row.program}`;
            if (!grouped.has(key)) grouped.set(key, []);
            grouped.get(key).push(row);
        });
        const ranked = [];
        [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0], "ar")).forEach(([, group]) => {
            group.sort((a, b) => Number(b.scoreWeighted) - Number(a.scoreWeighted));
            let lastScore = null;
            let lastRank = 0;
            let eligiblePosition = 0;
            group.forEach((row) => {
                const eligible = Boolean(row.eligible);
                const currentScore = Number(row.scoreWeighted);
                if (eligible) eligiblePosition += 1;
                if (eligible && currentScore !== lastScore) lastRank = eligiblePosition;
                if (eligible) lastScore = currentScore;
                ranked.push({ ...row, rank: eligible ? lastRank : "—" });
            });
        });
        return ranked;
    }

    function excellenceSummary(rows) {
        const eligible = rows.filter((row) => row.eligible);
        const groups = new Set(eligible.map((row) => `${row.degree}||${row.program}`));
        const eligibleResponses = eligible.reduce((sum, row) => sum + Number(row.respondents || 0), 0);
        const average = eligibleResponses
            ? eligible.reduce((sum, row) => sum + Number(row.scoreWeighted) * Number(row.respondents || 0), 0) / eligibleResponses
            : 0;
        const distribution = [
            { label: "ممتاز", range: "4.50 فأعلى", test: (value) => value >= 4.5 },
            { label: "جيد جداً", range: "4.00–4.49", test: (value) => value >= 4 && value < 4.5 },
            { label: "جيد", range: "3.50–3.99", test: (value) => value >= 3.5 && value < 4 },
            { label: "مقبول", range: "أقل من 3.50", test: (value) => value < 3.5 },
        ].map((item) => {
            const count = eligible.filter((row) => item.test(Number(row.scoreWeighted))).length;
            return { ...item, count, percent: eligible.length ? count / eligible.length * 100 : 0 };
        });
        const singleGroup = groups.size === 1;
        const topFive = singleGroup ? eligible.slice(0, 5) : [];

        return `
            <div class="ce-report-grid">
                <section class="ce-report-block">
                    <h4>التوزيع على فئات التميز</h4>
                    <div class="table-wrap"><table class="data-table"><thead><tr><th>الفئة</th><th>النطاق</th><th>العدد</th><th>النسبة</th></tr></thead><tbody>
                        ${distribution.map((item) => `<tr><td><strong>${item.label}</strong></td><td>${item.range}</td><td>${item.count.toLocaleString("ar-SA")}</td><td>${item.percent.toFixed(1)}٪</td></tr>`).join("")}
                    </tbody></table></div>
                </section>
                <section class="ce-report-block">
                    <h4>الإحصاءات الإجمالية</h4>
                    <div class="table-wrap"><table class="data-table"><tbody>
                        <tr><th>الأعضاء المؤهلون</th><td>${eligible.length.toLocaleString("ar-SA")}</td></tr>
                        <tr><th>الشعب المقيمة</th><td>${eligible.reduce((sum, row) => sum + Number(row.sections), 0).toLocaleString("ar-SA")}</td></tr>
                        <tr><th>الاستجابات</th><td>${eligible.reduce((sum, row) => sum + Number(row.respondents), 0).toLocaleString("ar-SA")}</td></tr>
                        <tr><th>المتوسط الموزون لجميع الاستجابات</th><td>${score(average)} من 5</td></tr>
                    </tbody></table></div>
                </section>
            </div>
            ${singleGroup ? `
                <section class="ce-report-block">
                    <h4>الأعضاء الخمسة الأعلى في المجموع الكلي</h4>
                    <div class="table-wrap"><table class="data-table"><thead><tr><th>#</th><th>اسم العضو</th><th>الشعب المحتسبة</th><th>المقررات</th><th>الاستجابات المحتسبة</th><th>الدرجة الموزونة</th></tr></thead><tbody>
                        ${topFive.map((row) => `<tr><td>${row.rank}</td><td><strong>${escapeHtml(row.instructorName)}</strong></td><td>${Number(row.sections).toLocaleString("ar-SA")}</td><td>${Number(row.courses).toLocaleString("ar-SA")}</td><td>${Number(row.respondents).toLocaleString("ar-SA")}</td><td>${scorePill(row.scoreWeighted)}</td></tr>`).join("")}
                    </tbody></table></div>
                </section>
                <section class="ce-report-block">
                    <h4>الأعلى خمسة في كل جانب من جوانب التدريس</h4>
                    <div class="table-wrap"><table class="data-table ce-aspects-table"><thead><tr><th>الجانب</th><th>الأعضاء الأعلى</th></tr></thead><tbody>
                        ${(meta.dimensions || []).map((item) => {
                            const leaders = [...eligible]
                                .filter((row) => row.dimensionScores?.[item.id] != null)
                                .sort((a, b) => Number(b.dimensionScores[item.id]) - Number(a.dimensionScores[item.id]))
                                .slice(0, 5);
                            return `<tr><td><strong>${escapeHtml(item.label)}</strong></td><td><div class="ce-top-list">${leaders.map((row, index) => `<span>${index + 1}. ${escapeHtml(row.instructorName)} (${score(row.dimensionScores[item.id])})</span>`).join("") || "—"}</div></td></tr>`;
                        }).join("")}
                    </tbody></table></div>
                </section>` : `
                <div class="ce-report-hint">اختر برنامجاً واحداً لإظهار جدولي الخمسة الأعلى والأعلى في كل جانب؛ المقارنة بين برامج أو درجات مختلفة غير معتمدة.</div>`}
        `;
    }

    function publicExcellenceMarkup(rows) {
        const periodLabel = state.year === "all" ? "تجميع السنوات" : `${state.year}هـ`;
        const programLabel = state.program === "all" ? "كل البرامج" : state.program;
        const departmentLabel = state.department === "all" ? "كل الأقسام" : state.department;
        const canPrint = state.program !== "all" && rows.length > 0;
        return `
            <div class="ce-private-toolbar ce-public-toolbar">
                <div><strong>قائمة عامة</strong><p class="meta-text">تقتصر على الخمسة الأعلى، ولا تعرض القائمة الكاملة أو الجوانب التفصيلية.</p></div>
                <div class="action-group">
                    <button class="btn" type="button" data-ce-action="csv-public-excellence"${canPrint ? "" : " disabled"}>تصدير CSV</button>
                    <button class="btn btn-primary" type="button" data-ce-action="pdf-public-excellence"${canPrint ? "" : " disabled"}>طباعة / PDF</button>
                </div>
            </div>
            <article id="cePublicTopFiveReport" class="card ce-private-report ce-public-report">
                <div class="card-head">
                    <div><p class="hero-kicker ce-report-kicker">كلية الشريعة والأنظمة</p><h3>الخمسة المتميزون في التدريس</h3><p class="meta-text">${escapeHtml(programLabel)} · ${escapeHtml(departmentLabel)} · ${escapeHtml(periodLabel)}. الترتيب مستقل داخل كل برنامج ودرجة.</p></div>
                    <span class="ce-public-badge">متاح للطباعة</span>
                </div>
                <section class="ce-report-block">
                    <div class="table-wrap"><table class="data-table ce-wide-table"><thead><tr>
                        <th>الترتيب</th><th>عضو هيئة التدريس</th><th>القسم</th><th>البرنامج</th><th>الدرجة</th><th>الدرجة الموزونة</th><th>التقدير</th><th>الاستجابات المحتسبة</th><th>الشعب المحتسبة</th><th>المقررات</th>
                    </tr></thead><tbody>
                        ${rows.map((row) => `<tr><td><strong>${escapeHtml(row.rank)}</strong></td><td><strong>${escapeHtml(row.instructorName)}</strong></td><td>${escapeHtml(row.facultyDepartment)}</td><td>${escapeHtml(row.program)}</td><td>${escapeHtml(row.degree)}</td><td>${scorePill(row.scoreWeighted)}</td><td>${escapeHtml(row.grade)}</td><td>${Number(row.respondents).toLocaleString("ar-SA")}</td><td>${Number(row.sections).toLocaleString("ar-SA")}</td><td>${Number(row.courses).toLocaleString("ar-SA")}</td></tr>`).join("") || `<tr><td colspan="10"><div class="empty-state">لا توجد قائمة خمسة مطابقة للفلاتر الحالية.</div></td></tr>`}
                    </tbody></table></div>
                </section>
                ${state.program === "all" ? `<div class="ce-report-hint">اختر برنامجاً واحداً لتفعيل طباعة تقرير الخمسة المتميزين.</div>` : ""}
            </article>`;
    }

    function renderExcellence(content) {
        const publicRows = filterPublicLeaderRows();
        const publicMarkup = publicExcellenceMarkup(publicRows);
        if (!state.privatePayload) {
            content.innerHTML = `${publicMarkup}
                <article class="card ce-lock-card">
                    <div class="ce-lock-icon" aria-hidden="true">🔒</div>
                    <div><h3>القائمة الكاملة والتفاصيل محمية</h3><p>الخمسة الأعلى ظاهرون أعلاه ومتاحون للطباعة للجميع. أدخل كلمة مرور مدير الموقع لعرض بقية الأعضاء، والجوانب التفصيلية، والسجلات غير المؤهلة.</p></div>
                    <div class="ce-unlock-form">
                        <label for="ceAdminPin">كلمة مرور المدير</label>
                        <input id="ceAdminPin" type="password" inputmode="numeric" autocomplete="current-password" aria-describedby="ceUnlockStatus">
                        <button class="btn btn-primary" type="button" data-ce-action="unlock">فتح التفاصيل الكاملة</button>
                    </div>
                    <p id="ceUnlockStatus" class="ce-lock-status" role="alert"></p>
                    <small>تنبيه: هذه بوابة خصوصية لموقع ثابت؛ صلاحية إدارية قوية تتطلب تسجيل دخول من الخادم.</small>
                </article>`;
            setNotice("المتاح للعامة هو الخمسة الأعلى فقط بحسب الفلاتر؛ بقية القائمة والتفاصيل مشفرة.", "secure");
            return;
        }

        const rows = rankFaculty(filterFacultyRows());
        const page = pageRows(rows);
        const eligibleCount = rows.filter((row) => row.eligible).length;
        const responseCount = rows.reduce((sum, row) => sum + Number(row.respondents || 0), 0);
        content.innerHTML = `${publicMarkup}
            <div class="ce-private-toolbar">
                <label class="ce-check"><input id="ceShowDescriptive" type="checkbox"${state.showDescriptive ? " checked" : ""}> إظهار السجلات الوصفية وغير المؤهلة</label>
                <div class="action-group">
                    <button class="btn" type="button" data-ce-action="export-excellence">Excel تفصيلي</button>
                    <button class="btn btn-primary" type="button" data-ce-action="pdf-excellence"${state.program === "all" ? " disabled" : ""}>PDF ملخص البرنامج</button>
                    <button class="btn btn-ghost" type="button" data-ce-action="lock">إغلاق التفاصيل</button>
                </div>
            </div>
            ${kpis([
                { label: "المؤهلون للمقارنة", value: eligibleCount.toLocaleString("ar-SA"), note: "داخل البرامج والفئات المتماثلة" },
                { label: "الاستجابات", value: responseCount.toLocaleString("ar-SA"), note: "للسجلات المعروضة" },
                { label: "الفترة", value: state.year === "all" ? "تجميع السنوات" : `${state.year}هـ`, note: "تجميع السنوات يفيد الدراسات العليا" },
                { label: "قسم العضو", value: state.department === "all" ? "كل الأقسام" : state.department, note: state.departmentAuto && state.program !== "all" ? "اختيار تلقائي من البرنامج" : "اختيار يدوي" },
                { label: "التفاصيل الإدارية", value: "مفتوحة مؤقتًا", note: "إغلاق تلقائي بعد 15 دقيقة" },
            ])}
            <article id="ceExcellenceReport" class="card ce-private-report">
                <div class="card-head"><div><p class="hero-kicker ce-report-kicker">كلية الشريعة والأنظمة</p><h3>تقرير التميز في التدريس</h3><p class="meta-text">الترتيب داخل البرنامج والدرجة، وفق المتوسط الموزون بعدد الاستجابات؛ كل طالب يمثل صوتًا واحدًا. قسم العضو: ${escapeHtml(state.department === "all" ? "كل الأقسام" : state.department)}.</p></div><span class="ce-confidential">سري — لمدير الموقع</span></div>
                ${excellenceSummary(rows)}
                <section class="ce-report-block ce-detailed-block">
                <h4>البيانات التفصيلية للأعضاء</h4>
                <div class="table-wrap"><table class="data-table ce-wide-table"><thead><tr>
                    <th>الترتيب</th><th>عضو هيئة التدريس</th><th>القسم</th><th>البرنامج</th><th>الدرجة</th><th>الدرجة الموزونة</th><th>التقدير</th><th>المستجيبون المحتسبون</th><th>الشعب المحتسبة</th><th>شعب البكالوريوس المستبعدة</th><th>الرسائل والمشروعات المحتسبة</th><th>المقررات</th><th>الحالة</th>${(meta.dimensions || []).map((item) => `<th>${escapeHtml(item.label)}</th>`).join("")}
                </tr></thead><tbody>
                    ${page.rows.map((row) => `<tr class="${row.eligible ? "" : "ce-row-muted"}"><td>${escapeHtml(row.rank)}</td><td><strong>${escapeHtml(row.instructorName)}</strong></td><td>${escapeHtml(row.facultyDepartment)}</td><td>${escapeHtml(row.program)}</td><td>${escapeHtml(row.degree)}</td><td>${row.respondents ? scorePill(row.scoreWeighted) : "—"}</td><td>${escapeHtml(row.grade)}</td><td>${Number(row.respondents).toLocaleString("ar-SA")}</td><td>${Number(row.sections).toLocaleString("ar-SA")}</td><td>${Number(row.excludedBachelorSections || 0).toLocaleString("ar-SA")}</td><td>${Number(row.researchSectionsIncluded || 0).toLocaleString("ar-SA")}</td><td>${Number(row.courses).toLocaleString("ar-SA")}</td><td><span class="status-pill">${escapeHtml(row.status)}</span></td>${(meta.dimensions || []).map((item) => `<td>${row.dimensionScores?.[item.id] == null ? "—" : score(row.dimensionScores[item.id])}</td>`).join("")}</tr>`).join("") || `<tr><td colspan="${13 + (meta.dimensions || []).length}"><div class="empty-state">لا توجد سجلات مؤهلة. فعّل عرض السجلات الوصفية أو غيّر الفلاتر.</div></td></tr>`}
                </tbody></table></div>
                ${pagination(rows.length, page.pages)}
                </section>
            </article>`;
        setNotice("الخمسة الأعلى متاحون للعامة؛ تم فتح القائمة الكاملة والتفاصيل الإدارية لهذه الجلسة فقط.", "secure");
    }

    function renderMethodology(content) {
        const quality = meta.dataQuality || {};
        content.innerHTML = `
            <div class="ce-method-grid">
                <article class="card ce-method-card"><span class="ce-method-number">${Number(quality.sections || 0).toLocaleString("ar-SA")}</span><h3>شعبة تدريسية</h3><p>من ثلاث سنوات دراسية، وتشمل الرسائل والمشروعات البحثية في حساب التميز.</p></article>
                <article class="card ce-method-card"><span class="ce-method-number">${Number(quality.excludedBachelorSmallGroups || 0).toLocaleString("ar-SA")}</span><h3>مجموعة بكالوريوس مستبعدة</h3><p>مجموع شعب المقرر للعضو في السنة والبرنامج بقي طالبًا أو طالبين.</p></article>
                <article class="card ce-method-card"><span class="ce-method-number">${Number(quality.recoveredBachelorSmallSections || 0).toLocaleString("ar-SA")}</span><h3>شعبة صغيرة استعيدت</h3><p>دخلت الحساب لوجود شعب أخرى للمقرر نفسه عند العضو في السنة نفسها.</p></article>
                <article class="card ce-method-card"><span class="ce-method-number">${Number(quality.excellenceBachelorSectionsBelowFive || 0).toLocaleString("ar-SA")}</span><h3>شعبة دون حد التميز</h3><p>شعبة بكالوريوس يقل عدد مستجيبيها عن 5؛ تبقى في التقرير السنوي ولا تدخل تقرير التميز.</p></article>
                <article class="card ce-method-card"><span class="ce-method-number">${Number(quality.researchSections || 0).toLocaleString("ar-SA")}</span><h3>رسالة أو مشروع بحثي</h3><p>تدخل ضمن الدرجة الموزونة للتميز وتخضع لحدود الأهلية بحسب الدرجة.</p></article>
            </div>
            <article class="card ce-methodology">
                <h3>قاعدة التقرير السنوي</h3>
                <p>في البكالوريوس تُجمع شعب المقرر للعضو نفسه داخل السنة والبرنامج أولًا؛ فإن بقي مجموع المجموعة طالبًا أو طالبين استُبعدت، وإن بلغ 3 فأكثر دخلت في التقرير السنوي. هذه القاعدة مستقلة عن أهلية تقرير التميز.</p>
                <h3>قواعد تقرير التميز</h3>
                <div class="ce-rule-grid">
                    <div><strong>شعبة البكالوريوس: 5 فأكثر</strong><p>لا تدخل الشعبة في حساب التميز إذا كان عدد مستجيبيها أقل من 5.</p></div>
                    <div><strong>عضو البكالوريوس: 20 فأكثر</strong><p>يشترط مجموع 20 استجابة محتسبة في الفترة، سواء جاءت من مقرر واحد أو مقررين أو أكثر.</p></div>
                    <div><strong>عضو الدراسات العليا: 4 فأكثر</strong><p>يشترط مجموع 4 استجابات محتسبة، ويطبق المتوسط الموزون نفسه بلا حد أدنى مستقل للشعبة.</p></div>
                    <div><strong>الرسائل والمشروعات</strong><p>تدخل ضمن المتوسط الموزون ودرجة التميز؛ وتُطبق عليها قاعدة البكالوريوس أو الدراسات العليا بحسب الدرجة.</p></div>
                </div>
                <h3>فلترة قسم عضو هيئة التدريس</h3>
                <p>عند اختيار برنامج يضبط الموقع القسم افتراضيًا على القسم الذي يتبعه البرنامج. يمكن بعد ذلك اختيار «كل الأقسام» أو قسم آخر. قسم العضو مأخوذ من سجل أعضاء هيئة التدريس بحسب سنة التقييم، وتظهر السجلات التي تعذر ربطها تحت «غير محدد» بدل نسبتها إلى قسم دون دليل.</p>
                <h3>حدود العرض العام والخاص</h3>
                <p>تظهر للعامة بيانات الخمسة الأعلى في البرنامج والدرجة ونطاق القسم المختار، مع إتاحة إخراجها بصيغتي PDF وCSV. أما بقية الأعضاء، والتوزيعات التفصيلية، والأعلى في كل جانب، والسجلات غير المؤهلة فتبقى محمية بكلمة مرور المدير.</p>
                <h3>طريقة حساب درجة العضو</h3>
                <p><strong>درجة العضو = مجموع (درجة كل شعبة × عدد مستجيبيها) ÷ مجموع المستجيبين المحتسبين.</strong> وبهذا يمثل كل طالب صوتًا واحدًا، ولا يُعطى المقرر الصغير الوزن نفسه للمقرر الكبير.</p>
                <p>مثال: 5 مستجيبين بدرجة 4.8، و35 مستجيبًا بدرجة 4.5؛ تكون النتيجة <span dir="ltr">(5×4.8 + 35×4.5) ÷ 40 = 4.5375</span>، وتعرض 4.54 من 5.</p>
                <h3>حدود الاستخدام</h3>
                <p>تقييم الطالب مؤشر مهم لتجربة التعلم، لكنه لا يُستخدم وحده لإثبات فاعلية التدريس. تقرير التميز النهائي ينبغي أن يضم قرائن أخرى مثل مراجعة الأقران، جودة ملف المقرر، التحصيل، والابتكار التدريسي.</p>
            </article>`;
        setNotice("التميز يحسب بالمتوسط الموزون بعدد الاستجابات: حد الشعبة في البكالوريوس 5، وحد العضو 20، وحد عضو الدراسات العليا 4.", "info");
    }

    function bytesFromBase64(value) {
        const binary = atob(value);
        return Uint8Array.from(binary, (character) => character.charCodeAt(0));
    }

    async function decryptPrivatePayload(pin) {
        const encrypted = packageData.encryptedFaculty;
        if (!encrypted || !window.crypto?.subtle) throw new Error("التشفير غير مدعوم في هذا المتصفح.");
        const encoder = new TextEncoder();
        const keyMaterial = await crypto.subtle.importKey("raw", encoder.encode(pin), "PBKDF2", false, ["deriveKey"]);
        const key = await crypto.subtle.deriveKey({
            name: "PBKDF2",
            salt: bytesFromBase64(encrypted.salt),
            iterations: encrypted.iterations,
            hash: "SHA-256",
        }, keyMaterial, { name: "AES-GCM", length: 256 }, false, ["decrypt"]);
        const plaintext = await crypto.subtle.decrypt({
            name: "AES-GCM",
            iv: bytesFromBase64(encrypted.iv),
            additionalData: encoder.encode(encrypted.aad),
        }, key, bytesFromBase64(encrypted.ciphertext));
        return JSON.parse(new TextDecoder().decode(plaintext));
    }

    async function unlockExcellence() {
        const input = $("#ceAdminPin");
        const status = $("#ceUnlockStatus");
        if (!input || !status) return;
        const pin = normalizeDigits(input.value);
        if (!pin) { status.textContent = "أدخل كلمة المرور."; return; }
        status.textContent = "جارٍ التحقق وفك التشفير...";
        input.disabled = true;
        try {
            state.privatePayload = await decryptPrivatePayload(pin);
            input.value = "";
            resetLockTimer();
            renderAll();
        } catch (error) {
            console.warn("تعذر فتح تقرير التميز:", error);
            status.textContent = "كلمة المرور غير صحيحة أو تعذر فك البيانات.";
            input.disabled = false;
            input.select();
        }
    }

    function resetLockTimer() {
        clearTimeout(state.lockTimer);
        state.lockTimer = setTimeout(() => lockExcellence("أُغلقت القائمة الكاملة والتفاصيل تلقائيًا بعد انتهاء الجلسة الآمنة."), LOCK_AFTER_MS);
    }

    function lockExcellence(message) {
        clearTimeout(state.lockTimer);
        state.lockTimer = null;
        state.privatePayload = null;
        state.showDescriptive = false;
        state.page = 1;
        renderAll();
        setNotice(message, "secure");
    }

    function exportWorkbook(rows, sheetName, filename) {
        if (!window.XLSX) { alert("مكتبة Excel غير متاحة حاليًا."); return; }
        const sheet = XLSX.utils.json_to_sheet(rows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, sheet, sheetName.slice(0, 31));
        sheet["!cols"] = Object.keys(rows[0] || {}).map((key) => ({ wch: Math.min(55, Math.max(14, key.length + 4)) }));
        XLSX.writeFile(workbook, filename);
    }

    function exportAnnual() {
        exportWorkbook(filterCourseRows().map((row) => ({
            "السنة": row.year,
            "الفصل": row.semester === "all" ? "كل الفصول" : row.semester,
            "رمز المقرر": row.courseCode || "",
            "اسم المقرر": row.courseName,
            "عدد الطلاب الذين قيموا المقرر": row.respondents,
            "نسبة المشاركين": "100٪",
            "نتيجة التقييم": row.score,
            "التوصيات التطويرية": row.recommendation,
        })), "التقرير السنوي", "تقييم_المقررات_للتقرير_السنوي.xlsx");
    }

    function exportIndirect() {
        if (!outcomeCategories.length) {
            alert("يلزم استكمال خريطة تصنيف البنود قبل تصدير التقييم غير المباشر.");
            return;
        }
        exportWorkbook(filterCourseRows().map((row) => {
            const output = {
                "السنة": row.year,
                "الفصل": row.semester === "all" ? "كل الفصول" : row.semester,
                "رمز المقرر": row.courseCode || "",
                "اسم المقرر": row.courseName,
                "عدد الطلاب": row.respondents,
                "جودة المقرر": row.score,
            };
            outcomeCategories.forEach((item) => { output[item.label] = row.outcomeScores?.[item.id] ?? ""; });
            return output;
        }), "التقييم غير المباشر", "التقييم_غير_المباشر_للمقررات.xlsx");
    }

    function excellenceExportRows() {
        return rankFaculty(filterFacultyRows()).map((row) => {
            const output = {
                "الترتيب داخل البرنامج والدرجة": row.rank,
                "عضو هيئة التدريس": row.instructorName,
                "القسم": row.facultyDepartment,
                "البرنامج": row.program,
                "الدرجة": row.degree,
                "الفترة": row.year === "all" ? "تجميع 1445-1447" : row.year,
                "الدرجة الموزونة": row.respondents ? row.scoreWeighted : "",
                "التقدير": row.grade,
                "المستجيبون المحتسبون": row.respondents,
                "الشعب المحتسبة": row.sections,
                "شعب البكالوريوس المستبعدة (أقل من 5)": row.excludedBachelorSections || 0,
                "شعب الرسائل والمشروعات المحتسبة": row.researchSectionsIncluded || 0,
                "المقررات": row.courses,
                "الحالة": row.status,
            };
            (meta.dimensions || []).forEach((item) => { output[item.label] = row.dimensionScores?.[item.id] ?? ""; });
            return output;
        });
    }

    function exportExcellence() {
        if (!state.privatePayload) return;
        exportWorkbook(excellenceExportRows(), "التميز في التدريس", "تقرير_التميز_في_التدريس_سري.xlsx");
        resetLockTimer();
    }

    function csvCell(value) {
        return `"${String(value ?? "").replace(/"/g, '""')}"`;
    }

    function exportPublicExcellenceCsv() {
        if (state.program === "all") {
            alert("اختر برنامجاً واحداً أولاً لتصدير الخمسة المتميزين.");
            return;
        }
        const rows = filterPublicLeaderRows();
        if (!rows.length) {
            alert("لا توجد قائمة خمسة مطابقة للفلاتر الحالية.");
            return;
        }
        const headers = [
            "الترتيب", "عضو هيئة التدريس", "القسم", "البرنامج", "الدرجة", "الفترة",
            "الدرجة الموزونة", "التقدير", "الاستجابات المحتسبة", "الشعب المحتسبة", "المقررات",
        ];
        const lines = [headers.map(csvCell).join(",")];
        rows.forEach((row) => {
            lines.push([
                row.rank,
                row.instructorName,
                row.facultyDepartment,
                row.program,
                row.degree,
                row.year === "all" ? "تجميع السنوات" : row.year,
                row.scoreWeighted,
                row.grade,
                row.respondents,
                row.sections,
                row.courses,
            ].map(csvCell).join(","));
        });
        const blob = new Blob(["\ufeff", lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "الخمسة_المتميزون_في_التدريس.csv";
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 0);
    }

    async function saveExcellencePdf({ reportSelector, action, filename, buttonLabel }) {
        const report = $(reportSelector);
        if (!report || !window.html2canvas || !window.jspdf?.jsPDF) {
            alert("تعذر تحميل أداة PDF. أعد تحميل الصفحة ثم حاول مرة أخرى.");
            return false;
        }
        const button = $(`[data-ce-action='${action}']`);
        if (button) { button.disabled = true; button.textContent = "جارٍ التوليد..."; }
        try {
            report.classList.add("ce-exporting");
            const canvas = await html2canvas(report, { scale: 1.1, backgroundColor: "#ffffff", useCORS: true });
            const pdf = new window.jspdf.jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
            const pageWidth = 287;
            const pageHeight = 190;
            const imageHeight = canvas.height * pageWidth / canvas.width;
            const image = canvas.toDataURL("image/jpeg", 0.9);
            let offset = 0;
            do {
                if (offset > 0) pdf.addPage();
                pdf.addImage(image, "JPEG", 5, 5 - offset, pageWidth, imageHeight);
                offset += pageHeight;
            } while (offset < imageHeight);
            pdf.save(filename);
            return true;
        } catch (error) {
            console.error("PDF export error:", error);
            alert("تعذر توليد PDF. قلّل النتائج بالفلاتر ثم حاول مرة أخرى.");
            return false;
        } finally {
            report.classList.remove("ce-exporting");
            const currentButton = $(`[data-ce-action='${action}']`);
            if (currentButton) { currentButton.disabled = false; currentButton.textContent = buttonLabel; }
        }
    }

    async function exportPublicExcellencePdf() {
        if (state.program === "all") {
            alert("اختر برنامجاً واحداً أولاً لطباعة الخمسة المتميزين.");
            return;
        }
        if (!filterPublicLeaderRows().length) {
            alert("لا توجد قائمة خمسة مطابقة للفلاتر الحالية.");
            return;
        }
        await saveExcellencePdf({
            reportSelector: "#cePublicTopFiveReport",
            action: "pdf-public-excellence",
            filename: "الخمسة_المتميزون_في_التدريس.pdf",
            buttonLabel: "طباعة / PDF",
        });
    }

    async function exportExcellencePdf() {
        if (!state.privatePayload) return;
        if (state.program === "all") {
            alert("اختر برنامجاً واحداً أولاً حتى يكون ترتيب تقرير التميز عادلاً ومحدداً.");
            return;
        }
        const saved = await saveExcellencePdf({
            reportSelector: "#ceExcellenceReport",
            action: "pdf-excellence",
            filename: "تقرير_التميز_في_التدريس_سري.pdf",
            buttonLabel: "PDF ملخص البرنامج",
        });
        if (saved) resetLockTimer();
    }

    function setNotice(message, tone = "info") {
        const notice = $("#ceNotice");
        if (!notice) return;
        notice.className = `ce-notice ce-notice--${tone}`;
        notice.textContent = message;
    }

    function render() {
        const root = $("#courseEvaluationsApp");
        if (!root) return;
        if (!state.initialized) createShell(root);
        renderAll();
    }

    window.CourseEvaluationsApp = { render };
})();
