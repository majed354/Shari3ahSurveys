(function () {
    "use strict";
    const $ = (id) => document.getElementById(id);
    const pageSize = 20;
    let records = [];
    let page = 1;
    const escape = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({"&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;"}[c]));
    const normalize = (s) => String(s).normalize("NFKC").replace(/[أإآ]/g, "ا").replace(/[\u064b-\u065fـ]/g, "").replace(/[()]/g, "").replace(/\s+/g, " ").trim();
    const fmt = (n) => Number(n).toLocaleString("ar-SA");
    function render() {
        const term = $("itemSemester").value;
        const department = $("itemDepartment").value;
        const query = normalize($("itemQuery").value);
        const selected = records.filter((r) => (term === "all" || r.semester === term) && (department === "all" || r.department === department) && normalize(r.courseName).includes(query));
        const pages = Math.max(1, Math.ceil(selected.length / pageSize));
        page = Math.min(page, pages);
        $("itemStatus").textContent = `${fmt(selected.length)} سجلًا بحسب المقرر والقسم والفصل؛ قد يظهر المقرر في الفصلين.`;
        $("itemPage").textContent = `صفحة ${fmt(page)} من ${fmt(pages)}`;
        $("itemPrevious").disabled = page <= 1;
        $("itemNext").disabled = page >= pages;
        $("itemRecords").innerHTML = selected.slice((page - 1) * pageSize, page * pageSize).map((r) => `
            <details><summary>${escape(r.courseName)} — ${escape(r.department)} — الفصل ${r.semester === "471" ? "الأول" : "الثاني"}</summary>
            <p class="item-context">${fmt(r.reportCount)} تقريرًا جُمعت لهذا المقرر. ${r.respondents == null ? "أعداد المستجيبين تختلف بين العبارات؛ راجع عدد الإجابات لكل عبارة." : `${fmt(r.respondents)} استجابة للمقرر عبر التقارير المجمعة؛ ليست حصرًا لطلاب فريدين على مستوى البرنامج.`}</p>
            <div class="table-wrap"><table><thead><tr><th>العبارة</th><th>المحور</th><th>عدد الإجابات</th><th>المتوسط من 5</th></tr></thead><tbody>
            ${r.items.map((i) => `<tr><td${i.statement == null ? ' class="missing-statement"' : ""}>${escape(i.statement || `العبارة ${i.number}: النص غير ظاهر في الأصل`)}</td><td>${escape(i.topic)}</td><td>${fmt(i.responses)}</td><td>${i.mean == null ? "—" : Number(i.mean).toFixed(2)}</td></tr>`).join("")}
            </tbody></table></div></details>`).join("") || "<p>لا توجد مقررات مطابقة للبحث.</p>";
    }
    ["itemSemester", "itemDepartment"].forEach((id) => $(id).addEventListener("change", () => { page = 1; render(); }));
    $("itemQuery").addEventListener("input", () => { page = 1; render(); });
    $("itemPrevious").addEventListener("click", () => { page--; render(); });
    $("itemNext").addEventListener("click", () => { page++; render(); });
    fetch("data/course-evaluations/1447/course-item-scores.json").then((r) => {
        if (!r.ok) throw new Error("http");
        return r.json();
    }).then((data) => {
        if (data.schemaVersion !== "course-item-evaluations-v1" || !Array.isArray(data.records)) throw new Error("schema");
        records = data.records;
        $("itemDepartment").innerHTML = '<option value="all">جميع الأقسام</option>' + [...new Set(records.map((r) => r.department))].sort().map((v) => `<option value="${escape(v)}">${escape(v)}</option>`).join("");
        render();
    }).catch(() => {
        $("itemStatus").textContent = "تعذر تحميل بيانات التقييمات. أعد تحميل الصفحة من الموقع أو شغّلها عبر خادم محلي.";
    });
})();
