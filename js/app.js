const PROGRAMS = [
    { id: "all", name: "جميع البرامج", degree: "all", dept: "كل الأقسام", years: [] },
    { id: "p01", name: "الأنظمة", degree: "بكالوريوس", dept: "الأنظمة", years: ["1446", "1445"] },
    { id: "p02", name: "الدراسات الإسلامية", degree: "بكالوريوس", dept: "الدراسات الإسلامية", years: ["1446", "1445"] },
    { id: "p03", name: "الشريعة", degree: "بكالوريوس", dept: "الشريعة", years: ["1446", "1445"] },
    { id: "p04", name: "القرآن وعلومه", degree: "بكالوريوس", dept: "القراءات", years: ["1446", "1445"] },
    { id: "p05", name: "القراءات", degree: "بكالوريوس", dept: "القراءات", years: ["1446", "1445"] },
    { id: "p06", name: "القانون", degree: "الماجستير", dept: "الأنظمة", years: ["1446", "1445"] },
    { id: "p07", name: "العقيدة", degree: "الماجستير", dept: "الدراسات الإسلامية", years: ["1446", "1445"] },
    { id: "p08", name: "أصول الفقه", degree: "الماجستير", dept: "الشريعة", years: ["1446", "1445"] },
    { id: "p09", name: "الفقه", degree: "الماجستير", dept: "الشريعة", years: ["1446", "1445"] },
    { id: "p10", name: "الدراسات القرآنية المعاصرة", degree: "الماجستير", dept: "القراءات", years: ["1446", "1445"] },
    { id: "p11", name: "القراءات", degree: "الماجستير", dept: "القراءات", years: ["1446", "1445"] },
    { id: "p12", name: "أصول الفقه", degree: "دكتوراه", dept: "الشريعة", years: ["1446", "1445"] },
    { id: "p13", name: "الفقه", degree: "دكتوراه", dept: "الشريعة", years: ["1446", "1445"] },
    { id: "p14", name: "الدراسات القرآنية", degree: "دكتوراه", dept: "القراءات", years: ["1446", "1445"] },
    { id: "p15", name: "القراءات", degree: "دكتوراه", dept: "القراءات", years: ["1446", "1445"] },
];

const PROGRAM_ORDER = new Map(PROGRAMS.map((program, index) => [program.id, index]));

const STAKEHOLDER_LABELS = {
    students: "الطلاب",
    faculty: "أعضاء هيئة التدريس",
    alumni: "الخريجون",
    employers: "أرباب العمل",
    advisory: "اللجنة الاستشارية",
    partners: "الشركاء",
};

const SECTION_META = [
    { id: "management", label: "إدارة البرنامج وضمان الجودة", shortLabel: "الإدارة", order: 1 },
    { id: "learning", label: "التعليم والتعلم", shortLabel: "التعلم", order: 2 },
    { id: "students", label: "الطلاب والخدمات", shortLabel: "الطلاب", order: 3 },
    { id: "faculty", label: "هيئة التدريس", shortLabel: "الهيئة", order: 4 },
    { id: "market", label: "الخريجون وسوق العمل", shortLabel: "السوق", order: 5 },
];

const CHART_COLORS = [
    "#1e5f58",
    "#c79e54",
    "#2e76b7",
    "#1b8a61",
    "#be4b3b",
    "#6b5bd6",
    "#d97706",
    "#0f766e",
];

const SURVEYS_PACKAGE = window.SURVEYS_DATA || {};
const SELF_STUDY_PACKAGE = window.SELF_STUDY_DATA || {};
const SOURCE_LABEL = SURVEYS_PACKAGE.sourceLabel || "منصة ذكاء الأعمال";
const EXTRACTED_DATA = SURVEYS_PACKAGE.extractedData || {};
const SELF_STUDY_LINKS = SELF_STUDY_PACKAGE.itemLinks || {};
const AVAILABLE_PROGRAM_YEARS = SURVEYS_PACKAGE.availableProgramYears || {};
const ITEM_RECORDS = buildItemRecords();
const CLOSURE_REPORT_SOURCE_LABEL = "المنظومة الجامعية";
const AVAILABLE_GENDERS = (SURVEYS_PACKAGE.availableGenders || []).length
    ? SURVEYS_PACKAGE.availableGenders
    : collectUnique(ITEM_RECORDS.map((record) => record.gender).filter(Boolean), (value) => normalizeText(value));

const ALL_AVAILABLE_YEARS = sortYears(
    collectUnique(
        Object.values(AVAILABLE_PROGRAM_YEARS).flatMap((years) => years),
        (year) => year
    )
);

const DEFAULT_YEAR = ALL_AVAILABLE_YEARS[0] || "all";
const DEFAULT_STAKEHOLDER = "students";
const DEFAULT_TOPIC_MODE = "general";
const SUMMARY_SURVEY_COUNT = SURVEYS_PACKAGE.surveyCount || countDistinctSurveys();
const TOTAL_PROGRAMS = Object.keys(AVAILABLE_PROGRAM_YEARS).length || PROGRAMS.filter((program) => program.id !== "all").length;

const refs = {};
let filterableDropdownEventsBound = false;

const state = {
    view: "overview",
    insightsSubTab: "analysis",
    charts: {},
    bottomSheetContext: "",
    latestClosurePayload: null,
    exploreFilters: {
        program: "all",
        year: DEFAULT_YEAR,
        stakeholder: DEFAULT_STAKEHOLDER,
        subject: "all",
        gender: "all",
        topicMode: DEFAULT_TOPIC_MODE,
        selfStudyTarget: "all",
        selfStudySearch: "",
    },
    compareFilters: {
        stakeholder: DEFAULT_STAKEHOLDER,
        subject: "all",
        topicMode: DEFAULT_TOPIC_MODE,
        selfStudyTarget: "all",
        selfStudySearch: "",
    },
    compareSlots: {
        a: { program: "", year: "" },
        b: { program: "", year: "" },
        c: { program: "", year: "" },
    },
    analysisFilters: {
        programs: new Set(getProgramsWithData().map((program) => program.id)),
        year: DEFAULT_YEAR,
        stakeholder: DEFAULT_STAKEHOLDER,
        subject: "all",
        topicMode: DEFAULT_TOPIC_MODE,
        selfStudyTarget: "all",
        selfStudySearch: "",
    },
    closureFilters: createClosureFilterState(),
    customFilters: createSingleFilterState(),
    customSelected: new Set(),
    searchMode: "precise",
    trendFilters: { program: "all", stakeholder: DEFAULT_STAKEHOLDER, topicMode: DEFAULT_TOPIC_MODE, selfStudyTarget: "all" },
    gapsFilters: { program: "all", year: DEFAULT_YEAR, target: 3.5, topicMode: DEFAULT_TOPIC_MODE, selfStudyTarget: "all" },
};

document.addEventListener("DOMContentLoaded", () => {
    try {
        cacheRefs();
        bindEvents();
        renderSummaryCards();
        refreshAllControls();
        renderViewState();
        renderCurrentView();
        requestAnimationFrame(() => {
            if (refs.loadingOverlay) refs.loadingOverlay.classList.add("is-hidden");
        });
    } catch (err) {
        console.error("Initialization error:", err);
        const overlay = document.getElementById("loadingOverlay");
        if (overlay) {
            const textEl = overlay.querySelector(".loading-text");
            const spinnerEl = overlay.querySelector(".loading-spinner");
            if (textEl) textEl.textContent = "حدث خطأ في تحميل البيانات. يرجى تحديث الصفحة.";
            if (spinnerEl) spinnerEl.style.display = "none";
        }
    }
});

function createSingleFilterState() {
    return {
        program: "all",
        year: DEFAULT_YEAR,
        stakeholder: DEFAULT_STAKEHOLDER,
        subject: "all",
        gender: "all",
        topicMode: DEFAULT_TOPIC_MODE,
        selfStudyTarget: "all",
        selfStudySearch: "",
    };
}

function createClosureFilterState(programId = "") {
    const pair = getDefaultClosureYearPair(programId);
    return {
        program: programId,
        fromYear: pair.fromYear,
        toYear: pair.toYear,
        subject: "all",
        gender: "all",
        topicMode: DEFAULT_TOPIC_MODE,
        selfStudyTarget: "all",
        selfStudySearch: "",
        level: "item",
        minImprovement: 2,
        displayMode: "details",
        reportShowResponses: true,
        reportShowStatement: true,
        reportShowSource: true,
    };
}

function getDefaultClosureYearPair(programId = "") {
    const availableYears = programId ? getAvailableYears(programId) : ALL_AVAILABLE_YEARS;
    const years = [...availableYears].sort((first, second) => Number(first) - Number(second));
    return {
        fromYear: years[0] || "",
        toYear: years[years.length - 1] || years[0] || "",
    };
}

function cacheRefs() {
    refs.summaryGrid = document.getElementById("summaryGrid");
    refs.footerSource = document.getElementById("footerSource");
    refs.navTabs = document.getElementById("navTabs");
    refs.loadingOverlay = document.getElementById("loadingOverlay");
    refs.bottomSheet = document.getElementById("bottomSheet");
    refs.bottomSheetCloseBtn = document.getElementById("bottomSheetCloseBtn");
    refs.bottomSheetTitle = document.getElementById("bottomSheetTitle");
    refs.bottomSheetSearch = document.getElementById("bottomSheetSearch");
    refs.bottomSheetSearchWrap = refs.bottomSheetSearch ? refs.bottomSheetSearch.closest(".bottom-sheet-search") : null;
    refs.bottomSheetOptions = document.getElementById("bottomSheetOptions");

    refs.overviewSection = document.getElementById("overviewSection");
    refs.exploreSection = document.getElementById("exploreSection");
    refs.insightsSection = document.getElementById("insightsSection");
    refs.searchSection = document.getElementById("searchSection");
    refs.searchMeta = document.getElementById("searchMeta");
    refs.searchInput = document.getElementById("searchInput");
    refs.searchModeChips = document.getElementById("searchModeChips");
    refs.searchResultCount = document.getElementById("searchResultCount");
    refs.searchResults = document.getElementById("searchResults");
    refs.searchEmpty = document.getElementById("searchEmpty");
    refs.searchNoResults = document.getElementById("searchNoResults");
    refs.searchNoResultsHint = document.getElementById("searchNoResultsHint");
    refs.searchResultsList = document.getElementById("searchResultsList");
    refs.exportSearchCsv = document.getElementById("exportSearchCsv");
    refs.exportSearchPdf = document.getElementById("exportSearchPdf");
    refs.analysisSubView = document.getElementById("analysisSubView");
    refs.compareSubView = document.getElementById("compareSubView");

    refs.overviewMeta = document.getElementById("overviewMeta");
    refs.overviewIndicators = document.getElementById("overviewIndicators");
    refs.overviewTableMeta = document.getElementById("overviewTableMeta");
    refs.overviewTableBody = document.getElementById("overviewTableBody");
    refs.overviewEmpty = document.getElementById("overviewEmpty");
    refs.overviewProgramChart = document.getElementById("overviewProgramChart");
    refs.overviewSectionChart = document.getElementById("overviewSectionChart");
    refs.overviewTrendChart = document.getElementById("overviewTrendChart");
    refs.overviewBandChart = document.getElementById("overviewBandChart");

    // Explore section
    refs.exploreMeta = document.getElementById("exploreMeta");
    refs.exploreResetBtn = document.getElementById("exploreResetBtn");
    // exploreSearchInput removed — search is now in dedicated tab
    refs.exploreProgram = document.getElementById("exploreProgram");
    refs.exploreYearChips = document.getElementById("exploreYearChips");
    refs.exploreStakeholderChips = document.getElementById("exploreStakeholderChips");
    refs.exploreGenderChips = document.getElementById("exploreGenderChips");
    refs.exploreTopicModeChips = document.getElementById("exploreTopicModeChips");
    refs.exploreGenderHint = document.getElementById("exploreGenderHint");
    refs.exploreGeneralTopicPanel = document.getElementById("exploreGeneralTopicPanel");
    refs.exploreSubject = document.getElementById("exploreSubject");
    refs.exploreSelfStudyPanel = document.getElementById("exploreSelfStudyPanel");
    refs.exploreSelfStudyFilter = document.getElementById("exploreSelfStudyFilter");
    refs.exploreActiveFilters = document.getElementById("exploreActiveFilters");
    refs.exploreIndicators = document.getElementById("exploreIndicators");
    refs.exploreSectionChart = document.getElementById("exploreSectionChart");
    refs.exploreItemChart = document.getElementById("exploreItemChart");
    refs.exploreTableMeta = document.getElementById("exploreTableMeta");
    refs.exploreTableBody = document.getElementById("exploreTableBody");
    refs.exploreEmpty = document.getElementById("exploreEmpty");
    refs.exportExploreExcel = document.getElementById("exportExploreExcel");
    refs.exportExploreCsv = document.getElementById("exportExploreCsv");
    refs.exportExplorePdf = document.getElementById("exportExplorePdf");

    // Insights section
    refs.insightsMeta = document.getElementById("insightsMeta");
    refs.insightsResetBtn = document.getElementById("insightsResetBtn");
    refs.analysisYearChips = document.getElementById("analysisYearChips");
    refs.analysisStakeholderChips = document.getElementById("analysisStakeholderChips");
    refs.analysisTopicModeChips = document.getElementById("analysisTopicModeChips");
    refs.compareStakeholderChips = document.getElementById("compareStakeholderChips");
    refs.compareTopicModeChips = document.getElementById("compareTopicModeChips");
    refs.exportInsightsExcel = document.getElementById("exportInsightsExcel");
    refs.exportInsightsCsv = document.getElementById("exportInsightsCsv");
    refs.exportInsightsPdf = document.getElementById("exportInsightsPdf");
    refs.exportClosureExcel = document.getElementById("exportClosureExcel");
    refs.exportClosureCsv = document.getElementById("exportClosureCsv");
    refs.exportClosurePdf = document.getElementById("exportClosurePdf");

    // Custom drawer
    refs.customSection = document.getElementById("customSection");
    refs.customSearchInput = document.getElementById("customSearchInput");
    refs.exportCustomExcel = document.getElementById("exportCustomExcel");
    refs.exportCustomCsv = document.getElementById("exportCustomCsv");
    refs.exportCustomPdf = document.getElementById("exportCustomPdf");
    refs.customYearChips = document.getElementById("customYearChips");
    refs.customStakeholderChips = document.getElementById("customStakeholderChips");
    refs.customGenderChips = document.getElementById("customGenderChips");

    // Compare section - kept elements
    refs.compareSubjectFilter = document.getElementById("compareSubjectFilter");
    refs.compareGeneralTopicPanel = document.getElementById("compareGeneralTopicPanel");
    refs.compareSelfStudyPanel = document.getElementById("compareSelfStudyPanel");
    refs.compareSelfStudyFilter = document.getElementById("compareSelfStudyFilter");
    refs.compareActiveFilters = document.getElementById("compareActiveFilters");
    refs.compareProgramA = document.getElementById("compareProgramA");
    refs.compareYearA = document.getElementById("compareYearA");
    refs.compareProgramB = document.getElementById("compareProgramB");
    refs.compareYearB = document.getElementById("compareYearB");
    refs.compareProgramC = document.getElementById("compareProgramC");
    refs.compareYearC = document.getElementById("compareYearC");
    refs.compareIndicators = document.getElementById("compareIndicators");
    refs.compareSectionChart = document.getElementById("compareSectionChart");
    refs.compareAverageChart = document.getElementById("compareAverageChart");
    refs.compareTableHead = document.getElementById("compareTableHead");
    refs.compareTableBody = document.getElementById("compareTableBody");
    refs.compareTableMeta = document.getElementById("compareTableMeta");
    refs.compareEmpty = document.getElementById("compareEmpty");

    // Analysis section - kept elements
    refs.analysisSubjectFilter = document.getElementById("analysisSubjectFilter");
    refs.analysisGeneralTopicPanel = document.getElementById("analysisGeneralTopicPanel");
    refs.analysisSelfStudyPanel = document.getElementById("analysisSelfStudyPanel");
    refs.analysisSelfStudyFilter = document.getElementById("analysisSelfStudyFilter");
    refs.analysisProgramsList = document.getElementById("analysisProgramsList");
    refs.analysisProgramsDropdown = document.getElementById("analysisProgramsDropdown");
    refs.analysisProgramsTrigger = document.getElementById("analysisProgramsTrigger");
    refs.analysisSelectAllProgramsBtn = document.getElementById("analysisSelectAllProgramsBtn");
    refs.analysisClearProgramsBtn = document.getElementById("analysisClearProgramsBtn");
    refs.analysisActiveFilters = document.getElementById("analysisActiveFilters");
    refs.analysisIndicators = document.getElementById("analysisIndicators");
    refs.analysisStrengths = document.getElementById("analysisStrengths");
    refs.analysisWatchlist = document.getElementById("analysisWatchlist");
    refs.analysisSectionChart = document.getElementById("analysisSectionChart");
    refs.analysisTableTitle = document.getElementById("analysisTableTitle");
    refs.analysisTableMeta = document.getElementById("analysisTableMeta");
    refs.analysisTableHead = document.getElementById("analysisTableHead");
    refs.analysisTableBody = document.getElementById("analysisTableBody");
    refs.analysisEmpty = document.getElementById("analysisEmpty");

    // Quality closure sub-view
    refs.closureSubView = document.getElementById("closureSubView");
    refs.closureProgramSelect = document.getElementById("closureProgramSelect");
    refs.closureFromYearSelect = document.getElementById("closureFromYearSelect");
    refs.closureToYearSelect = document.getElementById("closureToYearSelect");
    refs.closureLevelSelect = document.getElementById("closureLevelSelect");
    refs.closureMinImprovementInput = document.getElementById("closureMinImprovementInput");
    refs.closureTopicModeChips = document.getElementById("closureTopicModeChips");
    refs.closureGenderChips = document.getElementById("closureGenderChips");
    refs.closureGenderHint = document.getElementById("closureGenderHint");
    refs.closureGeneralTopicPanel = document.getElementById("closureGeneralTopicPanel");
    refs.closureSubjectFilter = document.getElementById("closureSubjectFilter");
    refs.closureSelfStudyPanel = document.getElementById("closureSelfStudyPanel");
    refs.closureSelfStudyFilter = document.getElementById("closureSelfStudyFilter");
    refs.closureActiveFilters = document.getElementById("closureActiveFilters");
    refs.closureIndicators = document.getElementById("closureIndicators");
    refs.closureHighlights = document.getElementById("closureHighlights");
    refs.closureNarrative = document.getElementById("closureNarrative");
    refs.closureDeltaChart = document.getElementById("closureDeltaChart");
    refs.closureDisplayTabs = document.getElementById("closureDisplayTabs");
    refs.closureTableCard = document.getElementById("closureTableCard");
    refs.closureTableMeta = document.getElementById("closureTableMeta");
    refs.closureTableHead = document.getElementById("closureTableHead");
    refs.closureTableBody = document.getElementById("closureTableBody");
    refs.closureEmpty = document.getElementById("closureEmpty");
    refs.closureReportCard = document.getElementById("closureReportCard");
    refs.closureReportMeta = document.getElementById("closureReportMeta");
    refs.closureReportOptionToggles = document.getElementById("closureReportOptionToggles");
    refs.closureReportEmpty = document.getElementById("closureReportEmpty");
    refs.closureReportPrintArea = document.getElementById("closureReportPrintArea");
    refs.closureReportSummary = document.getElementById("closureReportSummary");
    refs.closureReportHead = document.getElementById("closureReportHead");
    refs.closureReportBody = document.getElementById("closureReportBody");
    refs.exportClosureReportCsv = document.getElementById("exportClosureReportCsv");
    refs.exportClosureReportPdf = document.getElementById("exportClosureReportPdf");

    // Trend sub-view
    refs.trendSubView = document.getElementById("trendSubView");
    refs.trendProgramSelect = document.getElementById("trendProgramSelect");
    refs.trendStakeholderSelect = document.getElementById("trendStakeholderSelect");
    refs.trendIndicators = document.getElementById("trendIndicators");
    refs.trendLineChart = document.getElementById("trendLineChart");
    refs.trendChangeSummary = document.getElementById("trendChangeSummary");
    refs.trendEmpty = document.getElementById("trendEmpty");
    refs.trendTopicModeChips = document.getElementById("trendTopicModeChips");
    refs.trendSelfStudyPanel = document.getElementById("trendSelfStudyPanel");
    refs.trendSelfStudyFilter = document.getElementById("trendSelfStudyFilter");

    // Gaps sub-view
    refs.gapsSubView = document.getElementById("gapsSubView");
    refs.gapsProgramSelect = document.getElementById("gapsProgramSelect");
    refs.gapsYearSelect = document.getElementById("gapsYearSelect");
    refs.gapsTargetInput = document.getElementById("gapsTargetInput");
    refs.gapsIndicators = document.getElementById("gapsIndicators");
    refs.gapsBarChart = document.getElementById("gapsBarChart");
    refs.gapsTableBody = document.getElementById("gapsTableBody");
    refs.gapsCopyBtn = document.getElementById("gapsCopyBtn");
    refs.gapsEmpty = document.getElementById("gapsEmpty");
    refs.gapsNarrativeCard = document.getElementById("gapsNarrativeCard");
    refs.gapsNarrativeText = document.getElementById("gapsNarrativeText");
    refs.gapsTopicModeChips = document.getElementById("gapsTopicModeChips");
    refs.gapsSelfStudyPanel = document.getElementById("gapsSelfStudyPanel");
    refs.gapsSelfStudyFilter = document.getElementById("gapsSelfStudyFilter");

    // Custom section - kept elements
    refs.customMeta = document.getElementById("customMeta");
    refs.customProgramFilter = document.getElementById("customProgramFilter");
    refs.customSubjectFilter = document.getElementById("customSubjectFilter");
    refs.customGenderHint = document.getElementById("customGenderHint");
    refs.customActiveFilters = document.getElementById("customActiveFilters");
    refs.customIndicators = document.getElementById("customIndicators");
    refs.customAvailableMeta = document.getElementById("customAvailableMeta");
    refs.customOptionsList = document.getElementById("customOptionsList");
    refs.customSelectAllBtn = document.getElementById("customSelectAllBtn");
    refs.customClearSelectionBtn = document.getElementById("customClearSelectionBtn");
    refs.customSelectedMeta = document.getElementById("customSelectedMeta");
    refs.customSelectedTableBody = document.getElementById("customSelectedTableBody");
    refs.customEmpty = document.getElementById("customEmpty");
}

function bindEvents() {
    bindFilterableDropdownGlobalEvents();

    if (refs.navTabs) refs.navTabs.addEventListener("click", (event) => {
        const button = event.target.closest("[data-view]");
        if (!button) return;
        state.view = button.dataset.view;
        renderViewState();
        renderCurrentView();
    });

    // Insights sub-tabs
    if (refs.insightsSection) {
        refs.insightsSection.addEventListener("click", (event) => {
            const subTab = event.target.closest("[data-subtab]");
            if (!subTab) return;
            state.insightsSubTab = subTab.dataset.subtab;
            refs.insightsSection.querySelectorAll("[data-subtab]").forEach((btn) => {
                btn.classList.toggle("is-active", btn.dataset.subtab === state.insightsSubTab);
            });
            ["analysisSubView", "compareSubView", "closureSubView", "trendSubView", "gapsSubView"].forEach((key) => {
                if (refs[key]) refs[key].classList.toggle("hidden", state.insightsSubTab !== key.replace("SubView", ""));
            });
            renderCurrentView();
        });
    }

    /* Custom tab: search input for filtering available items */
    if (refs.customSearchInput) {
        refs.customSearchInput.addEventListener("input", () => {
            renderCustomSection();
        });
    }

    /* Custom tab: export buttons */
    if (refs.exportCustomExcel) refs.exportCustomExcel.addEventListener("click", () => exportCustomData("xlsx"));
    if (refs.exportCustomCsv) refs.exportCustomCsv.addEventListener("click", () => exportCustomData("csv"));
    if (refs.exportCustomPdf) refs.exportCustomPdf.addEventListener("click", () => exportCustomResultsPdf());

    // Explore section filters
    if (refs.exploreProgram) {
        refs.exploreProgram.addEventListener("change", (event) => {
            state.exploreFilters.program = event.target.value;
            renderExploreControls();
            renderExploreSection();
        });
    }

    // Dedicated search tab
    if (refs.searchInput) {
        let searchDebounce = null;
        refs.searchInput.addEventListener("input", () => {
            clearTimeout(searchDebounce);
            searchDebounce = setTimeout(() => renderSearchResults(), 200);
        });
    }
    if (refs.exportSearchCsv) refs.exportSearchCsv.addEventListener("click", () => exportSearchResults("csv"));
    if (refs.exportSearchPdf) refs.exportSearchPdf.addEventListener("click", () => exportSearchResults("pdf"));

    if (refs.exploreResetBtn) {
        refs.exploreResetBtn.addEventListener("click", () => {
            state.exploreFilters = {
                program: "all",
                year: DEFAULT_YEAR,
                stakeholder: DEFAULT_STAKEHOLDER,
                subject: "all",
                gender: "all",
                topicMode: DEFAULT_TOPIC_MODE,
                selfStudyTarget: "all",
                selfStudySearch: "",
                query: "",
            };
            renderExploreControls();
            renderExploreSection();
        });
    }

    if (refs.exploreSubject) {
        refs.exploreSubject.addEventListener("change", (event) => {
            state.exploreFilters.subject = event.target.value;
            renderFilterChips(refs.exploreActiveFilters, buildExploreFilterChips());
            renderExploreSection();
        });
    }

    if (refs.exploreSelfStudyFilter) {
        refs.exploreSelfStudyFilter.addEventListener("change", (event) => {
            state.exploreFilters.selfStudyTarget = event.target.value;
            renderFilterChips(refs.exploreActiveFilters, buildExploreFilterChips());
            renderExploreSection();
        });
    }

    // Custom controls - program and subject are chip-based in drawer
    if (refs.customProgramFilter) {
        refs.customProgramFilter.addEventListener("change", (event) => {
            state.customFilters.program = event.target.value;
            renderCustomControls();
            renderCustomSection();
        });
    }

    if (refs.customSubjectFilter) {
        refs.customSubjectFilter.addEventListener("change", (event) => {
            state.customFilters.subject = event.target.value;
            renderCustomSection();
        });
    }

    // Compare section filters
    if (refs.compareSubjectFilter) {
        refs.compareSubjectFilter.addEventListener("change", (event) => {
            state.compareFilters.subject = event.target.value;
            renderCompareControls();
            renderCompareSection();
        });
    }

    if (refs.compareSelfStudyFilter) {
        refs.compareSelfStudyFilter.addEventListener("change", (event) => {
            state.compareFilters.selfStudyTarget = event.target.value;
            renderCompareControls();
            renderCompareSection();
        });
    }

    if (refs.compareActiveFilters) {
        refs.compareActiveFilters.addEventListener("change", (event) => {
            const checkbox = event.target.closest("input[data-filter]");
            if (!checkbox) return;
            const filterKey = checkbox.dataset.filter;
            const filterValue = checkbox.dataset.value;
            if (checkbox.checked) {
                // Filter activated
                if (filterKey === "stakeholder") {
                    state.compareFilters.stakeholder = filterValue;
                } else if (filterKey === "topicMode") {
                    state.compareFilters.topicMode = filterValue;
                    if (filterValue === "general") {
                        state.compareFilters.selfStudyTarget = "all";
                        state.compareFilters.selfStudySearch = "";
                    } else {
                        state.compareFilters.subject = "all";
                    }
                }
            }
            renderCompareControls();
            renderCompareSection();
        });
    }

    // Insights reset button (for both compare and analysis)
    if (refs.insightsResetBtn) {
        refs.insightsResetBtn.addEventListener("click", () => {
            if (state.insightsSubTab === "compare") {
                state.compareFilters = {
                    stakeholder: DEFAULT_STAKEHOLDER,
                    subject: "all",
                    topicMode: DEFAULT_TOPIC_MODE,
                    selfStudyTarget: "all",
                    selfStudySearch: "",
                };
                state.compareSlots = {
                    a: { program: "", year: "" },
                    b: { program: "", year: "" },
                    c: { program: "", year: "" },
                };
                renderCompareControls();
                renderCompareSection();
            } else if (state.insightsSubTab === "analysis") {
                state.analysisFilters = {
                    programs: new Set(getProgramsWithData().map((program) => program.id)),
                    year: DEFAULT_YEAR,
                    stakeholder: DEFAULT_STAKEHOLDER,
                    subject: "all",
                    topicMode: DEFAULT_TOPIC_MODE,
                    selfStudyTarget: "all",
                    selfStudySearch: "",
                };
                renderAnalysisControls();
                renderAnalysisSection();
            } else if (state.insightsSubTab === "closure") {
                state.closureFilters = createClosureFilterState();
                renderClosureControls();
                renderClosureSection();
            }
        });
    }

    // Compare slot controls
    ["A", "B", "C"].forEach((slotKey) => {
        const slot = slotKey.toLowerCase();
        if (refs[`compareProgram${slotKey}`]) {
            refs[`compareProgram${slotKey}`].addEventListener("change", (event) => {
                state.compareSlots[slot].program = event.target.value;
                if (!state.compareSlots[slot].program) {
                    state.compareSlots[slot].year = "";
                } else if (!getAvailableYears(state.compareSlots[slot].program).includes(state.compareSlots[slot].year)) {
                    state.compareSlots[slot].year = "";
                }
                renderCompareControls();
                renderCompareSection();
            });
        }

        if (refs[`compareYear${slotKey}`]) {
            refs[`compareYear${slotKey}`].addEventListener("change", (event) => {
                state.compareSlots[slot].year = event.target.value;
                renderCompareSection();
            });
        }
    });

    // Analysis section filters - year and stakeholder now use chip-based controls
    if (refs.analysisActiveFilters) {
        refs.analysisActiveFilters.addEventListener("change", (event) => {
            const checkbox = event.target.closest("input[data-filter]");
            if (!checkbox) return;
            const filterKey = checkbox.dataset.filter;
            const filterValue = checkbox.dataset.value;
            if (checkbox.checked) {
                if (filterKey === "year") {
                    state.analysisFilters.year = filterValue;
                } else if (filterKey === "stakeholder") {
                    state.analysisFilters.stakeholder = filterValue;
                } else if (filterKey === "topicMode") {
                    state.analysisFilters.topicMode = filterValue;
                    if (filterValue === "general") {
                        state.analysisFilters.selfStudyTarget = "all";
                        state.analysisFilters.selfStudySearch = "";
                    } else {
                        state.analysisFilters.subject = "all";
                    }
                }
            }
            renderAnalysisControls();
            renderAnalysisSection();
        });
    }

    if (refs.analysisSubjectFilter) {
        refs.analysisSubjectFilter.addEventListener("change", (event) => {
            state.analysisFilters.subject = event.target.value;
            renderAnalysisControls();
            renderAnalysisSection();
        });
    }

    if (refs.analysisSelfStudyFilter) {
        refs.analysisSelfStudyFilter.addEventListener("change", (event) => {
            state.analysisFilters.selfStudyTarget = event.target.value;
            renderAnalysisControls();
            renderAnalysisSection();
        });
    }

    /* Multi-select dropdown toggle */
    if (refs.analysisProgramsTrigger && refs.analysisProgramsDropdown) {
        refs.analysisProgramsTrigger.addEventListener("click", () => {
            refs.analysisProgramsDropdown.classList.toggle("is-open");
        });
        document.addEventListener("click", (event) => {
            if (refs.analysisProgramsDropdown && !refs.analysisProgramsDropdown.contains(event.target)) {
                refs.analysisProgramsDropdown.classList.remove("is-open");
            }
        });
    }

    if (refs.analysisSelectAllProgramsBtn) {
        refs.analysisSelectAllProgramsBtn.addEventListener("click", () => {
            state.analysisFilters.programs = new Set(getProgramsWithData().map((program) => program.id));
            renderAnalysisControls();
            renderAnalysisSection();
        });
    }

    if (refs.analysisClearProgramsBtn) {
        refs.analysisClearProgramsBtn.addEventListener("click", () => {
            state.analysisFilters.programs = new Set();
            renderAnalysisControls();
            renderAnalysisSection();
        });
    }

    if (refs.analysisProgramsList) {
        refs.analysisProgramsList.addEventListener("change", (event) => {
            const checkbox = event.target.closest("input[data-program]");
            if (!checkbox) return;
            if (checkbox.checked) {
                state.analysisFilters.programs.add(checkbox.dataset.program);
            } else {
                state.analysisFilters.programs.delete(checkbox.dataset.program);
            }
            renderAnalysisControls();
            renderAnalysisSection();
        });
    }

    if (refs.customOptionsList) {
        refs.customOptionsList.addEventListener("change", (event) => {
            const checkbox = event.target.closest("input[data-row-id]");
            if (!checkbox) return;
            const rowId = checkbox.dataset.rowId;
            if (checkbox.checked) {
                state.customSelected.add(rowId);
            } else {
                state.customSelected.delete(rowId);
            }
            renderCustomSection();
        });
    }

    if (refs.customSelectAllBtn) {
        refs.customSelectAllBtn.addEventListener("click", () => {
            getCustomRows().forEach((row) => state.customSelected.add(row.uid));
            renderCustomSection();
        });
    }

    if (refs.customClearSelectionBtn) {
        refs.customClearSelectionBtn.addEventListener("click", () => {
            state.customSelected.clear();
            renderCustomSection();
        });
    }

    if (refs.exportExploreExcel) refs.exportExploreExcel.addEventListener("click", () => exportExploreData("xlsx"));
    if (refs.exportExploreCsv) refs.exportExploreCsv.addEventListener("click", () => exportExploreData("csv"));
    if (refs.exportExplorePdf) refs.exportExplorePdf.addEventListener("click", () => exportExploreResultsPdf());
    if (refs.exportClosureExcel) refs.exportClosureExcel.addEventListener("click", () => exportClosureData("xlsx"));
    if (refs.exportClosureCsv) refs.exportClosureCsv.addEventListener("click", () => exportClosureData("csv"));
    if (refs.exportClosurePdf) refs.exportClosurePdf.addEventListener("click", () => exportClosureResultsPdf());
    if (refs.exportClosureReportCsv) refs.exportClosureReportCsv.addEventListener("click", () => exportClosureReportData("csv"));
    if (refs.exportClosureReportPdf) refs.exportClosureReportPdf.addEventListener("click", () => exportClosureReportPdf());

    bindBottomSheetEvents();
    bindTrendEvents();
    bindClosureEvents();
    bindGapsEvents();

    if (refs.exportInsightsExcel) refs.exportInsightsExcel.addEventListener("click", () => {
        if (state.insightsSubTab === "compare") exportComparisonData("xlsx");
        else if (state.insightsSubTab === "analysis") exportAnalysisData("xlsx");
        else if (state.insightsSubTab === "closure") {
            if (state.closureFilters.displayMode === "report") exportClosureReportData("xlsx");
            else exportClosureData("xlsx");
        }
        else if (state.insightsSubTab === "trend") exportTrendData("xlsx");
        else if (state.insightsSubTab === "gaps") exportGapsData("xlsx");
    });
    if (refs.exportInsightsCsv) refs.exportInsightsCsv.addEventListener("click", () => {
        if (state.insightsSubTab === "compare") exportComparisonData("csv");
        else if (state.insightsSubTab === "analysis") exportAnalysisData("csv");
        else if (state.insightsSubTab === "closure") {
            if (state.closureFilters.displayMode === "report") exportClosureReportData("csv");
            else exportClosureData("csv");
        }
        else if (state.insightsSubTab === "trend") exportTrendData("csv");
        else if (state.insightsSubTab === "gaps") exportGapsData("csv");
    });
    if (refs.exportInsightsPdf) refs.exportInsightsPdf.addEventListener("click", () => {
        if (state.insightsSubTab === "compare") exportComparisonResultsPdf();
        else if (state.insightsSubTab === "analysis") exportAnalysisResultsPdf();
        else if (state.insightsSubTab === "closure") {
            if (state.closureFilters.displayMode === "report") exportClosureReportPdf();
            else exportClosureResultsPdf();
        }
        else if (state.insightsSubTab === "trend") exportTrendResultsPdf();
        else if (state.insightsSubTab === "gaps") exportGapsResultsPdf();
    });
}

function bindTrendEvents() {
    if (refs.trendProgramSelect) refs.trendProgramSelect.addEventListener("change", (e) => {
        state.trendFilters.program = e.target.value;
        renderTrendSection();
    });
    if (refs.trendStakeholderSelect) refs.trendStakeholderSelect.addEventListener("change", (e) => {
        state.trendFilters.stakeholder = e.target.value;
        renderTrendSection();
    });
    if (refs.trendSelfStudyFilter) refs.trendSelfStudyFilter.addEventListener("change", (e) => {
        state.trendFilters.selfStudyTarget = e.target.value;
        renderTrendSection();
    });
}

function bindBottomSheetEvents() {
    if (!refs.bottomSheet) return;

    const backdrop = refs.bottomSheet.querySelector(".bottom-sheet-backdrop");
    if (backdrop) {
        backdrop.addEventListener("click", () => closeBottomSheet());
    }

    if (refs.bottomSheetCloseBtn) {
        refs.bottomSheetCloseBtn.addEventListener("click", () => closeBottomSheet());
    }

    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && !refs.bottomSheet.classList.contains("hidden")) {
            closeBottomSheet();
        }
    });
}

function bindClosureEvents() {
    if (refs.closureProgramSelect) refs.closureProgramSelect.addEventListener("change", (event) => {
        state.closureFilters.program = event.target.value;
        const pair = getDefaultClosureYearPair(state.closureFilters.program);
        state.closureFilters.fromYear = pair.fromYear;
        state.closureFilters.toYear = pair.toYear;
        renderClosureControls();
        renderClosureSection();
    });

    if (refs.closureFromYearSelect) refs.closureFromYearSelect.addEventListener("change", (event) => {
        state.closureFilters.fromYear = event.target.value;
        normalizeClosureYearPair("from");
        renderClosureControls();
        renderClosureSection();
    });

    if (refs.closureToYearSelect) refs.closureToYearSelect.addEventListener("change", (event) => {
        state.closureFilters.toYear = event.target.value;
        normalizeClosureYearPair("to");
        renderClosureControls();
        renderClosureSection();
    });

    if (refs.closureLevelSelect) refs.closureLevelSelect.addEventListener("change", (event) => {
        state.closureFilters.level = event.target.value;
        renderClosureSection();
    });

    if (refs.closureMinImprovementInput) refs.closureMinImprovementInput.addEventListener("input", (event) => {
        const rawValue = Number(event.target.value);
        const value = Number.isFinite(rawValue) ? Math.max(0, rawValue) : 2;
        state.closureFilters.minImprovement = value;
        refs.closureMinImprovementInput.value = String(state.closureFilters.minImprovement);
        renderClosureSection();
    });

    if (refs.closureSubjectFilter) refs.closureSubjectFilter.addEventListener("change", (event) => {
        state.closureFilters.subject = event.target.value;
        renderClosureControls();
        renderClosureSection();
    });

    if (refs.closureSelfStudyFilter) refs.closureSelfStudyFilter.addEventListener("change", (event) => {
        state.closureFilters.selfStudyTarget = event.target.value;
        renderClosureControls();
        renderClosureSection();
    });

    if (refs.closureDisplayTabs) refs.closureDisplayTabs.addEventListener("click", (event) => {
        const button = event.target.closest("[data-closure-display]");
        if (!button) return;
        state.closureFilters.displayMode = button.dataset.closureDisplay;
        renderClosureControls();
        renderClosureSection();
    });

    if (refs.closureReportOptionToggles) refs.closureReportOptionToggles.addEventListener("click", (event) => {
        const button = event.target.closest("[data-report-option]");
        if (!button) return;
        const optionKey = button.dataset.reportOption;
        if (optionKey === "responses") state.closureFilters.reportShowResponses = !state.closureFilters.reportShowResponses;
        if (optionKey === "statement") state.closureFilters.reportShowStatement = !state.closureFilters.reportShowStatement;
        if (optionKey === "source") state.closureFilters.reportShowSource = !state.closureFilters.reportShowSource;
        renderClosureControls();
        renderClosureSection();
    });

    if (refs.closureHighlights) refs.closureHighlights.addEventListener("click", (event) => {
        const detailButton = event.target.closest("[data-closure-detail-index]");
        if (!detailButton) return;
        const rowIndex = Number(detailButton.dataset.closureDetailIndex);
        const row = state.latestClosurePayload?.qualifyingRows?.[rowIndex];
        if (!row) return;
        openClosureDetailSheet(row, state.latestClosurePayload);
    });
}

function bindGapsEvents() {
    if (refs.gapsProgramSelect) refs.gapsProgramSelect.addEventListener("change", (e) => {
        state.gapsFilters.program = e.target.value;
        renderGapsControls();
        renderGapsSection();
    });
    if (refs.gapsYearSelect) refs.gapsYearSelect.addEventListener("change", (e) => {
        state.gapsFilters.year = e.target.value;
        renderGapsSection();
    });
    if (refs.gapsTargetInput) refs.gapsTargetInput.addEventListener("input", (e) => {
        const v = parseFloat(e.target.value);
        if (!isNaN(v) && v >= 1 && v <= 5) {
            state.gapsFilters.target = v;
            renderGapsSection();
        }
    });
    if (refs.gapsSelfStudyFilter) refs.gapsSelfStudyFilter.addEventListener("change", (e) => {
        state.gapsFilters.selfStudyTarget = e.target.value;
        renderGapsSection();
    });
    if (refs.gapsCopyBtn) refs.gapsCopyBtn.addEventListener("click", () => {
        const text = refs.gapsNarrativeText ? refs.gapsNarrativeText.textContent : "";
        if (text && navigator.clipboard) {
            navigator.clipboard.writeText(text).then(() => {
                refs.gapsCopyBtn.textContent = "تم النسخ ✓";
                setTimeout(() => { refs.gapsCopyBtn.textContent = "نسخ للدراسة الذاتية"; }, 2000);
            });
        }
    });
}

function bindSingleFilterEvents(programRef, yearRef, stakeholderRef, subjectRef, genderRef, stateRef, resetRef, controlsRenderer, sectionRenderer) {
    programRef.addEventListener("change", (event) => {
        stateRef.program = event.target.value;
        controlsRenderer();
        sectionRenderer();
    });

    yearRef.addEventListener("change", (event) => {
        stateRef.year = event.target.value;
        controlsRenderer();
        sectionRenderer();
    });

    stakeholderRef.addEventListener("change", (event) => {
        stateRef.stakeholder = event.target.value;
        controlsRenderer();
        sectionRenderer();
    });

    subjectRef.addEventListener("change", (event) => {
        stateRef.subject = event.target.value;
        sectionRenderer();
    });

    genderRef.addEventListener("change", (event) => {
        stateRef.gender = event.target.value;
        controlsRenderer();
        sectionRenderer();
    });

    resetRef.addEventListener("click", () => {
        Object.assign(stateRef, createSingleFilterState());
        if (stateRef === state.customFilters) {
            state.customSelected.clear();
        }
        controlsRenderer();
        sectionRenderer();
    });
}

function bindTopicModeFilterEvents(modeRef, searchRef, selectRef, stateRef, controlsRenderer, sectionRenderer) {
    modeRef.addEventListener("change", (event) => {
        stateRef.topicMode = event.target.value;
        if (stateRef.topicMode === "general") {
            stateRef.selfStudyTarget = "all";
            stateRef.selfStudySearch = "";
        } else {
            stateRef.subject = "all";
        }
        controlsRenderer();
        sectionRenderer();
    });

    if (searchRef) {
        searchRef.addEventListener("input", (event) => {
            stateRef.selfStudySearch = event.target.value || "";
            controlsRenderer();
        });
    }

    selectRef.addEventListener("change", (event) => {
        stateRef.selfStudyTarget = event.target.value;
        controlsRenderer();
        sectionRenderer();
    });
}

function refreshAllControls() {
    renderExploreControls();
    renderCompareControls();
    renderAnalysisControls();
    renderClosureControls();
    renderSearchControls();
    renderTrendControls();
    renderGapsControls();
    // Don't render custom controls here - they render when drawer opens
}

function renderViewState() {
    if (refs.navTabs) refs.navTabs.querySelectorAll("[data-view]").forEach((button) => {
        button.classList.toggle("is-active", button.dataset.view === state.view);
        button.setAttribute("aria-current", button.dataset.view === state.view ? "page" : "false");
    });

    if (refs.overviewSection) refs.overviewSection.classList.toggle("hidden", state.view !== "overview");
    if (refs.exploreSection) refs.exploreSection.classList.toggle("hidden", state.view !== "explore");
    if (refs.insightsSection) refs.insightsSection.classList.toggle("hidden", state.view !== "insights");
    if (refs.customSection) refs.customSection.classList.toggle("hidden", state.view !== "custom");
    if (refs.searchSection) refs.searchSection.classList.toggle("hidden", state.view !== "search");
}

function renderCurrentView() {
    if (state.view === "overview") renderOverviewSection();
    if (state.view === "explore") renderExploreSection();
    if (state.view === "custom") { renderCustomControls(); renderCustomSection(); }
    if (state.view === "search") {
        renderSearchControls();
        if (refs.searchInput) refs.searchInput.focus();
    }
    if (state.view === "insights") {
        if (state.insightsSubTab === "analysis") renderAnalysisSection();
        if (state.insightsSubTab === "compare") renderCompareSection();
        if (state.insightsSubTab === "closure") renderClosureSection();
        if (state.insightsSubTab === "trend") renderTrendSection();
        if (state.insightsSubTab === "gaps") renderGapsSection();
    }
}

function renderChipOptions(containerEl, options, activeValue, onChange) {
    if (!containerEl) return containerEl;

    containerEl.innerHTML = options.map(opt =>
        `<button class="chip-option${opt.value === activeValue ? ' is-active' : ''}" type="button" data-value="${escapeHtml(opt.value)}">${escapeHtml(opt.label)}</button>`
    ).join("");

    /* Use a stored handler key to avoid duplicate listeners */
    if (containerEl._chipHandler) {
        containerEl.removeEventListener("click", containerEl._chipHandler);
    }
    containerEl._chipHandler = function (event) {
        const btn = event.target.closest("[data-value]");
        if (!btn) return;
        onChange(btn.dataset.value);
    };
    containerEl.addEventListener("click", containerEl._chipHandler);

    return containerEl;
}

function getSearchModeLabel(mode) {
    return mode === "precise" ? "مطابقة دقيقة" : "بحث مرن";
}

function renderSearchControls() {
    refs.searchModeChips = renderChipOptions(
        refs.searchModeChips,
        [
            { value: "precise", label: "مطابقة دقيقة" },
            { value: "flex", label: "بحث مرن" },
        ],
        state.searchMode,
        (value) => {
            state.searchMode = value;
            renderSearchControls();
            renderSearchResults();
        }
    );

    if (refs.searchMeta) {
        refs.searchMeta.textContent = state.searchMode === "precise"
            ? "مطابقة بالكلمات بعد التطبيع العربي مع تجاهل أل التعريف والفروق الإملائية الشائعة"
            : "بحث مرن يدعم التقارب اللغوي والمرادفات والمطابقة الموسعة";
    }
}

function renderSummaryCards() {
    const summaryCards = [
        { label: "عدد البرامج", value: toArabicNumber(TOTAL_PROGRAMS), note: "البرامج المتاحة في الملف الشامل" },
        { label: "عدد السنوات", value: toArabicNumber(ALL_AVAILABLE_YEARS.length), note: ALL_AVAILABLE_YEARS.map((year) => `${year}هـ`).join(" - ") },
        { label: "عدد الاستطلاعات", value: toArabicNumber(SUMMARY_SURVEY_COUNT), note: "الاستطلاعات الطلابية المستوردة" },
        { label: "مصدر الاستطلاعات", value: SOURCE_LABEL, note: "يمكن إضافة مصادر أخرى لاحقًا" },
    ];

    if (!refs.summaryGrid) return;
    refs.summaryGrid.innerHTML = summaryCards.map((card) => `
        <article class="summary-card">
            <div class="summary-label">${escapeHtml(card.label)}</div>
            <div class="summary-value">${escapeHtml(card.value)}</div>
            <div class="summary-note">${escapeHtml(card.note)}</div>
        </article>
    `).join("");

    if (refs.footerSource) refs.footerSource.textContent = `المصدر الحالي: ${SOURCE_LABEL}`;
}

function renderExploreControls() {
    if (!refs.exploreProgram) return;

    renderProgramOptionsFor(refs.exploreProgram, state.exploreFilters.program, true);

    const years = getAvailableYears(state.exploreFilters.program);
    refs.exploreYearChips = renderChipOptions(
        refs.exploreYearChips,
        [{value: "all", label: "كل السنوات"}, ...years.map(y => ({value: y, label: y + "هـ"}))],
        state.exploreFilters.year,
        (val) => { state.exploreFilters.year = val; renderExploreControls(); renderExploreSection(); }
    );

    const stakeholders = collectUnique(ITEM_RECORDS.map(r => r.stakeholder), v => v);
    refs.exploreStakeholderChips = renderChipOptions(
        refs.exploreStakeholderChips,
        [{value: "all", label: "كل الجهات"}, ...stakeholders.map(s => ({value: s, label: getStakeholderLabel(s)}))],
        state.exploreFilters.stakeholder,
        (val) => { state.exploreFilters.stakeholder = val; renderExploreControls(); renderExploreSection(); }
    );

    if (!AVAILABLE_GENDERS.length) {
        if (refs.exploreGenderChips) refs.exploreGenderChips.innerHTML = '<span class="chip-option is-active">غير متاح</span>';
        if (refs.exploreGenderHint) refs.exploreGenderHint.classList.remove("hidden");
    } else {
        if (refs.exploreGenderHint) refs.exploreGenderHint.classList.add("hidden");
        refs.exploreGenderChips = renderChipOptions(
            refs.exploreGenderChips,
            [{value: "all", label: "الكل"}, ...AVAILABLE_GENDERS.map(g => ({value: g, label: getGenderLabel(g)}))],
            state.exploreFilters.gender,
            (val) => { state.exploreFilters.gender = val; renderExploreControls(); renderExploreSection(); }
        );
    }

    refs.exploreTopicModeChips = renderChipOptions(
        refs.exploreTopicModeChips,
        [{value: "general", label: "موضوعات عامة"}, {value: "selfstudy", label: "محكات الدراسة الذاتية"}],
        state.exploreFilters.topicMode,
        (val) => {
            state.exploreFilters.topicMode = val;
            if (val === "general") { state.exploreFilters.selfStudyTarget = "all"; }
            else { state.exploreFilters.subject = "all"; }
            renderExploreControls(); renderExploreSection();
        }
    );

    renderSubjectOptionsForExplore();
    renderSelfStudyOptionsForExplore();
    toggleTopicModePanels(refs.exploreGeneralTopicPanel, refs.exploreSelfStudyPanel, state.exploreFilters.topicMode);
    renderFilterChips(refs.exploreActiveFilters, buildExploreFilterChips());
}

function renderSubjectOptionsForExplore() {
    if (!refs.exploreSubject) return;
    const baseRecords = ITEM_RECORDS.filter((record) => {
        if (state.exploreFilters.program !== "all" && record.programId !== state.exploreFilters.program) return false;
        if (state.exploreFilters.year !== "all" && record.year !== state.exploreFilters.year) return false;
        if (state.exploreFilters.stakeholder !== "all" && record.stakeholder !== state.exploreFilters.stakeholder) return false;
        if (state.exploreFilters.gender !== "all" && record.gender !== state.exploreFilters.gender) return false;
        return true;
    });
    renderSubjectOptions(refs.exploreSubject, baseRecords, state.exploreFilters, "subject");
}

function renderSelfStudyOptionsForExplore() {
    if (!refs.exploreSelfStudyFilter) return;
    const baseRecords = ITEM_RECORDS.filter((record) => {
        if (state.exploreFilters.program !== "all" && record.programId !== state.exploreFilters.program) return false;
        if (state.exploreFilters.year !== "all" && record.year !== state.exploreFilters.year) return false;
        if (state.exploreFilters.stakeholder !== "all" && record.stakeholder !== state.exploreFilters.stakeholder) return false;
        if (state.exploreFilters.gender !== "all" && record.gender !== state.exploreFilters.gender) return false;
        return true;
    });
    renderSelfStudyOptions(refs.exploreSelfStudyFilter, null, baseRecords, state.exploreFilters);
}

function buildExploreFilterChips() {
    return [
        { label: "البرنامج", value: state.exploreFilters.program === "all" ? "الكل" : formatProgramLabel(getProgramById(state.exploreFilters.program)) },
        { label: "السنة", value: state.exploreFilters.year === "all" ? "الكل" : `${state.exploreFilters.year}هـ` },
        { label: "الجهة", value: state.exploreFilters.stakeholder === "all" ? "الكل" : getStakeholderLabel(state.exploreFilters.stakeholder) },
        { label: "نوع الموضوع", value: getTopicModeLabel(state.exploreFilters.topicMode) },
        { label: "التحديد", value: getTopicSelectionLabel(state.exploreFilters) },
        { label: "الجنس", value: getGenderFilterLabel(state.exploreFilters.gender) },
    ];
}

// renderProgramControls removed — replaced by renderExploreControls

function renderCompareControls() {
    if (refs.compareStakeholderChips) {
        const stakeholders = collectUnique(ITEM_RECORDS.map(r => r.stakeholder), v => v);
        refs.compareStakeholderChips = renderChipOptions(
            refs.compareStakeholderChips,
            [{value: "all", label: "كل الجهات"}, ...stakeholders.map(s => ({value: s, label: getStakeholderLabel(s)}))],
            state.compareFilters.stakeholder,
            (val) => { state.compareFilters.stakeholder = val; renderCompareControls(); renderCompareSection(); }
        );
    }

    if (refs.compareTopicModeChips) {
        refs.compareTopicModeChips = renderChipOptions(
            refs.compareTopicModeChips,
            [{value: "general", label: "موضوعات عامة"}, {value: "selfstudy", label: "محكات الدراسة الذاتية"}],
            state.compareFilters.topicMode,
            (val) => {
                state.compareFilters.topicMode = val;
                if (val === "general") { state.compareFilters.selfStudyTarget = "all"; }
                else { state.compareFilters.subject = "all"; }
                renderCompareControls(); renderCompareSection();
            }
        );
    }

    renderSubjectOptionsForCompare();
    renderSelfStudyOptionsForCompare();
    toggleTopicModePanels(
        refs.compareGeneralTopicPanel,
        refs.compareSelfStudyPanel,
        state.compareFilters.topicMode
    );
    renderCompareSlotOptions("a");
    renderCompareSlotOptions("b");
    renderCompareSlotOptions("c");
    renderFilterChips(refs.compareActiveFilters, [
        { label: "الجهة", value: state.compareFilters.stakeholder === "all" ? "الكل" : getStakeholderLabel(state.compareFilters.stakeholder) },
        { label: "نوع الموضوع", value: getTopicModeLabel(state.compareFilters.topicMode) },
        { label: "التحديد", value: getTopicSelectionLabel(state.compareFilters) },
    ]);
}

function renderAnalysisControls() {
    if (refs.analysisYearChips) {
        refs.analysisYearChips = renderChipOptions(
            refs.analysisYearChips,
            [{value: "all", label: "كل السنوات"}, ...ALL_AVAILABLE_YEARS.map(y => ({value: y, label: y + "هـ"}))],
            state.analysisFilters.year,
            (val) => { state.analysisFilters.year = val; renderAnalysisControls(); renderAnalysisSection(); }
        );
    }

    if (refs.analysisStakeholderChips) {
        const stakeholders = collectUnique(ITEM_RECORDS.map(r => r.stakeholder), v => v);
        refs.analysisStakeholderChips = renderChipOptions(
            refs.analysisStakeholderChips,
            [{value: "all", label: "كل الجهات"}, ...stakeholders.map(s => ({value: s, label: getStakeholderLabel(s)}))],
            state.analysisFilters.stakeholder,
            (val) => { state.analysisFilters.stakeholder = val; renderAnalysisControls(); renderAnalysisSection(); }
        );
    }

    if (refs.analysisTopicModeChips) {
        refs.analysisTopicModeChips = renderChipOptions(
            refs.analysisTopicModeChips,
            [{value: "general", label: "موضوعات عامة"}, {value: "selfstudy", label: "محكات الدراسة الذاتية"}],
            state.analysisFilters.topicMode,
            (val) => {
                state.analysisFilters.topicMode = val;
                if (val === "general") { state.analysisFilters.selfStudyTarget = "all"; }
                else { state.analysisFilters.subject = "all"; }
                renderAnalysisControls(); renderAnalysisSection();
            }
        );
    }

    renderSubjectOptionsForAnalysis();
    renderSelfStudyOptionsForAnalysis();
    toggleTopicModePanels(
        refs.analysisGeneralTopicPanel,
        refs.analysisSelfStudyPanel,
        state.analysisFilters.topicMode
    );
    renderAnalysisProgramsChecklist();
    renderFilterChips(refs.analysisActiveFilters, [
        { label: "البرامج المحددة", value: state.analysisFilters.programs.size ? toArabicNumber(state.analysisFilters.programs.size) : "غير محدد" },
        { label: "السنة", value: state.analysisFilters.year === "all" ? "الكل" : `${state.analysisFilters.year}هـ` },
        { label: "الجهة", value: state.analysisFilters.stakeholder === "all" ? "الكل" : getStakeholderLabel(state.analysisFilters.stakeholder) },
        { label: "نوع الموضوع", value: getTopicModeLabel(state.analysisFilters.topicMode) },
        { label: "التحديد", value: getTopicSelectionLabel(state.analysisFilters) },
    ]);
}

function renderClosureControls() {
    if (!refs.closureProgramSelect) return;

    renderProgramOptionsFor(refs.closureProgramSelect, state.closureFilters.program, false, {
        includeEmpty: true,
        emptyLabel: "اختر البرنامج",
    });

    normalizeClosureYearPair();
    const years = state.closureFilters.program ? getAvailableYears(state.closureFilters.program) : ALL_AVAILABLE_YEARS;
    renderYearOptionsFromList(refs.closureFromYearSelect, state.closureFilters.fromYear, years, false);
    renderYearOptionsFromList(refs.closureToYearSelect, state.closureFilters.toYear, years, false);

    if (refs.closureLevelSelect) {
        refs.closureLevelSelect.innerHTML = [
            `<option value="item">مطابقة مضبوطة للبند الواحد</option>`,
            `<option value="expanded">المطابقة المتوسعة</option>`,
            `<option value="veryExpanded">المطابقة المتوسعة جدًا</option>`,
        ].join("");
        refs.closureLevelSelect.value = ["item", "expanded", "veryExpanded"].includes(state.closureFilters.level)
            ? state.closureFilters.level
            : "item";
        state.closureFilters.level = refs.closureLevelSelect.value;
    }

    if (refs.closureMinImprovementInput) {
        refs.closureMinImprovementInput.value = String(state.closureFilters.minImprovement);
    }

    refs.closureTopicModeChips = renderChipOptions(
        refs.closureTopicModeChips,
        [{ value: "general", label: "موضوعات عامة" }, { value: "selfstudy", label: "محكات الدراسة الذاتية" }],
        state.closureFilters.topicMode,
        (value) => {
            state.closureFilters.topicMode = value;
            if (value === "general") {
                state.closureFilters.selfStudyTarget = "all";
            } else {
                state.closureFilters.subject = "all";
            }
            renderClosureControls();
            renderClosureSection();
        }
    );

    if (!AVAILABLE_GENDERS.length) {
        if (refs.closureGenderChips) refs.closureGenderChips.innerHTML = '<span class="chip-option is-active">غير متاح</span>';
        if (refs.closureGenderHint) refs.closureGenderHint.classList.remove("hidden");
        state.closureFilters.gender = "all";
    } else {
        if (refs.closureGenderHint) refs.closureGenderHint.classList.add("hidden");
        refs.closureGenderChips = renderChipOptions(
            refs.closureGenderChips,
            [{ value: "all", label: "الكل" }, ...AVAILABLE_GENDERS.map((gender) => ({ value: gender, label: getGenderLabel(gender) }))],
            state.closureFilters.gender,
            (value) => {
                state.closureFilters.gender = value;
                renderClosureControls();
                renderClosureSection();
            }
        );
    }

    renderSubjectOptionsForClosure();
    renderSelfStudyOptionsForClosure();
    toggleTopicModePanels(refs.closureGeneralTopicPanel, refs.closureSelfStudyPanel, state.closureFilters.topicMode);
    renderClosureDisplayTabs();
    renderClosureReportOptionToggles();
    renderFilterChips(refs.closureActiveFilters, buildClosureFilterChips());
}

function renderSubjectOptionsForClosure() {
    renderSubjectOptions(refs.closureSubjectFilter, getClosureBaseRecords(), state.closureFilters, "subject");
}

function renderSelfStudyOptionsForClosure() {
    renderSelfStudyOptions(refs.closureSelfStudyFilter, null, getClosureBaseRecords(), state.closureFilters);
}

function getClosureBaseRecords() {
    if (!state.closureFilters.program) return [];
    const years = new Set([state.closureFilters.fromYear, state.closureFilters.toYear].filter(Boolean));
    return ITEM_RECORDS.filter((record) =>
        record.programId === state.closureFilters.program &&
        years.has(record.year) &&
        record.stakeholder === "students" &&
        (state.closureFilters.gender === "all" || record.gender === state.closureFilters.gender)
    );
}

function normalizeClosureYearPair(changedKey = "") {
    const years = [...(state.closureFilters.program ? getAvailableYears(state.closureFilters.program) : ALL_AVAILABLE_YEARS)]
        .sort((first, second) => Number(first) - Number(second));

    if (!years.length) {
        state.closureFilters.fromYear = "";
        state.closureFilters.toYear = "";
        return;
    }

    if (!years.includes(state.closureFilters.fromYear)) {
        state.closureFilters.fromYear = years[0];
    }

    if (!years.includes(state.closureFilters.toYear)) {
        state.closureFilters.toYear = years[years.length - 1];
    }

    if (years.length > 1 && state.closureFilters.fromYear === state.closureFilters.toYear) {
        if (changedKey === "from") {
            state.closureFilters.toYear = years.find((year) => year !== state.closureFilters.fromYear) || state.closureFilters.toYear;
        } else {
            state.closureFilters.fromYear = [...years].reverse().find((year) => year !== state.closureFilters.toYear) || state.closureFilters.fromYear;
        }
    }

    if (state.closureFilters.fromYear && state.closureFilters.toYear && Number(state.closureFilters.fromYear) > Number(state.closureFilters.toYear)) {
        const olderYear = state.closureFilters.toYear;
        state.closureFilters.toYear = state.closureFilters.fromYear;
        state.closureFilters.fromYear = olderYear;
    }
}

function buildClosureFilterChips() {
    return [
        { label: "البرنامج", value: state.closureFilters.program ? formatProgramLabel(getProgramById(state.closureFilters.program)) : "غير محدد" },
        { label: "الفترة", value: state.closureFilters.fromYear && state.closureFilters.toYear ? `${state.closureFilters.fromYear}هـ ← ${state.closureFilters.toYear}هـ` : "غير مكتملة" },
        { label: "نوع الموضوع", value: getTopicModeLabel(state.closureFilters.topicMode) },
        { label: "التحديد", value: getTopicSelectionLabel(state.closureFilters) },
        { label: "الجنس", value: getGenderFilterLabel(state.closureFilters.gender) },
        { label: "المطابقة", value: getClosureLevelLabel(state.closureFilters.level) },
        { label: "الحد الأدنى", value: formatClosureThreshold(state.closureFilters.minImprovement) },
    ];
}

function getClosureLevelLabel(level) {
    if (level === "item") return "مطابقة مضبوطة للبند الواحد";
    if (level === "expanded") return "المطابقة المتوسعة";
    if (level === "veryExpanded") return "المطابقة المتوسعة جدًا";
    if (level === "survey") return "الاستطلاع الكامل";
    return "الموضوع الدقيق";
}

function getClosureThreshold() {
    const rawValue = Number(state.closureFilters.minImprovement);
    return Number.isFinite(rawValue) ? Math.max(0, rawValue) : 2;
}

function renderClosureDisplayTabs() {
    if (!refs.closureDisplayTabs) return;
    refs.closureDisplayTabs.querySelectorAll("[data-closure-display]").forEach((button) => {
        button.classList.toggle("is-active", button.dataset.closureDisplay === state.closureFilters.displayMode);
    });
}

function renderClosureReportOptionToggles() {
    if (!refs.closureReportOptionToggles) return;
    const options = [
        { key: "responses", label: "عدد المقيمين", active: state.closureFilters.reportShowResponses },
        { key: "statement", label: "العبارة", active: state.closureFilters.reportShowStatement },
        { key: "source", label: "المصدر", active: state.closureFilters.reportShowSource },
    ];
    refs.closureReportOptionToggles.innerHTML = options.map((option) => `
        <button
            class="chip-option${option.active ? " is-active" : ""}"
            type="button"
            data-report-option="${option.key}"
        >${escapeHtml(option.label)}</button>
    `).join("");
}

function renderCustomControls() {
    renderProgramOptionsFor(refs.customProgramFilter, state.customFilters.program, true);

    if (refs.customYearChips) {
        const years = getAvailableYears(state.customFilters.program);
        refs.customYearChips = renderChipOptions(
            refs.customYearChips,
            [{value: "all", label: "كل السنوات"}, ...years.map(y => ({value: y, label: y + "هـ"}))],
            state.customFilters.year,
            (val) => { state.customFilters.year = val; renderCustomControls(); renderCustomSection(); }
        );
    }

    if (refs.customStakeholderChips) {
        const stakeholders = collectUnique(ITEM_RECORDS.map(r => r.stakeholder), v => v);
        refs.customStakeholderChips = renderChipOptions(
            refs.customStakeholderChips,
            [{value: "all", label: "كل الجهات"}, ...stakeholders.map(s => ({value: s, label: getStakeholderLabel(s)}))],
            state.customFilters.stakeholder,
            (val) => { state.customFilters.stakeholder = val; renderCustomControls(); renderCustomSection(); }
        );
    }

    renderSubjectOptionsForSingle(refs.customSubjectFilter, state.customFilters);

    if (refs.customGenderChips) {
        if (!AVAILABLE_GENDERS.length) {
            refs.customGenderChips.innerHTML = '<span class="chip-option is-active">غير متاح</span>';
            if (refs.customGenderHint) refs.customGenderHint.classList.remove("hidden");
        } else {
            if (refs.customGenderHint) refs.customGenderHint.classList.add("hidden");
            refs.customGenderChips = renderChipOptions(
                refs.customGenderChips,
                [{value: "all", label: "الكل"}, ...AVAILABLE_GENDERS.map(g => ({value: g, label: getGenderLabel(g)}))],
                state.customFilters.gender,
                (val) => { state.customFilters.gender = val; renderCustomControls(); renderCustomSection(); }
            );
        }
    }

    renderFilterChips(refs.customActiveFilters, buildSingleFilterChips(state.customFilters));
    pruneCustomSelection();
}

function renderProgramOptionsFor(selectRef, currentValue, includeAll, options = {}) {
    const programs = getProgramsWithData();
    selectRef.innerHTML = [
        options.includeEmpty ? `<option value="">${escapeHtml(options.emptyLabel || "غير محدد")}</option>` : "",
        includeAll ? `<option value="all">كل البرامج</option>` : "",
        ...programs.map((program) => `<option value="${program.id}">${escapeHtml(formatProgramLabel(program))}</option>`),
    ].join("");
    const hasCurrentProgram = programs.some((program) => program.id === currentValue);
    if (hasCurrentProgram || (includeAll && currentValue === "all") || (options.includeEmpty && currentValue === "")) {
        selectRef.value = currentValue;
    } else {
        selectRef.value = options.includeEmpty ? "" : "all";
    }
    renderFilterableSelect(selectRef, {
        kind: "program",
        searchPlaceholder: "ابحث في البرامج",
        emptyLabel: "لا توجد برامج مطابقة",
    });
}

function renderYearOptionsFor(selectRef, currentValue, programId, includeAll) {
    const years = getAvailableYears(programId);
    if (currentValue !== "all" && !years.includes(currentValue)) {
        currentValue = years[0] || "all";
    }
    renderYearOptionsFromList(selectRef, currentValue, years, includeAll);
}

function renderYearOptionsFromList(selectRef, currentValue, years, includeAll, options = {}) {
    selectRef.innerHTML = [
        options.includeEmpty ? `<option value="">${escapeHtml(options.emptyLabel || "غير محدد")}</option>` : "",
        includeAll ? `<option value="all">كل السنوات</option>` : "",
        ...years.map((year) => `<option value="${year}">${year}هـ</option>`),
    ].join("");
    if (currentValue === "all" || years.includes(currentValue) || (options.includeEmpty && currentValue === "")) {
        selectRef.value = currentValue;
    } else {
        selectRef.value = options.includeEmpty ? "" : (includeAll ? "all" : (years[0] || ""));
    }
}

function renderStakeholderOptionsFor(selectRef, currentValue) {
    const stakeholders = collectUnique(ITEM_RECORDS.map((record) => record.stakeholder), (value) => value);
    selectRef.innerHTML = [
        `<option value="all">كل الجهات</option>`,
        ...stakeholders.map((stakeholder) => `<option value="${stakeholder}">${escapeHtml(getStakeholderLabel(stakeholder))}</option>`),
    ].join("");
    selectRef.value = stakeholders.includes(currentValue) || currentValue === "all" ? currentValue : "all";
}

function renderTopicModeOptionsFor(selectRef, currentValue) {
    const options = [
        `<option value="general">موضوعات عامة</option>`,
        `<option value="selfstudy">محكات الدراسة الذاتية</option>`,
    ];
    selectRef.innerHTML = options.join("");
    selectRef.value = ["general", "selfstudy"].includes(currentValue) ? currentValue : "general";
}

function renderSubjectOptionsForSingle(selectRef, filters) {
    const baseRecords = ITEM_RECORDS.filter((record) => {
        if (filters.program !== "all" && record.programId !== filters.program) return false;
        if (filters.year !== "all" && record.year !== filters.year) return false;
        if (filters.stakeholder !== "all" && record.stakeholder !== filters.stakeholder) return false;
        if (filters.gender !== "all" && record.gender !== filters.gender) return false;
        return true;
    });
    renderSubjectOptions(selectRef, baseRecords, filters, "subject");
}

function renderSelfStudyOptionsForSingle(selectRef, searchRef, filters) {
    const baseRecords = ITEM_RECORDS.filter((record) => {
        if (filters.program !== "all" && record.programId !== filters.program) return false;
        if (filters.year !== "all" && record.year !== filters.year) return false;
        if (filters.stakeholder !== "all" && record.stakeholder !== filters.stakeholder) return false;
        if (filters.gender !== "all" && record.gender !== filters.gender) return false;
        return true;
    });
    renderSelfStudyOptions(selectRef, searchRef, baseRecords, filters);
}

function renderSubjectOptionsForCompare() {
    renderSubjectOptions(refs.compareSubjectFilter, getCompareBaseRecords(), state.compareFilters, "subject");
}

function renderSelfStudyOptionsForCompare() {
    renderSelfStudyOptions(refs.compareSelfStudyFilter, refs.compareSelfStudySearch, getCompareBaseRecords(), state.compareFilters);
}

function getCompareBaseRecords() {
    const activeSlots = getActiveCompareSlots();
    return ITEM_RECORDS.filter((record) => {
        if (state.compareFilters.stakeholder !== "all" && record.stakeholder !== state.compareFilters.stakeholder) return false;
        if (!activeSlots.length) return true;
        return activeSlots.some((slot) => record.programId === slot.program && record.year === slot.year);
    });
}

function renderSubjectOptionsForAnalysis() {
    if (!state.analysisFilters.programs.size) {
        renderSubjectOptions(refs.analysisSubjectFilter, [], state.analysisFilters, "subject");
        return;
    }

    const baseRecords = ITEM_RECORDS.filter((record) => {
        if (state.analysisFilters.programs.size && !state.analysisFilters.programs.has(record.programId)) return false;
        if (state.analysisFilters.year !== "all" && record.year !== state.analysisFilters.year) return false;
        if (state.analysisFilters.stakeholder !== "all" && record.stakeholder !== state.analysisFilters.stakeholder) return false;
        return true;
    });
    renderSubjectOptions(refs.analysisSubjectFilter, baseRecords, state.analysisFilters, "subject");
}

function renderSelfStudyOptionsForAnalysis() {
    if (!state.analysisFilters.programs.size) {
        renderSelfStudyOptions(refs.analysisSelfStudyFilter, refs.analysisSelfStudySearch, [], state.analysisFilters);
        return;
    }

    const baseRecords = ITEM_RECORDS.filter((record) => {
        if (state.analysisFilters.programs.size && !state.analysisFilters.programs.has(record.programId)) return false;
        if (state.analysisFilters.year !== "all" && record.year !== state.analysisFilters.year) return false;
        if (state.analysisFilters.stakeholder !== "all" && record.stakeholder !== state.analysisFilters.stakeholder) return false;
        return true;
    });
    renderSelfStudyOptions(refs.analysisSelfStudyFilter, refs.analysisSelfStudySearch, baseRecords, state.analysisFilters);
}

function renderSubjectOptions(selectRef, baseRecords, targetState, key) {
    const sections = collectUnique(baseRecords.map((record) => record.sectionId), (value) => value);
    const surveys = collectUnique(baseRecords.map((record) => record.surveyTitle), (value) => normalizeText(value));
    const topics = collectUnique(baseRecords.map((record) => record.topicLabel), (value) => normalizeText(value));

    const allowedValues = new Set([
        "all",
        ...sections.map((sectionId) => `section:${sectionId}`),
        ...surveys.map((surveyTitle) => `survey:${surveyTitle}`),
        ...topics.map((topicLabel) => `topic:${topicLabel}`),
    ]);

    if (!allowedValues.has(targetState[key])) {
        targetState[key] = "all";
    }

    selectRef.innerHTML = [
        `<option value="all">كل الموضوعات</option>`,
        sections.length ? `
            <optgroup label="المحاور">
                ${sections.map((sectionId) => `<option value="section:${sectionId}">${escapeHtml(getSectionLabel(sectionId))}</option>`).join("")}
            </optgroup>
        ` : "",
        surveys.length ? `
            <optgroup label="الاستطلاعات">
                ${surveys.map((surveyTitle) => `<option value="survey:${escapeHtml(surveyTitle)}">${escapeHtml(surveyTitle)}</option>`).join("")}
            </optgroup>
        ` : "",
        topics.length ? `
            <optgroup label="الموضوعات">
                ${topics.map((topicLabel) => `<option value="topic:${escapeHtml(topicLabel)}">${escapeHtml(topicLabel)}</option>`).join("")}
            </optgroup>
        ` : "",
    ].join("");
    selectRef.value = targetState[key];
    renderFilterableSelect(selectRef, {
        kind: "general",
        searchPlaceholder: "ابحث في الموضوعات العامة",
        emptyLabel: "لا توجد موضوعات مطابقة",
    });
}

function renderSelfStudyOptions(selectRef, searchRef, baseRecords, targetState) {
    if (!selectRef) return;
    const allOptions = collectSelfStudyOptions(baseRecords);
    const hasUnlinkedItems = baseRecords.some((record) => !(record.selfStudyEntries || []).length);
    if (hasUnlinkedItems) {
        allOptions.unshift({
            value: "unlinked",
            label: "غير مرتبط بمحك | عبارات بلا ربط بمحك",
            searchText: normalizeText("غير مرتبط بمحك عبارات بلا ربط بمحك"),
        });
    }
    const allowedValues = new Set(["all", ...allOptions.map((option) => option.value)]);
    if (!allowedValues.has(targetState.selfStudyTarget)) {
        targetState.selfStudyTarget = "all";
    }

    if (searchRef) {
        searchRef.value = "";
        searchRef.disabled = true;
    }
    targetState.selfStudySearch = "";

    selectRef.disabled = !allOptions.length;
    selectRef.innerHTML = [
        `<option value="all">${allOptions.length ? "كل محكات الدراسة الذاتية" : "لا توجد محكات مطابقة"}</option>`,
        ...allOptions.map((option) => `<option value="${escapeHtml(option.value)}">${escapeHtml(option.label)}</option>`),
    ].join("");
    selectRef.value = targetState.selfStudyTarget;
    renderFilterableSelect(selectRef, {
        kind: "selfstudy",
        searchPlaceholder: "ابحث بالمحك أو الجانب أو نص العبارة",
        emptyLabel: "لا توجد محكات مطابقة",
    });
}

function bindFilterableDropdownGlobalEvents() {
    if (filterableDropdownEventsBound) return;
    filterableDropdownEventsBound = true;

    document.addEventListener("click", (event) => {
        document.querySelectorAll(".filterable-select.is-open").forEach((dropdown) => {
            if (!dropdown.contains(event.target)) {
                dropdown.classList.remove("is-open");
            }
        });
    });

    document.addEventListener("keydown", (event) => {
        if (event.key !== "Escape") return;
        document.querySelectorAll(".filterable-select.is-open").forEach((dropdown) => {
            dropdown.classList.remove("is-open");
        });
    });
}

function renderFilterableSelect(selectRef, config = {}) {
    if (!selectRef) return;

    const fieldGroup = selectRef.closest(".field-group");
    if (!fieldGroup) return;

    selectRef.classList.add("native-select-hidden");

    let wrapper = fieldGroup.querySelector(".filterable-select");
    if (!wrapper) {
        wrapper = document.createElement("div");
        wrapper.className = "filterable-select";
        fieldGroup.appendChild(wrapper);
    }

    const items = extractFilterableItems(selectRef, config.kind);
    const selectedItem = items.find((item) => item.value === selectRef.value) || items[0] || null;
    const triggerLabel = selectedItem?.triggerLabel || selectedItem?.title || config.emptyLabel || "لا توجد خيارات";

    wrapper.classList.toggle("is-open", false);
    wrapper.innerHTML = `
        <button class="filterable-select__trigger" type="button" ${selectRef.disabled ? "disabled" : ""}>
            <span class="filterable-select__trigger-label">${escapeHtml(triggerLabel)}</span>
            <span class="filterable-select__trigger-icon">▾</span>
        </button>
        <div class="filterable-select__menu">
            <div class="filterable-select__search">
                <input
                    type="search"
                    placeholder="${escapeHtml(config.searchPlaceholder || "ابحث في القائمة")}"
                    ${items.length ? "" : "disabled"}
                >
            </div>
            <div class="filterable-select__options"></div>
        </div>
    `;

    const trigger = wrapper.querySelector(".filterable-select__trigger");
    const searchInput = wrapper.querySelector(".filterable-select__search input");
    const optionsHost = wrapper.querySelector(".filterable-select__options");

    const paintOptions = () => {
        const query = searchInput.value || "";
        const visibleItems = items.filter((item) => item.isAll || !query || searchMatches(query, item.searchText));

        if (!visibleItems.length) {
            optionsHost.innerHTML = `<div class="filterable-select__empty">${escapeHtml(config.emptyLabel || "لا توجد نتائج مطابقة")}</div>`;
            return;
        }

        optionsHost.innerHTML = visibleItems
            .map((item) => renderFilterableItem(item, selectRef.value))
            .join("");
    };

    trigger.addEventListener("click", () => {
        if (trigger.disabled) return;
        const willOpen = !wrapper.classList.contains("is-open");
        document.querySelectorAll(".filterable-select.is-open").forEach((dropdown) => {
            if (dropdown !== wrapper) dropdown.classList.remove("is-open");
        });
        wrapper.classList.toggle("is-open", willOpen);
        if (willOpen) {
            searchInput.value = "";
            paintOptions();
            searchInput.focus();
        }
    });

    searchInput.addEventListener("input", () => {
        paintOptions();
    });

    optionsHost.addEventListener("click", (event) => {
        const optionButton = event.target.closest("[data-value]");
        if (!optionButton) return;

        const nextValue = optionButton.dataset.value;
        if (!nextValue) return;

        selectRef.value = nextValue;

        /* Update trigger label to reflect the new selection */
        const selectedItem = items.find((item) => item.value === nextValue);
        if (selectedItem) {
            const triggerLabelEl = wrapper.querySelector(".filterable-select__trigger-label");
            if (triggerLabelEl) triggerLabelEl.textContent = selectedItem.triggerLabel || selectedItem.title;
        }

        /* Mark selected option visually */
        wrapper.querySelectorAll(".filterable-select__option").forEach((btn) => {
            btn.classList.toggle("is-selected", btn.dataset.value === nextValue);
        });

        selectRef.dispatchEvent(new Event("change", { bubbles: true }));
        wrapper.classList.remove("is-open");
    });

    paintOptions();
}

function extractFilterableItems(selectRef, kind) {
    return Array.from(selectRef.children).flatMap((node) => {
        if (node.tagName === "OPTGROUP") {
            return Array.from(node.children).map((option) => buildFilterableItem(option, kind, node.label || ""));
        }
        if (node.tagName === "OPTION") {
            return [buildFilterableItem(node, kind, "")];
        }
        return [];
    }).filter(Boolean);
}

function buildFilterableItem(optionNode, kind, groupLabel) {
    const label = normalizeText(optionNode.textContent || "");
    if (!label) return null;

    if (optionNode.value === "all") {
        return {
            value: optionNode.value,
            tone: "all",
            badge: "الكل",
            title: label,
            meta: kind === "selfstudy" ? "يعرض جميع العناصر المرتبطة بالمحكات وغير المرتبطة بها" : "يعرض جميع الخيارات المتاحة",
            triggerLabel: label,
            searchText: label,
            isAll: true,
        };
    }

    if (kind === "selfstudy") {
        return buildSelfStudyFilterableItem(optionNode.value, label);
    }

    if (kind === "program") {
        const parts = label.split(" - ");
        return {
            value: optionNode.value,
            tone: "program",
            badge: parts[1] || "برنامج",
            title: parts[0] || label,
            meta: parts[1] || "",
            triggerLabel: label,
            searchText: normalizeText(label),
            isAll: false,
        };
    }

    return buildGeneralFilterableItem(optionNode.value, label, groupLabel);
}

function buildGeneralFilterableItem(value, label, groupLabel) {
    const toneMeta = {
        المحاور: { tone: "section", badge: "محور" },
        الاستطلاعات: { tone: "survey", badge: "استطلاع" },
        الموضوعات: { tone: "topic", badge: "موضوع" },
    };

    const meta = toneMeta[groupLabel] || { tone: "topic", badge: "موضوع" };
    return {
        value,
        tone: meta.tone,
        badge: meta.badge,
        title: label,
        meta: groupLabel || "",
        triggerLabel: label,
        searchText: normalizeText([label, groupLabel].join(" ")),
        isAll: false,
    };
}

function buildSelfStudyFilterableItem(value, label) {
    if (value === "unlinked") {
        return {
            value,
            tone: "phrase",
            badge: "بدون محك",
            title: "غير مرتبط بمحك",
            meta: "يعرض العبارات التي لم تُربط بأي محك",
            triggerLabel: "غير مرتبط بمحك",
            searchText: normalizeText([label, value].join(" ")),
            isAll: false,
        };
    }

    const [head, ...tailParts] = label.split("|");
    const mainHead = normalizeText(head.replace(/^↳+/g, "").replace(/^العبارة/, "نص العبارة").trim());
    const detail = normalizeText(tailParts.join("|"));
    const valueSearchText = normalizeText(value.replace(/^[a-z]+:/, "").replace(/\|\|/g, " "));

    if (value.startsWith("criterion:")) {
        return {
            value,
            tone: "criterion",
            badge: "محك",
            title: mainHead,
            meta: detail,
            triggerLabel: detail ? `${mainHead} - ${detail}` : mainHead,
            searchText: normalizeText([mainHead, detail, valueSearchText].join(" ")),
            isAll: false,
        };
    }

    if (value.startsWith("side:")) {
        return {
            value,
            tone: "side",
            badge: "جانب",
            title: "الجانب المدعوم",
            meta: detail || mainHead.replace(/^الجانب/, "").trim(),
            triggerLabel: detail || mainHead,
            searchText: normalizeText([mainHead, detail, valueSearchText].join(" ")),
            isAll: false,
        };
    }

    return {
        value,
        tone: "phrase",
        badge: "عبارة",
        title: "نص العبارة",
        meta: detail || mainHead,
        triggerLabel: detail || mainHead,
        searchText: normalizeText([mainHead, detail, valueSearchText].join(" ")),
        isAll: false,
    };
}

function renderFilterableItem(item, selectedValue) {
    return `
        <button
            class="filterable-select__option${item.value === selectedValue ? " is-selected" : ""}"
            data-tone="${escapeHtml(item.tone)}"
            data-value="${escapeHtml(item.value)}"
            type="button"
        >
            <span class="filterable-select__option-head">
                <span class="filterable-select__title">${escapeHtml(item.title)}</span>
                <span class="filterable-select__badge" data-tone="${escapeHtml(item.tone)}">${escapeHtml(item.badge)}</span>
            </span>
            ${item.meta ? `<span class="filterable-select__meta">${escapeHtml(item.meta)}</span>` : ""}
        </button>
    `;
}

function collectSelfStudyOptions(baseRecords) {
    const criteria = new Map();

    baseRecords.forEach((record) => {
        (record.selfStudyEntries || []).forEach((entry) => {
            const criterionValue = buildSelfStudyCriterionValue(entry);
            if (!criteria.has(criterionValue)) {
                criteria.set(criterionValue, {
                    type: "criterion",
                    value: criterionValue,
                    label: `المحك ${entry.criterionCode} | ${entry.criterionText}`,
                    searchText: normalizeText([entry.criterionCode, entry.criterionText, entry.standard, entry.section].join(" ")),
                    criterionCode: entry.criterionCode,
                    sides: new Map(),
                });
            }

            const criterionEntry = criteria.get(criterionValue);
            const sideValue = buildSelfStudySideValue(entry);
            if (!criterionEntry.sides.has(sideValue)) {
                criterionEntry.sides.set(sideValue, {
                    type: "side",
                    value: sideValue,
                    label: `↳ الجانب | ${entry.supportedSide}`,
                    searchText: normalizeText([entry.criterionCode, entry.criterionText, entry.supportedSide].join(" ")),
                    supportedSide: entry.supportedSide,
                    phrases: new Map(),
                });
            }

            const sideEntry = criterionEntry.sides.get(sideValue);
            const phraseValue = buildSelfStudyPhraseValue(entry, record.itemLabel);
            if (!sideEntry.phrases.has(phraseValue)) {
                sideEntry.phrases.set(phraseValue, {
                    type: "phrase",
                    value: phraseValue,
                    label: `↳↳ العبارة | ${record.itemLabel}`,
                    searchText: normalizeText([entry.criterionCode, entry.criterionText, entry.supportedSide, record.itemLabel].join(" ")),
                });
            }
        });
    });

    return Array.from(criteria.values())
        .sort(compareSelfStudyOptionNodes)
        .flatMap((criterionEntry) => {
            const criterionOptions = [{ value: criterionEntry.value, label: criterionEntry.label, searchText: criterionEntry.searchText }];
            const sideOptions = Array.from(criterionEntry.sides.values())
                .sort(compareSelfStudyOptionNodes)
                .flatMap((sideEntry) => {
                    const phraseOptions = Array.from(sideEntry.phrases.values())
                        .sort(compareSelfStudyOptionNodes)
                        .map((phraseEntry) => ({
                            value: phraseEntry.value,
                            label: phraseEntry.label,
                            searchText: phraseEntry.searchText,
                        }));
                    return [
                        { value: sideEntry.value, label: sideEntry.label, searchText: sideEntry.searchText },
                        ...phraseOptions,
                    ];
                });
            return [...criterionOptions, ...sideOptions];
        });
}

function compareSelfStudyOptionNodes(first, second) {
    if ((first.criterionCode || "") !== (second.criterionCode || "")) {
        return (first.criterionCode || "").localeCompare(second.criterionCode || "", "ar", { numeric: true });
    }
    const firstText = first.supportedSide || first.criterionText || first.label || "";
    const secondText = second.supportedSide || second.criterionText || second.label || "";
    return firstText.localeCompare(secondText, "ar");
}

function toggleTopicModePanels(generalPanel, selfStudyPanel, topicMode) {
    const showSelfStudy = topicMode === "selfstudy";
    generalPanel.classList.toggle("hidden", showSelfStudy);
    selfStudyPanel.classList.toggle("hidden", !showSelfStudy);
}

function renderGenderOptionsFor(selectRef, hintRef, targetState) {
    if (!AVAILABLE_GENDERS.length) {
        selectRef.innerHTML = `<option value="all">غير متاح في الملف الحالي</option>`;
        selectRef.disabled = true;
        hintRef.classList.remove("hidden");
        targetState.gender = "all";
        return;
    }

    hintRef.classList.add("hidden");
    selectRef.disabled = false;
    selectRef.innerHTML = [
        `<option value="all">كل الأنواع</option>`,
        ...AVAILABLE_GENDERS.map((gender) => `<option value="${gender}">${escapeHtml(getGenderLabel(gender))}</option>`),
    ].join("");
    selectRef.value = AVAILABLE_GENDERS.includes(targetState.gender) || targetState.gender === "all" ? targetState.gender : "all";
}

function renderCompareSlotOptions(slot) {
    const slotKey = slot.toUpperCase();
    const programRef = refs[`compareProgram${slotKey}`];
    const yearRef = refs[`compareYear${slotKey}`];
    const selection = state.compareSlots[slot];
    const programs = getProgramsWithData();

    programRef.innerHTML = [
        `<option value="">غير محدد</option>`,
        ...programs.map((program) => `<option value="${program.id}">${escapeHtml(formatProgramLabel(program))}</option>`),
    ].join("");
    programRef.value = programs.some((program) => program.id === selection.program) ? selection.program : "";

    if (!selection.program) {
        yearRef.innerHTML = `<option value="">غير محدد</option>`;
        yearRef.value = "";
        return;
    }

    const years = getAvailableYears(selection.program);
    if (!years.includes(selection.year)) {
        selection.year = "";
    }

    yearRef.innerHTML = [
        `<option value="">غير محدد</option>`,
        ...years.map((year) => `<option value="${year}">${year}هـ</option>`),
    ].join("");
    yearRef.value = selection.year;
}

function renderAnalysisProgramsChecklist() {
    const programs = getProgramsWithData();
    refs.analysisProgramsList.innerHTML = programs.map((program) => `
        <label class="program-option">
            <input type="checkbox" data-program="${program.id}" ${state.analysisFilters.programs.has(program.id) ? "checked" : ""}>
            <span>
                <strong>${escapeHtml(program.name)}</strong>
                <small>${escapeHtml(program.degree)}</small>
            </span>
        </label>
    `).join("");

    /* Update dropdown trigger label */
    if (refs.analysisProgramsTrigger) {
        const labelEl = refs.analysisProgramsTrigger.querySelector(".multi-select__trigger-label");
        if (labelEl) {
            const count = state.analysisFilters.programs.size;
            const total = programs.length;
            if (count === 0) labelEl.textContent = "لم يُحدد أي برنامج";
            else if (count === total) labelEl.textContent = "كل البرامج";
            else labelEl.textContent = `${toArabicNumber(count)} من ${toArabicNumber(total)} برنامج`;
        }
    }
}

function renderOverviewSection() {
    const surveyRows = aggregateSurveyRows(ITEM_RECORDS);
    const overallAverage = averageScore(surveyRows);
    const latestYearRows = surveyRows.filter((row) => row.year === DEFAULT_YEAR);
    const latestYearAverage = averageScore(latestYearRows);
    const topSurvey = getExtremeRow(surveyRows, "max");
    const lowSurvey = getExtremeRow(surveyRows, "min");

    refs.overviewMeta.textContent = `${toArabicNumber(surveyRows.length)} استطلاع رئيس عبر جميع البرامج والسنوات`;

    renderMetricCards(refs.overviewIndicators, [
        {
            label: "المتوسط العام",
            value: formatScore(overallAverage),
            note: "على مستوى الاستطلاعات الرئيسة",
            tone: toneForScore(overallAverage),
        },
        {
            label: `متوسط ${DEFAULT_YEAR}هـ`,
            value: formatScore(latestYearAverage),
            note: "أحدث سنة متاحة",
            tone: toneForScore(latestYearAverage),
        },
        {
            label: "أعلى استطلاع",
            value: topSurvey ? formatScore(topSurvey.average) : "—",
            note: topSurvey ? topSurvey.title : "لا توجد بيانات",
            tone: "good",
        },
        {
            label: "أولوية تحسين",
            value: lowSurvey ? formatScore(lowSurvey.average) : "—",
            note: lowSurvey ? lowSurvey.title : "لا توجد بيانات",
            tone: lowSurvey ? "danger" : "info",
        },
    ]);

    const programSeries = groupAverageRows(
        surveyRows,
        (row) => row.programId,
        (row) => getProgramById(row.programId).name
    ).slice(0, 10);

    const sectionSeries = SECTION_META.map((section) => ({
        label: section.shortLabel,
        average: averageScore(surveyRows.filter((row) => row.sectionId === section.id)),
    })).filter((item) => item.average != null);

    const trendSeries = ALL_AVAILABLE_YEARS
        .slice()
        .reverse()
        .map((year) => ({
            label: year,
            average: averageScore(surveyRows.filter((row) => row.year === year)),
        }))
        .filter((item) => item.average != null);

    const bandSeries = buildScoreBandSeries(surveyRows);
    const rankingRows = buildProgramRankingRows(surveyRows);

    drawChart("overviewProgramChart", refs.overviewProgramChart, {
        type: "bar",
        data: {
            labels: programSeries.map((item) => truncateLabel(item.label, 20)),
            datasets: [{
                label: "متوسط البرامج",
                data: programSeries.map((item) => item.average),
                backgroundColor: CHART_COLORS[0],
                borderRadius: 8,
            }],
        },
        options: baseChartOptions({
            indexAxis: "y",
            plugins: { legend: { display: false } },
            scales: buildScoreScales("x"),
        }),
    });

    drawChart("overviewSectionChart", refs.overviewSectionChart, {
        type: "radar",
        data: {
            labels: sectionSeries.map((item) => item.label),
            datasets: [{
                label: "متوسط المحاور",
                data: sectionSeries.map((item) => item.average),
                backgroundColor: "rgba(30, 95, 88, 0.16)",
                borderColor: CHART_COLORS[0],
                borderWidth: 2,
                pointBackgroundColor: CHART_COLORS[1],
            }],
        },
        options: baseChartOptions({
            scales: {
                r: {
                    min: 0,
                    max: 5,
                    ticks: { stepSize: 1, backdropColor: "transparent" },
                    grid: { color: "rgba(18, 57, 54, 0.14)" },
                    angleLines: { color: "rgba(18, 57, 54, 0.14)" },
                    pointLabels: { font: { family: "Tajawal", size: 12 } },
                },
            },
        }),
    });

    drawChart("overviewTrendChart", refs.overviewTrendChart, {
        type: "line",
        data: {
            labels: trendSeries.map((item) => `${item.label}هـ`),
            datasets: [{
                label: "المتوسط العام",
                data: trendSeries.map((item) => item.average),
                borderColor: CHART_COLORS[2],
                backgroundColor: "rgba(46, 118, 183, 0.16)",
                fill: true,
                tension: 0.28,
            }],
        },
        options: baseChartOptions({
            plugins: { legend: { display: false } },
            scales: buildScoreScales("y"),
        }),
    });

    drawChart("overviewBandChart", refs.overviewBandChart, {
        type: "doughnut",
        data: {
            labels: bandSeries.map((item) => item.label),
            datasets: [{
                data: bandSeries.map((item) => item.value),
                backgroundColor: bandSeries.map((item) => item.color),
                borderWidth: 0,
            }],
        },
        options: baseChartOptions({
            cutout: "58%",
        }),
    });

    refs.overviewTableMeta.textContent = `${toArabicNumber(rankingRows.length)} برنامج`;
    refs.overviewTableBody.innerHTML = rankingRows.map((item, index) => `
        <tr>
            <td>${toArabicNumber(index + 1)}</td>
            <td>
                <span class="cell-title">${escapeHtml(item.program.name)}</span>
                <span class="cell-subtitle">${escapeHtml(item.program.dept)}</span>
            </td>
            <td>${escapeHtml(item.program.degree)}</td>
            <td>${toArabicNumber(item.count)}</td>
            <td>${renderScorePill(item.average)}</td>
            <td>${item.bestSection ? escapeHtml(item.bestSection) : "—"}</td>
        </tr>
    `).join("");
    refs.overviewEmpty.classList.toggle("hidden", rankingRows.length > 0);
}

function renderExploreSection() {
    const records = getItemRecordsForExploreFilters();
    const surveyRows = aggregateSurveyRows(records);
    const programRows = buildExploreItemRows(records);
    const topItemRows = [...programRows]
        .filter((row) => row.average != null)
        .sort((first, second) => {
            if (second.average !== first.average) return second.average - first.average;
            if (second.respondentCount !== first.respondentCount) return second.respondentCount - first.respondentCount;
            return first.title.localeCompare(second.title, "ar");
        })
        .slice(0, 10);
    const sectionAverage = surveyRows.length ? averageScore(surveyRows) : null;

    refs.exploreMeta.textContent = `${toArabicNumber(surveyRows.length)} استطلاع · ${toArabicNumber(programRows.length)} عبارة`;

    renderMetricCards(refs.exploreIndicators, [
        {
            label: "المتوسط العام",
            value: formatScore(sectionAverage),
            note: records.length ? formatResponseCount(records.reduce((sum, r) => sum + r.responses, 0)) : "لا توجد بيانات",
            tone: toneForScore(sectionAverage),
        },
        {
            label: "أفضل استطلاع",
            value: surveyRows.length ? formatScore(getExtremeRow(surveyRows, "max").average) : "—",
            note: surveyRows.length ? truncateLabel(getExtremeRow(surveyRows, "max").title, 30) : "لا توجد بيانات",
            tone: surveyRows.length ? "good" : "info",
        },
        {
            label: "أدنى استطلاع",
            value: surveyRows.length ? formatScore(getExtremeRow(surveyRows, "min").average) : "—",
            note: surveyRows.length ? truncateLabel(getExtremeRow(surveyRows, "min").title, 30) : "لا توجد بيانات",
            tone: surveyRows.length ? (getExtremeRow(surveyRows, "min").average < 3.25 ? "danger" : "warning") : "info",
        },
        {
            label: "عدد الموضوعات",
            value: toArabicNumber(collectUnique(records.map(r => r.topicLabel), v => normalizeText(v)).length),
            note: "الموضوعات المتاحة بالفلاتر الحالية",
            tone: "info",
        },
    ]);

    const sectionDatasets = SECTION_META.map((section) => ({
        label: section.shortLabel,
        data: [averageScore(surveyRows.filter((row) => row.sectionId === section.id))],
        backgroundColor: CHART_COLORS[0],
        borderRadius: 8,
    })).filter(d => d.data[0] != null);

    drawChart("exploreSectionChart", refs.exploreSectionChart, {
        type: "bar",
        data: {
            labels: sectionDatasets.map(d => d.label),
            datasets: [{
                label: "متوسط المحاور",
                data: sectionDatasets.map(d => d.data[0]),
                backgroundColor: CHART_COLORS.slice(0, sectionDatasets.length),
                borderRadius: 10,
            }],
        },
        options: baseChartOptions({
            plugins: { legend: { display: false } },
            scales: buildScoreScales("y"),
        }),
    });

    drawChart("exploreItemChart", refs.exploreItemChart, {
        type: "bar",
        data: {
            labels: topItemRows.map((row) => truncateLabel(row.title, 34)),
            datasets: [{
                label: "متوسط البند",
                data: topItemRows.map((row) => row.average),
                backgroundColor: "rgba(30, 95, 88, 0.82)",
                borderRadius: 10,
            }],
        },
        options: baseChartOptions({
            indexAxis: "y",
            plugins: { legend: { display: false } },
            scales: buildScoreScales("x"),
        }),
    });

    if (programRows.length) {
        refs.exploreTableMeta.textContent = `${toArabicNumber(programRows.length)} بند`;
        refs.exploreTableBody.innerHTML = programRows.map((row) => `
            <tr>
                <td>${escapeHtml(row.program.name)}</td>
                <td>${escapeHtml(row.year)}</td>
                <td>${escapeHtml(row.sectionLabel)}</td>
                <td>${escapeHtml(row.surveyTitle)}</td>
                <td>
                    <span class="cell-title">${escapeHtml(row.title)}</span>
                    ${row.subtitle ? `<span class="cell-subtitle">${escapeHtml(row.subtitle)}</span>` : ""}
                </td>
                <td>${escapeHtml(row.type)}</td>
                <td>${escapeHtml(row.genderLabel)}</td>
                <td>${toArabicNumber(row.respondentCount)}</td>
                <td>${renderScorePill(row.average)}</td>
                <td>${renderStatusPill(row.average)}</td>
            </tr>
        `).join("");
        refs.exploreEmpty.classList.add("hidden");
    } else {
        refs.exploreTableMeta.textContent = "0 عناصر";
        refs.exploreTableBody.innerHTML = "";
        refs.exploreEmpty.classList.remove("hidden");
    }
}

function getItemRecordsForExploreFilters() {
    return ITEM_RECORDS.filter((record) => matchesSingleFilters(record, state.exploreFilters));
}

function buildExploreItemRows(records) {
    const genderLabel = state.exploreFilters.gender === "all"
        ? "كل الاستجابات"
        : getGenderFilterLabel(state.exploreFilters.gender);

    return aggregateItemRows(records).map((row) => ({
        uid: row.uid,
        program: getProgramById(row.programId),
        year: row.year,
        sectionId: row.sectionId,
        sectionLabel: row.sectionLabel,
        surveyTitle: row.surveyTitle,
        title: row.title,
        subtitle: row.parentTitle || "",
        type: row.rowKind,
        genderLabel,
        respondentCount: row.respondentCount,
        average: row.average,
    }));
}

// renderProgramSection removed — replaced by renderExploreSection

function renderCompareSection() {
    const selectedSlots = getActiveCompareSlots();
    const slotGroups = selectedSlots.map((slot) => {
        const records = getRecordsForCompareSlot(slot);
        return {
            slot,
            surveyRows: aggregateSurveyRows(records),
        };
    });
    const comparisonRows = slotGroups.length >= 2 ? buildComparisonTableRows(slotGroups.map((group) => group.surveyRows)) : [];

    if (refs.insightsMeta) {
        refs.insightsMeta.textContent = slotGroups.length
            ? `${toArabicNumber(slotGroups.length)} خانات محددة`
            : "اختر برنامجًا وسنة في أي خانة لتبدأ المقارنة";
    }

    renderMetricCards(refs.compareIndicators, ["a", "b", "c"].map((slotKey, index) => {
        const slot = state.compareSlots[slotKey];
        const group = slotGroups.find((item) => item.slot === slot);
        const active = Boolean(slot.program && slot.year && group);
        const average = active ? averageScore(group.surveyRows) : null;
        const totalResponses = active ? group.surveyRows.reduce((sum, row) => sum + row.respondentCount, 0) : 0;
        return {
            label: `الخانة ${toArabicNumber(index + 1)}`,
            value: active ? formatScore(average) : "—",
            note: active ? `${formatCompareSlotLabel(slot)} · ${formatResponseCount(totalResponses)}` : "غير محدد",
            tone: active ? toneForScore(average) : "info",
        };
    }));

    const sectionDatasets = slotGroups.map((group, index) => ({
        label: formatCompareSlotLabel(group.slot),
        data: SECTION_META.map((section) => averageScore(group.surveyRows.filter((row) => row.sectionId === section.id))),
        backgroundColor: CHART_COLORS[index],
        borderColor: CHART_COLORS[index],
        borderRadius: 8,
    }));

    drawChart("compareSectionChart", refs.compareSectionChart, {
        type: "bar",
        data: {
            labels: SECTION_META.map((section) => section.shortLabel),
            datasets: sectionDatasets,
        },
        options: baseChartOptions({
            plugins: { legend: { position: "bottom" } },
            scales: buildScoreScales("y"),
        }),
    });

    drawChart("compareAverageChart", refs.compareAverageChart, {
        type: "bar",
        data: {
            labels: slotGroups.map((group) => truncateLabel(formatCompareSlotLabel(group.slot), 22)),
            datasets: [{
                label: "المتوسط العام",
                data: slotGroups.map((group) => averageScore(group.surveyRows)),
                backgroundColor: CHART_COLORS.slice(0, slotGroups.length),
                borderRadius: 10,
            }],
        },
        options: baseChartOptions({
            plugins: { legend: { display: false } },
            scales: buildScoreScales("y"),
        }),
    });

    if (slotGroups.length >= 2) {
        refs.compareTableHead.innerHTML = `
            <tr>
                <th>المحور</th>
                <th>الاستطلاع</th>
                ${slotGroups.map((group) => `<th>${escapeHtml(formatCompareSlotLabel(group.slot))}</th>`).join("")}
            </tr>
        `;
        refs.compareTableBody.innerHTML = comparisonRows.map((row) => `
            <tr>
                <td>${escapeHtml(row.meta.sectionLabel)}</td>
                <td>
                    <span class="cell-title">${escapeHtml(row.meta.title)}</span>
                    <span class="cell-subtitle">${toArabicNumber(row.meta.itemCount)} عبارة</span>
                </td>
                ${row.values.map((value) => renderComparisonCell(value)).join("")}
            </tr>
        `).join("");
        refs.compareTableMeta.textContent = `${toArabicNumber(comparisonRows.length)} استطلاع مقارنة`;
    } else {
        refs.compareTableHead.innerHTML = "";
        refs.compareTableBody.innerHTML = "";
        refs.compareTableMeta.textContent = "0 عناصر";
    }

    refs.compareEmpty.classList.toggle("hidden", !(slotGroups.length < 2 || !comparisonRows.length));
}

function renderAnalysisSection() {
    const records = getRecordsForAnalysis();
    const surveyRows = aggregateSurveyRows(records);
    const programRows = buildAnalysisProgramRows(surveyRows);
    const sectionGroups = SECTION_META.map((section) => ({
        ...section,
        average: averageScore(surveyRows.filter((row) => row.sectionId === section.id)),
    })).filter((item) => item.average != null);
    const topProgram = [...programRows].sort((first, second) => second.average - first.average)[0];
    const lowProgram = [...programRows].sort((first, second) => first.average - second.average)[0];
    const topSurvey = getExtremeRow(surveyRows, "max");
    const lowSurvey = getExtremeRow(surveyRows, "min");
    const strengths = [...surveyRows].sort((first, second) => second.average - first.average).slice(0, 5);
    const watchlist = [...surveyRows].sort((first, second) => first.average - second.average).slice(0, 5);
    const analysisTable = buildAnalysisTable(programRows);

    if (refs.insightsMeta) {
        refs.insightsMeta.textContent = `${buildAnalysisScopeLabel()} · ${toArabicNumber(surveyRows.length)} استطلاع`;
    }

    renderMetricCards(refs.analysisIndicators, [
        {
            label: "أفضل برنامج",
            value: topProgram ? formatScore(topProgram.average) : "—",
            note: topProgram ? topProgram.program.name : "لا توجد بيانات",
            tone: "good",
        },
        {
            label: "أولوية تحسين",
            value: lowProgram ? formatScore(lowProgram.average) : "—",
            note: lowProgram ? lowProgram.program.name : "لا توجد بيانات",
            tone: lowProgram ? "warning" : "info",
        },
        {
            label: "أعلى استطلاع",
            value: topSurvey ? formatScore(topSurvey.average) : "—",
            note: topSurvey ? topSurvey.title : "لا توجد بيانات",
            tone: "good",
        },
        {
            label: "أدنى استطلاع",
            value: lowSurvey ? formatScore(lowSurvey.average) : "—",
            note: lowSurvey ? lowSurvey.title : "لا توجد بيانات",
            tone: lowSurvey ? "danger" : "info",
        },
    ]);

    renderInsightList(refs.analysisStrengths, strengths, "لا توجد عناصر قوة في النطاق الحالي.");
    renderInsightList(refs.analysisWatchlist, watchlist, "لا توجد عناصر بحاجة متابعة في النطاق الحالي.");

    drawChart("analysisSectionChart", refs.analysisSectionChart, {
        type: "bar",
        data: {
            labels: sectionGroups.map((item) => item.shortLabel),
            datasets: [{
                label: "متوسط المحاور",
                data: sectionGroups.map((item) => item.average),
                backgroundColor: CHART_COLORS.slice(0, sectionGroups.length),
                borderRadius: 10,
            }],
        },
        options: baseChartOptions({
            plugins: { legend: { display: false } },
            scales: buildScoreScales("y"),
        }),
    });

    refs.analysisTableTitle.textContent = analysisTable.title;
    refs.analysisTableMeta.textContent = analysisTable.meta;
    refs.analysisTableHead.innerHTML = analysisTable.head;
    refs.analysisTableBody.innerHTML = analysisTable.body;
    refs.analysisEmpty.classList.toggle("hidden", analysisTable.hasRows);
}

function renderClosureSection() {
    renderClosureControls();
    const payload = getClosureComparisonPayload();
    state.latestClosurePayload = payload;

    if (refs.insightsMeta) {
        refs.insightsMeta.textContent = payload.metaText;
    }

    renderMetricCards(refs.closureIndicators, payload.cards);
    renderClosureHighlights(refs.closureHighlights, payload.qualifyingRows, payload.emptyMessage);
    if (refs.closureNarrative) {
        refs.closureNarrative.innerHTML = buildClosureNarrativeHtml(payload);
    }

    drawClosureChart(payload.qualifyingRows);
    renderClosureDetailTable(payload);
    renderClosureReportTable(payload);
    syncClosureDisplayCards();
}

function syncClosureDisplayCards() {
    const isReport = state.closureFilters.displayMode === "report";
    if (refs.closureTableCard) refs.closureTableCard.classList.toggle("hidden", isReport);
    if (refs.closureReportCard) refs.closureReportCard.classList.toggle("hidden", !isReport);
}

function renderClosureDetailTable(payload) {
    if (payload.ready && payload.qualifyingRows.length) {
        if (refs.closureTableCard) refs.closureTableCard.classList.remove("hidden");
        if (refs.closureTableHead) {
            refs.closureTableHead.innerHTML = `
                <tr>
                    <th>المحور</th>
                    <th>التحسن</th>
                    <th>نوع المطابقة</th>
                    <th>تفاصيل الربط</th>
                    <th>${escapeHtml(payload.fromYearLabel)}</th>
                    <th>${escapeHtml(payload.toYearLabel)}</th>
                </tr>
            `;
        }
        if (refs.closureTableBody) {
            refs.closureTableBody.innerHTML = payload.qualifyingRows.map((row) => `
                <tr>
                    <td>
                        <span class="cell-title">${escapeHtml(row.sectionLabel)}</span>
                        <span class="cell-subtitle">${escapeHtml(row.title)}</span>
                    </td>
                    <td>
                        <span class="delta-pill positive">${escapeHtml(formatClosureImprovement(row.deltaHundred))}</span>
                        <span class="cell-subtitle">${escapeHtml(formatPercent(row.fromPercent))} → ${escapeHtml(formatPercent(row.toPercent))}</span>
                    </td>
                    <td>
                        <span class="cell-title">${escapeHtml(getClosureMatchModeLabel(row.matchMode))}</span>
                        <span class="cell-subtitle">${escapeHtml(formatClosureMatchScore(row.matchMode, row.matchScore))}</span>
                    </td>
                    <td>${buildClosureLinkingCellHtml(row)}</td>
                    <td>${buildClosureYearTableCellHtml(row.fromRow, payload.fromYearLabel, row.fromAverage, row.fromPercent, row.fromResponses)}</td>
                    <td>${buildClosureYearTableCellHtml(row.toRow, payload.toYearLabel, row.toAverage, row.toPercent, row.toResponses)}</td>
                </tr>
            `).join("");
        }
        if (refs.closureTableMeta) {
            refs.closureTableMeta.textContent = `${toArabicNumber(payload.qualifyingRows.length)} مؤشر تحسن مؤثر مع تفاصيل الربط والبنود`;
        }
        if (refs.closureEmpty) refs.closureEmpty.classList.add("hidden");
    } else {
        if (refs.closureTableHead) refs.closureTableHead.innerHTML = "";
        if (refs.closureTableBody) refs.closureTableBody.innerHTML = "";
        if (refs.closureTableMeta) refs.closureTableMeta.textContent = payload.ready ? "0 مؤشرات مؤهلة" : "اختر نطاق المقارنة";
        if (refs.closureEmpty) {
            refs.closureEmpty.textContent = payload.emptyMessage;
            refs.closureEmpty.classList.remove("hidden");
        }
    }
}

function renderClosureReportTable(payload) {
    if (!refs.closureReportHead || !refs.closureReportBody) return;

    if (payload.ready && payload.qualifyingRows.length) {
        refs.closureReportHead.innerHTML = `
            <tr>
                <th>البند أو الموضوع</th>
                <th>${escapeHtml(payload.fromYearLabel)}</th>
                <th>${escapeHtml(payload.toYearLabel)}</th>
                <th>التحسن</th>
            </tr>
        `;
        refs.closureReportBody.innerHTML = payload.qualifyingRows.map((row) => `
            <tr>
                <td>${buildClosureReportFocusCellHtml(row)}</td>
                <td>${buildClosureReportYearCellHtml(row.fromRow, row.fromAverage, row.fromPercent, row.fromResponses)}</td>
                <td>${buildClosureReportYearCellHtml(row.toRow, row.toAverage, row.toPercent, row.toResponses)}</td>
                <td>${buildClosureReportImprovementCellHtml(row)}</td>
            </tr>
        `).join("");
        if (refs.closureReportMeta) {
            refs.closureReportMeta.textContent = `${toArabicNumber(payload.qualifyingRows.length)} صفًا جاهزًا للتوثيق والطباعة`;
        }
        if (refs.closureReportSummary) {
            refs.closureReportSummary.innerHTML = buildClosureReportSummaryHtml(payload);
        }
        if (refs.closureReportEmpty) refs.closureReportEmpty.classList.add("hidden");
        if (refs.closureReportPrintArea) refs.closureReportPrintArea.classList.remove("hidden");
    } else {
        refs.closureReportHead.innerHTML = "";
        refs.closureReportBody.innerHTML = "";
        if (refs.closureReportMeta) refs.closureReportMeta.textContent = payload.ready ? "0 صفوف تقرير" : "اختر نطاق المقارنة";
        if (refs.closureReportSummary) refs.closureReportSummary.innerHTML = "";
        if (refs.closureReportEmpty) {
            refs.closureReportEmpty.textContent = payload.emptyMessage;
            refs.closureReportEmpty.classList.remove("hidden");
        }
        if (refs.closureReportPrintArea) refs.closureReportPrintArea.classList.add("hidden");
    }
}

function getClosureComparisonPayload() {
    const matchMode = ["item", "expanded", "veryExpanded"].includes(state.closureFilters.level)
        ? state.closureFilters.level
        : "item";
    const program = state.closureFilters.program ? getProgramById(state.closureFilters.program) : null;
    const fromYearLabel = state.closureFilters.fromYear ? `${state.closureFilters.fromYear}هـ` : "السنة الأولى";
    const toYearLabel = state.closureFilters.toYear ? `${state.closureFilters.toYear}هـ` : "السنة الثانية";
    const threshold = getClosureThreshold();

    if (!program || !state.closureFilters.fromYear || !state.closureFilters.toYear) {
        return buildClosureEmptyPayload({
            metaText: "اختر برنامجًا وسنتين للبحث عن دلائل إغلاق دائرة الجودة.",
            emptyMessage: "اختر برنامجًا وسنتين لعرض دلائل إغلاق دائرة الجودة.",
            fromYearLabel,
            toYearLabel,
        });
    }

    const fromRecords = ITEM_RECORDS.filter((record) =>
        record.programId === state.closureFilters.program &&
        record.year === state.closureFilters.fromYear &&
        record.stakeholder === "students" &&
        (state.closureFilters.gender === "all" || record.gender === state.closureFilters.gender) &&
        matchesTopicFilter(record, state.closureFilters)
    );
    const toRecords = ITEM_RECORDS.filter((record) =>
        record.programId === state.closureFilters.program &&
        record.year === state.closureFilters.toYear &&
        record.stakeholder === "students" &&
        (state.closureFilters.gender === "all" || record.gender === state.closureFilters.gender) &&
        matchesTopicFilter(record, state.closureFilters)
    );

    const fromRows = aggregateItemRows(fromRecords);
    const toRows = aggregateItemRows(toRecords);
    const comparableRows = buildClosureComparableRows(fromRows, toRows, matchMode)
        .map((pair) => buildClosureRow(pair.fromRow, pair.toRow, matchMode, pair.matchMeta))
        .sort((first, second) => second.deltaHundred - first.deltaHundred || second.toAverage - first.toAverage);

    const qualifyingRows = comparableRows.filter((row) => row.deltaHundred >= threshold);
    const positiveBelowThresholdCount = comparableRows.filter((row) => row.deltaHundred > 0 && row.deltaHundred < threshold).length;
    const topRow = qualifyingRows[0] || null;
    const averageImprovement = qualifyingRows.length
        ? roundNumber(qualifyingRows.reduce((sum, row) => sum + row.deltaHundred, 0) / qualifyingRows.length)
        : null;

    let emptyMessage = "لا توجد مؤشرات قابلة للمقارنة بين السنتين ضمن النطاق الحالي.";
    if (comparableRows.length && !qualifyingRows.length) {
        emptyMessage = `توجد ${toArabicNumber(comparableRows.length)} مؤشرات قابلة للمقارنة، لكن لم يتجاوز أيٌّ منها حد التحسن الأدنى (${formatClosureThreshold(threshold)}).`;
    } else if (!comparableRows.length && matchMode === "veryExpanded") {
        emptyMessage = "لم تظهر مطابقة متوسعة جدًا كافية بين السنتين ضمن هذا النطاق. يمكنك توسيع الموضوع أو تخفيف الحد الأدنى للتحسن.";
    } else if (!comparableRows.length && matchMode === "expanded") {
        emptyMessage = "لم تظهر مطابقة متوسعة كافية بين السنتين ضمن هذا النطاق. يمكنك تغيير الموضوع أو تخفيف الحد الأدنى للتحسن.";
    } else if (!comparableRows.length) {
        emptyMessage = "لم تظهر بنود قابلة للمطابقة المضبوطة بين السنتين ضمن هذا النطاق. جرّب التحويل إلى المطابقة المتوسعة لالتقاط المزيد من التحسن.";
    }

    return {
        ready: true,
        program,
        level: matchMode,
        threshold,
        fromYear: state.closureFilters.fromYear,
        toYear: state.closureFilters.toYear,
        fromYearLabel,
        toYearLabel,
        comparableRows,
        qualifyingRows,
        positiveBelowThresholdCount,
        averageImprovement,
        topRow,
        emptyMessage,
        metaText: `${formatProgramLabel(program)} · ${fromYearLabel} ← ${toYearLabel} · ${getGenderFilterLabel(state.closureFilters.gender)} · ${getClosureLevelLabel(matchMode)} · ${getTopicSelectionLabel(state.closureFilters)}`,
        cards: [
            {
                label: "المؤشرات القابلة للمقارنة",
                value: toArabicNumber(comparableRows.length),
                note: "المؤشرات التي أمكن ربطها بين السنتين",
                tone: comparableRows.length ? "info" : "warning",
            },
            {
                label: "المؤشرات المؤهلة",
                value: toArabicNumber(qualifyingRows.length),
                note: `بعد استبعاد أي تحسن أقل من ${formatClosureThreshold(threshold)}`,
                tone: qualifyingRows.length ? "good" : "warning",
            },
            {
                label: "أكبر تحسن",
                value: topRow ? formatClosureImprovement(topRow.deltaHundred) : "—",
                note: topRow ? truncateLabel(topRow.title, 42) : "لا يوجد تحسن مؤثر حتى الآن",
                tone: topRow ? "good" : "info",
            },
            {
                label: "متوسط التحسن",
                value: averageImprovement != null ? formatClosureImprovement(averageImprovement) : "—",
                note: qualifyingRows.length
                    ? `تم تجاهل ${toArabicNumber(positiveBelowThresholdCount)} تحسنات طفيفة غير مؤثرة`
                    : "لا توجد نتائج مؤهلة للحساب",
                tone: averageImprovement != null ? "info" : "warning",
            },
        ],
    };
}

function buildClosureEmptyPayload({ metaText, emptyMessage, fromYearLabel, toYearLabel }) {
    return {
        ready: false,
        program: null,
        level: state.closureFilters.level,
        threshold: getClosureThreshold(),
        fromYear: state.closureFilters.fromYear,
        toYear: state.closureFilters.toYear,
        fromYearLabel,
        toYearLabel,
        comparableRows: [],
        qualifyingRows: [],
        positiveBelowThresholdCount: 0,
        averageImprovement: null,
        topRow: null,
        emptyMessage,
        metaText,
        cards: [
            { label: "المؤشرات القابلة للمقارنة", value: "—", note: "حدّد البرنامج والفترة أولًا", tone: "warning" },
            { label: "المؤشرات المؤهلة", value: "—", note: "لن يظهر إلا التحسن المؤثر", tone: "warning" },
            { label: "أكبر تحسن", value: "—", note: "سيظهر بعد اكتمال النطاق", tone: "info" },
            { label: "متوسط التحسن", value: "—", note: "يُحسب بعد ظهور نتائج مؤهلة", tone: "info" },
        ],
    };
}

function aggregateRowsByLevel(records, level) {
    if (level === "survey") return aggregateSurveyRows(records);
    if (level === "item") return aggregateItemRows(records);
    return aggregateTopicRows(records);
}

function buildClosureComparableRows(fromRows, toRows, level) {
    if (level === "survey") {
        const fromMap = new Map(fromRows.map((row) => [getClosureComparisonKey(row, level), row]));
        const toMap = new Map(toRows.map((row) => [getClosureComparisonKey(row, level), row]));
        return Array.from(fromMap.entries())
            .filter(([key]) => toMap.has(key))
            .map(([key, fromRow]) => ({
                fromRow,
                toRow: toMap.get(key),
                matchMeta: { mode: "exact", score: 1 },
            }));
    }

    const fromGroups = groupClosureRows(fromRows, level);
    const toGroups = groupClosureRows(toRows, level);
    const groupKeys = collectUnique([...fromGroups.keys(), ...toGroups.keys()], (value) => value);
    return groupKeys.flatMap((groupKey) => pairClosureRowsWithinGroup(
        fromGroups.get(groupKey) || [],
        toGroups.get(groupKey) || [],
        level
    ));
}

function groupClosureRows(rows, level) {
    const groups = new Map();
    rows.forEach((row) => {
        const groupKey = level === "expanded" || level === "veryExpanded"
            ? row.sectionId
            : [row.sectionId, normalizeText(row.surveyTitle)].join("||");
        if (!groups.has(groupKey)) {
            groups.set(groupKey, []);
        }
        groups.get(groupKey).push(row);
    });
    return groups;
}

function pairClosureRowsWithinGroup(fromRows, toRows, level) {
    if (!fromRows.length || !toRows.length) return [];

    const matchedFrom = new Set();
    const matchedTo = new Set();
    const pairs = [];
    const exactMap = new Map();

    toRows.forEach((row, index) => {
        exactMap.set(getClosureExactKey(row, level), { row, index });
    });

    fromRows.forEach((row, index) => {
        const exactMatch = exactMap.get(getClosureExactKey(row, level));
        if (!exactMatch || matchedTo.has(exactMatch.index)) return;
        matchedFrom.add(index);
        matchedTo.add(exactMatch.index);
        pairs.push({
            fromRow: row,
            toRow: exactMatch.row,
            matchMeta: { mode: "exact", score: 1 },
        });
    });

    const candidates = [];
    fromRows.forEach((fromRow, fromIndex) => {
        if (matchedFrom.has(fromIndex)) return;
        toRows.forEach((toRow, toIndex) => {
            if (matchedTo.has(toIndex)) return;
            const score = getClosureFlexibleMatchScore(fromRow, toRow, level);
            if (score <= 0) return;
            candidates.push({
                fromIndex,
                toIndex,
                fromRow,
                toRow,
                score,
            });
        });
    });

    candidates
        .sort((first, second) => second.score - first.score)
        .forEach((candidate) => {
            if (matchedFrom.has(candidate.fromIndex) || matchedTo.has(candidate.toIndex)) return;
            matchedFrom.add(candidate.fromIndex);
            matchedTo.add(candidate.toIndex);
            pairs.push({
                fromRow: candidate.fromRow,
                toRow: candidate.toRow,
                matchMeta: {
                    mode: level === "veryExpanded" ? "veryExpanded" : level === "expanded" ? "expanded" : "flexible",
                    score: candidate.score,
                },
            });
        });

    return pairs;
}

function getClosureExactKey(row, level) {
    if (level === "topic") {
        return normalizeText(row.title);
    }

    if (level === "item" || level === "expanded" || level === "veryExpanded") {
        const orderKey = row.itemOrder < 900000 ? `#${row.itemOrder}` : "";
        return [normalizeText(row.surveyTitle), normalizeText(row.parentTitle), orderKey, normalizeText(row.title)].join("||");
    }

    return getClosureComparisonKey(row, level);
}

function getClosureComparisonKey(row, level) {
    if (level === "survey") {
        return [row.sectionId, normalizeText(row.title)].join("||");
    }

    if (level === "item" || level === "expanded" || level === "veryExpanded") {
        const itemKey = row.itemOrder < 900000
            ? `item:${row.itemOrder}`
            : `label:${normalizeText(row.title)}`;
        return [
            row.sectionId,
            normalizeText(row.surveyTitle),
            normalizeText(row.parentTitle),
            itemKey,
        ].join("||");
    }

    return [
        row.sectionId,
        normalizeText(row.surveyTitle),
        normalizeText(row.title),
    ].join("||");
}

function getClosureFlexibleMatchScore(fromRow, toRow, level) {
    const surveyScore = buildFlexibleTextScore(fromRow.surveyTitle, toRow.surveyTitle);
    if (level === "veryExpanded") {
        const topicScore = buildFlexibleTextScore(fromRow.parentTitle, toRow.parentTitle);
        const titleScore = buildFlexibleTextScore(fromRow.title, toRow.title);
        const exactOrder = fromRow.itemOrder < 900000 && toRow.itemOrder < 900000 && fromRow.itemOrder === toRow.itemOrder;
        const nearOrder = fromRow.itemOrder < 900000 && toRow.itemOrder < 900000 && Math.abs(fromRow.itemOrder - toRow.itemOrder) <= 4;
        const broadOrder = fromRow.itemOrder < 900000 && toRow.itemOrder < 900000 && Math.abs(fromRow.itemOrder - toRow.itemOrder) <= 8;

        if (Math.max(surveyScore, topicScore, titleScore) < 0.1 && !nearOrder) return 0;
        if (surveyScore < 0.3 && topicScore < 0.12 && titleScore < 0.08 && !broadOrder) return 0;

        const orderBonus = exactOrder ? 0.24 : nearOrder ? 0.15 : broadOrder ? 0.07 : 0;
        const structureBonus = surveyScore >= 0.8 ? 0.08 : surveyScore >= 0.55 ? 0.04 : 0;
        const semanticBonus = topicScore >= 0.25 || titleScore >= 0.22 ? 0.05 : 0;

        return Math.min(
            1,
            roundNumber((surveyScore * 0.22) + (topicScore * 0.24) + (titleScore * 0.1) + orderBonus + structureBonus + semanticBonus)
        );
    }

    if (level === "expanded") {
        const topicScore = buildFlexibleTextScore(fromRow.parentTitle, toRow.parentTitle);
        const titleScore = buildFlexibleTextScore(fromRow.title, toRow.title);
        const exactOrder = fromRow.itemOrder < 900000 && toRow.itemOrder < 900000 && fromRow.itemOrder === toRow.itemOrder;
        const nearOrder = fromRow.itemOrder < 900000 && toRow.itemOrder < 900000 && Math.abs(fromRow.itemOrder - toRow.itemOrder) <= 2;
        const farButStructured = surveyScore >= 0.75 && (topicScore >= 0.18 || nearOrder || exactOrder);

        if (surveyScore < 0.5 && topicScore < 0.5 && !exactOrder && !nearOrder) return 0;
        if (titleScore < 0.18 && topicScore < 0.18 && !farButStructured) return 0;

        const orderBonus = exactOrder ? 0.22 : nearOrder ? 0.12 : 0;
        const structureBonus = surveyScore >= 0.9 ? 0.08 : surveyScore >= 0.75 ? 0.04 : 0;
        return Math.min(
            1,
            roundNumber((surveyScore * 0.28) + (topicScore * 0.34) + (titleScore * 0.18) + orderBonus + structureBonus)
        );
    }

    if (surveyScore < 0.88) return 0;

    if (level === "topic") {
        const titleScore = buildFlexibleTextScore(fromRow.title, toRow.title);
        return titleScore >= 0.45 ? titleScore : 0;
    }

    if (level === "item") {
        const topicScore = buildFlexibleTextScore(fromRow.parentTitle, toRow.parentTitle);
        const titleScore = buildFlexibleTextScore(fromRow.title, toRow.title);
        const exactOrder = fromRow.itemOrder < 900000 && toRow.itemOrder < 900000 && fromRow.itemOrder === toRow.itemOrder;
        const nearOrder = fromRow.itemOrder < 900000 && toRow.itemOrder < 900000 && Math.abs(fromRow.itemOrder - toRow.itemOrder) <= 1;

        if (topicScore < 0.2 && titleScore < 0.7 && !exactOrder) return 0;
        if (titleScore < 0.34 && topicScore < 0.7 && !exactOrder) return 0;

        const orderBonus = exactOrder ? 0.18 : nearOrder ? 0.06 : 0;
        return Math.min(1, roundNumber((titleScore * 0.65) + (topicScore * 0.35) + orderBonus));
    }

    return 0;
}

function buildFlexibleTextScore(firstValue, secondValue) {
    const first = normalizeForSearch(firstValue);
    const second = normalizeForSearch(secondValue);
    if (!first || !second) return 0;
    if (first === second) return 1;
    if (first.includes(second) || second.includes(first)) return 0.92;

    const preciseForward = searchMatchesPrecisely(first, second);
    const preciseBackward = searchMatchesPrecisely(second, first);
    if (preciseForward && preciseBackward) return 0.9;

    const softForward = searchMatches(first, second);
    const softBackward = searchMatches(second, first);

    const firstTokens = collectUnique(first.split(" ").filter(Boolean), (value) => value);
    const secondTokens = collectUnique(second.split(" ").filter(Boolean), (value) => value);
    const secondSet = new Set(secondTokens);
    const sharedCount = firstTokens.filter((token) => secondSet.has(token)).length;
    const overlapScore = Math.max(
        firstTokens.length ? sharedCount / firstTokens.length : 0,
        secondTokens.length ? sharedCount / secondTokens.length : 0,
    );

    if (softForward && softBackward) return Math.max(0.78, overlapScore);
    if (softForward || softBackward) return Math.max(0.64, overlapScore);
    return overlapScore;
}

function buildClosureRow(fromRow, toRow, level, matchMeta = { mode: "exact", score: 1 }) {
    const fromAverage = Number(fromRow.average || 0);
    const toAverage = Number(toRow.average || 0);
    const deltaScore = roundNumber(toAverage - fromAverage);
    const deltaHundred = roundNumber(deltaScore * 100);
    const deltaPercent = roundNumber(deltaScore * 20);
    const fromPercent = roundNumber(fromAverage * 20);
    const toPercent = roundNumber(toAverage * 20);
    const coverageCount = Math.max(fromRow.itemCount || 0, toRow.itemCount || 0);

    let subtitle = "";
    if (level === "survey") {
        subtitle = `${toArabicNumber(coverageCount)} عبارة ضمن الاستطلاع`;
    } else if (level === "topic") {
        subtitle = `${toRow.surveyTitle || fromRow.surveyTitle} · ${toArabicNumber(coverageCount)} عبارة في الموضوع`;
    } else {
        subtitle = `${toRow.parentTitle || fromRow.parentTitle} · ${toRow.surveyTitle || fromRow.surveyTitle}`;
    }

    return {
        sectionId: toRow.sectionId || fromRow.sectionId,
        sectionLabel: toRow.sectionLabel || fromRow.sectionLabel,
        title: toRow.title || fromRow.title,
        subtitle: matchMeta.mode !== "exact"
            ? `${subtitle} · ${getClosureMatchModeLabel(matchMeta.mode)}`
            : subtitle,
        fromAverage,
        toAverage,
        fromPercent,
        toPercent,
        deltaHundred,
        deltaScore,
        deltaPercent,
        fromResponses: fromRow.respondentCount,
        toResponses: toRow.respondentCount,
        matchMode: matchMeta.mode,
        matchScore: matchMeta.score,
        fromRow,
        toRow,
    };
}

function getClosureMatchModeLabel(mode) {
    if (mode === "veryExpanded") return "مطابقة متوسعة جدًا";
    if (mode === "expanded") return "مطابقة متوسعة";
    if (mode === "flexible") return "ربط مرن";
    return "مطابقة مباشرة";
}

function formatClosureMatchScore(mode, score) {
    if (mode === "exact") return "البندان متطابقان مباشرة بين السنتين";
    if (!Number.isFinite(Number(score))) return "ربط سياقي";
    const percent = roundNumber(Number(score) * 100);
    return `قوة الربط ${toArabicNumber(percent)}%`;
}

function buildClosureLinkingCellHtml(row) {
    return `
        <div class="table-detail-stack">
            <span class="cell-title">${escapeHtml(row.fromRow.parentTitle || row.toRow.parentTitle || "—")}</span>
            <span class="cell-subtitle">الاستطلاع: ${escapeHtml(row.fromRow.surveyTitle || row.toRow.surveyTitle || "—")}</span>
            <span class="cell-subtitle">رقم العبارة: ${escapeHtml(getItemNumberLabel(row.fromRow))} → ${escapeHtml(getItemNumberLabel(row.toRow))}</span>
        </div>
    `;
}

function buildClosureYearTableCellHtml(sourceRow, yearLabel, average, percent, responses) {
    return `
        <div class="table-detail-stack">
            <span class="cell-title">${escapeHtml(yearLabel)}</span>
            <span class="cell-subtitle">الاستطلاع: ${escapeHtml(sourceRow.surveyTitle || "—")}</span>
            <span class="cell-subtitle">الموضوع: ${escapeHtml(sourceRow.parentTitle || "—")}</span>
            <span class="cell-subtitle">رقم العبارة: ${escapeHtml(getItemNumberLabel(sourceRow))}</span>
            <span class="cell-subtitle">البند: ${escapeHtml(sourceRow.title || "—")}</span>
            <span class="cell-subtitle">المتوسط: ${escapeHtml(formatScore(average))} · النسبة: ${escapeHtml(formatPercent(percent))}</span>
            <span class="cell-subtitle">عدد المقيمين: ${toArabicNumber(responses)}</span>
        </div>
    `;
}

function buildClosureReportSummaryHtml(payload) {
    const items = [
        `البرنامج: ${formatProgramLabel(payload.program)}`,
        `الفترة: ${payload.fromYearLabel} ← ${payload.toYearLabel}`,
        `الجنس: ${getGenderFilterLabel(state.closureFilters.gender)}`,
        `المطابقة: ${getClosureLevelLabel(payload.level)}`,
        `الحد الأدنى: ${formatClosureThreshold(payload.threshold)}`,
        `المصدر: ${CLOSURE_REPORT_SOURCE_LABEL}`,
    ];

    return items.map((item) => `<span class="report-summary-pill">${escapeHtml(item)}</span>`).join("");
}

function buildClosureReportFocusCellHtml(row) {
    const topicLabel = row.fromRow.parentTitle || row.toRow.parentTitle || "—";
    const surveyLabel = row.fromRow.surveyTitle || row.toRow.surveyTitle || "—";

    return `
        <div class="report-focus-cell">
            <span class="report-focus-title">${escapeHtml(row.title)}</span>
            <span class="report-focus-meta">الموضوع: ${escapeHtml(topicLabel)}</span>
            <span class="report-focus-meta">الاستطلاع: ${escapeHtml(surveyLabel)}</span>
            <span class="report-focus-meta">نوع المطابقة: ${escapeHtml(getClosureMatchModeLabel(row.matchMode))}</span>
        </div>
    `;
}

function buildClosureReportYearCellHtml(sourceRow, average, percent, responses) {
    const detailRows = [];

    if (state.closureFilters.reportShowResponses) {
        detailRows.push(`عدد المقيمين: ${toArabicNumber(responses)}`);
    }

    if (state.closureFilters.reportShowStatement) {
        detailRows.push(`العبارة: ${sourceRow.title || "—"}`);
    }

    if (state.closureFilters.reportShowSource) {
        detailRows.push(`المصدر: ${CLOSURE_REPORT_SOURCE_LABEL}`);
    }

    return `
        <div class="report-year-cell">
            <span class="report-score-pill">${escapeHtml(formatScore(average))}</span>
            ${detailRows.map((detail) => `<span class="report-year-detail">${escapeHtml(detail)}</span>`).join("")}
        </div>
    `;
}

function buildClosureReportImprovementCellHtml(row) {
    return `
        <div class="report-improvement-cell">
            <span class="report-score-pill is-improvement">${escapeHtml(formatClosureImprovement(row.deltaHundred))}</span>
            <span class="report-year-detail">${escapeHtml(formatScore(row.fromAverage))} → ${escapeHtml(formatScore(row.toAverage))}</span>
            <span class="report-year-detail">${escapeHtml(formatPercent(row.fromPercent))} → ${escapeHtml(formatPercent(row.toPercent))}</span>
        </div>
    `;
}

function renderClosureHighlights(container, rows, emptyText) {
    if (!container) return;
    if (!rows.length) {
        container.innerHTML = `<div class="empty-state">${escapeHtml(emptyText)}</div>`;
        return;
    }

    container.innerHTML = rows.slice(0, 5).map((row, index) => `
        <button class="insight-item insight-item-button" type="button" data-closure-detail-index="${index}">
            <div class="insight-item-title">${toArabicNumber(index + 1)}. ${escapeHtml(row.title)}</div>
            <div class="insight-item-meta">${escapeHtml(row.sectionLabel)} · ${escapeHtml(row.subtitle)}</div>
            <div class="insight-item-meta">${escapeHtml(formatPercent(row.fromPercent))} → ${escapeHtml(formatPercent(row.toPercent))} · ${toArabicNumber(row.fromResponses)} / ${toArabicNumber(row.toResponses)} استجابة</div>
            <div class="insight-item-value">${escapeHtml(formatClosureImprovement(row.deltaHundred))}</div>
            <div class="insight-item-foot">
                <span>اضغط لعرض تفاصيل البند في كل سنة</span>
                <span class="insight-item-arrow" aria-hidden="true">‹</span>
            </div>
        </button>
    `).join("");
}

function buildClosureNarrativeHtml(payload) {
    if (!payload.ready) {
        return `<p>${escapeHtml(payload.emptyMessage)}</p>`;
    }

    if (!payload.comparableRows.length) {
        return `<p>${escapeHtml(payload.emptyMessage)}</p>`;
    }

    if (!payload.qualifyingRows.length) {
        return `
            <p>تمت مقارنة <strong>${escapeHtml(toArabicNumber(payload.comparableRows.length))}</strong> مؤشرًا بين ${escapeHtml(payload.fromYearLabel)} و${escapeHtml(payload.toYearLabel)} في برنامج <strong>${escapeHtml(payload.program.name)}</strong>، لكن لم يظهر تحسن يتجاوز الحد الأدنى المعتمد (${escapeHtml(formatClosureThreshold(payload.threshold))}).</p>
            <p>يمكن توسيع فرص الالتقاط بتخفيف الحد الأدنى قليلًا أو تغيير نطاق الموضوع حتى تظهر مقارنات أكثر داخل الاستطلاعات نفسها.</p>
        `;
    }

    const topRows = payload.qualifyingRows.slice(0, 3);
    return `
        <p>أظهر برنامج <strong>${escapeHtml(payload.program.name)}</strong> عدد <strong>${escapeHtml(toArabicNumber(payload.qualifyingRows.length))}</strong> من دلائل التحسن المؤثرة بين ${escapeHtml(payload.fromYearLabel)} و${escapeHtml(payload.toYearLabel)}، بعد استبعاد أي زيادة تقل عن <strong>${escapeHtml(formatClosureThreshold(payload.threshold))}</strong>.</p>
        <p>أقوى مؤشرات الإغلاق كانت في: ${topRows.map((row) => `<strong>${escapeHtml(row.title)}</strong> (${escapeHtml(formatClosureImprovement(row.deltaHundred))})`).join("، ")}.</p>
        <p>متوسط التحسن في المؤشرات المؤهلة بلغ <strong>${escapeHtml(formatClosureImprovement(payload.averageImprovement))}</strong>، وهذا يمنح البرنامج شواهد مباشرة على تحسن النتائج بين السنتين في النطاق المحدد.</p>
    `;
}

function openClosureDetailSheet(row, payload) {
    if (!refs.bottomSheet || !refs.bottomSheetOptions) return;

    state.bottomSheetContext = "closure-detail";
    refs.bottomSheetTitle.textContent = row.title;
    if (refs.bottomSheetSearchWrap) refs.bottomSheetSearchWrap.classList.add("hidden");
    if (refs.bottomSheetSearch) refs.bottomSheetSearch.value = "";
    refs.bottomSheetOptions.innerHTML = buildClosureDetailHtml(row, payload);
    refs.bottomSheet.classList.remove("hidden");
    document.body.classList.add("sheet-open");
}

function closeBottomSheet() {
    if (!refs.bottomSheet) return;
    refs.bottomSheet.classList.add("hidden");
    if (refs.bottomSheetSearchWrap) refs.bottomSheetSearchWrap.classList.add("hidden");
    if (refs.bottomSheetSearch) refs.bottomSheetSearch.value = "";
    if (refs.bottomSheetOptions) refs.bottomSheetOptions.innerHTML = "";
    document.body.classList.remove("sheet-open");
    state.bottomSheetContext = "";
}

function buildClosureDetailHtml(row, payload) {
    const comparisonLabel = getClosureMatchModeLabel(row.matchMode);
    const summaryText = row.matchMode === "veryExpanded"
        ? `استخدم هذا العرض مطابقة متوسعة جدًا لالتقاط التحسن حتى مع تباعد أكبر في الصياغة، مع إبقاء الربط داخل السياق الأقرب في ${payload.fromYearLabel} و${payload.toYearLabel}.`
        : row.matchMode === "expanded"
        ? `استخدم هذا العرض مطابقة متوسعة لالتقاط التحسن بين بندين متقاربين في السياق حتى لو كانت الصياغة أبعد من المطابقة المضبوطة بين ${payload.fromYearLabel} و${payload.toYearLabel}.`
        : row.matchMode === "flexible"
            ? `تم ربط بندين متقاربين داخل الاستطلاع نفسه، ثم قياس التحسن بين ${payload.fromYearLabel} و${payload.toYearLabel}.`
            : `تمت مطابقة البند نفسه بين ${payload.fromYearLabel} و${payload.toYearLabel} ثم قياس مقدار التحسن.`;

    return `
        <section class="detail-sheet-summary">
            <div class="detail-sheet-badges">
                <span class="detail-sheet-chip">${escapeHtml(row.sectionLabel)}</span>
                <span class="detail-sheet-chip">${escapeHtml(comparisonLabel)}</span>
                <span class="detail-sheet-chip positive">${escapeHtml(formatClosureImprovement(row.deltaHundred))}</span>
            </div>
            <p>${escapeHtml(summaryText)}</p>
        </section>
        <section class="detail-sheet-grid">
            ${buildClosureYearCardHtml(row.fromRow, payload.fromYearLabel, row.fromAverage, row.fromPercent, row.fromResponses, "from")}
            ${buildClosureYearCardHtml(row.toRow, payload.toYearLabel, row.toAverage, row.toPercent, row.toResponses, "to")}
        </section>
    `;
}

function buildClosureYearCardHtml(sourceRow, yearLabel, average, percent, responses, tone) {
    return `
        <article class="detail-sheet-card theme-${escapeHtml(tone)}">
            <header class="detail-sheet-card-head">
                <div class="detail-sheet-year">${escapeHtml(yearLabel)}</div>
                <div class="detail-sheet-score">${escapeHtml(formatScore(average))}</div>
            </header>
            <div class="detail-sheet-stats">
                <div class="detail-sheet-stat">
                    <span class="detail-sheet-stat-label">النسبة</span>
                    <span class="detail-sheet-stat-value">${escapeHtml(formatPercent(percent))}</span>
                </div>
                <div class="detail-sheet-stat">
                    <span class="detail-sheet-stat-label">عدد المقيمين</span>
                    <span class="detail-sheet-stat-value">${toArabicNumber(responses)}</span>
                </div>
            </div>
            <div class="detail-sheet-fields">
                <div class="detail-sheet-field">
                    <span class="detail-sheet-field-label">الاستطلاع</span>
                    <span class="detail-sheet-field-value">${escapeHtml(sourceRow.surveyTitle || "—")}</span>
                </div>
                <div class="detail-sheet-field">
                    <span class="detail-sheet-field-label">الموضوع الدقيق</span>
                    <span class="detail-sheet-field-value">${escapeHtml(sourceRow.parentTitle || "—")}</span>
                </div>
                <div class="detail-sheet-field">
                    <span class="detail-sheet-field-label">رقم العبارة</span>
                    <span class="detail-sheet-field-value">${escapeHtml(getItemNumberLabel(sourceRow))}</span>
                </div>
                <div class="detail-sheet-field">
                    <span class="detail-sheet-field-label">نص البند</span>
                    <span class="detail-sheet-field-value">${escapeHtml(sourceRow.title || "—")}</span>
                </div>
            </div>
        </article>
    `;
}

function drawClosureChart(rows) {
    const topRows = rows.slice(0, 10);
    drawChart("closureDeltaChart", refs.closureDeltaChart, {
        type: "bar",
        data: {
            labels: topRows.map((row) => truncateLabel(row.title, 42)),
            datasets: [{
                label: "التحسن من 100",
                data: topRows.map((row) => row.deltaHundred),
                backgroundColor: "#1b8a61",
                borderRadius: 10,
            }],
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: "y",
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label(context) {
                            return `التحسن ${formatClosureImprovement(context.raw)}`;
                        },
                    },
                },
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: {
                        callback(value) {
                            return toArabicNumber(value);
                        },
                    },
                    grid: {
                        color: "rgba(18, 57, 54, 0.1)",
                    },
                },
                y: {
                    grid: { display: false },
                },
            },
        },
    });
}

function renderCustomSection() {
    const allRows = getCustomRows();
    const customQuery = refs.customSearchInput ? refs.customSearchInput.value.trim() : "";
    const rows = customQuery
        ? allRows.filter((row) => searchMatches(customQuery, [row.title, row.surveyTitle, row.topicLabel, row.sectionLabel, row.programName].join(" ")))
        : allRows;
    const selectedRows = allRows.filter((row) => state.customSelected.has(row.uid));
    const selectedAverage = averageScore(selectedRows);
    const totalSelectedResponses = selectedRows.reduce((sum, row) => sum + row.respondentCount, 0);

    refs.customMeta.textContent = `${buildSingleFilterScopeLabel(state.customFilters)} · ${toArabicNumber(rows.length)} عبارة متاحة`;
    refs.customAvailableMeta.textContent = `${toArabicNumber(rows.length)} عبارة`;
    refs.customSelectedMeta.textContent = `${toArabicNumber(selectedRows.length)} عبارة محددة`;

    renderMetricCards(refs.customIndicators, [
        {
            label: "العبارات المتاحة",
            value: toArabicNumber(rows.length),
            note: "بعد تطبيق الفلاتر الحالية",
            tone: "info",
        },
        {
            label: "العناصر المحددة",
            value: toArabicNumber(selectedRows.length),
            note: "المختارة يدويًا من القائمة",
            tone: "info",
        },
        {
            label: "متوسط المحدد",
            value: formatScore(selectedAverage),
            note: selectedRows.length ? formatResponseCount(totalSelectedResponses) : "لا توجد عناصر محددة",
            tone: toneForScore(selectedAverage),
        },
        {
            label: "أعلى عبارة محددة",
            value: selectedRows.length ? formatScore(getExtremeRow(selectedRows, "max").average) : "—",
            note: selectedRows.length ? truncateLabel(getExtremeRow(selectedRows, "max").title, 30) : "لا توجد عناصر محددة",
            tone: selectedRows.length ? "good" : "info",
        },
    ]);

    refs.customOptionsList.innerHTML = rows.map((row) => `
        <label class="custom-option">
            <input type="checkbox" data-row-id="${row.uid}" ${state.customSelected.has(row.uid) ? "checked" : ""}>
            <span class="custom-option-copy">
                <span class="custom-option-title">${escapeHtml(row.title)}</span>
                <span class="custom-option-meta">${escapeHtml(row.sectionLabel)} · ${escapeHtml(row.surveyTitle)} · ${escapeHtml(row.topicLabel)} · ${escapeHtml(row.programName)} · ${escapeHtml(`${row.year}هـ`)} · ${formatResponseCount(row.respondentCount)}</span>
            </span>
        </label>
    `).join("");

    refs.customSelectedTableBody.innerHTML = selectedRows.map((row) => `
        <tr>
            <td>${escapeHtml(row.sectionLabel)}</td>
            <td>${escapeHtml(row.surveyTitle)}</td>
            <td>
                <span class="cell-title">${escapeHtml(row.title)}</span>
                <span class="cell-subtitle">${escapeHtml(row.topicLabel)}</span>
            </td>
            <td>${escapeHtml(getGenderFilterLabel(state.customFilters.gender))}</td>
            <td>${toArabicNumber(row.respondentCount)}</td>
            <td>${renderScorePill(row.average)}</td>
        </tr>
    `).join("");

    refs.customEmpty.classList.toggle("hidden", selectedRows.length > 0);
}

function buildItemRecords() {
    const records = [];

    Object.entries(EXTRACTED_DATA).forEach(([datasetKey, dataset]) => {
        const [programId, year] = datasetKey.split("::");
        const program = getProgramById(programId);

        (dataset.surveys || []).forEach((survey, surveyIndex) => {
            const sectionLabel = getSectionLabel(survey.sectionId);
            (survey.topics || []).forEach((topic, topicIndex) => {
                (topic.items || []).forEach((item, itemIndex) => {
                    (item.genders || []).forEach((genderEntry, genderIndex) => {
                        const responses = Number(genderEntry.responses || 0);
                        const scoreTotal = Number(genderEntry.scoreTotal || 0);
                        if (!responses) return;

                        records.push({
                            uid: `${datasetKey}:${survey.id}:${topicIndex}:${itemIndex}:${genderIndex}`,
                            datasetKey,
                            programId,
                            programName: program.name,
                            degree: program.degree,
                            dept: program.dept,
                            year,
                            stakeholder: survey.stakeholder || "students",
                            stakeholderLabel: getStakeholderLabel(survey.stakeholder || "students"),
                            sectionId: survey.sectionId,
                            sectionLabel,
                            surveyId: survey.id,
                            surveyTitle: survey.title,
                            topicLabel: topic.label,
                            itemNumber: item.number || "",
                            itemLabel: item.label,
                            gender: genderEntry.gender || "",
                            responses,
                            scoreTotal,
                            average: roundNumber(scoreTotal / responses),
                            surveyIndex,
                            topicIndex,
                            itemOrder: parseItemOrder(item.number, itemIndex),
                            sortRank: getSectionOrder(survey.sectionId),
                            selfStudyEntries: getSelfStudyEntries(programId, year, survey.title, item.label),
                        });
                    });
                });
            });
        });
    });

    return records.sort((first, second) => {
        if (first.year !== second.year) return Number(second.year) - Number(first.year);
        if (first.programId !== second.programId) return (PROGRAM_ORDER.get(first.programId) || 99) - (PROGRAM_ORDER.get(second.programId) || 99);
        if (first.sortRank !== second.sortRank) return first.sortRank - second.sortRank;
        if (first.surveyIndex !== second.surveyIndex) return first.surveyIndex - second.surveyIndex;
        if (first.topicIndex !== second.topicIndex) return first.topicIndex - second.topicIndex;
        if (first.itemOrder !== second.itemOrder) return first.itemOrder - second.itemOrder;
        if (first.gender !== second.gender) return first.gender.localeCompare(second.gender, "ar");
        return first.itemLabel.localeCompare(second.itemLabel, "ar");
    });
}

function getItemRecordsForSingleFilters(filters) {
    return ITEM_RECORDS.filter((record) => matchesSingleFilters(record, filters));
}

function matchesSingleFilters(record, filters) {
    if (filters.program !== "all" && record.programId !== filters.program) return false;
    if (filters.year !== "all" && record.year !== filters.year) return false;
    if (filters.stakeholder !== "all" && record.stakeholder !== filters.stakeholder) return false;
    if (filters.gender !== "all" && record.gender !== filters.gender) return false;
    return matchesTopicFilter(record, filters);
}

function getRecordsForCompareSlot(slot) {
    return ITEM_RECORDS.filter((record) => {
        if (record.programId !== slot.program) return false;
        if (record.year !== slot.year) return false;
        if (state.compareFilters.stakeholder !== "all" && record.stakeholder !== state.compareFilters.stakeholder) return false;
        return matchesTopicFilter(record, state.compareFilters);
    });
}

function getRecordsForAnalysis() {
    if (!state.analysisFilters.programs.size) return [];
    return ITEM_RECORDS.filter((record) => {
        if (state.analysisFilters.programs.size && !state.analysisFilters.programs.has(record.programId)) return false;
        if (state.analysisFilters.year !== "all" && record.year !== state.analysisFilters.year) return false;
        if (state.analysisFilters.stakeholder !== "all" && record.stakeholder !== state.analysisFilters.stakeholder) return false;
        return matchesTopicFilter(record, state.analysisFilters);
    });
}

function getCustomRows() {
    return aggregateItemRows(getItemRecordsForSingleFilters(state.customFilters));
}

function matchesSubjectFilter(record, subjectFilter) {
    if (!subjectFilter || subjectFilter === "all") return true;
    if (subjectFilter.startsWith("section:")) {
        return record.sectionId === subjectFilter.replace("section:", "");
    }
    if (subjectFilter.startsWith("survey:")) {
        return normalizeText(record.surveyTitle) === normalizeText(subjectFilter.replace("survey:", ""));
    }
    if (subjectFilter.startsWith("topic:")) {
        return normalizeText(record.topicLabel) === normalizeText(subjectFilter.replace("topic:", ""));
    }
    return true;
}

function matchesTopicFilter(record, filters) {
    if ((filters.topicMode || DEFAULT_TOPIC_MODE) === "selfstudy") {
        return matchesSelfStudyFilter(record, filters.selfStudyTarget || "all");
    }
    return matchesSubjectFilter(record, filters.subject);
}

function matchesSelfStudyFilter(record, selfStudyTarget) {
    const entries = record.selfStudyEntries || [];
    if (selfStudyTarget === "all") return true;
    if (selfStudyTarget === "unlinked") return !entries.length;
    if (!entries.length) return false;

    if (selfStudyTarget.startsWith("criterion:")) {
        return entries.some((entry) => buildSelfStudyCriterionValue(entry) === selfStudyTarget);
    }

    if (selfStudyTarget.startsWith("side:")) {
        return entries.some((entry) => buildSelfStudySideValue(entry) === selfStudyTarget);
    }

    if (selfStudyTarget.startsWith("phrase:")) {
        return entries.some((entry) => buildSelfStudyPhraseValue(entry, record.itemLabel) === selfStudyTarget);
    }

    return true;
}

function getActiveCompareSlots() {
    return Object.values(state.compareSlots).filter((slot) => slot.program && slot.year);
}

function aggregateSurveyRows(records) {
    return aggregateRecords(records, "survey");
}

function aggregateTopicRows(records) {
    return aggregateRecords(records, "topic");
}

function aggregateItemRows(records) {
    return aggregateRecords(records, "item");
}

function aggregateRecords(records, level) {
    const map = new Map();

    records.forEach((record) => {
        const key = getAggregateKey(record, level);
        if (!map.has(key)) {
            map.set(key, createAggregateEntry(record, level, key));
        }

        const entry = map.get(key);
        entry._scoreTotal += record.scoreTotal;
        entry._responseTotal += record.responses;
        entry._itemKeys.add(`${normalizeText(record.topicLabel)}||${record.itemNumber}||${normalizeText(record.itemLabel)}`);
        entry._topicKeys.add(normalizeText(record.topicLabel));

        if (level === "item") {
            entry.respondentCount += record.responses;
        } else {
            const itemKey = `${normalizeText(record.topicLabel)}||${record.itemNumber}||${normalizeText(record.itemLabel)}`;
            entry._itemResponseMap.set(itemKey, (entry._itemResponseMap.get(itemKey) || 0) + record.responses);
        }
    });

    return Array.from(map.values())
        .map((entry) => finalizeAggregateEntry(entry))
        .sort(compareAggregateRows);
}

function createAggregateEntry(record, level, key) {
    const base = {
        uid: key,
        key,
        programId: record.programId,
        programName: record.programName,
        degree: record.degree,
        dept: record.dept,
        year: record.year,
        stakeholder: record.stakeholder,
        stakeholderLabel: record.stakeholderLabel,
        sectionId: record.sectionId,
        sectionLabel: record.sectionLabel,
        surveyTitle: record.surveyTitle,
        topicLabel: record.topicLabel,
        title: record.surveyTitle,
        displayTitle: record.surveyTitle,
        parentTitle: "",
        rowKind: "استطلاع",
        isPrimary: level === "survey",
        rowDepth: 0,
        surveyIndex: record.surveyIndex,
        topicIndex: record.topicIndex,
        itemNumber: record.itemNumber,
        itemOrder: record.itemOrder,
        sortRank: record.sortRank,
        respondentCount: 0,
        _scoreTotal: 0,
        _responseTotal: 0,
        _itemKeys: new Set(),
        _topicKeys: new Set(),
        _itemResponseMap: new Map(),
    };

    if (level === "topic") {
        base.title = record.topicLabel;
        base.displayTitle = record.topicLabel;
        base.parentTitle = record.surveyTitle;
        base.rowKind = "موضوع";
        base.isPrimary = false;
        base.rowDepth = 1;
    }

    if (level === "item") {
        base.title = record.itemLabel;
        base.displayTitle = record.itemLabel;
        base.parentTitle = record.topicLabel;
        base.rowKind = "عبارة";
        base.isPrimary = false;
        base.rowDepth = 2;
    }

    return base;
}

function finalizeAggregateEntry(entry) {
    const respondentCount = entry.rowKind === "عبارة"
        ? entry.respondentCount
        : Math.max(0, ...entry._itemResponseMap.values());

    return {
        uid: entry.uid,
        programId: entry.programId,
        programName: entry.programName,
        degree: entry.degree,
        dept: entry.dept,
        year: entry.year,
        stakeholder: entry.stakeholder,
        stakeholderLabel: entry.stakeholderLabel,
        sectionId: entry.sectionId,
        sectionLabel: entry.sectionLabel,
        surveyTitle: entry.surveyTitle,
        topicLabel: entry.topicLabel,
        title: entry.title,
        displayTitle: entry.displayTitle,
        parentTitle: entry.parentTitle,
        rowKind: entry.rowKind,
        isPrimary: entry.isPrimary,
        rowDepth: entry.rowDepth,
        surveyIndex: entry.surveyIndex,
        topicIndex: entry.topicIndex,
        itemNumber: entry.itemNumber,
        itemOrder: entry.itemOrder,
        sortRank: entry.sortRank,
        average: entry._responseTotal ? roundNumber(entry._scoreTotal / entry._responseTotal) : null,
        respondentCount,
        responseTotal: entry._responseTotal,
        itemCount: entry._itemKeys.size,
        topicCount: entry._topicKeys.size,
    };
}

function buildProgramTableRows(surveyRows, topicRows, itemRows) {
    return [...surveyRows, ...topicRows, ...itemRows].sort(compareAggregateRows);
}

function compareAggregateRows(first, second) {
    if (first.year !== second.year) return Number(second.year) - Number(first.year);
    if (first.programId !== second.programId) return (PROGRAM_ORDER.get(first.programId) || 99) - (PROGRAM_ORDER.get(second.programId) || 99);
    if (first.sortRank !== second.sortRank) return first.sortRank - second.sortRank;
    if (first.surveyIndex !== second.surveyIndex) return first.surveyIndex - second.surveyIndex;
    if (first.rowDepth !== second.rowDepth) return first.rowDepth - second.rowDepth;
    if (first.topicIndex !== second.topicIndex) return first.topicIndex - second.topicIndex;
    if (first.itemOrder !== second.itemOrder) return first.itemOrder - second.itemOrder;
    return first.displayTitle.localeCompare(second.displayTitle, "ar");
}

function buildComparisonTableRows(selectionRows) {
    const metaMap = new Map();
    const valueMaps = selectionRows.map((rows) => {
        const map = new Map();
        rows.forEach((row) => {
            const key = getComparisonKey(row);
            map.set(key, {
                average: row.average,
                respondentCount: row.respondentCount,
            });
            if (!metaMap.has(key)) metaMap.set(key, row);
        });
        return map;
    });

    return Array.from(metaMap.keys())
        .sort((firstKey, secondKey) => compareAggregateRows(metaMap.get(firstKey), metaMap.get(secondKey)))
        .map((key) => ({
            meta: metaMap.get(key),
            values: valueMaps.map((map) => map.get(key) || null),
        }));
}

function buildAnalysisProgramRows(surveyRows) {
    return getProgramsWithData()
        .filter((program) => state.analysisFilters.programs.has(program.id))
        .map((program) => {
            const programSurveyRows = surveyRows.filter((row) => row.programId === program.id);
            if (!programSurveyRows.length) return null;

            return {
                program,
                average: averageScore(programSurveyRows),
                count: programSurveyRows.length,
                totalResponses: programSurveyRows.reduce((sum, row) => sum + row.respondentCount, 0),
                sectionAverages: SECTION_META.map((section) => ({
                    id: section.id,
                    label: section.shortLabel,
                    average: averageScore(programSurveyRows.filter((row) => row.sectionId === section.id)),
                })),
            };
        })
        .filter(Boolean)
        .sort((first, second) => second.average - first.average);
}

function buildAnalysisTable(programRows) {
    if (!programRows.length) {
        return {
            title: "مصفوفة التحليل",
            meta: "0 صف",
            head: "",
            body: "",
            hasRows: false,
        };
    }

    return {
        title: "مصفوفة المحاور حسب البرامج المحددة",
        meta: `${toArabicNumber(programRows.length)} برنامج`,
        head: `
            <tr>
                <th>البرنامج</th>
                <th>عدد الاستطلاعات</th>
                <th>عدد الاستجابات</th>
                ${SECTION_META.map((section) => `<th>${escapeHtml(section.shortLabel)}</th>`).join("")}
            </tr>
        `,
        body: programRows.map((item) => `
            <tr>
                <td>
                    <span class="cell-title">${escapeHtml(item.program.name)}</span>
                    <span class="cell-subtitle">${escapeHtml(item.program.degree)}</span>
                </td>
                <td>${toArabicNumber(item.count)}</td>
                <td>${toArabicNumber(item.totalResponses)}</td>
                ${item.sectionAverages.map((section) => `<td>${section.average == null ? "—" : renderScorePill(section.average)}</td>`).join("")}
            </tr>
        `).join(""),
        hasRows: true,
    };
}

function pruneCustomSelection() {
    const visibleIds = new Set(getCustomRows().map((row) => row.uid));
    state.customSelected.forEach((rowId) => {
        if (!visibleIds.has(rowId)) {
            state.customSelected.delete(rowId);
        }
    });
}

function buildProgramRankingRows(surveyRows) {
    return getProgramsWithData()
        .map((program) => {
            const programRows = surveyRows.filter((row) => row.programId === program.id);
            if (!programRows.length) return null;

            const bestSection = SECTION_META
                .map((section) => ({
                    label: section.shortLabel,
                    average: averageScore(programRows.filter((row) => row.sectionId === section.id)),
                }))
                .filter((item) => item.average != null)
                .sort((first, second) => second.average - first.average)[0];

            return {
                program,
                count: programRows.length,
                average: averageScore(programRows),
                bestSection: bestSection ? bestSection.label : "",
            };
        })
        .filter(Boolean)
        .sort((first, second) => second.average - first.average);
}

function buildScoreBandSeries(rows) {
    const bands = [
        { label: "قوي", test: (score) => score >= 4.25, color: "#1b8a61" },
        { label: "جيد", test: (score) => score >= 3.75 && score < 4.25, color: "#2e76b7" },
        { label: "متوسط", test: (score) => score >= 3.25 && score < 3.75, color: "#c17b1f" },
        { label: "بحاجة متابعة", test: (score) => score < 3.25, color: "#be4b3b" },
    ];

    return bands.map((band) => ({
        label: band.label,
        value: rows.filter((row) => band.test(row.average)).length,
        color: band.color,
    }));
}

function renderMetricCards(container, cards) {
    container.innerHTML = cards.map((card) => `
        <article class="kpi-card ${escapeHtml(card.tone || "info")}">
            <div class="kpi-label">${escapeHtml(card.label)}</div>
            <div class="kpi-value">${escapeHtml(card.value)}</div>
            <div class="kpi-note">${escapeHtml(card.note || "")}</div>
        </article>
    `).join("");
}

/* ===== SEARCH TAB ===== */

function getVisibleSearchCriteria(record, queryInfo) {
    return (record.selfStudyEntries || [])
        .filter((entry) => entry.criterionCode)
        .filter((entry) => !queryInfo.criterionCode || normalizeCriterionCodeForSearch(entry.criterionCode).startsWith(queryInfo.criterionCode));
}

function buildSearchCriterionSummary(uniqueCriterionCodes) {
    if (!uniqueCriterionCodes.length) return "غير مرتبط بمحك";
    return `المحكات: ${escapeHtml(uniqueCriterionCodes.slice(0, 4).join("، "))}${uniqueCriterionCodes.length > 4 ? ` +${toArabicNumber(uniqueCriterionCodes.length - 4)}` : ""}`;
}

function buildSearchResultMarkup(record, raw, queryInfo) {
    const scoreTone = record.average >= 3.5 ? "good" : record.average >= 2.5 ? "ok" : "low";
    const visibleCriteria = getVisibleSearchCriteria(record, queryInfo);
    const uniqueCriterionCodes = Array.from(new Set(visibleCriteria.map((entry) => entry.criterionCode)));
    const genderLabel = getSearchResultGenderLabel(record.gender);
    const criterionBadges = uniqueCriterionCodes
        .map((code) => `<span class="search-result-badge is-criterion">${escapeHtml(code)}</span>`)
        .join("");
    const criterionSummary = buildSearchCriterionSummary(uniqueCriterionCodes);
    const unlinkedBadge = uniqueCriterionCodes.length ? "" : `<span class="search-result-badge">غير مرتبط بمحك</span>`;

    let contextHtml = escapeHtml(record.itemLabel);
    if (raw.length >= 2) {
        const escaped = escapeRegExp(raw);
        try {
            contextHtml = contextHtml.replace(new RegExp(`(${escaped})`, "gi"), "<mark>$1</mark>");
        } catch (error) {
            /* ignore regex error */
        }
    }

    return `<div class="search-result-item">
        <div class="search-result-head">
            <span class="search-result-title">${contextHtml}</span>
            <span class="search-result-score" data-tone="${scoreTone}">المتوسط ${formatScore(record.average)}</span>
        </div>
        <div class="search-result-badges">
            <span class="search-result-badge is-program">${escapeHtml(record.programName)}</span>
            <span class="search-result-badge">${escapeHtml(record.degree)}</span>
            <span class="search-result-badge is-year">${record.year}هـ</span>
            <span class="search-result-badge is-gender">${escapeHtml(genderLabel)}</span>
            <span class="search-result-badge">${escapeHtml(record.stakeholderLabel)}</span>
            <span class="search-result-badge">${escapeHtml(record.sectionLabel)}</span>
            ${unlinkedBadge}
            ${criterionBadges}
        </div>
        <div class="search-result-context">
            ${escapeHtml(record.surveyTitle)}${record.topicLabel ? " · " + escapeHtml(record.topicLabel) : ""}
        </div>
        <div class="search-result-metrics">
            <span class="search-result-metric">الجنس: ${escapeHtml(genderLabel)}</span>
            <span class="search-result-metric">عدد الاستجابات: ${formatResponseCount(record.responses)}</span>
            <span class="search-result-metric">درجة المتوسط: ${escapeHtml(formatScore(record.average))}</span>
        </div>
        ${criterionSummary ? `<div class="search-result-context">${criterionSummary}</div>` : ""}
    </div>`;
}

function getSearchResultsPayload(rawInput) {
    const raw = String(rawInput || "").trim();
    const queryInfo = analyzeSearchQuery(raw);
    const hasStructuredQuery = Boolean(queryInfo.criterionCode || queryInfo.years.size || queryInfo.degrees.size);
    const results = [];

    if (!raw) {
        return {
            raw,
            queryInfo,
            hasStructuredQuery,
            results,
            displayResults: results,
            truncated: false,
        };
    }

    ITEM_RECORDS.forEach((record) => {
        const entries = record.selfStudyEntries || [];
        let matched = true;

        if (queryInfo.degrees.size && !degreeMatchesQuery(record.degree, queryInfo.degrees)) {
            matched = false;
        }

        if (matched && queryInfo.years.size && !queryInfo.years.has(String(record.year))) {
            matched = false;
        }

        if (matched && queryInfo.criterionCode) {
            matched = entries.some((entry) => {
                if (!entry.criterionCode) return false;
                return normalizeCriterionCodeForSearch(entry.criterionCode).startsWith(queryInfo.criterionCode);
            });
        }

        if (matched && queryInfo.textQuery) {
            const searchText = [
                record.programName, record.surveyTitle, record.topicLabel,
                record.itemLabel, record.sectionLabel, record.stakeholderLabel,
                record.year, record.degree,
                `${record.year}ه`,
                ...(entries.length ? entries.map((entry) => `${entry.criterionCode} ${entry.criterionText} ${entry.supportedSide}`) : ["غير مرتبط بمحك"]),
            ].join(" ");
            matched = state.searchMode === "precise"
                ? searchMatchesPrecisely(queryInfo.textQuery, searchText)
                : searchMatches(queryInfo.textQuery, searchText);
        }

        if (!hasStructuredQuery && !queryInfo.textQuery) matched = false;
        if (matched) results.push(record);
    });

    return {
        raw,
        queryInfo,
        hasStructuredQuery,
        results,
        displayResults: results.slice(0, 200),
        truncated: results.length > 200,
    };
}

function renderSearchResults() {
    const payload = getSearchResultsPayload(refs.searchInput ? refs.searchInput.value : "");
    if (!payload.raw) {
        if (refs.searchEmpty) refs.searchEmpty.classList.remove("hidden");
        if (refs.searchNoResults) refs.searchNoResults.classList.add("hidden");
        if (refs.searchResultsList) { refs.searchResultsList.classList.add("hidden"); refs.searchResultsList.innerHTML = ""; }
        if (refs.searchResultCount) refs.searchResultCount.textContent = "";
        return;
    }
    if (refs.searchEmpty) refs.searchEmpty.classList.add("hidden");

    /* Show results */
    if (!payload.results.length) {
        if (refs.searchNoResults) refs.searchNoResults.classList.remove("hidden");
        if (refs.searchNoResultsHint) refs.searchNoResultsHint.textContent = `لم يُعثر على نتائج لـ "${payload.raw}"`;
        if (refs.searchResultsList) { refs.searchResultsList.classList.add("hidden"); refs.searchResultsList.innerHTML = ""; }
        if (refs.searchResultCount) refs.searchResultCount.textContent = "";
        return;
    }

    if (refs.searchNoResults) refs.searchNoResults.classList.add("hidden");
    if (refs.searchResultsList) refs.searchResultsList.classList.remove("hidden");
    if (refs.searchResultCount) refs.searchResultCount.textContent = `${toArabicNumber(payload.results.length)} نتيجة`;

    if (refs.searchResultsList) {
        refs.searchResultsList.innerHTML = payload.displayResults
            .map((record) => buildSearchResultMarkup(record, payload.raw, payload.queryInfo))
            .join("");

        if (payload.truncated) {
            refs.searchResultsList.innerHTML += `<div class="search-empty-state"><p>يُعرض أول ٢٠٠ نتيجة من ${toArabicNumber(payload.results.length)}. حاول تضييق البحث.</p></div>`;
        }
    }
}

function renderFilterChips(container, chips) {
    container.innerHTML = chips.map((chip) => `
        <span class="chip">${escapeHtml(chip.label)}: ${escapeHtml(chip.value)}</span>
    `).join("");
}

function renderInsightList(container, items, emptyText) {
    if (!items.length) {
        container.innerHTML = `<div class="empty-state">${escapeHtml(emptyText)}</div>`;
        return;
    }

    container.innerHTML = items.map((item) => `
        <article class="insight-item">
            <div class="insight-item-title">${escapeHtml(item.title)}</div>
            <div class="insight-item-meta">${escapeHtml(item.sectionLabel)} · ${escapeHtml(item.programName)} · ${escapeHtml(item.year)}هـ · ${formatResponseCount(item.respondentCount)}</div>
            <div class="insight-item-value">${formatScore(item.average)}</div>
        </article>
    `).join("");
}

function drawChart(key, canvas, config) {
    if (!canvas || !window.Chart) return;
    const card = canvas.closest(".chart-card");
    const wrap = canvas.parentElement;
    let emptyLabel = card ? card.querySelector(".chart-empty") : null;

    if (!emptyLabel && wrap) {
        emptyLabel = document.createElement("div");
        emptyLabel.className = "chart-empty hidden";
        wrap.insertAdjacentElement("afterend", emptyLabel);
    }

    const hasData = chartHasData(config);
    if (emptyLabel) {
        emptyLabel.classList.toggle("hidden", hasData);
        emptyLabel.textContent = hasData ? "" : "لا توجد بيانات كافية للرسم في هذا النطاق.";
    }
    canvas.classList.toggle("hidden", !hasData);

    if (!hasData) {
        if (state.charts[key]) {
            state.charts[key].destroy();
            delete state.charts[key];
        }
        return;
    }

    if (state.charts[key]) {
        state.charts[key].destroy();
    }

    state.charts[key] = new window.Chart(canvas.getContext("2d"), config);
}

function baseChartOptions(overrides = {}) {
    return {
        responsive: true,
        maintainAspectRatio: false,
        locale: "ar-SA",
        plugins: {
            legend: {
                labels: {
                    font: {
                        family: "Tajawal",
                        size: 12,
                    },
                },
            },
            tooltip: {
                bodyFont: { family: "Tajawal" },
                titleFont: { family: "Tajawal" },
                callbacks: {
                    label(context) {
                        const value = context.raw;
                        if (typeof value === "number") {
                            return `${context.dataset.label || ""} ${formatScore(value)}`.trim();
                        }
                        return context.dataset.label || "";
                    },
                },
            },
        },
        scales: {},
        ...overrides,
    };
}

function buildScoreScales(axis) {
    const scales = {
        [axis]: {
            min: 0,
            max: 5,
            ticks: {
                stepSize: 1,
            },
            grid: {
                color: "rgba(18, 57, 54, 0.1)",
            },
        },
    };

    if (axis === "x") {
        scales.y = {
            grid: { display: false },
        };
    } else {
        scales.x = {
            grid: { display: false },
        };
    }

    return scales;
}

/* ===== TREND SECTION — التطور الزمني ===== */
function renderTrendControls() {
    if (!refs.trendProgramSelect) return;
    renderProgramOptionsFor(refs.trendProgramSelect, state.trendFilters.program, false);
    if (refs.trendStakeholderSelect) {
        const stakeholders = [{ value: "all", label: "كل الجهات" }, ...Object.entries(STAKEHOLDER_LABELS).map(([k, v]) => ({ value: k, label: v }))];
        refs.trendStakeholderSelect.innerHTML = stakeholders.map(s =>
            `<option value="${s.value}"${s.value === state.trendFilters.stakeholder ? " selected" : ""}>${escapeHtml(s.label)}</option>`
        ).join("");
    }
    renderChipOptions(refs.trendTopicModeChips, [
        { value: "general", label: "موضوعات عامة" },
        { value: "selfstudy", label: "محكات الدراسة الذاتية" },
    ], state.trendFilters.topicMode, (val) => {
        state.trendFilters.topicMode = val;
        renderTrendControls();
        renderTrendSection();
    });
    if (refs.trendSelfStudyPanel) refs.trendSelfStudyPanel.classList.toggle("hidden", state.trendFilters.topicMode !== "selfstudy");
    if (state.trendFilters.topicMode === "selfstudy") {
        const baseRecords = ITEM_RECORDS.filter(r => state.trendFilters.program === "all" || r.programId === state.trendFilters.program);
        renderSelfStudyOptions(refs.trendSelfStudyFilter, null, baseRecords, state.trendFilters);
    }
}

function renderTrendSection() {
    renderTrendControls();
    const programId = state.trendFilters.program;
    if (refs.insightsMeta) {
        refs.insightsMeta.textContent = programId && programId !== "all"
            ? `${formatProgramLabel(getProgramById(programId))} · ${getTopicSelectionLabel(state.trendFilters)}`
            : "اختر برنامجاً لعرض التطور الزمني";
    }
    if (!programId || programId === "all") {
        if (refs.trendEmpty) { refs.trendEmpty.classList.remove("hidden"); }
        if (refs.trendChangeSummary) refs.trendChangeSummary.innerHTML = "";
        if (refs.trendIndicators) refs.trendIndicators.innerHTML = "";
        return;
    }
    if (refs.trendEmpty) refs.trendEmpty.classList.add("hidden");

    const years = sortYears(getAvailableYears(programId));
    const stakeholder = state.trendFilters.stakeholder;

    /* Build data per year per section */
    const sectionYearData = {};
    SECTION_META.forEach(sec => { sectionYearData[sec.id] = {}; });

    years.forEach(year => {
        const records = ITEM_RECORDS.filter(r =>
            r.programId === programId && r.year === year &&
            (stakeholder === "all" || r.stakeholder === stakeholder) &&
            matchesTopicFilter(r, state.trendFilters)
        );
        const surveys = aggregateSurveyRows(records);
        SECTION_META.forEach(sec => {
            const secSurveys = surveys.filter(s => s.sectionId === sec.id);
            if (secSurveys.length) {
                const avg = roundNumber(secSurveys.reduce((s, r) => s + r.average, 0) / secSurveys.length);
                sectionYearData[sec.id][year] = avg;
            }
        });
    });

    /* Overall average per year */
    const overallByYear = {};
    years.forEach(year => {
        const vals = SECTION_META.map(sec => sectionYearData[sec.id][year]).filter(v => v !== undefined);
        if (vals.length) overallByYear[year] = roundNumber(vals.reduce((a, b) => a + b, 0) / vals.length);
    });

    /* KPI cards */
    const latestYear = years[0];
    const prevYear = years[1];
    const latestAvg = overallByYear[latestYear];
    const prevAvg = prevYear ? overallByYear[prevYear] : null;
    const delta = prevAvg != null && latestAvg != null ? roundNumber(latestAvg - prevAvg) : null;

    renderMetricCards(refs.trendIndicators, [
        { label: `المتوسط ${latestYear || ""}هـ`, value: latestAvg != null ? formatScore(latestAvg) : "—", note: "المتوسط العام للبرنامج" },
        { label: prevYear ? `المتوسط ${prevYear}هـ` : "السنة السابقة", value: prevAvg != null ? formatScore(prevAvg) : "—", note: "المتوسط العام السابق" },
        { label: "التغيّر", value: delta != null ? (delta >= 0 ? `+${formatScore(delta)}` : formatScore(delta)) : "—", note: delta != null ? (delta > 0 ? "تحسّن" : delta < 0 ? "تراجع" : "مستقر") : "لا توجد بيانات مقارنة" },
    ]);

    /* Line chart */
    if (refs.trendLineChart && window.Chart) {
        const datasets = SECTION_META.map((sec, i) => ({
            label: sec.shortLabel,
            data: years.map(y => sectionYearData[sec.id][y] || null),
            borderColor: CHART_COLORS[i % CHART_COLORS.length],
            backgroundColor: CHART_COLORS[i % CHART_COLORS.length] + "22",
            tension: 0.3,
            fill: false,
            pointRadius: 5,
            pointHoverRadius: 7,
        }));
        drawChart("trendLine", refs.trendLineChart, {
            type: "line",
            data: { labels: years.map(y => `${y}هـ`), datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: "top", rtl: true, labels: { font: { family: "inherit" } } } },
                scales: { y: { min: 1, max: 5, ticks: { stepSize: 0.5 } } },
            },
        });
    }

    /* Change summary */
    if (refs.trendChangeSummary && years.length >= 2) {
        const changes = SECTION_META.map(sec => {
            const latest = sectionYearData[sec.id][latestYear];
            const prev = sectionYearData[sec.id][prevYear];
            if (latest == null || prev == null) return null;
            const d = roundNumber(latest - prev);
            return { label: sec.label, latest, prev, delta: d };
        }).filter(Boolean).sort((a, b) => b.delta - a.delta);

        refs.trendChangeSummary.innerHTML = changes.map(c => {
            const cls = c.delta > 0.05 ? "positive" : c.delta < -0.05 ? "negative" : "neutral";
            const icon = c.delta > 0.05 ? "📈" : c.delta < -0.05 ? "📉" : "➡️";
            return `<div class="trend-change-item">
                <span class="change-icon">${icon}</span>
                <span class="change-label">${escapeHtml(c.label)}</span>
                <span class="change-values">${formatScore(c.prev)} → ${formatScore(c.latest)}</span>
                <span class="change-delta ${cls}">${c.delta >= 0 ? "+" : ""}${formatScore(c.delta)}</span>
            </div>`;
        }).join("") || '<div class="empty-state">لا توجد بيانات كافية للمقارنة.</div>';
    }
}

/* ===== GAPS SECTION — تقرير الفجوات ===== */
function renderGapsControls() {
    if (refs.gapsProgramSelect) renderProgramOptionsFor(refs.gapsProgramSelect, state.gapsFilters.program, false);
    if (refs.gapsYearSelect) {
        const prog = state.gapsFilters.program;
        const years = prog !== "all" ? getAvailableYears(prog) : ALL_AVAILABLE_YEARS;
        refs.gapsYearSelect.innerHTML = years.map(y =>
            `<option value="${y}"${y === state.gapsFilters.year ? " selected" : ""}>${y}هـ</option>`
        ).join("");
        if (!years.includes(state.gapsFilters.year) && years.length) {
            state.gapsFilters.year = years[0];
        }
    }
    renderChipOptions(refs.gapsTopicModeChips, [
        { value: "general", label: "موضوعات عامة" },
        { value: "selfstudy", label: "محكات الدراسة الذاتية" },
    ], state.gapsFilters.topicMode, (val) => {
        state.gapsFilters.topicMode = val;
        renderGapsControls();
        renderGapsSection();
    });
    if (refs.gapsSelfStudyPanel) refs.gapsSelfStudyPanel.classList.toggle("hidden", state.gapsFilters.topicMode !== "selfstudy");
    if (state.gapsFilters.topicMode === "selfstudy") {
        const baseRecords = ITEM_RECORDS.filter(r =>
            (state.gapsFilters.program === "all" || r.programId === state.gapsFilters.program) &&
            (state.gapsFilters.year === "all" || r.year === state.gapsFilters.year)
        );
        renderSelfStudyOptions(refs.gapsSelfStudyFilter, null, baseRecords, state.gapsFilters);
    }
}

function renderGapsSection() {
    renderGapsControls();
    const { program, year, target } = state.gapsFilters;
    if (refs.insightsMeta) {
        refs.insightsMeta.textContent = program && program !== "all"
            ? `${formatProgramLabel(getProgramById(program))} · ${year}هـ · المستهدف ${formatScore(target)}`
            : "اختر برنامجاً وسنة لعرض تقرير الفجوات";
    }
    if (!program || program === "all") {
        if (refs.gapsEmpty) refs.gapsEmpty.classList.remove("hidden");
        if (refs.gapsTableBody) refs.gapsTableBody.innerHTML = "";
        if (refs.gapsIndicators) refs.gapsIndicators.innerHTML = "";
        if (refs.gapsNarrativeCard) refs.gapsNarrativeCard.classList.add("hidden");
        return;
    }
    if (refs.gapsEmpty) refs.gapsEmpty.classList.add("hidden");

    const records = ITEM_RECORDS.filter(r =>
        r.programId === program && r.year === year && r.stakeholder === "students" &&
        matchesTopicFilter(r, state.gapsFilters)
    );
    const surveys = aggregateSurveyRows(records);

    const gapRows = surveys.map(s => {
        const gap = roundNumber(s.average - target);
        const status = gap >= 0 ? "achieved" : gap >= -0.3 ? "close" : "below";
        const statusLabel = gap >= 0 ? "محقق" : gap >= -0.3 ? "قريب" : "دون المستهدف";
        return { ...s, gap, status, statusLabel };
    }).sort((a, b) => a.gap - b.gap);

    /* KPI */
    const achieved = gapRows.filter(r => r.status === "achieved").length;
    const close = gapRows.filter(r => r.status === "close").length;
    const below = gapRows.filter(r => r.status === "below").length;
    const overallAvg = gapRows.length ? roundNumber(gapRows.reduce((s, r) => s + r.average, 0) / gapRows.length) : 0;

    renderMetricCards(refs.gapsIndicators, [
        { label: "المتوسط العام", value: formatScore(overallAvg), note: `المستهدف: ${formatScore(target)}` },
        { label: "محقق", value: toArabicNumber(achieved), note: `من ${toArabicNumber(gapRows.length)} استطلاع` },
        { label: "قريب من المستهدف", value: toArabicNumber(close), note: "فجوة أقل من 0.3" },
        { label: "دون المستهدف", value: toArabicNumber(below), note: "يحتاج تحسين" },
    ]);

    /* Bar chart */
    if (refs.gapsBarChart && window.Chart) {
        const sectionGaps = SECTION_META.map(sec => {
            const secRows = gapRows.filter(r => r.sectionId === sec.id);
            if (!secRows.length) return null;
            const avg = roundNumber(secRows.reduce((s, r) => s + r.average, 0) / secRows.length);
            return { label: sec.shortLabel, avg, gap: roundNumber(avg - target) };
        }).filter(Boolean);

        drawChart("gapsBar", refs.gapsBarChart, {
            type: "bar",
            data: {
                labels: sectionGaps.map(s => s.label),
                datasets: [
                    { label: "المتوسط", data: sectionGaps.map(s => s.avg), backgroundColor: sectionGaps.map(s => s.gap >= 0 ? "#16a34a88" : "#dc262688"), borderRadius: 6 },
                    { label: "المستهدف", data: sectionGaps.map(() => target), type: "line", borderColor: "#c79e54", borderDash: [6, 3], pointRadius: 0, fill: false },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: "top", rtl: true, labels: { font: { family: "inherit" } } } },
                scales: { y: { min: 1, max: 5, ticks: { stepSize: 0.5 } } },
            },
        });
    }

    /* Table */
    if (refs.gapsTableBody) {
        refs.gapsTableBody.innerHTML = gapRows.map(r =>
            `<tr>
                <td>${escapeHtml(r.sectionLabel)}</td>
                <td>${escapeHtml(r.title)}</td>
                <td>${formatScore(r.average)}</td>
                <td>${formatScore(target)}</td>
                <td>${r.gap >= 0 ? "+" : ""}${formatScore(r.gap)}</td>
                <td><span class="gap-status ${r.status}">${r.statusLabel}</span></td>
            </tr>`
        ).join("");
    }

    /* Narrative for self-study */
    if (refs.gapsNarrativeCard && refs.gapsNarrativeText && gapRows.length) {
        refs.gapsNarrativeCard.classList.remove("hidden");
        const prog = getProgramById(program);
        const strengths = gapRows.filter(r => r.status === "achieved").slice(-3);
        const weaknesses = gapRows.filter(r => r.status === "below").slice(0, 3);

        let text = `بلغ المتوسط العام لاستطلاعات برنامج ${prog.name} (${prog.degree}) للعام ${year}هـ (${formatScore(overallAvg)} من 5)، `;
        text += overallAvg >= target
            ? `وهو أعلى من المعيار المستهدف (${formatScore(target)}).\n`
            : `وهو أقل من المعيار المستهدف (${formatScore(target)}) بفجوة (${formatScore(Math.abs(roundNumber(overallAvg - target)))}).\n`;

        if (strengths.length) {
            text += `\nمن جوانب القوة: ${strengths.map(s => `"${s.title}" بمتوسط (${formatScore(s.average)})`).join("، ")}.\n`;
        }
        if (weaknesses.length) {
            text += `\nومن جوانب التحسين: ${weaknesses.map(s => `"${s.title}" بمتوسط (${formatScore(s.average)}) بفجوة (${formatScore(Math.abs(s.gap))})`).join("، ")}.\n`;
        }
        text += `\nبلغت نسبة الاستطلاعات المحققة للمستهدف ${toArabicNumber(achieved)} من ${toArabicNumber(gapRows.length)} (${toArabicNumber(Math.round(achieved / gapRows.length * 100))}%).`;

        refs.gapsNarrativeText.textContent = text;
    }
}

function buildExportCardHtml(payload) {
    const filtersHtml = (payload.filters || []).length
        ? `<div class="pdf-export-badges">${payload.filters.map((item) => `<span class="pdf-export-badge">${escapeHtml(item)}</span>`).join("")}</div>`
        : "";
    const metricsHtml = (payload.metrics || []).length
        ? `<div class="pdf-export-metrics">${payload.metrics.map((item) => `
            <div class="pdf-export-metric">
                <span class="pdf-export-metric-label">${escapeHtml(item.label)}</span>
                <strong>${escapeHtml(item.value)}</strong>
            </div>
        `).join("")}</div>`
        : "";
    const noteHtml = payload.note ? `<div class="pdf-export-note">${escapeHtml(payload.note)}</div>` : "";
    const tableHead = payload.headers.map((header) => `<th scope="col">${escapeHtml(header)}</th>`).join("");
    const tableRows = payload.data.map((row) => `
        <tr>${row.map((value) => `<td>${escapeHtml(value)}</td>`).join("")}</tr>
    `).join("");

    return `
        <div class="pdf-export-card">
            <div class="pdf-export-head">
                <div>
                    <h1>${escapeHtml(payload.title)}</h1>
                    ${payload.subtitle ? `<p>${escapeHtml(payload.subtitle)}</p>` : ""}
                </div>
                <div class="pdf-export-date">${escapeHtml(new Date().toLocaleString("ar-SA"))}</div>
            </div>
            ${filtersHtml}
            ${metricsHtml}
            ${noteHtml}
            <div class="pdf-export-table-wrap">
                <table class="pdf-export-table">
                    <thead><tr>${tableHead}</tr></thead>
                    <tbody>${tableRows}</tbody>
                </table>
            </div>
        </div>
    `;
}

async function exportPayloadAsPdfCard(payload) {
    if (!payload) return;
    if (!payload.data || !payload.data.length) {
        window.alert("لا توجد نتائج قابلة للتصدير.");
        return;
    }

    const container = document.createElement("div");
    container.id = `pdfExportCard-${Date.now()}`;
    container.className = "pdf-export-shell";
    container.innerHTML = buildExportCardHtml(payload);
    document.body.appendChild(container);

    try {
        await exportSectionAsPdf(container.id, payload.filename);
    } finally {
        container.remove();
    }
}

function buildExploreExportPayload() {
    const records = getItemRecordsForExploreFilters();
    const surveyRows = aggregateSurveyRows(records);
    if (!records.length) {
        window.alert("لا توجد بيانات قابلة للتصدير في هذا النطاق.");
        return null;
    }

    return {
        filename: "استكشاف-البرامج.pdf",
        title: "بطاقة نتائج الاستكشاف",
        subtitle: "النتائج الحالية وفق المرشحات المختارة",
        filters: buildSingleFilterChips(state.exploreFilters).map((chip) => `${chip.label}: ${chip.value}`),
        metrics: [
            { label: "عدد السجلات", value: toArabicNumber(records.length) },
            { label: "عدد الاستطلاعات", value: toArabicNumber(surveyRows.length) },
            { label: "إجمالي الاستجابات", value: toArabicNumber(records.reduce((sum, row) => sum + row.responses, 0)) },
        ],
        headers: ["البرنامج", "الدرجة", "السنة", "المحور", "الاستطلاع", "العبارة", "الجنس", "عدد الاستجابات", "المتوسط"],
        data: records.map((row) => [
            row.programName,
            row.degree,
            `${row.year}هـ`,
            row.sectionLabel,
            row.surveyTitle,
            row.itemLabel,
            getSearchResultGenderLabel(row.gender),
            formatResponseCount(row.responses),
            formatScore(row.average),
        ]),
    };
}

function buildComparisonExportPayload() {
    const slotGroups = getActiveCompareSlots().map((slot) => ({
        slot,
        surveyRows: aggregateSurveyRows(getRecordsForCompareSlot(slot)),
    }));

    if (slotGroups.length < 2) {
        window.alert("يلزم تحديد خانتين على الأقل لتصدير المقارنة.");
        return null;
    }

    const comparisonRows = buildComparisonTableRows(slotGroups.map((group) => group.surveyRows));
    return {
        filename: "مقارنة-الاستطلاعات.pdf",
        title: "بطاقة نتائج المقارنة",
        subtitle: slotGroups.map((group) => formatCompareSlotLabel(group.slot)).join(" | "),
        filters: [
            `الجهة: ${state.compareFilters.stakeholder === "all" ? "الكل" : getStakeholderLabel(state.compareFilters.stakeholder)}`,
            `الموضوع: ${getTopicSelectionLabel(state.compareFilters)}`,
        ],
        metrics: [
            { label: "عدد المقارنات", value: toArabicNumber(slotGroups.length) },
            { label: "عدد الصفوف", value: toArabicNumber(comparisonRows.length) },
        ],
        headers: [
            "المحور",
            "الاستطلاع",
            ...slotGroups.flatMap((group) => [
                `${formatCompareSlotLabel(group.slot)} - المتوسط`,
                `${formatCompareSlotLabel(group.slot)} - عدد الاستجابات`,
            ]),
        ],
        data: comparisonRows.map((row) => [
            row.meta.sectionLabel,
            row.meta.title,
            ...row.values.flatMap((value) => [
                value ? formatScore(value.average) : "—",
                value ? formatResponseCount(value.respondentCount) : "—",
            ]),
        ]),
    };
}

function buildAnalysisExportPayload() {
    const records = getRecordsForAnalysis();
    const surveyRows = aggregateSurveyRows(records);
    if (!surveyRows.length) {
        window.alert("لا توجد بيانات قابلة للتصدير في هذا النطاق.");
        return null;
    }

    return {
        filename: "تحليل-الاستطلاعات.pdf",
        title: "بطاقة نتائج التحليل",
        subtitle: "ملخص الاستطلاعات ضمن نطاق التحليل الحالي",
        filters: [
            `البرامج: ${state.analysisFilters.programs.size ? `${toArabicNumber(state.analysisFilters.programs.size)} برنامج` : "لا توجد برامج محددة"}`,
            `السنة: ${state.analysisFilters.year === "all" ? "كل السنوات" : `${state.analysisFilters.year}هـ`}`,
            `الجهة: ${state.analysisFilters.stakeholder === "all" ? "الكل" : getStakeholderLabel(state.analysisFilters.stakeholder)}`,
            `الموضوع: ${getTopicSelectionLabel(state.analysisFilters)}`,
        ],
        metrics: [
            { label: "عدد الاستطلاعات", value: toArabicNumber(surveyRows.length) },
            { label: "إجمالي الاستجابات", value: toArabicNumber(surveyRows.reduce((sum, row) => sum + row.respondentCount, 0)) },
        ],
        headers: ["المحور", "الاستطلاع", "عدد الاستجابات", "المتوسط"],
        data: surveyRows.map((row) => [
            row.sectionLabel,
            row.title,
            formatResponseCount(row.respondentCount),
            formatScore(row.average),
        ]),
    };
}

function buildClosureExportPayload() {
    const payload = getClosureComparisonPayload();
    if (!payload.ready) {
        window.alert("اختر برنامجًا وسنتين أولًا.");
        return null;
    }
    if (!payload.qualifyingRows.length) {
        window.alert("لا توجد مؤشرات تحسن مؤثرة قابلة للتصدير في النطاق الحالي.");
        return null;
    }

    return {
        filename: `اغلاق-دائرة-الجودة-${sanitizeFileName(payload.program.name)}-${payload.fromYear}-${payload.toYear}.pdf`,
        title: "بطاقة إغلاق دائرة الجودة",
        subtitle: `${payload.program.name} - ${payload.program.degree} - ${payload.fromYearLabel} إلى ${payload.toYearLabel}`,
        filters: buildClosureFilterChips().map((chip) => `${chip.label}: ${chip.value}`),
        metrics: [
            { label: "المؤشرات القابلة للمقارنة", value: toArabicNumber(payload.comparableRows.length) },
            { label: "المؤشرات المؤهلة", value: toArabicNumber(payload.qualifyingRows.length) },
            { label: "أكبر تحسن", value: payload.topRow ? formatClosureImprovement(payload.topRow.deltaHundred) : "—" },
        ],
        note: `يعرض هذا التصدير المؤشرات التي تحسنت فقط، مع استبعاد أي تحسن أقل من ${formatClosureThreshold(payload.threshold)} وعدم إظهار حالات التراجع.`,
        headers: [
            "المحور",
            "نوع المطابقة",
            "المؤشر",
            `${payload.fromYearLabel} المتوسط`,
            `${payload.toYearLabel} المتوسط`,
            `${payload.fromYearLabel} النسبة`,
            `${payload.toYearLabel} النسبة`,
            "التحسن (من 100)",
            `${payload.fromYearLabel} الاستجابات`,
            `${payload.toYearLabel} الاستجابات`,
        ],
        data: payload.qualifyingRows.map((row) => [
            row.sectionLabel,
            getClosureMatchModeLabel(row.matchMode),
            `${row.title} — ${row.subtitle}`,
            formatScore(row.fromAverage),
            formatScore(row.toAverage),
            formatPercent(row.fromPercent),
            formatPercent(row.toPercent),
            formatClosureImprovement(row.deltaHundred),
            toArabicNumber(row.fromResponses),
            toArabicNumber(row.toResponses),
        ]),
    };
}

function buildClosureReportExportModel() {
    const payload = getClosureComparisonPayload();
    if (!payload.ready) {
        window.alert("اختر برنامجًا وسنتين أولًا.");
        return null;
    }
    if (!payload.qualifyingRows.length) {
        window.alert("لا توجد مؤشرات تحسن مؤثرة قابلة للتوثيق في النطاق الحالي.");
        return null;
    }

    const headers = [
        "البند أو الموضوع",
        payload.fromYearLabel,
        payload.toYearLabel,
        "التحسن (من 100)",
    ];
    const data = payload.qualifyingRows.map((row) => [
        buildClosureReportFocusExportText(row),
        buildClosureReportYearExportText(row.fromRow, row.fromAverage, row.fromPercent, row.fromResponses),
        buildClosureReportYearExportText(row.toRow, row.toAverage, row.toPercent, row.toResponses),
        formatClosureImprovement(row.deltaHundred),
    ]);
    const filename = `تقرير-اغلاق-دائرة-الجودة-${sanitizeFileName(payload.program.name)}-${payload.fromYear}-${payload.toYear}`;

    return { payload, headers, data, filename };
}

function buildClosureReportFocusExportText(row) {
    const topicLabel = row.fromRow.parentTitle || row.toRow.parentTitle || "—";
    const surveyLabel = row.fromRow.surveyTitle || row.toRow.surveyTitle || "—";
    return [
        `البند: ${row.title}`,
        `الموضوع: ${topicLabel}`,
        `الاستطلاع: ${surveyLabel}`,
        `نوع المطابقة: ${getClosureMatchModeLabel(row.matchMode)}`,
    ].join(" | ");
}

function buildClosureReportYearExportText(sourceRow, average, percent, responses) {
    const values = [formatScore(average)];

    if (state.closureFilters.reportShowResponses) {
        values.push(`عدد المقيمين: ${toArabicNumber(responses)}`);
    }
    if (state.closureFilters.reportShowStatement) {
        values.push(`العبارة: ${sourceRow.title || "—"}`);
    }
    if (state.closureFilters.reportShowSource) {
        values.push(`المصدر: ${CLOSURE_REPORT_SOURCE_LABEL}`);
    }

    return values.join(" | ");
}

function buildClosureReportPdfTitle(payload) {
    return `تقرير متابعة التحسن لبرنامج ${payload.program.name} (${payload.program.degree}) لسنتي ${payload.fromYear}هـ و${payload.toYear}هـ`;
}

function buildClosureReportPdfShell(payload) {
    const headerBadges = [
        `البرنامج: ${formatProgramLabel(payload.program)}`,
        `الفترة: ${payload.fromYearLabel} ← ${payload.toYearLabel}`,
        `الجنس: ${getGenderFilterLabel(state.closureFilters.gender)}`,
    ];
    const metrics = [
        { label: "المؤشرات المؤهلة", value: toArabicNumber(payload.qualifyingRows.length) },
        { label: "المؤشرات القابلة للمقارنة", value: toArabicNumber(payload.comparableRows.length) },
        { label: "أكبر تحسن", value: payload.topRow ? formatClosureImprovement(payload.topRow.deltaHundred) : "—" },
    ];
    const rowsHtml = payload.qualifyingRows.map((row) => `
        <tr data-closure-pdf-row>
            <td>${buildClosureReportFocusCellHtml(row)}</td>
            <td>${buildClosureReportYearCellHtml(row.fromRow, row.fromAverage, row.fromPercent, row.fromResponses)}</td>
            <td>${buildClosureReportYearCellHtml(row.toRow, row.toAverage, row.toPercent, row.toResponses)}</td>
            <td>${buildClosureReportImprovementCellHtml(row)}</td>
        </tr>
    `).join("");

    return `
        <div class="pdf-export-card" style="gap: 18px; padding: 28px;">
            <div data-closure-pdf-header>
                <div class="pdf-export-head">
                    <div>
                        <h1>${escapeHtml(buildClosureReportPdfTitle(payload))}</h1>
                        <p style="margin-top: 8px; font-size: 0.82rem;">المصدر: استطلاعات المنظومة الجامعية</p>
                    </div>
                </div>
                <div class="pdf-export-badges" style="margin-top: 8px;">${headerBadges.map((item) => `<span class="pdf-export-badge">${escapeHtml(item)}</span>`).join("")}</div>
                <div class="pdf-export-metrics" style="margin-top: 10px;">${metrics.map((item) => `
                    <div class="pdf-export-metric">
                        <span class="pdf-export-metric-label">${escapeHtml(item.label)}</span>
                        <strong>${escapeHtml(item.value)}</strong>
                    </div>
                `).join("")}</div>
            </div>

            <div class="table-wrap report-table-wrap" style="overflow: visible; border: 1px solid rgba(18, 57, 54, 0.1); border-radius: 18px;">
                <table class="data-table" style="width: 100%; table-layout: fixed; border-collapse: collapse; background: #ffffff;">
                    <thead data-closure-pdf-table-head>
                        <tr>
                            <th style="width: 32%;">البند أو الموضوع</th>
                            <th style="width: 24%;">${escapeHtml(payload.fromYearLabel)}</th>
                            <th style="width: 24%;">${escapeHtml(payload.toYearLabel)}</th>
                            <th style="width: 20%;">التحسن</th>
                        </tr>
                    </thead>
                    <tbody>${rowsHtml}</tbody>
                </table>
            </div>
        </div>
    `;
}

function waitForNextPaint() {
    return new Promise((resolve) => {
        window.requestAnimationFrame(() => {
            window.requestAnimationFrame(resolve);
        });
    });
}

async function captureElementAsPdfImage(element, targetWidth) {
    const canvas = await window.html2canvas(element, {
        scale: 2,
        backgroundColor: "#ffffff",
        useCORS: true,
        scrollX: 0,
        scrollY: 0,
    });

    return {
        imageData: canvas.toDataURL("image/png"),
        width: targetWidth,
        height: (canvas.height * targetWidth) / canvas.width,
    };
}

function buildTrendExportPayload() {
    const programId = state.trendFilters.program;
    if (!programId || programId === "all") {
        window.alert("اختر برنامجاً أولاً.");
        return null;
    }

    const program = getProgramById(programId);
    const years = sortYears(getAvailableYears(programId));
    const headers = ["المحور", ...years.map((year) => `${year}هـ`), "التغيّر"];
    const data = SECTION_META.map((section) => {
        const values = years.map((year) => {
            const records = ITEM_RECORDS.filter((record) =>
                record.programId === programId && record.year === year &&
                (state.trendFilters.stakeholder === "all" || record.stakeholder === state.trendFilters.stakeholder) &&
                matchesTopicFilter(record, state.trendFilters)
            );
            const surveys = aggregateSurveyRows(records).filter((row) => row.sectionId === section.id);
            if (!surveys.length) return null;
            return roundNumber(surveys.reduce((sum, row) => sum + row.average, 0) / surveys.length);
        });
        const delta = values.length >= 2 && values[0] != null && values[1] != null ? roundNumber(values[0] - values[1]) : null;
        return [section.label, ...values.map((value) => value != null ? formatScore(value) : "—"), delta != null ? formatScore(delta) : "—"];
    });

    return {
        filename: `التطور-الزمني-${sanitizeFileName(program.name)}.pdf`,
        title: "بطاقة نتائج التطور الزمني",
        subtitle: `${program.name} - ${program.degree}`,
        filters: [
            `الجهة: ${state.trendFilters.stakeholder === "all" ? "الكل" : getStakeholderLabel(state.trendFilters.stakeholder)}`,
            `الموضوع: ${getTopicSelectionLabel(state.trendFilters)}`,
        ],
        metrics: [
            { label: "عدد السنوات", value: toArabicNumber(years.length) },
            { label: "البرنامج", value: formatProgramLabel(program) },
        ],
        headers,
        data,
    };
}

function buildGapsExportPayload() {
    const { program, year, target } = state.gapsFilters;
    if (!program || program === "all") {
        window.alert("اختر برنامجاً أولاً.");
        return null;
    }

    const programInfo = getProgramById(program);
    const records = ITEM_RECORDS.filter((record) =>
        record.programId === program && record.year === year && record.stakeholder === "students" &&
        matchesTopicFilter(record, state.gapsFilters)
    );
    const surveys = aggregateSurveyRows(records);
    if (!surveys.length) {
        window.alert("لا توجد بيانات قابلة للتصدير.");
        return null;
    }

    const gapRows = surveys.map((row) => {
        const gap = roundNumber(row.average - target);
        const status = gap >= 0 ? "محقق" : gap >= -0.3 ? "قريب" : "دون المستهدف";
        return [row.sectionLabel, row.title, formatScore(row.average), formatScore(target), formatScore(gap), status];
    }).sort((first, second) => parseFloat(first[4]) - parseFloat(second[4]));

    return {
        filename: `تقرير-الفجوات-${sanitizeFileName(programInfo.name)}-${year}.pdf`,
        title: "بطاقة نتائج الفجوات",
        subtitle: `${programInfo.name} - ${programInfo.degree} - ${year}هـ`,
        filters: [`الموضوع: ${getTopicSelectionLabel(state.gapsFilters)}`],
        metrics: [
            { label: "عدد الاستطلاعات", value: toArabicNumber(surveys.length) },
            { label: "المستهدف", value: formatScore(target) },
        ],
        headers: ["المحور", "الاستطلاع", "المتوسط", "المستهدف", "الفجوة", "الحالة"],
        data: gapRows,
    };
}

function buildCustomExportPayload() {
    const selectedRows = getCustomRows().filter((row) => state.customSelected.has(row.uid));
    if (!selectedRows.length) {
        window.alert("لا توجد عناصر محددة للتصدير. حدّد بنوداً أولاً.");
        return null;
    }

    return {
        filename: "استطلاع-مخصص.pdf",
        title: "بطاقة نتائج الاستطلاع المخصص",
        subtitle: "العناصر المحددة حاليًا",
        filters: buildSingleFilterChips(state.customFilters).map((chip) => `${chip.label}: ${chip.value}`),
        metrics: [
            { label: "عدد العناصر", value: toArabicNumber(selectedRows.length) },
            { label: "إجمالي الاستجابات", value: toArabicNumber(selectedRows.reduce((sum, row) => sum + row.respondentCount, 0)) },
        ],
        headers: ["المحور", "الاستطلاع", "العبارة", "الموضوع", "البرنامج", "السنة", "عدد الاستجابات", "المتوسط"],
        data: selectedRows.map((row) => [
            row.sectionLabel,
            row.surveyTitle,
            row.title,
            row.topicLabel,
            row.programName,
            `${row.year}هـ`,
            formatResponseCount(row.respondentCount),
            formatScore(row.average),
        ]),
    };
}

function buildSearchExportPayload(limitToDisplay = false) {
    const payload = getSearchResultsPayload(refs.searchInput ? refs.searchInput.value : "");
    if (!payload.raw) {
        window.alert("اكتب عبارة بحث أولًا.");
        return null;
    }
    if (!payload.results.length) {
        window.alert("لا توجد نتائج مطابقة قابلة للتصدير.");
        return null;
    }

    const records = limitToDisplay ? payload.displayResults : payload.results;
    return {
        filename: `نتائج-البحث-${sanitizeFileName(payload.raw)}.pdf`,
        title: "بطاقة نتائج البحث",
        subtitle: `عبارة البحث: ${payload.raw}`,
        filters: [
            `نمط البحث: ${getSearchModeLabel(state.searchMode)}`,
            `عدد النتائج: ${toArabicNumber(payload.results.length)}`,
        ],
        metrics: [
            { label: "النتائج المصدرة", value: toArabicNumber(records.length) },
            { label: "إجمالي الاستجابات", value: toArabicNumber(records.reduce((sum, row) => sum + row.responses, 0)) },
        ],
        note: payload.truncated && limitToDisplay ? `اقتصر تصدير PDF على أول ${toArabicNumber(records.length)} نتيجة من أصل ${toArabicNumber(payload.results.length)} نتيجة.` : "",
        headers: ["البرنامج", "الدرجة", "السنة", "الاستطلاع", "الموضوع", "العبارة", "الجنس", "عدد الاستجابات", "المتوسط", "المحكات"],
        data: records.map((record) => {
            const visibleCriteria = getVisibleSearchCriteria(record, payload.queryInfo);
            const criterionCodes = Array.from(new Set(visibleCriteria.map((entry) => entry.criterionCode))).filter(Boolean);
            return [
                record.programName,
                record.degree,
                `${record.year}هـ`,
                record.surveyTitle,
                record.topicLabel || "—",
                record.itemLabel,
                getSearchResultGenderLabel(record.gender),
                formatResponseCount(record.responses),
                formatScore(record.average),
                criterionCodes.length ? criterionCodes.join("، ") : "غير مرتبط بمحك",
            ];
        }),
    };
}

function exportSearchResults(type) {
    const payload = buildSearchExportPayload(type === "pdf");
    if (!payload) return;

    if (type === "csv") {
        exportCsv(payload.filename.replace(/\.pdf$/i, ".csv"), payload.headers, payload.data);
        return;
    }

    exportPayloadAsPdfCard(payload);
}

function exportExploreResultsPdf() {
    return exportPayloadAsPdfCard(buildExploreExportPayload());
}

function exportComparisonResultsPdf() {
    return exportPayloadAsPdfCard(buildComparisonExportPayload());
}

function exportAnalysisResultsPdf() {
    return exportPayloadAsPdfCard(buildAnalysisExportPayload());
}

function exportClosureResultsPdf() {
    return exportPayloadAsPdfCard(buildClosureExportPayload());
}

async function exportClosureReportPdf() {
    const reportModel = buildClosureReportExportModel();
    if (!reportModel) return;
    if (!window.html2canvas || !window.jspdf || !window.jspdf.jsPDF) {
        window.alert("أدوات PDF غير متاحة حاليًا.");
        return;
    }

    const exportShell = document.createElement("div");
    exportShell.className = "pdf-export-shell";
    exportShell.style.width = "1280px";
    exportShell.innerHTML = buildClosureReportPdfShell(reportModel.payload);
    document.body.appendChild(exportShell);

    try {
        await waitForNextPaint();

        const headerElement = exportShell.querySelector("[data-closure-pdf-header]");
        const tableHeadElement = exportShell.querySelector("[data-closure-pdf-table-head]");
        const rowElements = Array.from(exportShell.querySelectorAll("[data-closure-pdf-row]"));
        if (!headerElement || !tableHeadElement || !rowElements.length) {
            window.alert("تعذر تجهيز التقرير للطباعة.");
            return;
        }

        const pdf = new window.jspdf.jsPDF("l", "pt", "a4");
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const marginX = 26;
        const marginTop = 26;
        const marginBottom = 22;
        const sectionGap = 10;
        const contentWidth = pageWidth - (marginX * 2);
        let cursorY = marginTop;

        const headerImage = await captureElementAsPdfImage(headerElement, contentWidth);
        pdf.addImage(headerImage.imageData, "PNG", marginX, cursorY, contentWidth, headerImage.height, undefined, "FAST");
        cursorY += headerImage.height + sectionGap;

        const tableHeadImage = await captureElementAsPdfImage(tableHeadElement, contentWidth);
        if (cursorY + tableHeadImage.height > pageHeight - marginBottom) {
            pdf.addPage();
            cursorY = marginTop;
        }
        pdf.addImage(tableHeadImage.imageData, "PNG", marginX, cursorY, contentWidth, tableHeadImage.height, undefined, "FAST");
        cursorY += tableHeadImage.height;

        for (const rowElement of rowElements) {
            const rowImage = await captureElementAsPdfImage(rowElement, contentWidth);
            const remainingHeight = pageHeight - marginBottom - cursorY;

            if (rowImage.height > remainingHeight) {
                pdf.addPage();
                cursorY = marginTop;
                pdf.addImage(tableHeadImage.imageData, "PNG", marginX, cursorY, contentWidth, tableHeadImage.height, undefined, "FAST");
                cursorY += tableHeadImage.height;
            }

            let renderWidth = contentWidth;
            let renderHeight = rowImage.height;
            const maxRowHeight = pageHeight - marginBottom - cursorY;
            if (renderHeight > maxRowHeight && maxRowHeight > 0) {
                const scale = maxRowHeight / renderHeight;
                renderWidth *= scale;
                renderHeight *= scale;
            }

            pdf.addImage(rowImage.imageData, "PNG", marginX, cursorY, renderWidth, renderHeight, undefined, "FAST");
            cursorY += renderHeight;
        }

        pdf.save(`${reportModel.filename}.pdf`);
    } finally {
        exportShell.remove();
    }
}

function exportTrendResultsPdf() {
    return exportPayloadAsPdfCard(buildTrendExportPayload());
}

function exportGapsResultsPdf() {
    return exportPayloadAsPdfCard(buildGapsExportPayload());
}

function exportCustomResultsPdf() {
    return exportPayloadAsPdfCard(buildCustomExportPayload());
}

function exportExploreData(type) {
    const records = getItemRecordsForExploreFilters();
    const surveyRows = aggregateSurveyRows(records);
    const headers = ["البرنامج", "السنة", "المحور", "الاستطلاع", "العبارة", "الجنس", "عدد الاستجابات", "المتوسط"];
    const data = records.map((row) => [
        row.programName,
        row.year,
        row.sectionLabel,
        row.surveyTitle,
        row.itemLabel,
        getGenderLabel(row.gender),
        row.responses,
        formatScore(row.average),
    ]);

    if (type === "csv") {
        exportCsv("استكشاف-البرامج.csv", headers, data);
        return;
    }

    exportExcel("استكشاف-البرامج.xlsx", "الاستكشاف", headers, data);
}

/* exportProgramData removed — merged into exportExploreData */

function exportComparisonData(type) {
    const slotGroups = getActiveCompareSlots().map((slot) => ({
        slot,
        surveyRows: aggregateSurveyRows(getRecordsForCompareSlot(slot)),
    }));

    if (slotGroups.length < 2) {
        window.alert("يلزم تحديد خانتين على الأقل لتصدير المقارنة.");
        return;
    }

    const comparisonRows = buildComparisonTableRows(slotGroups.map((group) => group.surveyRows));
    const headers = [
        "المحور",
        "الاستطلاع",
        ...slotGroups.flatMap((group) => [
            `${formatCompareSlotLabel(group.slot)} - المتوسط`,
            `${formatCompareSlotLabel(group.slot)} - عدد الاستجابات`,
        ]),
    ];
    const data = comparisonRows.map((row) => [
        row.meta.sectionLabel,
        row.meta.title,
        ...row.values.flatMap((value) => [
            value ? formatScore(value.average) : "",
            value ? value.respondentCount : "",
        ]),
    ]);

    if (type === "csv") {
        exportCsv("مقارنة-الاستطلاعات.csv", headers, data);
        return;
    }

    exportExcel("مقارنة-الاستطلاعات.xlsx", "المقارنة", headers, data);
}

function exportAnalysisData(type) {
    const records = ITEM_RECORDS.filter((record) => {
        if (!state.analysisFilters.programs.has(record.programId)) return false;
        if (state.analysisFilters.year !== "all" && record.year !== state.analysisFilters.year) return false;
        if (state.analysisFilters.stakeholder !== "all" && record.stakeholder !== state.analysisFilters.stakeholder) return false;
        return true;
    });
    const surveyRows = aggregateSurveyRows(records);

    if (!surveyRows.length) {
        window.alert("لا توجد بيانات قابلة للتصدير في هذا النطاق.");
        return;
    }

    const headers = ["المحور", "الاستطلاع", "عدد الاستجابات", "المتوسط"];
    const data = surveyRows.map((row) => [
        row.sectionLabel,
        row.title,
        row.respondentCount,
        formatScore(row.average),
    ]);

    if (type === "csv") {
        exportCsv("تحليل-الاستطلاعات.csv", headers, data);
        return;
    }
    exportExcel("تحليل-الاستطلاعات.xlsx", "التحليل", headers, data);
}

function exportClosureData(type) {
    const payload = getClosureComparisonPayload();
    if (!payload.ready) {
        window.alert("اختر برنامجًا وسنتين أولًا.");
        return;
    }
    if (!payload.qualifyingRows.length) {
        window.alert("لا توجد مؤشرات تحسن مؤثرة قابلة للتصدير في النطاق الحالي.");
        return;
    }

    const headers = [
        "المحور",
        "نوع المطابقة",
        "المؤشر",
        "التفصيل",
        `${payload.fromYearLabel} المتوسط`,
        `${payload.toYearLabel} المتوسط`,
        `${payload.fromYearLabel} النسبة`,
        `${payload.toYearLabel} النسبة`,
        "التحسن (من 100)",
        `${payload.fromYearLabel} الاستجابات`,
        `${payload.toYearLabel} الاستجابات`,
    ];
    const data = payload.qualifyingRows.map((row) => [
        row.sectionLabel,
        getClosureMatchModeLabel(row.matchMode),
        row.title,
        row.subtitle,
        formatScore(row.fromAverage),
        formatScore(row.toAverage),
        formatPercent(row.fromPercent),
        formatPercent(row.toPercent),
        formatClosureImprovement(row.deltaHundred),
        row.fromResponses,
        row.toResponses,
    ]);
    const filename = `اغلاق-دائرة-الجودة-${sanitizeFileName(payload.program.name)}-${payload.fromYear}-${payload.toYear}`;

    if (type === "csv") {
        exportCsv(`${filename}.csv`, headers, data);
        return;
    }

    exportExcel(`${filename}.xlsx`, "إغلاق دائرة الجودة", headers, data);
}

function exportClosureReportData(type) {
    const reportModel = buildClosureReportExportModel();
    if (!reportModel) return;

    if (type === "csv") {
        exportCsv(`${reportModel.filename}.csv`, reportModel.headers, reportModel.data);
        return;
    }

    exportExcel(`${reportModel.filename}.xlsx`, "تقرير الإغلاق", reportModel.headers, reportModel.data);
}

function exportTrendData(type) {
    const programId = state.trendFilters.program;
    if (!programId || programId === "all") {
        window.alert("اختر برنامجاً أولاً.");
        return;
    }
    const prog = getProgramById(programId);
    const years = sortYears(getAvailableYears(programId));
    const stakeholder = state.trendFilters.stakeholder;

    const headers = ["المحور", ...years.map(y => `${y}هـ`), "التغيّر"];
    const data = SECTION_META.map(sec => {
        const vals = years.map(year => {
            const records = ITEM_RECORDS.filter(r =>
                r.programId === programId && r.year === year &&
                (stakeholder === "all" || r.stakeholder === stakeholder) &&
                matchesTopicFilter(r, state.trendFilters)
            );
            const surveys = aggregateSurveyRows(records).filter(s => s.sectionId === sec.id);
            if (!surveys.length) return null;
            return roundNumber(surveys.reduce((s, r) => s + r.average, 0) / surveys.length);
        });
        const delta = vals.length >= 2 && vals[0] != null && vals[1] != null ? roundNumber(vals[0] - vals[1]) : null;
        return [sec.label, ...vals.map(v => v != null ? formatScore(v) : "—"), delta != null ? formatScore(delta) : "—"];
    });

    const filename = `التطور-الزمني-${sanitizeFileName(prog.name)}`;
    if (type === "csv") {
        exportCsv(`${filename}.csv`, headers, data);
        return;
    }
    exportExcel(`${filename}.xlsx`, "التطور الزمني", headers, data);
}

function exportGapsData(type) {
    const { program, year, target } = state.gapsFilters;
    if (!program || program === "all") {
        window.alert("اختر برنامجاً أولاً.");
        return;
    }
    const prog = getProgramById(program);
    const records = ITEM_RECORDS.filter(r =>
        r.programId === program && r.year === year && r.stakeholder === "students" &&
        matchesTopicFilter(r, state.gapsFilters)
    );
    const surveys = aggregateSurveyRows(records);

    if (!surveys.length) {
        window.alert("لا توجد بيانات قابلة للتصدير.");
        return;
    }

    const headers = ["المحور", "الاستطلاع", "المتوسط", "المستهدف", "الفجوة", "الحالة"];
    const data = surveys.map(s => {
        const gap = roundNumber(s.average - target);
        const status = gap >= 0 ? "محقق" : gap >= -0.3 ? "قريب" : "دون المستهدف";
        return [s.sectionLabel, s.title, formatScore(s.average), formatScore(target), formatScore(gap), status];
    }).sort((a, b) => parseFloat(a[4]) - parseFloat(b[4]));

    const filename = `تقرير-الفجوات-${sanitizeFileName(prog.name)}-${year}`;
    if (type === "csv") {
        exportCsv(`${filename}.csv`, headers, data);
        return;
    }
    exportExcel(`${filename}.xlsx`, "تقرير الفجوات", headers, data);
}

function exportCustomData(type) {
    const allRows = getCustomRows();
    const selectedRows = allRows.filter((row) => state.customSelected.has(row.uid));
    if (!selectedRows.length) {
        window.alert("لا توجد عناصر محددة للتصدير. حدّد بنوداً أولاً.");
        return;
    }
    const headers = ["المحور", "الاستطلاع", "العبارة", "الموضوع", "البرنامج", "السنة", "عدد الاستجابات", "المتوسط"];
    const data = selectedRows.map((row) => [
        row.sectionLabel, row.surveyTitle, row.title, row.topicLabel,
        row.programName, `${row.year}هـ`, row.respondentCount, formatScore(row.average),
    ]);
    const filename = "استطلاع-مخصص";
    if (type === "csv") { exportCsv(`${filename}.csv`, headers, data); return; }
    exportExcel(`${filename}.xlsx`, "استطلاع مخصص", headers, data);
}

function exportExcel(filename, sheetName, headers, data) {
    if (!window.XLSX) {
        window.alert("مكتبة Excel غير متاحة حاليًا.");
        return;
    }

    const workbook = window.XLSX.utils.book_new();
    const worksheet = window.XLSX.utils.aoa_to_sheet([headers, ...data]);
    window.XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    window.XLSX.writeFile(workbook, filename);
}

function exportCsv(filename, headers, data) {
    const csvRows = [headers, ...data];
    const csv = "\uFEFF" + csvRows.map((row) => row.map(csvEscape).join(",")).join("\r\n");
    downloadBlob(filename, new Blob([csv], { type: "text/csv;charset=utf-8;" }));
}

async function exportSectionAsPdf(elementId, filename) {
    if (!window.html2canvas || !window.jspdf || !window.jspdf.jsPDF) {
        window.alert("أدوات PDF غير متاحة حاليًا.");
        return;
    }

    const element = document.getElementById(elementId);
    if (!element) return;

    const canvas = await window.html2canvas(element, {
        scale: 2,
        backgroundColor: "#ffffff",
        useCORS: true,
        scrollX: 0,
        scrollY: -window.scrollY,
    });

    const imageData = canvas.toDataURL("image/png");
    const pdf = new window.jspdf.jsPDF("l", "pt", "a4");
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = 18;
    const imageWidth = pageWidth - (margin * 2);
    const imageHeight = (canvas.height * imageWidth) / canvas.width;
    let heightLeft = imageHeight;
    let position = margin;

    pdf.addImage(imageData, "PNG", margin, position, imageWidth, imageHeight);
    heightLeft -= (pageHeight - (margin * 2));

    while (heightLeft > 0) {
        position = heightLeft - imageHeight + margin;
        pdf.addPage();
        pdf.addImage(imageData, "PNG", margin, position, imageWidth, imageHeight);
        heightLeft -= (pageHeight - (margin * 2));
    }

    pdf.save(filename);
}

function getProgramsWithData() {
    return PROGRAMS.filter((program) => program.id !== "all" && getAvailableYears(program.id).length);
}

function getAvailableYears(programId) {
    if (programId === "all") return ALL_AVAILABLE_YEARS;
    const years = AVAILABLE_PROGRAM_YEARS[programId];
    if (years && years.length) return sortYears(years);
    return [];
}

function getProgramById(programId) {
    return PROGRAMS.find((program) => program.id === programId) || PROGRAMS[0];
}

function getSectionLabel(sectionId) {
    return SECTION_META.find((section) => section.id === sectionId)?.label || sectionId;
}

function getSectionOrder(sectionId) {
    return SECTION_META.find((section) => section.id === sectionId)?.order || 99;
}

function getStakeholderLabel(stakeholder) {
    return STAKEHOLDER_LABELS[stakeholder] || stakeholder;
}

function getGenderLabel(gender) {
    const map = {
        male: "ذكر",
        m: "ذكر",
        ذكر: "ذكر",
        female: "إناث",
        f: "إناث",
        "أنثى": "إناث",
        "إناث": "إناث",
    };
    return map[gender] || gender;
}

function getSearchResultGenderLabel(gender) {
    const label = getGenderLabel(gender);
    if (label === "إناث") return "أنثى";
    return label || "غير محدد";
}

function getSelfStudyEntries(programId, year, surveyTitle, itemLabel) {
    const key = buildSelfStudyItemKey(programId, year, surveyTitle, itemLabel);
    return SELF_STUDY_LINKS[key] || [];
}

function buildSelfStudyItemKey(programId, year, surveyTitle, itemLabel) {
    return [programId, year, surveyTitle, itemLabel].map((value) => normalizeText(value)).join("||");
}

function buildSelfStudyCriterionValue(entry) {
    return `criterion:${entry.criterionCode}||${normalizeText(entry.criterionText)}`;
}

function buildSelfStudySideValue(entry) {
    return `side:${entry.criterionCode}||${normalizeText(entry.criterionText)}||${normalizeText(entry.supportedSide)}`;
}

function buildSelfStudyPhraseValue(entry, phraseText) {
    return `phrase:${entry.criterionCode}||${normalizeText(entry.criterionText)}||${normalizeText(entry.supportedSide)}||${normalizeText(phraseText)}`;
}

function buildSingleFilterChips(filters) {
    return [
        { label: "البرنامج", value: filters.program === "all" ? "الكل" : formatProgramLabel(getProgramById(filters.program)) },
        { label: "السنة", value: filters.year === "all" ? "الكل" : `${filters.year}هـ` },
        { label: "الجهة", value: filters.stakeholder === "all" ? "الكل" : getStakeholderLabel(filters.stakeholder) },
        { label: "الموضوع", value: getSubjectLabel(filters.subject) },
        { label: "الجنس", value: getGenderFilterLabel(filters.gender) },
    ];
}

/* buildProgramFilterChips removed — no longer used */

function getGenderFilterLabel(gender) {
    if (!AVAILABLE_GENDERS.length) return "غير متاح";
    if (gender === "all") return "الكل";
    return getGenderLabel(gender);
}

function getTopicModeLabel(topicMode) {
    return topicMode === "selfstudy" ? "محكات الدراسة الذاتية" : "موضوعات عامة";
}

function getSubjectLabel(subjectValue) {
    if (subjectValue === "all") return "كل الموضوعات";
    if (subjectValue.startsWith("section:")) {
        return getSectionLabel(subjectValue.replace("section:", ""));
    }
    if (subjectValue.startsWith("survey:")) {
        return subjectValue.replace("survey:", "");
    }
    if (subjectValue.startsWith("topic:")) {
        return subjectValue.replace("topic:", "");
    }
    return subjectValue;
}

function getSelfStudyLabel(value) {
    if (value === "all") return "كل محكات الدراسة الذاتية";
    if (value === "unlinked") return "غير مرتبط بمحك";

    if (value.startsWith("criterion:")) {
        const [criterionCode, criterionText] = value.replace("criterion:", "").split("||");
        return `المحك ${criterionCode} | ${criterionText}`;
    }

    if (value.startsWith("side:")) {
        const [criterionCode, criterionText, supportedSide] = value.replace("side:", "").split("||");
        return `المحك ${criterionCode} | ${criterionText} | الجانب: ${supportedSide}`;
    }

    if (value.startsWith("phrase:")) {
        const [criterionCode, criterionText, supportedSide, phraseText] = value.replace("phrase:", "").split("||");
        return `المحك ${criterionCode} | ${criterionText} | الجانب: ${supportedSide} | العبارة: ${phraseText}`;
    }

    return value;
}

function getTopicSelectionLabel(filters) {
    if ((filters.topicMode || DEFAULT_TOPIC_MODE) === "selfstudy") {
        return getSelfStudyLabel(filters.selfStudyTarget || "all");
    }
    return getSubjectLabel(filters.subject || "all");
}

function buildSingleFilterScopeLabel(filters) {
    return [
        filters.program === "all" ? "جميع البرامج" : formatProgramLabel(getProgramById(filters.program)),
        filters.year === "all" ? "كل السنوات" : `${filters.year}هـ`,
        filters.stakeholder === "all" ? "كل الجهات" : getStakeholderLabel(filters.stakeholder),
        getSubjectLabel(filters.subject),
        getGenderFilterLabel(filters.gender),
    ].join(" · ");
}

/* buildProgramScopeLabel removed — no longer used */

function buildAnalysisScopeLabel() {
    return [
        state.analysisFilters.programs.size ? `${toArabicNumber(state.analysisFilters.programs.size)} برنامج` : "لا توجد برامج محددة",
        state.analysisFilters.year === "all" ? "كل السنوات" : `${state.analysisFilters.year}هـ`,
        state.analysisFilters.stakeholder === "all" ? "كل الجهات" : getStakeholderLabel(state.analysisFilters.stakeholder),
        getTopicModeLabel(state.analysisFilters.topicMode),
        getTopicSelectionLabel(state.analysisFilters),
    ].join(" · ");
}

function formatProgramLabel(program) {
    return `${program.name} - ${program.degree}`;
}

function formatCompareSlotLabel(slot) {
    const program = getProgramById(slot.program);
    return `${program.name} - ${slot.year}هـ`;
}

/* buildProgramFileName removed — no longer used */

function getComparisonKey(row) {
    return [row.sectionId, normalizeText(row.title)].join("||");
}

function averageScore(rows) {
    if (!rows.length) return null;
    return roundNumber(rows.reduce((sum, row) => sum + Number(row.average || 0), 0) / rows.length);
}

function getExtremeRow(rows, direction) {
    if (!rows.length) return null;
    const sorted = [...rows].sort((first, second) => direction === "max" ? second.average - first.average : first.average - second.average);
    return sorted[0];
}

function groupAverageRows(rows, keyFn, labelFn) {
    const map = new Map();
    rows.forEach((row) => {
        const key = keyFn(row);
        if (!map.has(key)) {
            map.set(key, { key, label: labelFn(row), sum: 0, count: 0 });
        }
        const item = map.get(key);
        item.sum += row.average;
        item.count += 1;
    });

    return Array.from(map.values())
        .map((item) => ({
            key: item.key,
            label: item.label,
            average: roundNumber(item.sum / item.count),
            count: item.count,
        }))
        .sort((first, second) => second.average - first.average);
}

function toneForScore(score) {
    if (score == null) return "info";
    if (score >= 4.25) return "good";
    if (score >= 3.75) return "info";
    if (score >= 3.25) return "warning";
    return "danger";
}

function scoreStatusLabel(score) {
    if (score == null) return "غير متاح";
    if (score >= 4.25) return "قوي";
    if (score >= 3.75) return "جيد";
    if (score >= 3.25) return "متوسط";
    return "بحاجة متابعة";
}

function scoreStatusClass(score) {
    if (score == null) return "status-medium";
    if (score >= 4.25) return "status-excellent";
    if (score >= 3.75) return "status-good";
    if (score >= 3.25) return "status-medium";
    return "status-watch";
}

function renderScorePill(score) {
    return `<span class="score-pill">${escapeHtml(formatScore(score))}</span>`;
}

function renderStatusPill(score) {
    return `<span class="status-pill ${scoreStatusClass(score)}">${escapeHtml(scoreStatusLabel(score))}</span>`;
}

function renderComparisonCell(value) {
    if (!value) return "<td>—</td>";
    return `
        <td>
            <span class="cell-title">${escapeHtml(formatScore(value.average))}</span>
            <span class="cell-subtitle">${formatResponseCount(value.respondentCount)}</span>
        </td>
    `;
}

function formatResponseCount(value) {
    return `${toArabicNumber(value)} استجابة`;
}

function buildProgramRowSubtitle(row) {
    if (row.rowKind === "موضوع") {
        return `<span class="cell-subtitle">${toArabicNumber(row.itemCount)} عبارة · ${escapeHtml(row.parentTitle)}</span>`;
    }
    if (row.rowKind === "عبارة") {
        const itemNumber = getItemNumberLabel(row) !== "—"
            ? `رقم ${getItemNumberLabel(row)}`
            : "عبارة تفصيلية";
        return `<span class="cell-subtitle">${escapeHtml(row.parentTitle)} · ${escapeHtml(itemNumber)}</span>`;
    }
    return "";
}

function getItemNumberLabel(row) {
    const rawValue = String(row?.itemNumber || "").trim();
    if (rawValue) return rawValue;
    if (Number.isFinite(Number(row?.itemOrder)) && Number(row.itemOrder) < 900000) {
        return toArabicNumber(Number(row.itemOrder));
    }
    return "—";
}

function chartHasData(config) {
    if (!config || !config.data || !Array.isArray(config.data.datasets)) return false;
    return config.data.datasets.some((dataset) => Array.isArray(dataset.data) && dataset.data.some((value) => value != null && value !== 0));
}

function truncateLabel(value, maxLength) {
    const text = String(value || "");
    if (text.length <= maxLength) return text;
    return `${text.slice(0, maxLength - 1)}…`;
}

function sortYears(years) {
    return [...years].sort((first, second) => Number(second) - Number(first));
}

function collectUnique(items, keyFn) {
    const seen = new Set();
    const unique = [];
    items.forEach((item) => {
        if (item == null || item === "") return;
        const key = keyFn(item);
        if (seen.has(key)) return;
        seen.add(key);
        unique.push(item);
    });
    return unique;
}

function countDistinctSurveys() {
    const keys = new Set(ITEM_RECORDS.map((record) => `${record.programId}||${record.year}||${normalizeText(record.surveyTitle)}`));
    return keys.size;
}

function parseItemOrder(itemNumber, fallbackIndex) {
    const digits = String(itemNumber || "").replace(/[^\d]/g, "");
    if (digits) return Number(digits);
    return 900000 + fallbackIndex;
}

function getAggregateKey(record, level) {
    if (level === "survey") {
        return [
            record.programId,
            record.year,
            record.stakeholder,
            record.sectionId,
            normalizeText(record.surveyTitle),
        ].join("||");
    }
    if (level === "topic") {
        return [
            record.programId,
            record.year,
            record.stakeholder,
            record.sectionId,
            normalizeText(record.surveyTitle),
            normalizeText(record.topicLabel),
        ].join("||");
    }
    return [
        record.programId,
        record.year,
        record.stakeholder,
        record.sectionId,
        normalizeText(record.surveyTitle),
        normalizeText(record.topicLabel),
        record.itemNumber,
        normalizeText(record.itemLabel),
    ].join("||");
}

function normalizeText(value) {
    return String(value || "").trim().replace(/\s+/g, " ");
}

function normalizeForSearch(value) {
    return normalizeText(value)
        .toLowerCase()
        .replace(/[\u064b-\u065f\u0670\u0640]/g, "")
        .replace(/[أإآٱ]/g, "ا")
        .replace(/ى/g, "ي")
        .replace(/ة/g, "ه")
        .replace(/ؤ/g, "و")
        .replace(/ئ/g, "ي")
        .replace(/[٠-٩]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0x0660 + 48))
        .replace(/[^0-9\u0621-\u063a\u0641-\u064a\s]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function normalizeCriterionCodeForSearch(value) {
    return String(value || "")
        .replace(/[٠-٩]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0x0660 + 48))
        .replace(/[^0-9]+/g, ".")
        .replace(/\.+/g, ".")
        .replace(/^\.|\.$/g, "");
}

function escapeRegExp(value) {
    return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function analyzeSearchQuery(raw) {
    const normalizedRaw = normalizeForSearch(raw);
    const info = {
        raw,
        normalizedRaw,
        criterionCode: "",
        years: new Set(),
        degrees: new Set(),
        textQuery: "",
    };

    const criterionMatch = raw.match(/([0-9٠-٩]+(?:\s*[.\-\u2013\u2014٫ـ]\s*[0-9٠-٩]+){2})/);
    if (criterionMatch) {
        info.criterionCode = normalizeCriterionCodeForSearch(criterionMatch[1]);
    }

    normalizedRaw.replace(/(^|\s)(14\d{2})(?:\s*ه)?(?=\s|$)/g, (_, __, year) => {
        info.years.add(year);
        return _;
    });

    let textQuery = normalizedRaw;
    DEGREE_SEARCH_RULES.forEach((rule) => {
        if (rule.pattern.test(normalizedRaw)) {
            info.degrees.add(rule.value);
            textQuery = textQuery.replace(rule.pattern, " ");
        }
    });
    if (info.criterionCode) {
        const criterionToken = info.criterionCode.replace(/\./g, " ");
        textQuery = textQuery.replace(new RegExp(`(^|\\s)${escapeRegExp(criterionToken)}(?=\\s|$)`, "g"), " ");
    }
    info.years.forEach((year) => {
        textQuery = textQuery.replace(new RegExp(`(^|\\s)${year}(?:\\s*ه)?(?=\\s|$)`, "g"), " ");
    });
    textQuery = textQuery
        .replace(/(^|\s)محك(?=\s|$)/g, " ")
        .replace(/\s+/g, " ")
        .trim();

    info.textQuery = textQuery;
    return info;
}

const DEGREE_SEARCH_RULES = [
    { pattern: /(^|\s)بكالوريوس انتساب(?=\s|$)/, value: "بكالوريوس" },
    { pattern: /(^|\s)بكالوريوس(?=\s|$)/, value: "بكالوريوس" },
    { pattern: /(^|\s)الماجستير(?=\s|$)/, value: "الماجستير" },
    { pattern: /(^|\s)ماجستير(?=\s|$)/, value: "الماجستير" },
    { pattern: /(^|\s)الدكتوراه(?=\s|$)/, value: "دكتوراه" },
    { pattern: /(^|\s)دكتوراه(?=\s|$)/, value: "دكتوراه" },
    { pattern: /(^|\s)دراسات عليا(?=\s|$)/, value: "دراسات عليا" },
];

function degreeMatchesQuery(recordDegree, degreeFilters) {
    if (!degreeFilters || !degreeFilters.size) return true;
    const degree = normalizeDegreeForSearch(recordDegree);
    if (!degree) return false;
    return Array.from(degreeFilters).some((filterValue) => {
        if (filterValue === "دراسات عليا") {
            return degree === "الماجستير" || degree === "دكتوراه";
        }
        return degree === filterValue;
    });
}

function normalizeDegreeForSearch(value) {
    const degree = normalizeText(value);
    if (degree === "ماجستير" || degree === "الماجستير") return "الماجستير";
    if (degree === "دكتوراه" || degree === "الدكتوراه" || degree === "دكتوراة") return "دكتوراه";
    if (degree === "بكالوريوس" || degree === "بكالوريوس انتساب") return "بكالوريوس";
    return degree;
}

function stripCommonArabicPrefix(token) {
    return token.startsWith("ال") && token.length > 2 ? token.slice(2) : token;
}

function stripCommonArabicSuffix(token) {
    const suffixes = ["يات", "يون", "تين", "تين", "ية", "يه", "ات", "ان", "ين", "ون", "تي", "ها", "هم", "هن", "كم", "نا", "يه", "ية", "ي", "ه", "ة"];
    for (const suffix of suffixes) {
        if (token.endsWith(suffix) && token.length - suffix.length >= 3) {
            return token.slice(0, -suffix.length);
        }
    }
    return token;
}

function simplifyArabicToken(token) {
    return stripCommonArabicSuffix(stripCommonArabicPrefix(token));
}

function sharedPrefixLength(first, second) {
    const limit = Math.min(first.length, second.length);
    let count = 0;
    while (count < limit && first[count] === second[count]) {
        count += 1;
    }
    return count;
}

function levenshteinDistance(first, second) {
    if (first === second) return 0;
    if (!first.length) return second.length;
    if (!second.length) return first.length;

    const previous = Array.from({ length: second.length + 1 }, (_, index) => index);
    const current = new Array(second.length + 1);

    for (let row = 1; row <= first.length; row += 1) {
        current[0] = row;
        for (let column = 1; column <= second.length; column += 1) {
            const cost = first[row - 1] === second[column - 1] ? 0 : 1;
            current[column] = Math.min(
                current[column - 1] + 1,
                previous[column] + 1,
                previous[column - 1] + cost,
            );
        }
        for (let column = 0; column <= second.length; column += 1) {
            previous[column] = current[column];
        }
    }

    return previous[second.length];
}

function isCloseTokenMatch(queryToken, textToken) {
    const candidates = [queryToken, simplifyArabicToken(queryToken)].filter(Boolean);
    const targets = [textToken, simplifyArabicToken(textToken)].filter(Boolean);

    return candidates.some((candidate) =>
        targets.some((target) => {
            if (!candidate || !target) return false;
            if (candidate === target) return true;
            if (candidate.length >= 3 && (target.includes(candidate) || candidate.includes(target))) return true;

            const prefixLength = sharedPrefixLength(candidate, target);
            if (prefixLength >= Math.min(4, candidate.length, target.length)) return true;

            if (Math.abs(candidate.length - target.length) > 2) return false;
            const threshold = Math.min(candidate.length, target.length) >= 6 ? 2 : 1;
            return levenshteinDistance(candidate, target) <= threshold;
        }),
    );
}

function searchMatchesPrecisely(query, text) {
    const normalizedQuery = normalizeForSearch(query);
    const normalizedText = normalizeForSearch(text);

    if (!normalizedQuery) return true;
    if (!normalizedText) return false;

    const queryTokens = normalizedQuery.split(" ").filter(Boolean);
    const textTokens = normalizedText.split(" ").filter(Boolean);
    if (!queryTokens.length) return true;
    if (!textTokens.length) return false;

    const exactTextTokens = new Set();
    textTokens.forEach((token) => {
        exactTextTokens.add(token);
        const strippedToken = stripCommonArabicPrefix(token);
        if (strippedToken && strippedToken !== token) {
            exactTextTokens.add(strippedToken);
        }
    });

    return queryTokens.every((token) => {
        if (exactTextTokens.has(token)) return true;
        const strippedToken = stripCommonArabicPrefix(token);
        return Boolean(strippedToken && exactTextTokens.has(strippedToken));
    });
}

function searchMatches(query, text) {
    const normalizedQuery = normalizeForSearch(query);
    const normalizedText = normalizeForSearch(text);

    if (!normalizedQuery) return true;
    if (!normalizedText) return false;
    if (normalizedText.includes(normalizedQuery)) return true;

    const queryTokens = normalizedQuery.split(" ").filter(Boolean);
    const textTokens = normalizedText.split(" ").filter(Boolean);
    if (!queryTokens.length) return true;
    if (queryTokens.every((token) => normalizedText.includes(token))) return true;
    if (!textTokens.length) return false;
    if (queryTokens.every((token) => textTokens.some((textToken) => isCloseTokenMatch(token, textToken)))) return true;

    return false;
}

function roundNumber(value) {
    return Math.round(Number(value) * 100) / 100;
}

function formatScore(value) {
    if (value == null || Number.isNaN(Number(value))) return "—";
    return Number(value).toLocaleString("ar-SA", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    });
}

function formatPercent(value) {
    if (value == null || Number.isNaN(Number(value))) return "—";
    return `${Number(value).toLocaleString("ar-SA", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    })}%`;
}

function formatDeltaPoints(value) {
    if (value == null || Number.isNaN(Number(value))) return "—";
    const numeric = Number(value);
    const sign = numeric > 0 ? "+" : "";
    return `${sign}${numeric.toLocaleString("ar-SA", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    })}`;
}

function formatClosureImprovement(value) {
    if (value == null || Number.isNaN(Number(value))) return "—";
    return `${formatDeltaPoints(value)} من 100`;
}

function formatClosureThreshold(value) {
    if (value == null || Number.isNaN(Number(value))) return "—";
    return `${Number(value).toLocaleString("ar-SA", {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2,
    })} من 100`;
}

function toArabicNumber(value) {
    return Number(value).toLocaleString("ar-SA");
}

function csvEscape(value) {
    const text = String(value ?? "");
    if (text.includes(",") || text.includes("\"") || text.includes("\n")) {
        return `"${text.replace(/"/g, "\"\"")}"`;
    }
    return text;
}

function sanitizeFileName(value) {
    return String(value).replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, "-");
}

function downloadBlob(filename, blob) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}
