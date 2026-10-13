/*
 * مؤشر «تقويم الطلاب لجودة خبرات التعلم في البرنامج» (KPI-P-01 / KPI-PG-01).
 *
 * هذه الوحدة هي المصدر الوحيد لحساب المؤشر. يستعملها سكربت
 * scripts/build_program_experience_kpi.cjs لتوليد الملف الجاهز
 * data/kpi/program-experience.json الذي تقرؤه لوحة مؤشرات الأداء.
 *
 * مراتب الاستخراج (تُجرَّب بالترتيب، وتُعتمد أول مرتبة متاحة):
 *   1. بند «التقويم العام» في استبانة تقويم البرنامج — المطابق لتعريف المركز الوطني
 *      («متوسط التقييم العام لجودة خبرة التعلم في البرنامج على مقياس خماسي»).
 *   2. متوسط بنود استبانة تقويم البرنامج نفسها بعد استبعاد المحاور غير المصنفة
 *      والبنود الدخيلة من استبانات أخرى (مثل التدريب الميداني).
 *   3. محور «التعلم والتعليم» في الاستبانة المكافئة (استبانة تقويم خبرة الطالب).
 *
 * التعرف على استبانة تقويم البرنامج لا يعتمد على العنوان وحده: تُقبل إذا طابق
 * عنوانها إحدى الصيغ المعروفة، أو إذا غطت بنودها نسبة كافية من بصمة الاستبانة المرجعية.
 */
(function (root, factory) {
    "use strict";
    const api = factory();
    if (typeof module === "object" && module.exports) module.exports = api;
    else root.ProgramExperienceKpi = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
    "use strict";

    const RULES = Object.freeze({
        version: 1,
        stakeholder: "students",
        programEvaluation: {
            titles: [
                "استبانة تقويم برنامج",
                "استبانة تقويم البرنامج",
                "استطلاع تقويم البرنامج",
                "تقويم البرنامج",
                "تقييم البرنامج",
                "تقويم جودة البرنامج",
                "تقييم جودة البرنامج",
                "تقويم خبرة البرنامج",
                "تقييم خبرة البرنامج",
                "استطلاع الخبرة التعليمية في البرنامج",
            ],
            minFingerprintCoverage: 0.6,
            fingerprint: [
                "أتيح لي الإرشاد الأكاديمي والمهني المناسب خلال فترة دراستي بالبرنامج",
                "كانت هيئة التدريس متاحة للإرشاد والمشورة عندما كنت أحتاج للتحدث إليهم",
                "شجعتني هيئة التدريس بالقسم على أن أقدم أفضل ما عندي",
                "قدمت هيئة التدريس بالقسم تغذية راجعة على عملي",
                "لدى هيئة التدريس بالقسم معرفة كبيرة بمحتوى المقررات التي يدرسونها",
                "كانت هيئة التدريس بالقسم متحمسة للعمل",
                "اهتمت هيئة التدريس بمدى تقدم طلبتهم",
                "كانت المواد الدراسية المساعدة بالمقررات حديثة ومفيدة",
                "كانت مصادر المكتبة مناسبة ومتاحة كلما كنت أحتاج لها",
                "اتسمت تجهيزات القاعات للمحاضرات والمعامل والدروس بالجودة",
                "كانت تجهيزات الحاسب كافية لاحتياجاتي",
                "توفرت التجهيزات المناسبة للأنشطة اللامنهجية بما في ذلك التجهيزات الخاصة بالرياضة والترفيه",
                "هناك مرافق مناسبة لأداء الشعائر الدينية",
                "كانت برامج التدريب الميداني أو سنة الامتياز فعالة في تطوير مهاراتي",
                "ما تعلمته في هذا البرنامج القسم سيكون مهما لمستقبلي",
                "لقد ساعدني البرنامج في تطوير الاهتمام الكافي لدي للسعي في الاستمرار في تحديث معلوماتي حسبما يستجد في مجال دراستي",
                "لقد طور البرنامج قدرتي على استقصاء وحل المشكلات الجديدة",
                "لقد طور البرنامج قدرتي على العمل بفاعلية مع المجموعات",
                "لقد حسن البرنامج مهاراتي في الاتصال",
                "لقد ساعدني البرنامج في تطوير مهاراتي الأساسية في استخدام التقنية لدراسة القضايا والتعبير عن النتائج",
                "لقد طورت المعارف والمهارات اللازمة لمهنتي التي اخترتها",
                "أشعر بالرضا بشكل عام عن مستوى جودة خبرتي التعليمية في الجامعة",
            ],
        },
        overallItem: {
            topics: ["التقويم العام", "التقييم العام", "الرضا العام"],
            phrases: ["بشكل عام", "عموما", "بوجه عام"],
            context: ["خبرتي التعليمية", "خبرة التعلم", "خبرات التعلم", "جودة التعلم", "البرنامج"],
            exclude: ["المقرر", "حياتي الطلابية", "عضو هيئة التدريس"],
        },
        excludedTopics: ["بنود متنوعة", "بنود غير مصنفة", "غير مصنف"],
        foreignItems: ["التدريب الميداني", "اماكن التدريب", "المشرف الميداني", "اثناء التدريب", "قبل التدريب", "الميثاق الاخلاقي للمهنة"],
        equivalent: {
            titles: ["استبانة تقويم خبرة الطالب/ الطالبة", "استبانة تقويم خبرة الطالب", "استبانة خبرة الطالب"],
            topics: ["التعلم والتعليم"],
        },
        sharedScopeMinPrograms: 3,
    });

    const METHOD_LABELS = Object.freeze({
        1: "بند التقويم العام في استبانة تقويم البرنامج",
        2: "متوسط بنود استبانة تقويم البرنامج بعد التنقية",
        3: "محور التعلم والتعليم في استبانة تقويم خبرة الطالب",
    });

    function normalize(value) {
        return String(value ?? "")
            .normalize("NFC")
            .replace(/[ً-ٰٟـ]/g, "")
            .replace(/[أإآٱ]/g, "ا")
            .replace(/ى/g, "ي")
            .replace(/ة/g, "ه")
            .replace(/[‌-‏‪-‮]/g, "")
            .replace(/^[\s\d٠-٩\-–.)(]+/, "")
            .replace(/[^\p{L}\p{N}\s]/gu, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    const compiled = (rules) => {
        const list = (values) => values.map(normalize);
        return {
            rules,
            pesTitles: new Set(list(rules.programEvaluation.titles)),
            pesPrint: new Set(list(rules.programEvaluation.fingerprint)),
            overallTopics: list(rules.overallItem.topics),
            overallPhrases: list(rules.overallItem.phrases),
            overallContext: list(rules.overallItem.context),
            overallExclude: list(rules.overallItem.exclude),
            excludedTopics: new Set(list(rules.excludedTopics)),
            foreignItems: list(rules.foreignItems),
            equivalentTitles: new Set(list(rules.equivalent.titles)),
            equivalentTopics: new Set(list(rules.equivalent.topics)),
        };
    };

    const includesAny = (text, needles) => needles.some((needle) => needle && text.includes(needle));

    function genderOf(value) {
        const name = normalize(value);
        if (["ذكر", "ذكور", "طالب", "طلاب", "male"].includes(name)) return "male";
        if (["اناث", "انثي", "طالبه", "طالبات", "female"].includes(name)) return "female";
        return null;
    }

    function measurement(entry) {
        const count = Number(entry.responses), total = Number(entry.scoreTotal);
        if (Number.isFinite(count) && count > 0 && Number.isFinite(total) && total >= count && total <= count * 5) {
            return { weight: count, numerator: total, respondents: count, mode: "responses" };
        }
        // المتوسطات الجاهزة بلا أعداد: measurementCount يعدّ تكرار البرنامج بأسماء بديلة
        // في المصدر، لا عدد مجيبين؛ فلكل ملاحظة وزن واحد.
        const average = Number(entry.average);
        if (entry.average != null && entry.average !== "" && Number.isFinite(average) && average >= 1 && average <= 5) {
            return { weight: 1, numerator: average, respondents: null, mode: "averages" };
        }
        return null;
    }

    function itemCells(item) {
        return (item.genders || []).filter((entry) => genderOf(entry.gender)).map(measurement).filter(Boolean);
    }

    function aggregate(items) {
        let weight = 0, numerator = 0, responsesOnly = true, maxRespondents = 0, cells = 0;
        const signature = [];
        items.forEach((item) => {
            const values = itemCells(item);
            let itemRespondents = 0;
            values.forEach((cell) => {
                weight += cell.weight;
                numerator += cell.numerator;
                cells += 1;
                if (cell.mode !== "responses") responsesOnly = false;
                else itemRespondents += cell.respondents;
            });
            maxRespondents = Math.max(maxRespondents, itemRespondents);
            signature.push([normalize(item.label), values.map((cell) => Number((cell.numerator / cell.weight).toFixed(6)))]);
        });
        if (!cells || !weight) return null;
        return {
            value: numerator / weight,
            respondents: responsesOnly ? maxRespondents : null,
            weighting: responsesOnly ? "responses" : "averages",
            signature: JSON.stringify(signature),
        };
    }

    function identifyProgramEvaluation(survey, c) {
        if (survey.stakeholder && survey.stakeholder !== c.rules.stakeholder) return null;
        const labels = (survey.topics || []).flatMap((topic) => (topic.items || []).map((item) => normalize(item.label)));
        const matched = new Set(labels.filter((label) => c.pesPrint.has(label)));
        const coverage = c.pesPrint.size ? matched.size / c.pesPrint.size : 0;
        if (coverage >= c.rules.programEvaluation.minFingerprintCoverage) return { identifiedBy: "fingerprint", coverage };
        if (c.pesTitles.has(normalize(survey.title))) return { identifiedBy: "title", coverage };
        return null;
    }

    function overallRank(topicLabel, itemLabel, c) {
        const topic = normalize(topicLabel), label = normalize(itemLabel);
        if (includesAny(label, c.overallExclude)) return 0;
        const phrased = includesAny(label, c.overallPhrases);
        if (phrased && includesAny(topic, c.overallTopics)) return 1;
        if (phrased && includesAny(label, c.overallContext)) return 2;
        return 0;
    }

    function coreItems(survey, c) {
        return (survey.topics || [])
            .filter((topic) => !c.excludedTopics.has(normalize(topic.label)))
            .flatMap((topic) => topic.items || [])
            .filter((item) => {
                const label = normalize(item.label);
                return c.pesPrint.has(label) || !includesAny(label, c.foreignItems);
            });
    }

    function resolveDataset(dataset, c) {
        const surveys = dataset.surveys || [];
        const candidates = surveys
            .map((survey) => ({ survey, id: identifyProgramEvaluation(survey, c) }))
            .filter((entry) => entry.id)
            .sort((a, b) => b.id.coverage - a.id.coverage);

        for (const { survey, id } of candidates) {
            const ranked = [];
            (survey.topics || []).forEach((topic) => (topic.items || []).forEach((item) => {
                const rank = overallRank(topic.label, item.label, c);
                if (rank) ranked.push({ rank, item });
            }));
            ranked.sort((a, b) => a.rank - b.rank);
            for (const { item } of ranked) {
                const result = aggregate([item]);
                if (!result) continue;
                const composite = aggregate(coreItems(survey, c));
                return { tier: 1, survey: survey.title, item: item.label, identifiedBy: id.identifiedBy, result, composite };
            }
        }
        for (const { survey, id } of candidates) {
            const items = coreItems(survey, c);
            const result = aggregate(items);
            if (result) return { tier: 2, survey: survey.title, item: null, identifiedBy: id.identifiedBy, itemCount: items.length, result, composite: result };
        }
        for (const survey of surveys) {
            if (survey.stakeholder && survey.stakeholder !== c.rules.stakeholder) continue;
            if (!c.equivalentTitles.has(normalize(survey.title))) continue;
            const items = (survey.topics || []).filter((topic) => c.equivalentTopics.has(normalize(topic.label))).flatMap((topic) => topic.items || []);
            const result = aggregate(items);
            if (result) return { tier: 3, survey: survey.title, item: null, identifiedBy: "title", itemCount: items.length, result, composite: result };
        }
        return null;
    }

    const round = (value, digits) => (value == null ? null : Number(value.toFixed(digits)));

    function resolveAll(payload, rules = RULES) {
        const c = compiled(rules);
        const records = {};
        const signatures = {};
        Object.keys((payload && payload.extractedData) || {}).sort().forEach((key) => {
            const [programId, year] = key.split("::");
            const resolved = resolveDataset(payload.extractedData[key], c);
            if (!resolved) return;
            const { result, composite } = resolved;
            records[key] = {
                programId,
                year,
                value: round(result.value, 2),
                tier: resolved.tier,
                method: METHOD_LABELS[resolved.tier],
                survey: resolved.survey,
                item: resolved.item,
                identifiedBy: resolved.identifiedBy,
                respondents: result.respondents,
                weighting: result.weighting,
                itemsMean: composite ? round(composite.value, 2) : null,
                scope: "program",
            };
            const group = `${year}::${resolved.tier}::${normalize(resolved.survey)}`;
            (signatures[group] = signatures[group] || []).push({ key, signature: result.signature });
        });
        // إذا كانت القيمة مأخوذة من بيانات متطابقة حرفيًا في عدة برامج للسنة نفسها
        // فهي بيانات مجمعة على مستوى أعلى من البرنامج؛ تُوثَّق دون أن تُحجب.
        Object.values(signatures).forEach((entries) => {
            if (entries.length < rules.sharedScopeMinPrograms) return;
            if (new Set(entries.map((entry) => entry.signature)).size !== 1) return;
            entries.forEach((entry) => { records[entry.key].scope = "shared"; });
        });
        return records;
    }

    function buildKpiFile(payload, rules = RULES) {
        return {
            kpi: "تقويم الطلاب لجودة خبرات التعلم في البرنامج",
            codes: ["KPI-P-01", "KPI-PG-01"],
            scale: "1-5",
            rulesVersion: rules.version,
            sourceFile: "js/surveys-data.js",
            sourceGeneratedAt: (payload && payload.generatedAt) || null,
            methods: METHOD_LABELS,
            fields: {
                value: "قيمة المؤشر مقربة إلى منزلتين",
                tier: "مرتبة الاستخراج المعتمدة (1 أدقها)",
                respondents: "عدد المجيبين عن البند المعتمد، أو null إذا كان المصدر متوسطات بلا أعداد",
                weighting: "responses: مجموع الدرجات ÷ مجموع الإجابات، averages: متوسط المتوسطات الجاهزة بوزن متساوٍ",
                itemsMean: "متوسط بنود الاستبانة بعد التنقية، للتحليل لا للمؤشر",
                scope: "program: بيانات خاصة بالبرنامج، shared: قيم متطابقة بين البرامج في المصدر",
            },
            records: resolveAll(payload, rules),
        };
    }

    return { RULES, METHOD_LABELS, normalize, resolveAll, buildKpiFile };
});
