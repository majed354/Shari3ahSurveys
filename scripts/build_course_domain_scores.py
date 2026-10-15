"""Build per-course learning-domain and quality scores from the public v2 API.

Every item is classified only through data/course-evaluations/item-classification.json
(semantic map, not keywords). An item absent from the map stops the build.
"""
import argparse
import difflib
import hashlib
import json
from pathlib import Path
import re
import unicodedata

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "data/course-evaluations/v2/student-course-evaluations.json"
CLASSIFICATION = ROOT / "data/course-evaluations/item-classification.json"
OUTPUT = ROOT / "data/course-evaluations/course-domain-scores.json"
# Course codes shown on the dashboard (resolved there from the programmes' course lists).
SITE_DATA = ROOT / "js/course-evaluations-data.js"
LINKABLE_CODE_STATUSES = {"matched", "cross-program"}
# Human-approved code links where the survey name differs and the dashboard gives no code.
COURSE_LINKS = ROOT / "data/course-evaluations/course-links.json"
# overall_satisfaction is the "overall" subset of quality; it is scored separately.
SCORED = ("knowledge", "skills", "values", "overall_satisfaction", "quality", "other", "unclassified", "review")
COHORTS = ("total", "male", "female")


def normalized(text):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(text or ""))).strip()


_LETTERS = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ة": "ه", "ى": "ي", "ئ": "ي", "ؤ": "و", "ء": "",
                          **{chr(0x0660 + n): str(n) for n in range(10)}, **{chr(0x06F0 + n): str(n) for n in range(10)}})
# Words whose addition, removal or change alters meaning; any difference touching them stops the match.
NEGATIONS = {"غير", "لم", "لا", "ليس", "ليست", "لن", "عدم", "ما", "دون", "بدون"}
FILLERS = {"هذا", "هذه"}


def tokens(text):
    """Spelling-insensitive words: same rule as the automation pipeline (program_survey_domains)."""
    text = unicodedata.normalize("NFKC", str(text or "")).replace("\u200f", "").replace("\u200e", "")
    text = re.sub(r"[\u064b-\u065f\u0670\u0640]", "", text).translate(_LETTERS).casefold()
    return re.sub(r"[\W_]+", " ", text).split()


def match_key(text):
    return "".join(tokens(text))


def _bare(word):
    return word[2:] if word.startswith("ال") and len(word) > 4 else word


def same_word(a, b, equivalents=frozenset()):
    """Same word: equal, differing only by the article ال, or a curated spelling-equivalent pair.

    No open letter-level tolerance: Arabic words one letter apart often differ in meaning
    (مهم/مهمل، الجيد/الجديد، بعد/بعدم). Numbers and negations never count as variants.
    """
    if a == b:
        return True
    if any(ch.isdigit() for ch in a + b) or a in NEGATIONS or b in NEGATIONS:
        return False
    return _bare(a) == _bare(b) or frozenset((a, b)) in equivalents or frozenset((_bare(a), _bare(b))) in equivalents


def is_variant(words, item_words, equivalents=frozenset()):
    """Accept only spelling differences and the fillers هذا/هذه; any other word change is a new meaning."""
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, words, item_words, autojunk=False).get_opcodes():
        if op == "equal":
            continue
        changed = words[i1:i2] + item_words[j1:j2]
        if NEGATIONS & set(changed):
            return False
        if op == "replace":
            if i2 - i1 != j2 - j1 or not all(same_word(a, b, equivalents) for a, b in zip(words[i1:i2], item_words[j1:j2])):
                return False
        elif not set(changed) <= FILLERS:
            return False
    return True


def equivalents_of(classification):
    pairs = set()
    for a, b in classification["matching"].get("spellingEquivalents", []):
        ta, tb = tokens(a), tokens(b)
        if len(ta) != 1 or len(tb) != 1 or ta == tb:
            raise ValueError(f"Spelling equivalent must be two different single words: {a} / {b}")
        pairs.add(frozenset((ta[0], tb[0])))
    return frozenset(pairs)


def resolve(statement, keyed, equivalents=frozenset()):
    """Return (item, method, nearest) for a statement; item None means a human must classify it."""
    key = match_key(statement)
    if not key:
        return None, "unmatched", None
    if key in keyed:
        return keyed[key], "normalized", None
    words = tokens(statement)
    candidates = {id(item): item for item in keyed.values() if is_variant(words, tokens(item["statement"]), equivalents)}
    if len(candidates) == 1:
        return next(iter(candidates.values())), "variant", None
    nearest = max(keyed, key=lambda other: difflib.SequenceMatcher(None, key, other).ratio())
    return None, "unmatched", keyed[nearest]


def load_map(classification):
    if classification.get("schemaVersion") != "course-item-classification-v1":
        raise ValueError("Expected course-item-classification-v1")
    allowed = {c["id"] for c in classification["categories"]}
    by_id, by_statement, keyed = {}, {}, {}
    for item in classification["items"]:
        if item["category"] not in allowed:
            raise ValueError(f"Unknown category: {item['category']}")
        if item["overall"] and item["category"] != "quality":
            raise ValueError("overall=true is only valid for quality items")
        if (item["category"] == "review") != (item["status"] == "needs_review"):
            raise ValueError("review category and needs_review status must agree")
        learning = item["category"] in ("knowledge", "skills", "values")
        if learning != (item.get("measurementBasis") in ("direct_gain", "proxy")):
            raise ValueError("learning-domain items need measurementBasis direct_gain or proxy; others none")
        key = normalized(item["statement"])
        if key in by_statement:
            raise ValueError(f"Duplicate statement in classification: {key}")
        by_statement[key] = item
        for text in [item["statement"], *item.get("aliases", [])]:
            # Human-approved wording variants (aliases) match exactly from then on.
            if match_key(text) in keyed:
                raise ValueError(f"Two classified statements share one normalized form: {text}")
            keyed[match_key(text)] = item
        for item_id in item["itemIds"]:
            by_id[item_id] = item
    return by_id, by_statement, keyed


def categories_of(item):
    if item["overall"]:
        return ("quality", "overall_satisfaction")
    return (item["category"],)


def item_counts(item):
    """Accepted counts of one item: complete counts, or the accepted part of a partial item.

    A partial item keeps the counts of evaluations that were not quarantined
    (knownResponseCounts); the conflicting reports themselves never enter.
    """
    if item["status"] == "partial":
        counts = item.get("knownResponseCounts") or [0, 0, 0, 0, 0]
        return sum(i * n for i, n in enumerate(counts, 1)), sum(counts), 1
    return item.get("scoreNumerator") or 0, item.get("responses") or 0, 0


def load_site_codes(path=SITE_DATA):
    """(year, semester|all, programme name, course name) -> course code, as on the dashboard."""
    if not path.exists():
        return {}
    text = path.read_text(encoding="utf-8")
    data = json.loads(text[text.index("{"):text.rindex("}") + 1])
    codes = {}
    for row in data.get("courseRecords", []):
        if row.get("courseCode") and row.get("codeStatus") in LINKABLE_CODE_STATUSES:
            # Spelling-insensitive like the dashboard's own grouping (اصول/أصول، الفقة/الفقه).
            key = (str(row["year"]), str(row["semester"]), row["program"], match_key(row["courseName"]))
            codes[key] = (row["courseCode"], row["codeStatus"])
    return codes


def load_manual_links(path=COURSE_LINKS):
    """(programId, course name) -> course code, from approved entries only."""
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("schemaVersion") != "course-links-v1":
        raise ValueError("Expected course-links-v1")
    links = {}
    for link in data["links"]:
        if link.get("status") not in ("approved", "proposed"):
            raise ValueError(f"Unknown link status: {link.get('status')}")
        if link["status"] != "approved":
            continue
        for name in link["surveyNames"]:
            key = (link["programId"], normalized(name))
            if links.get(key, link["courseCode"]) != link["courseCode"]:
                raise ValueError(f"Survey name linked to two codes: {name}")
            links[key] = link["courseCode"]
    return links


def course_code(codes, year, semester, programme, course):
    """Semester row first, then the year row; the same dashboard rule for every record."""
    for period in (semester, "all"):
        found = codes.get((year, period, programme, match_key(course)))
        if found:
            return found
    return None, None


def build_payload(source, classification, classification_bytes=None, site_codes=None, manual_links=None):
    if source.get("schemaVersion") != "course-student-evaluations-v2":
        raise ValueError("Expected course-student-evaluations-v2")
    by_id, by_statement, keyed = load_map(classification)
    equivalents = equivalents_of(classification)
    item_category = {}
    proxy_items = set()
    matches, unmatched = [], []
    for item_id, definition in source["itemDefinitions"].items():
        statement = definition.get("statement")
        if not statement:
            item_category[item_id] = None
            continue
        mapped = by_id.get(item_id)
        if mapped is not None and match_key(mapped["statement"]) != match_key(statement):
            mapped = None  # a reused id with new wording is matched by its text
        if mapped is None:
            mapped, method, nearest = resolve(statement, keyed, equivalents)
            if mapped is None:
                hint = f" — الأقرب «{nearest['statement']}» ({nearest['category']})؛ إن كان صيغة له فأضفه إلى aliases" if nearest else ""
                unmatched.append(f"{item_id}: «{statement}»{hint}")
                continue
            if normalized(mapped["statement"]) != normalized(statement):
                matches.append({"itemId": item_id, "statement": statement, "matchedStatement": mapped["statement"],
                                "method": method, "category": mapped["category"]})
        item_category[item_id] = categories_of(mapped)
        if mapped.get("measurementBasis") == "proxy":
            proxy_items.add(item_id)
    if unmatched:
        raise ValueError("بنود جديدة تحتاج تصنيفًا في item-classification.json قبل النشر:\n" + "\n".join(unmatched))

    degree_names = {"البكالوريوس": "بكالوريوس", "الماجستير": "ماجستير", "الدكتوراه": "دكتوراه"}
    programme_names = {p["id"]: f"{degree_names[p['degree']]} {p['specialty']}" for p in source["programs"]}
    site_codes = load_site_codes() if site_codes is None else site_codes
    manual_links = load_manual_links() if manual_links is None else manual_links
    records = []
    for record in source["records"]:
        cohorts = {}
        for cohort_name in COHORTS:
            cohort = record["cohorts"].get(cohort_name)
            if cohort is None:
                continue
            # [scoreNumerator, responses, itemCount, partialItems, unscoredItems]
            totals = {}
            unnamed = 0
            proxy = set()
            for item in cohort["items"]:
                categories = item_category[item["itemId"]]
                if categories is None or item.get("statementSource") == "missing_in_source":
                    unnamed += 1
                    continue
                numerator, responses, partial = item_counts(item)
                if item["itemId"] in proxy_items:
                    proxy.update(categories)
                for category in categories:
                    value = totals.setdefault(category, [0, 0, 0, 0, 0])
                    value[0] += numerator
                    value[1] += responses
                    value[2] += 1
                    value[3] += partial
                    value[4] += responses == 0
            cohorts[cohort_name] = {"categories": {c: totals[c] for c in SCORED if c in totals},
                                    # Categories measured (wholly or partly) through proxy items.
                                    **({"proxyCategories": [c for c in SCORED if c in proxy]} if proxy else {}),
                                    **({"unnamedItems": unnamed} if unnamed else {})}
        manual = manual_links.get((record["programId"], normalized(record["courseName"])))
        code, code_status = (manual, "approved_link") if manual else course_code(
            site_codes, record["academicYear"], record["semester"], programme_names[record["programId"]], record["courseName"])
        records.append({"year": record["academicYear"], "semester": record["semester"],
                        "programId": record["programId"], "courseName": record["courseName"],
                        **({"courseCode": code, "courseCodeSource": code_status if manual else "dashboard_" + code_status} if code else {}),
                        "recordId": record["id"], "cohorts": cohorts})
    records.sort(key=lambda r: (r["year"], r["semester"], r["programId"], r["courseName"]))
    labels = {c["id"]: c["label"] for c in classification["categories"]}
    return {"schemaVersion": "course-domain-scores-v1", "source": "course-student-evaluations-v2",
            "classificationSha256": hashlib.sha256(classification_bytes).hexdigest() if classification_bytes else None,
            "tuple": ["scoreNumerator", "responses", "itemCount", "partialItems", "unscoredItems"],
            "categories": [{"id": c, "label": "الرضا العام عن المقرر" if c == "overall_satisfaction" else labels[c]}
                           for c in SCORED],
            "scoring": classification["scoring"],
            # Wording variants accepted automatically; exact statements are not listed.
            "itemMatches": sorted(matches, key=lambda m: m["itemId"]),
            "programs": [{**p, "name": f"{degree_names[p['degree']]} {p['specialty']}"} for p in source["programs"]],
            "years": sorted({r["academicYear"] for r in source["records"]}, reverse=True),
            "records": records}


def category_score(values):
    """Score of one category from its accepted responses; None only without any."""
    if not values:
        return None
    numerator, responses, items = values[:3]
    return numerator / responses if items and responses > 0 else None


def is_partial(values):
    """True when part of the sample was excluded (quarantined report or missing item)."""
    return bool(values) and bool(values[3] or values[4])


def serialize(payload):
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"), allow_nan=False) + "\n"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=SOURCE)
    parser.add_argument("--classification", type=Path, default=CLASSIFICATION)
    parser.add_argument("--output", type=Path, default=OUTPUT)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    classification_bytes = args.classification.read_bytes()
    output = serialize(build_payload(json.loads(args.source.read_text()), json.loads(classification_bytes), classification_bytes))
    if args.check:
        if not args.output.exists() or args.output.read_text() != output:
            raise SystemExit("Course domain scores are out of date")
        print("Course domain scores match public v2 source and classification")
    else:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output)
        print(f"Wrote {args.output} ({len(output.encode()):,} bytes)")


if __name__ == "__main__":
    main()
