# engine/classifier.py
"""
Column Classifier – Stage 2 of the extraction pipeline.
Given a FileStructure and a company header dictionary (tokenised keyword sets),
produces a MappingResult for each table, strictly respecting the company dictionary.
No content‑based fallback; ambiguous/missing cases are flagged for user review.
"""

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple

import polars as pl

from engine.inspector import TableInfo


# ---------- Data classes ----------
@dataclass
class ColumnMapping:
    category: str              # "surname", "firstname", "fullname", "sex", "dob", "policynum"
    column_index: int          # absolute column index in the DataFrame
    header_raw: str            # original header text (for display)
    confidence: float          # 0.0 – 1.0 (higher = more certain)
    ambiguous_candidates: List[int] = field(default_factory=list)


@dataclass
class MappingResult:
    table: TableInfo
    mappings: List[ColumnMapping]    # final selected mappings
    missing_categories: List[str]   # e.g. ["fullname", "dob"]
    requires_user_input: bool       # True if any required field is missing or ambiguous


# ---------- Configuration ----------
REQUIRED_CATEGORIES = ["fullname", "dob", "sex", "policynum"]  # at least one name type is required
NAME_CATEGORIES = {"fullname", "surname", "firstname"}
STOP_WORDS = {"of", "the", "a", "an", "in", "on", "at", "to", "for", "and", "or", "is", "are", "was", "were", "be"}

# Mapping from header dictionary keys to our internal categories
DICT_TO_CATEGORY = {
    "name": "fullname",
    "firstlastname": "firstlast",   # temporary; will be split into surname/firstname
    "sex": "sex",
    "dob": "dob",
    "policynum": "policynum",
}


# ==================== Helpers ====================
def _clean_string(text: str) -> str:
    """Return lowercased alpha‑only version of text."""
    if not text:
        return ""
    return re.sub(r'[^a-z]', '', text.lower())


def _tokenise(text: str) -> Set[str]:
    """Split cleaned text into words (tokens), excluding stop words and short tokens."""
    cleaned = _clean_string(text)
    tokens = set()
    for word in cleaned.split():
        if len(word) >= 2 and word not in STOP_WORDS:
            tokens.add(word)
    return tokens


def _fuzzy_ratio(a: str, b: str) -> float:
    """Quick similarity ratio without importing difflib if possible."""
    import difflib
    return difflib.SequenceMatcher(None, a, b).ratio()


def _content_score_for_category(col_series: pl.Series, category: str) -> float:
    """
    Estimate how well the data in a column fits the expected pattern for a category.
    Returns 0.0 – 1.0.
    """
    non_null = col_series.drop_nulls()
    if len(non_null) == 0:
        return 0.0
    sample = non_null.head(100).cast(pl.Utf8).to_list()
    total = len(sample)

    if category in ("fullname", "surname", "firstname"):
        # Name: expect mostly alphabetic, possibly with hyphens/apostrophes, spaces
        name_pattern = re.compile(r"^[A-Za-zÀ-ÿ'\-\. ]{2,}$")
        matches = sum(1 for v in sample if name_pattern.match(str(v)))
        return matches / total if total else 0.0

    if category == "sex":
        sex_indicators = {'m', 'f', 'male', 'female'}
        matches = sum(1 for v in sample if str(v).strip().lower() in sex_indicators)
        return matches / total if total else 0.0

    if category == "dob":
        # Lazy import to avoid circular dependency
        from utils.cleaners import parse_date_to_mmddyyyy
        matches = sum(1 for v in sample if parse_date_to_mmddyyyy(v) != '')
        return matches / total if total else 0.0

    if category == "policynum":
        # Policy numbers: mix of letters/digits, 6‑30 chars, no spaces
        pol_pattern = re.compile(r'^[A-Za-z0-9\-./]{6,30}$')
        matches = sum(1 for v in sample if pol_pattern.match(str(v)))
        return matches / total if total else 0.0

    return 0.0


def _best_header_match_score(header_text: str, keyword_set: Set[str]) -> Tuple[float, str]:
    """Return highest fuzzy ratio between cleaned header and any keyword in the set."""
    cleaned = _clean_string(header_text)
    if not cleaned:
        return 0.0, ""

    best = 0.0
    best_kw = ""
    for kw in keyword_set:
        ratio = _fuzzy_ratio(cleaned, kw)
        if ratio > best:
            best = ratio
            best_kw = kw
    return best, best_kw


# ==================== Main classification function ====================
def classify_columns(
    table: TableInfo,
    company_header_dict: Dict[str, Set[str]]
) -> MappingResult:
    """
    Apply strict three‑pass classification to a single table.
    Returns a MappingResult.
    """
    headers = table.headers

    # -------- Prepare keyword sets for each category --------
    kw_sets = {
        "fullname": company_header_dict.get("name", set()),
        "firstlast": company_header_dict.get("firstlastname", set()),
        "sex": company_header_dict.get("sex", set()),
        "dob": company_header_dict.get("dob", set()),
        "policynum": company_header_dict.get("policynum", set()),
    }

    # Track per‑column matches: col_idx -> list of (category, confidence)
    col_matches = {}

    def add_match(col_idx, category, confidence):
        col_matches.setdefault(col_idx, []).append((category, confidence))

    # -------- Pass 1: Exact cleaned full‑string match ---------
    for idx, raw_header in enumerate(headers):
        cleaned_full = _clean_string(raw_header)
        if not cleaned_full:
            continue

        for dict_key, cat in DICT_TO_CATEGORY.items():
            kw_set = kw_sets[cat]
            if cleaned_full in kw_set:
                add_match(idx, cat, 0.95)

    # -------- Pass 2: Token intersection ---------
    for idx, raw_header in enumerate(headers):
        tokens = _tokenise(raw_header)
        if not tokens:
            continue

        for dict_key, cat in DICT_TO_CATEGORY.items():
            kw_set = kw_sets[cat]
            if any(token in kw_set for token in tokens):
                add_match(idx, cat, 0.75)

    # -------- Pass 3: Adjacency heuristic ---------
    matched_categories = {
        cat for idx, lst in col_matches.items() for cat, _ in lst
    }
    if matched_categories & {"fullname", "firstlast", "sex", "dob", "policynum"}:
        missing = [cat for cat in REQUIRED_CATEGORIES if cat not in matched_categories]
        for miss_cat in missing:
            for idx, raw_header in enumerate(headers):
                if idx in col_matches:
                    continue
                # Use data sample to score
                if idx >= table.data_sample.width:
                    continue
                col_data = table.data_sample[:, idx]
                score = _content_score_for_category(col_data, miss_cat)
                if score > 0.7:
                    add_match(idx, miss_cat, 0.4)  # low confidence

    # -------- Consolidate matches per category --------
    final_mappings: List[ColumnMapping] = []
    ambiguous = False

    single_cats = ["fullname", "sex", "dob", "policynum"]
    for cat in single_cats:
        candidates = [(idx, conf) for idx, lst in col_matches.items()
                      for c, conf in lst if c == cat]
        if not candidates:
            continue

        if len(candidates) == 1:
            idx, conf = candidates[0]
            final_mappings.append(ColumnMapping(
                category=cat,
                column_index=idx,
                header_raw=headers[idx],
                confidence=conf,
            ))
        else:
            scored = []
            for idx, conf in candidates:
                header_score, _ = _best_header_match_score(headers[idx], kw_sets[cat])
                col_data = table.data_sample[:, idx] if idx < table.data_sample.width else pl.Series([], dtype=pl.Utf8)
                content_score = _content_score_for_category(col_data, cat)
                combined = header_score * 0.6 + content_score * 0.4
                scored.append((idx, conf, combined))

            scored.sort(key=lambda x: x[2], reverse=True)
            best_idx, best_conf, best_comb = scored[0]
            second_best_comb = scored[1][2] if len(scored) > 1 else 0.0

            if best_comb - second_best_comb < 0.1:
                ambiguous = True
                for idx, conf, comb in scored:
                    if best_comb - comb <= 0.1:
                        final_mappings.append(ColumnMapping(
                            category=cat,
                            column_index=idx,
                            header_raw=headers[idx],
                            confidence=conf,
                            ambiguous_candidates=[
                                other_idx for other_idx, _, other_comb in scored
                                if other_idx != idx and best_comb - other_comb <= 0.1
                            ]
                        ))
            else:
                final_mappings.append(ColumnMapping(
                    category=cat,
                    column_index=best_idx,
                    header_raw=headers[best_idx],
                    confidence=best_conf,
                ))

    # -------- Handle firstlast (surname + firstname) --------
    firstlast_candidates = sorted(
        [idx for idx, lst in col_matches.items() for c, _ in lst if c == "firstlast"]
    )
    if firstlast_candidates:
        pairs = []
        i = 0
        while i < len(firstlast_candidates):
            if i + 1 < len(firstlast_candidates):
                pairs.append((firstlast_candidates[i], firstlast_candidates[i+1]))
                i += 2
            else:
                ambiguous = True
                idx = firstlast_candidates[i]
                final_mappings.append(ColumnMapping(
                    category="surname",
                    column_index=idx,
                    header_raw=headers[idx],
                    confidence=0.5,
                ))
                i += 1

        for sur_idx, fn_idx in pairs:
            final_mappings.append(ColumnMapping(
                category="surname",
                column_index=sur_idx,
                header_raw=headers[sur_idx],
                confidence=0.85,
            ))
            final_mappings.append(ColumnMapping(
                category="firstname",
                column_index=fn_idx,
                header_raw=headers[fn_idx],
                confidence=0.85,
            ))

    # -------- Determine missing categories and user input flag --------
    mapped_cats = {m.category for m in final_mappings}
    has_name = "fullname" in mapped_cats or ("surname" in mapped_cats and "firstname" in mapped_cats)
    missing_cats = []
    if not has_name:
        missing_cats.append("fullname")
    for req in ["sex", "dob", "policynum"]:
        if req not in mapped_cats:
            missing_cats.append(req)

    requires_user_input = bool(missing_cats) or ambiguous

    return MappingResult(
        table=table,
        mappings=final_mappings,
        missing_categories=missing_cats,
        requires_user_input=requires_user_input,
    )