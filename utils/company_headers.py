# utils/company_headers.py
"""
Company Header Dictionary Loader (with tokenisation)
Loads header keywords from Excel files per company.
Supports multi‑word cells – splits into tokens for better matching.
The headers folder can be set dynamically via set_header_path().
"""

import os
import re
from typing import Dict, Set, Tuple

import polars as pl

# ---------- DEFAULT PATH (will be overridden by GUI) ----------
COMPANY_HEADERS_PATH = os.environ.get(
    "OFAC_HEADERS_PATH",
    r"C:\Users\s0055198\OneDrive - RGA Reinsurance Company\Documents\OFAC CODE\header"
)

def set_header_path(new_path: str):
    """Allow the GUI (or other callers) to change the header files folder at runtime."""
    global COMPANY_HEADERS_PATH
    COMPANY_HEADERS_PATH = new_path

# ---------- CONSTANTS ----------
HEADER_SHEETS = ["name", "firstlastname", "sex", "dob", "policynum"]

STOP_WORDS = {"of", "the", "a", "an", "in", "on", "at", "to", "for", "and", "or", "is", "are", "was", "were", "be"}


def _clean_text(text: str) -> str:
    """Lowercase, strip, keep only alpha characters and spaces."""
    if not text:
        return ""
    text = text.lower().strip()
    text = re.sub(r'[^a-z\s]', '', text)
    return text


def _tokenise_cell(raw_value) -> Set[str]:
    """
    Given a cell value, return a set of keywords:
    - The full cleaned string (run‑on word)
    - Individual tokens (words) excluding stop words and single letters.
    """
    if raw_value is None:
        return set()

    text = _clean_text(str(raw_value))
    if not text:
        return set()

    # Full run‑on string (no spaces) – matches the old behaviour
    full_runon = text.replace(' ', '')
    tokens = {full_runon} if full_runon else set()

    # Individual words (split on whitespace)
    words = text.split()
    for word in words:
        w = word.strip()
        if len(w) >= 2 and w not in STOP_WORDS:
            tokens.add(w)
    return tokens


def get_company_header(company_code: str) -> Tuple[Dict[str, Set[str]], str]:
    """
    Load header keywords for a given company code.
    First tries: ...\header\by_company\header_{COMPANY_CODE}.xlsx
    Falls back to: ...\header\header.xlsx
    Returns (header_dict, used_file_path)
        header_dict: keys = "name", "firstlastname", "sex", "dob", "policynum"
                     values = set of lowercase token strings
    """
    company_path = os.path.join(
        COMPANY_HEADERS_PATH, "by_company",
        f"header_{company_code.upper().strip()}.xlsx"
    )
    default_path = os.path.join(COMPANY_HEADERS_PATH, "header.xlsx")
    file_path = company_path if os.path.exists(company_path) else default_path

    header_dict: Dict[str, Set[str]] = {}

    for sheet in HEADER_SHEETS:
        tokens_set = set()
        try:
            df = pl.read_excel(
                file_path, sheet_name=sheet, has_header=False, raise_if_empty=False
            )
            if df.is_empty():
                header_dict[sheet] = tokens_set
                continue

            for col in df.columns:
                values = df[col].drop_nulls().to_list()
                for val in values:
                    tokens_set.update(_tokenise_cell(val))

        except Exception:
            pass

        header_dict[sheet] = tokens_set

    return header_dict, file_path