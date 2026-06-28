# utils/cleaners.py
"""
Field‑level cleaning and normalisation functions for OFAC extraction.
- Name cleaning (no aggressive digit stripping)
- Full‑name parsing (using nameparser if available)
- Sex normalisation
- DOB parsing (format MM/DD/YYYY)
- Policy number cleaning
"""

import re
from datetime import datetime
from typing import Optional, Tuple

from dateutil import parser as dateutil_parser


# ==================== Name cleaning ====================
def clean_name(text: Optional[str]) -> str:
    """
    Clean a name field.
    - Removes leading/trailing whitespace.
    - Removes standalone numbers (e.g., '123') but keeps digits attached
      to letters (like '3rd', 'Smith2').
    - Collapses multiple spaces.
    """
    if text is None:
        return ""

    s = str(text).strip()
    if not s:
        return ""

    # Remove isolated numbers that are not part of a word
    # Matches a whole token that consists only of digits (possibly with punctuation)
    s = re.sub(r'\b\d+\b', '', s)

    # Collapse whitespace
    s = re.sub(r'\s+', ' ', s).strip()

    return s


def split_full_name(full_name: str) -> Tuple[str, str]:
    """
    Parse a single full‑name string into (SURNAME, FIRST_NAME).
    Uses the nameparser library if installed; otherwise falls back to a
    simple last‑word‑as‑surname heuristic.
    """
    if not full_name or not full_name.strip():
        return "", ""

    try:
        from nameparser import HumanName
        name = HumanName(full_name)
        # nameparser: .last, .first, .middle, .title, .suffix
        last = name.last or ""
        first = " ".join(filter(None, [name.first, name.middle])).strip()
        # If nameparser couldn't split (e.g., single word), treat whole as surname
        if not last and not first:
            return full_name.strip(), ""
        return last, first
    except ImportError:
        pass

    # Fallback: last word = surname, rest = first name
    parts = full_name.strip().split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[-1], " ".join(parts[:-1])


# ==================== Sex normalisation ====================
def normalize_sex(value: Optional[str]) -> str:
    """
    Convert sex to 'M', 'F', or 'U' (unknown).
    Handles common variations (male, female, m, f, etc.).
    """
    if value is None:
        return ''

    v = str(value).strip().lower()
    if v in ('male', 'm'):
        return 'M'
    if v in ('female', 'f'):
        return 'F'
    # Already single letter? Return uppercase
    if len(v) == 1 and v.isalpha():
        return v.upper()
    return 'U'


# ==================== Date of Birth ====================
def parse_date_to_mmddyyyy(value) -> str:
    """
    Parse a DOB value and return it as a string in MM/DD/YYYY format.
    Accepts datetime objects, strings, or anything dateutil can parse.
    Returns empty string if parsing fails.
    """
    if value is None:
        return ''

    if isinstance(value, datetime):
        return value.strftime('%m/%d/%Y')

    s = str(value).strip()
    if not s:
        return ''

    try:
        dt = dateutil_parser.parse(s, fuzzy=False)
        return dt.strftime('%m/%d/%Y')
    except Exception:
        pass

    # Try a set of common formats as fallback
    formats = [
        '%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y',
        '%d-%b-%Y', '%d-%b-%y', '%b %d, %Y', '%Y%m%d'
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime('%m/%d/%Y')
        except ValueError:
            continue

    return ''


# ==================== Policy number ====================
def clean_policy_number(value: Optional[str]) -> str:
    """
    Trim whitespace from a policy number; keep as original string.
    """
    if value is None:
        return ''
    return str(value).strip()