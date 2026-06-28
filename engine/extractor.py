# engine/extractor.py
"""
Extractor – Stage 3 of the extraction pipeline.
Takes a raw DataFrame (full file) and a validated MappingResult,
applies cleaners, and produces the standard OFAC output DataFrame.
"""

from typing import List, Optional

import polars as pl

from engine.inspector import TableInfo
from engine.classifier import MappingResult, ColumnMapping
from utils.cleaners import (
    clean_name,
    split_full_name,
    normalize_sex,
    parse_date_to_mmddyyyy,
    clean_policy_number,
)

# ---------- Constants ----------
OUTPUT_COLUMNS = [
    "SURNAME", "FIRST_NAME", "COMPLETE_NAME", "SEX", "DATE_OF_BIRTH",
    "CMPY_NO", "POLICY_NUMBER", "FILE_PATH", "SHEET"
]


def _get_column(df: pl.DataFrame, col_index: Optional[int]) -> pl.Series:
    """Return the column as a Utf8 Series, or an all‑null Utf8 Series if index is None."""
    if col_index is None or col_index >= df.width:
        return pl.Series("null_col", [None] * df.height, dtype=pl.Utf8)
    return df[:, col_index].cast(pl.Utf8)


def extract_from_table(
    df: pl.DataFrame,
    table_info: TableInfo,
    mapping: MappingResult,
    company_code: str,
    file_path: str,
) -> pl.DataFrame:
    """
    Extract and clean OFAC fields from a single table.

    Args:
        df: Full raw DataFrame (including header row).
        table_info: TableInfo for this block (contains col_start, col_end,
                    data_start, etc.)
        mapping: MappingResult (user‑verified, no missing required categories).
        company_code: Company code string.
        file_path: Original file path for output record.

    Returns:
        Polars DataFrame with columns: OUTPUT_COLUMNS.
    """
    # 1. Slice to data rows only (skip header)
    data_df = df.slice(table_info.data_start)
    if data_df.is_empty():
        return pl.DataFrame(schema={col: pl.Utf8 for col in OUTPUT_COLUMNS})

    # 2. Build helper to find mapping by category
    mapping_dict = {m.category: m.column_index for m in mapping.mappings}

    # 3. Extract raw columns
    surname_idx = mapping_dict.get("surname")
    firstname_idx = mapping_dict.get("firstname")
    fullname_idx = mapping_dict.get("fullname")
    sex_idx = mapping_dict.get("sex")
    dob_idx = mapping_dict.get("dob")
    pol_idx = mapping_dict.get("policynum")

    surname_raw = _get_column(data_df, surname_idx)
    firstname_raw = _get_column(data_df, firstname_idx)
    fullname_raw = _get_column(data_df, fullname_idx)
    sex_raw = _get_column(data_df, sex_idx)
    dob_raw = _get_column(data_df, dob_idx)
    pol_raw = _get_column(data_df, pol_idx)

    # 4. Build SURNAME and FIRST_NAME
    # Priority: explicit surname/firstname mapping > split from fullname > empty
    if surname_idx is not None and firstname_idx is not None:
        # Use explicit columns
        surname_clean = surname_raw.map_elements(clean_name, return_dtype=pl.Utf8)
        firstname_clean = firstname_raw.map_elements(clean_name, return_dtype=pl.Utf8)
    elif fullname_idx is not None:
        # Split fullname into surname and first_name
        # We'll apply split_full_name row‑wise via a lambda
        def split_fn(name_str):
            if name_str is None:
                return ("", "")
            cleaned = clean_name(str(name_str))
            return split_full_name(cleaned)

        # map_elements returns a struct; we can extract fields
        splitted = fullname_raw.map_elements(
            lambda x: split_fn(x),
            return_dtype=pl.Struct([pl.Field("surname", pl.Utf8), pl.Field("firstname", pl.Utf8)]),
        )
        surname_clean = splitted.struct.field("surname")
        firstname_clean = splitted.struct.field("firstname")
    else:
        # No name columns at all – should not happen if mapping is validated
        surname_clean = pl.Series("surname", [""] * data_df.height, dtype=pl.Utf8)
        firstname_clean = pl.Series("firstname", [""] * data_df.height, dtype=pl.Utf8)

    # 5. Build COMPLETE_NAME
    complete_name = (
        pl.concat_str([surname_clean, firstname_clean], separator=" ")
        .str.strip_chars()
        .fill_null("")
    )

    # 6. Clean sex, dob, policy
    sex_clean = sex_raw.map_elements(normalize_sex, return_dtype=pl.Utf8).fill_null("")
    dob_clean = dob_raw.map_elements(parse_date_to_mmddyyyy, return_dtype=pl.Utf8).fill_null("")
    pol_clean = pol_raw.map_elements(clean_policy_number, return_dtype=pl.Utf8).fill_null("")

    # 7. Assemble result
    result = pl.DataFrame({
        "SURNAME": surname_clean,
        "FIRST_NAME": firstname_clean,
        "COMPLETE_NAME": complete_name,
        "SEX": sex_clean,
        "DATE_OF_BIRTH": dob_clean,
        "CMPY_NO": pl.lit(company_code).cast(pl.Utf8),
        "POLICY_NUMBER": pol_clean,
        "FILE_PATH": pl.lit(file_path).cast(pl.Utf8),
        "SHEET": pl.lit(table_info.sheet_name).cast(pl.Utf8),
    })

    # 8. Drop rows with empty COMPLETE_NAME (could happen if both names missing)
    missing_name_count = result.filter(pl.col("COMPLETE_NAME") == "").height
    result = result.filter(pl.col("COMPLETE_NAME") != "")

    # 9. Deduplicate within this table (optional, but helpful)
    result = result.unique()

    return result