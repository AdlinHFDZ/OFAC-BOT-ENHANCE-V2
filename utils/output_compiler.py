# utils/output_compiler.py
"""
Output Compiler
- Reads all generated OFAC CSV output files
- Concatenates, deduplicates, and writes to Excel (.xlsx)
- Respects EXCEL_MAX_ROWS (default 1,048,576) and splits into multiple files if needed
- Removes Mark of the Web (Zone.Identifier) from final Excel files
"""

import os
from typing import List

import polars as pl

# ---------- Constants ----------
OUTPUT_COLUMNS = [
    "SURNAME", "FIRST_NAME", "COMPLETE_NAME", "SEX", "DATE_OF_BIRTH",
    "CMPY_NO", "POLICY_NUMBER", "FILE_PATH", "SHEET"
]
EXCEL_MAX_ROWS = 1_000_000   # stay well under Excel's 1,048,576 limit


def _remove_motw(file_path: str) -> None:
    """
    Remove Mark of the Web (Zone.Identifier alternate data stream)
    from an Excel file to prevent 'file is dangerous' warning on shared drives.
    Only applies on Windows; silently passes on other OS.
    """
    try:
        # Windows alternate data stream syntax
        os.remove(file_path + ":Zone.Identifier:$DATA")
    except Exception:
        pass  # stream may not exist or OS doesn't support it


def compile_outputs(
    csv_paths: List[str],
    output_dir: str,
    base_filename: str,
    max_rows: int = EXCEL_MAX_ROWS,
) -> List[str]:
    """
    Read a list of OFAC output CSV files, concatenate, deduplicate,
    and write to one or more Excel files (splitting if row limit exceeded).

    Args:
        csv_paths: List of full paths to individual output CSV files.
        output_dir: Destination folder for compiled Excel files.
        base_filename: Base name for output file(s) (without extension).
                       e.g., "OFAC_ABS_Log_20260101_ABC"
        max_rows: Maximum rows per Excel file.

    Returns:
        List of paths to the created Excel files.
    """
    os.makedirs(output_dir, exist_ok=True)

    # Read and combine all CSV files
    df_combined = None
    for path in csv_paths:
        if not os.path.exists(path):
            continue
        try:
            # Ensure POLICY_NUMBER is read as string
            chunk = pl.read_csv(path, schema_overrides={'POLICY_NUMBER': pl.Utf8})
            if chunk.is_empty():
                continue
            if df_combined is None:
                df_combined = chunk
            else:
                df_combined = pl.concat([df_combined, chunk], how='diagonal_relaxed')
        except Exception as e:
            # Log to console; later integrate with logger
            print(f"Warning: Could not read {path}: {e}")

    if df_combined is None or df_combined.is_empty():
        return []

    # Deduplicate
    df_combined = df_combined.unique()

    # Ensure OUTPUT_COLUMNS exist; add missing as None
    for col in OUTPUT_COLUMNS:
        if col not in df_combined.columns:
            df_combined = df_combined.with_columns(pl.lit(None).alias(col))

    # Reorder columns
    df_combined = df_combined.select(OUTPUT_COLUMNS)

    total_rows = df_combined.height
    output_files = []

    # Split into chunks if needed
    file_index = 1
    chunk_start = 0
    while chunk_start < total_rows:
        chunk_end = min(chunk_start + max_rows, total_rows)
        chunk_df = df_combined.slice(chunk_start, chunk_end - chunk_start)

        if file_index == 1 and total_rows <= max_rows:
            out_path = os.path.join(output_dir, f"{base_filename}.xlsx")
        else:
            out_path = os.path.join(output_dir, f"{base_filename}_{file_index}.xlsx")

        # Ensure unique path (add counter if already exists)
        from utils.file_handler import get_unique_save_path
        out_path = get_unique_save_path(out_path)

        chunk_df.write_excel(out_path)
        _remove_motw(out_path)
        output_files.append(out_path)

        file_index += 1
        chunk_start = chunk_end

    return output_files