# engine/inspector.py
"""
File Inspector – Stage 1 of the extraction pipeline.
Reads Excel/CSV files (decrypting if necessary) and returns a
FileStructure object containing all discovered tables and their raw headers.
"""

import os
import io
import time
from dataclasses import dataclass, field
from typing import List, Optional, Union

import polars as pl

from utils.file_handler import (
    unlock_excel,
    is_file_extension,
    FILE_EXTENSIONS_EXCEL,
    FILE_EXTENSIONS_TEXT,
)

# ---------- Constants ----------
MAX_SEARCH_ROWS = 50        # rows to scan for header detection
SAMPLE_PREVIEW_ROWS = 20    # rows to keep for UI preview


# ---------- Data classes ----------
@dataclass
class TableInfo:
    sheet_name: str
    col_start: int
    col_end: int
    header_row: int
    data_start: int
    headers: List[str]
    data_sample: pl.DataFrame   # first SAMPLE_PREVIEW_ROWS rows of data


@dataclass
class FileStructure:
    file_path: str
    tables: List[TableInfo] = field(default_factory=list)


# ---------- Helper functions ----------
def _is_empty_cell(val) -> bool:
    """True if the cell is None, NaN, or empty string after stripping."""
    if val is None:
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    try:
        if pl.Series([val]).null_count() == 1:
            return True
    except Exception:
        pass
    return False


def _column_is_completely_empty(df: pl.DataFrame, col_idx: int, max_rows: int = MAX_SEARCH_ROWS) -> bool:
    """Check if all cells in the first `max_rows` of a column are empty."""
    n = min(df.height, max_rows)
    if n == 0:
        return True
    col_data = df[:n, col_idx].to_list()
    return all(_is_empty_cell(v) for v in col_data)


def _find_table_blocks(df: pl.DataFrame) -> List[tuple]:
    """
    Split a DataFrame into contiguous column blocks (tables).
    A new table begins when we encounter a non‑empty column after a gap
    of at least one completely empty column (within first MAX_SEARCH_ROWS rows).
    """
    non_empty_cols = []
    for col_idx in range(df.width):
        if not _column_is_completely_empty(df, col_idx):
            non_empty_cols.append(col_idx)

    if not non_empty_cols:
        return []

    blocks = []
    start = non_empty_cols[0]
    prev = non_empty_cols[0]
    for c in non_empty_cols[1:]:
        if c > prev + 1:  # gap found
            blocks.append((start, prev))
            start = c
        prev = c
    blocks.append((start, prev))
    return blocks


def _is_likely_numeric(series: pl.Series, sample_size: int = 30) -> bool:
    """Heuristic: if >80% of non‑null sample cells can be parsed as float, treat as numeric."""
    non_null = series.drop_nulls()
    if len(non_null) == 0:
        return False
    sample = non_null.head(sample_size).cast(pl.Utf8).to_list()
    numeric_count = 0
    for v in sample:
        try:
            float(v)
            numeric_count += 1
        except (ValueError, TypeError):
            pass
    return (numeric_count / len(sample)) >= 0.8


def _score_header_row(df: pl.DataFrame, col_start: int, col_end: int, row_idx: int) -> int:
    """
    Score a potential header row by the number of non‑empty cells
    that are NOT purely numeric. Numeric cells are penalised
    (they are likely data, not headers).
    """
    row_vals = df.row(row_idx)[col_start:col_end+1]
    score = 0
    for v in row_vals:
        if _is_empty_cell(v):
            continue
        s = str(v).strip()
        # Full‑numeric strings are less likely to be headers
        try:
            float(s)
            score += 1          # small credit, could be a year like "2025"
        except ValueError:
            score += 3          # text cells are strong header candidates
    return score


def _detect_header_row(df: pl.DataFrame, col_start: int, col_end: int) -> int:
    """
    Find the best header row within MAX_SEARCH_ROWS.
    Returns the row index (0‑based) with the highest score.
    If tie, return the earliest.
    """
    n_rows = min(df.height, MAX_SEARCH_ROWS)
    best_row = 0
    best_score = -1
    for r in range(n_rows):
        score = _score_header_row(df, col_start, col_end, r)
        if score > best_score:
            best_score = score
            best_row = r
    return best_row


# ---------- Main inspection function ----------
def inspect_file(
    file_path: str,
    passwords: List[str],
    temp_dir: str,
) -> FileStructure:
    """
    Inspect a single file (Excel/CSV) and return its FileStructure.
    Handles decryption automatically using the provided passwords.
    """
    structure = FileStructure(file_path=file_path)
    temp_path = None
    actual_path = file_path
    file_name = os.path.basename(file_path)

    try:
        # 1. Decrypt if Excel (CSV/TXT never encrypted)
        if is_file_extension(file_name, FILE_EXTENSIONS_EXCEL):
            decrypted, _ = unlock_excel(file_path, passwords)
            if isinstance(decrypted, io.BytesIO):
                temp_path = os.path.join(temp_dir, f"temp_insp_{int(time.time())}_{file_name}")
                with open(temp_path, 'wb') as f:
                    f.write(decrypted.getvalue())
                actual_path = temp_path
            else:
                actual_path = decrypted

            # Read all sheets
            sheets_dict = pl.read_excel(
                actual_path, has_header=False, sheet_id=0, raise_if_empty=False
            )
        elif is_file_extension(file_name, FILE_EXTENSIONS_TEXT):
            # Read as CSV (try automatic delimiter)
            actual_path = file_path
            df = pl.read_csv(actual_path, has_header=False, truncate_ragged_lines=True)
            sheets_dict = {'': df}
        else:
            # Unsupported format – return empty structure
            return structure

        # 2. Process each sheet
        for sheet_name, df in sheets_dict.items():
            if df.is_empty():
                continue

            # Split into column blocks (tables)
            blocks = _find_table_blocks(df)
            if not blocks:
                continue

            # 3. For each block, detect header row and extract info
            for table_idx, (col_start, col_end) in enumerate(blocks):
                # Detect header row inside this block
                hdr_row = _detect_header_row(df, col_start, col_end)
                data_start = hdr_row + 1

                # Extract raw header values (clean None → empty string)
                raw_headers = []
                for c in range(col_start, col_end + 1):
                    val = df[hdr_row, c]
                    raw_headers.append(str(val) if not _is_empty_cell(val) else "")

                # Grab a preview sample of the data rows
                sample_end = min(df.height, data_start + SAMPLE_PREVIEW_ROWS)
                data_sample = df[data_start:sample_end, col_start:col_end+1]

                # Rename sample columns using header values for readability
                sample_col_names = [
                    h if h else f"Col_{c}"
                    for c, h in enumerate(raw_headers, start=col_start)
                ]
                data_sample.columns = sample_col_names

                # For sheets with multiple tables, name them Table 1, Table 2, ...
                table_label = f"{sheet_name}_table{table_idx+1}" if len(blocks) > 1 else sheet_name

                structure.tables.append(TableInfo(
                    sheet_name=table_label,
                    col_start=col_start,
                    col_end=col_end,
                    header_row=hdr_row,
                    data_start=data_start,
                    headers=raw_headers,
                    data_sample=data_sample,
                ))

    except Exception as e:
        # In a real run we'd log the exception; for now just return whatever was found
        print(f"Inspector error on {file_path}: {e}")
    finally:
        # Clean up temp decrypted file
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

    return structure