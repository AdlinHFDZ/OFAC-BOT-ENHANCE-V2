# engine/job.py
"""
Extraction Job – Orchestrator that runs the full pipeline for a batch of files.
Replaces the old ScannerJob + process_files_direct logic.
Now supports user‑defined mapping overrides, configurable output root,
archive progress messages, and mid‑file cancellation via stop flag.
"""

import os
import io
import time
import traceback
from datetime import datetime
from typing import Callable, Dict, List, Optional, Tuple

import polars as pl

from utils.file_handler import (
    unlock_excel,
    get_unique_save_path,
    move_file_to_archived,
    extract_archive,
    is_file_extension,
    FILE_EXTENSIONS_EXCEL,
    FILE_EXTENSIONS_TEXT,
    FILE_EXTENSIONS_ARCHIVE,
)
from utils.company_headers import get_company_header
from utils.logger import SCAN_LOG_HEADERS, write_log_row
from utils.output_compiler import compile_outputs, OUTPUT_COLUMNS

from engine.inspector import inspect_file, FileStructure, TableInfo
from engine.classifier import classify_columns, MappingResult, ColumnMapping
from engine.extractor import extract_from_table

# ---------- Constants ----------
MAX_ROWS_PER_OUTPUT_CSV = 500_000

# Fallback output folder – used only if not provided by the GUI
DEFAULT_OUTPUT_FOLDER = os.environ.get(
    "OFAC_OUTPUT_FOLDER",
    os.path.expanduser("~\\OFAC_Output")
)


class ExtractionJob:
    """
    Holds the state of a scan session and executes the pipeline.
    """

    def __init__(
        self,
        input_folder: str,
        company_code: str,
        passwords: List[str],
        email_received_date: str,
        file_names: List[str],
        mapping_overrides: Optional[Dict[str, Dict[Tuple[str, int, int], List[ColumnMapping]]]] = None,
        output_root: Optional[str] = None,
    ):
        self.input_folder = input_folder
        self.company_code = company_code
        self.passwords = passwords
        self.email_received_date = email_received_date
        self.file_names = file_names
        self.mapping_overrides = mapping_overrides or {}

        # Date formatting
        try:
            dt = datetime.strptime(email_received_date, "%Y-%m-%d")
            self.date_display = dt.strftime("%Y%m%d")
        except ValueError:
            self.date_display = email_received_date.replace("-", "")

        # Use provided output root, or fall back to default
        self.output_root = os.path.join(
            output_root or DEFAULT_OUTPUT_FOLDER,
            self.date_display,
            self.company_code
        )
        self.csv_folder = os.path.join(self.output_root, "CSVs")
        self.archived_folder = os.path.join(self.output_root, "Archived")
        self.unzipped_folder = os.path.join(self.output_root, "Unzipped")
        self.compiled_folder = os.path.join(self.output_root, "Compiled")
        self.log_file = os.path.join(
            self.output_root, f"Log_{self.date_display}.csv"
        )

        for d in [self.csv_folder, self.archived_folder, self.unzipped_folder, self.compiled_folder]:
            os.makedirs(d, exist_ok=True)

    def run(
        self,
        progress_callback: Optional[Callable[[str], None]] = None,
        stop_flag: Optional[Callable[[], bool]] = None,
        progress_update: Optional[Callable[[int, int], None]] = None,
    ) -> bool:
        """Execute the extraction pipeline."""
        company_header_dict, used_header_file = get_company_header(self.company_code)
        if progress_callback:
            progress_callback(f"Loaded header dictionary from {used_header_file}")

        if not os.path.exists(self.log_file) or os.path.getsize(self.log_file) == 0:
            write_log_row(self.log_file, {"Remarks": "Log started"}, SCAN_LOG_HEADERS)

        all_file_paths = [
            os.path.join(self.input_folder, f)
            for f in self.file_names
            if os.path.exists(os.path.join(self.input_folder, f))
        ]
        idx = 0
        processed_csvs = []

        while idx < len(all_file_paths):
            if stop_flag and stop_flag():
                write_log_row(self.log_file, {"Remarks": "Scan stopped by user"}, SCAN_LOG_HEADERS)
                return False

            src_path = all_file_paths[idx]
            file_name = os.path.basename(src_path)

            if progress_callback:
                progress_callback(f"Processing {file_name}...")
            if progress_update:
                progress_update(idx + 1, len(all_file_paths))

            ext = os.path.splitext(file_name)[1].lower().lstrip('.')

            # Archives
            if is_file_extension(file_name, FILE_EXTENSIONS_ARCHIVE):
                if progress_callback:
                    progress_callback(f"Extracting archive: {file_name}")
                try:
                    extracted = extract_archive(src_path, self.passwords, self.unzipped_folder)
                    all_file_paths.extend(extracted)
                    if progress_callback:
                        progress_callback(f"Extracted {len(extracted)} files from {file_name}")
                    write_log_row(self.log_file, {
                        "File Path": src_path, "File Name": file_name,
                        "Extension": ext, "Remarks": f"Extracted {len(extracted)} files"
                    }, SCAN_LOG_HEADERS)
                    try:
                        move_file_to_archived(src_path, self.archived_folder)
                    except Exception:
                        pass
                except Exception as e:
                    write_log_row(self.log_file, {
                        "File Path": src_path, "File Name": file_name,
                        "Extension": ext, "Error Msg": str(e)
                    }, SCAN_LOG_HEADERS)
                idx += 1
                continue

            # Excel / CSV
            if not is_file_extension(file_name, FILE_EXTENSIONS_EXCEL + FILE_EXTENSIONS_TEXT):
                idx += 1
                continue

            try:
                file_csvs = self._process_single_file(
                    src_path, file_name, company_header_dict, progress_callback,
                    stop_flag=stop_flag
                )
                processed_csvs.extend(file_csvs)

                # If stop was requested during file processing, break out of the loop
                if stop_flag and stop_flag():
                    write_log_row(self.log_file, {"Remarks": "Scan stopped by user"}, SCAN_LOG_HEADERS)
                    return False

                try:
                    move_file_to_archived(src_path, self.archived_folder)
                except Exception as e:
                    if progress_callback:
                        progress_callback(f"Warning: could not archive {file_name}: {e}")

            except Exception as e:
                write_log_row(self.log_file, {
                    "File Path": src_path, "File Name": file_name,
                    "Extension": ext,
                    "Error Msg": f"{type(e).__name__}: {e}",
                    "Remarks": traceback.format_exc()[-2000:],
                }, SCAN_LOG_HEADERS)
                if progress_callback:
                    progress_callback(f"Error processing {file_name}: {e}")

            idx += 1

        # Final compilation
        if processed_csvs:
            if progress_callback:
                progress_callback("Compiling output files...")
            try:
                compiled_files = compile_outputs(
                    csv_paths=processed_csvs,
                    output_dir=self.compiled_folder,
                    base_filename=f"OFAC_ABS_Log_{self.date_display}_{self.company_code}",
                )
                write_log_row(self.log_file, {
                    "Remarks": f"Compiled {len(compiled_files)} Excel file(s)"
                }, SCAN_LOG_HEADERS)
            except Exception as e:
                write_log_row(self.log_file, {
                    "Error Msg": f"Compile failed: {e}"
                }, SCAN_LOG_HEADERS)

        return True

    # ==================== Private methods ====================
    def _process_single_file(
        self,
        file_path: str,
        file_name: str,
        header_dict: Dict[str, set],
        progress_callback: Optional[Callable] = None,
        stop_flag: Optional[Callable[[], bool]] = None,
    ) -> List[str]:
        """Inspect, classify, and extract all tables in a single file, using overrides if available."""
        structure = inspect_file(file_path, self.passwords, self.unzipped_folder)
        if not structure.tables:
            write_log_row(self.log_file, {
                "File Path": file_path, "File Name": file_name,
                "Error Msg": "No tables found"
            }, SCAN_LOG_HEADERS)
            return []

        full_df_dict = self._load_full_file(file_path, file_name)
        if full_df_dict is None:
            return []

        generated_csvs = []

        for table in structure.tables:
            # Allow stopping between tables
            if stop_flag and stop_flag():
                write_log_row(self.log_file, {
                    "File Path": file_path, "File Name": file_name,
                    "Remarks": "Scan stopped by user during table processing"
                }, SCAN_LOG_HEADERS)
                return generated_csvs

            # Determine base sheet name for full_df_dict
            if "_table" in table.sheet_name:
                base_sheet = table.sheet_name.rsplit("_table", 1)[0]
            else:
                base_sheet = table.sheet_name

            if base_sheet not in full_df_dict:
                continue

            full_df = full_df_dict[base_sheet]

            # --- Mapping override logic ---
            mapping = None
            override_key = (table.sheet_name, table.col_start, table.col_end)
            if file_path in self.mapping_overrides:
                overrides_for_file = self.mapping_overrides[file_path]
                if override_key in overrides_for_file:
                    mapping_list = overrides_for_file[override_key]
                    mapping = MappingResult(
                        table=table,
                        mappings=mapping_list,
                        missing_categories=[],
                        requires_user_input=False,
                    )
                    if progress_callback:
                        progress_callback(f"Using user‑defined mapping for {table.sheet_name}")

            if mapping is None:
                mapping = classify_columns(table, header_dict)

                if mapping.requires_user_input:
                    write_log_row(self.log_file, {
                        "File Path": file_path,
                        "File Name": file_name,
                        "Sheet Name": table.sheet_name,
                        "Error Msg": f"Missing categories: {mapping.missing_categories}",
                        "Remarks": "Extraction will proceed with incomplete mapping",
                    }, SCAN_LOG_HEADERS)

            output_df = extract_from_table(
                df=full_df,
                table_info=table,
                mapping=mapping,
                company_code=self.company_code,
                file_path=file_path,
            )

            if output_df.is_empty():
                write_log_row(self.log_file, {
                    "File Path": file_path,
                    "File Name": file_name,
                    "Sheet Name": table.sheet_name,
                    "Error Msg": "No data rows after extraction",
                }, SCAN_LOG_HEADERS)
                continue

            csv_paths = self._write_output_csvs(output_df, file_name, table.sheet_name)
            generated_csvs.extend(csv_paths)

            write_log_row(self.log_file, {
                "File Path": file_path,
                "File Name": file_name,
                "Sheet Name": table.sheet_name,
                "Row Count": full_df.height,
                "Output Row Count": output_df.height,
                "Output CSV": ", ".join(csv_paths),
                "Identified Headers": ", ".join(table.headers),
                "Multiple Name": False,
            }, SCAN_LOG_HEADERS)

        return generated_csvs

    def _load_full_file(self, file_path: str, file_name: str) -> Optional[Dict[str, pl.DataFrame]]:
        """Load full DataFrames for all sheets of a file (handles decryption)."""
        try:
            if is_file_extension(file_name, FILE_EXTENSIONS_EXCEL):
                decrypted, _ = unlock_excel(file_path, self.passwords)
                if isinstance(decrypted, io.BytesIO):
                    temp_path = os.path.join(self.unzipped_folder, f"temp_full_{int(time.time())}_{file_name}")
                    with open(temp_path, 'wb') as f:
                        f.write(decrypted.getvalue())
                    sheets = pl.read_excel(temp_path, has_header=False, sheet_id=0, raise_if_empty=False)
                    os.remove(temp_path)
                else:
                    sheets = pl.read_excel(decrypted, has_header=False, sheet_id=0, raise_if_empty=False)
            else:
                df = pl.read_csv(file_path, has_header=False, truncate_ragged_lines=True)
                sheets = {'': df}
            return sheets
        except Exception as e:
            return None

    def _write_output_csvs(self, df: pl.DataFrame, file_name: str, sheet_label: str) -> List[str]:
        """Split a result DataFrame into chunks and write to CSV files."""
        total = df.height
        max_rows = MAX_ROWS_PER_OUTPUT_CSV
        chunk_start = 0
        part = 1
        csv_paths = []

        while chunk_start < total:
            chunk = df.slice(chunk_start, max_rows)
            out_path = os.path.join(
                self.csv_folder,
                f"{file_name}[{sheet_label}]_part{part}_OFAC_OUTPUT.csv"
            )
            out_path = get_unique_save_path(out_path)
            chunk.write_csv(out_path)
            csv_paths.append(out_path)
            chunk_start += max_rows
            part += 1
        return csv_paths