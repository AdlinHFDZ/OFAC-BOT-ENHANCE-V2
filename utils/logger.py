# utils/logger.py
"""
CSV logging utility for OFAC scans and header extractions.
"""

import csv
import os
from datetime import datetime

# ---------- Log headers ----------
SCAN_LOG_HEADERS = [
    "File Path", "File Name", "Scan Date", "Extension", "Company Code",
    "Password", "Sheet Name", "Error Msg", "Identified Headers",
    "Multiple Name", "Row Count", "Output Row Count", "Output CSV",
    "First Last Name Header", "Full Name Header", "Policy Number Header",
    "DOB Header", "Sex Header", "Remarks"
]

EXTRACT_LOG_HEADERS = [
    "File Path", "File Name", "Sheet Name",
    "Headers Extracted", "Error Msg", "Remarks"
]


def write_log_row(
    log_file_path: str,
    row_dict: dict,
    fieldnames: list
) -> None:
    """
    Append a row to a CSV log file.
    If the file does not exist or is empty, write the header first.
    Missing keys are filled with None.
    """
    # Ensure the output folder exists
    os.makedirs(os.path.dirname(log_file_path), exist_ok=True)

    # Build a complete row with all fields
    row = {field: None for field in fieldnames}
    row.update(row_dict)

    try:
        file_exists = os.path.exists(log_file_path) and os.path.getsize(log_file_path) > 0
        with open(log_file_path, 'a', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            if not file_exists:
                writer.writeheader()
            writer.writerow(row)
    except Exception as e:
        # Fallback to console if logging fails
        print(f"Log write error: {e}")