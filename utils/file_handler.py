# utils/file_handler.py
"""
File‑level operations shared across the OFAC pipeline.
- Excel decryption
- 7‑Zip archive extraction
- Safe file moving / archiving
- Path helpers and extension constants
"""

import os
import io
import subprocess
import shutil
import time
import tempfile
from typing import List, Tuple, Union

import msoffcrypto

# ---------- Constants ----------
SEVEN_ZIP_PATH = os.environ.get(
    "SEVEN_ZIP_PATH", r"C:\Program Files\7-zip\7z.exe"
)

# File extension groups (used for filtering)
FILE_EXTENSIONS_EXCEL = ["xlsx", "xls", "xlsm", "xlsb"]
FILE_EXTENSIONS_TEXT   = ["csv", "txt", "rpt"]
FILE_EXTENSIONS_ARCHIVE = ["zip", "zipx", "tar", "7z", "rar"]


# ==================== Path helpers ====================
def get_unique_save_path(file_path: str) -> str:
    """
    Return a non‑existent file path by appending a counter.
    """
    folder = os.path.dirname(file_path)
    base, ext = os.path.splitext(os.path.basename(file_path))
    safe_name = base[:200]  # avoid long paths
    new_path = os.path.join(folder, f"{safe_name}{ext}")
    count = 1
    while os.path.exists(new_path):
        new_path = os.path.join(folder, f"{safe_name}_{count}{ext}")
        count += 1
    return new_path


def get_all_files(directory: str) -> set:
    """Return a set of full paths to all files inside directory (recursive)."""
    return {
        os.path.join(root, f)
        for root, _, files in os.walk(directory)
        for f in files
    }


# ==================== Excel decryption ====================
def unlock_excel(file_path: str, passwords: List[str]) -> Tuple[Union[str, io.BytesIO], str]:
    """
    Try to decrypt an Excel file using a list of passwords.
    Returns (decrypted_file_path_or_BytesIO, successful_password).
    If file is not encrypted, returns (original_path, "").
    """
    with open(file_path, 'rb') as f:
        office_file = msoffcrypto.OfficeFile(f)
        if not office_file.is_encrypted():
            return file_path, ""

        for pwd in passwords:
            try:
                office_file.load_key(password=pwd)
                decrypted = io.BytesIO()
                office_file.decrypt(decrypted)
                decrypted.seek(0)
                return decrypted, pwd
            except Exception:
                continue

    raise ValueError(f"Could not decrypt {file_path} with any provided password.")


# ==================== Archive extraction ====================
def extract_archive(
    archive_path: str,
    passwords: List[str],
    output_dir: str
) -> List[str]:
    """
    Extract an archive (zip, 7z, rar, tar) into output_dir.
    Tries each password with 7‑Zip; if all fail, attempts without password.
    Returns list of extracted file paths.
    """
    extracted_files = []
    # Create an isolated temp folder inside output_dir to avoid clutter
    extract_dir = tempfile.mkdtemp(prefix="ext_", dir=output_dir)

    try:
        success = False
        for pwd in passwords:
            cmd = [
                SEVEN_ZIP_PATH, 'x', archive_path,
                '-aou', f'-o{extract_dir}', f'-p{pwd}'
            ]
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                success = True
                break

        if not success:
            # Try without password as last resort
            cmd = [SEVEN_ZIP_PATH, 'x', archive_path, '-aou', f'-o{extract_dir}']
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode != 0:
                raise RuntimeError("Archive extraction failed – no valid password")

        # Collect all files inside the extract directory
        extracted_files = [
            os.path.join(root, f)
            for root, _, files in os.walk(extract_dir)
            for f in files
        ]
        if not extracted_files:
            raise RuntimeError("No files found after extraction")

    except Exception:
        # Cleanup failed extraction dir
        shutil.rmtree(extract_dir, ignore_errors=True)
        raise

    return extracted_files


# ==================== Safe move / archive ====================
def move_file_to_archived(
    src: str,
    archived_folder: str,
    retries: int = 3
) -> str:
    """
    Move a file to the archived folder, ensuring unique name.
    Retries on permission error.
    Returns the final destination path.
    """
    dest = get_unique_save_path(
        os.path.join(archived_folder, os.path.basename(src))
    )
    for attempt in range(retries):
        try:
            shutil.move(src, dest)
            return dest
        except (PermissionError, OSError):
            time.sleep(10)
    raise OSError(f"Failed to move {src} after {retries} attempts")


# ==================== Quick check ====================
def is_file_extension(filename: str, extensions: List[str]) -> bool:
    """Return True if the file extension (without dot) is in the given list."""
    ext = os.path.splitext(filename)[1].lower().lstrip('.')
    return ext in extensions