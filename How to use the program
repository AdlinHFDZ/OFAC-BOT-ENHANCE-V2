Here is your user guide reformatted for GitHub (Markdown). It has been structured with clear heading hierarchies, proper code blocks, scannable bold text, bullet points, and clean tables for enhanced readability.

---

# OFAC Scanner – User Guide

Welcome to the **OFAC Scanner User Guide**. This document explains how to use the scanner day‑to‑day after it has been installed and configured.

## What the Scanner Does

The OFAC Scanner reads Excel and CSV files from the **Watch Folder**, finds the required columns (name, date of birth, sex, policy number), cleans the data, and produces a compiled Excel file ready for OFAC screening.

You can process files for multiple companies one after another using the built-in **Job Queue**.

---

## 1. Understanding the Main Window

When you open the scanner, you see four tabs at the top:

* **Scan:** The main workspace where you process files.
* **Extract Headers:** A tool to catalogue the column headers in your files (useful for building header dictionaries).
* **Settings:** Where you set the watch folder, passwords file, output folder, and header folder.
* **History:** A list of past scans with convenient links to their output folders.

### The Scan Tab Layout

The **Scan** tab is where you will spend most of your time. It is divided into three distinct areas:

1. **Left Sidebar:** Company code, email received date, and password selection.
2. **Top Right:** File list (with filters) and a preview pane that shows the first 20 rows of a selected file.
3. **Bottom Right:** "Detect Headers" button, queue controls, the job queue list, and a live log output.

> 💡 **Status Bar:** At the very bottom of the entire window, there is a master progress bar and a status message.

---

## 2. Everyday Scan – Step by Step

### Step 1: Place Files in the Watch Folder

Copy the Excel / CSV files you want to scan into your configured **Watch Folder**.

* If you have the **Auto‑refresh** checkbox ticked, the file list will update automatically.
* Otherwise, click **Refresh**.

### Step 2: Select a Company

In the left sidebar, type the company code in the **Company Code** box. A dropdown will appear as you type—click on the correct code or press `Enter` when only one choice remains.

### Step 3: Choose Passwords

Below the company code, you’ll see a list of passwords registered for that company.

* Tick the passwords that should be tried when opening protected Excel files.
* Use the **Search** box to filter a long list, or the **Select All / Clear All** buttons for speed.
* If you need to add a new password, click **+ New**.

### Step 4: Select Files

In the file list (top right), tick the files you want to scan.

* Use the filter checkboxes (**Excel, CSV, Archives**) to isolate specific file types.
* **Tip:** Click on a file name to see a preview of its first 20 rows. If the file is an Excel workbook with multiple sheets, use the sheet selector dropdown above the preview pane.

### Step 5: Review Column Mappings (Optional)

Click the **Detect Headers** button. A pop‑up window will show you how the scanner has identified each column.

* Each row shows: *File, Sheet, Table, Header*, and the assigned **Category** (*Surname, First Name, Full Name, Sex, DOB, Policy Number, or Ignore*).
* If a mapping is incorrect, **double‑click** the category cell and pick the correct one from the dropdown.
* To start over, use the **Clear All Mappings** button.
* Click **Save & Close** when done.

*(Note: You can skip this step—the scanner will still extract data automatically based on pre-configured rules).*

### Step 6: Add to the Queue

Click **Add to Queue**. Your current selection (company, passwords, files, and mapping corrections) is now saved as a job.

You can immediately change the company and select different files to add another job to the queue.

### Step 7: Run the Queue

When you are ready, click **Run Queue**. The scanner will process each job sequentially.

* Monitor progress via the log output and the main status bar.
* If you need to cancel, click **Stop**. The current file will finish its active table, and all remaining jobs in the queue will be cleared.

### Step 8: Open the Output

After all jobs finish, click the **Open Output** button to immediately open the output folder for today’s date.

* The final compiled Excel file is located inside the `Compiled` subfolder.
* You can also review past results via the **History** tab by double‑clicking any historical folder row.

---

## 3. The Job Queue

The queue allows you to line up scans for multiple companies without waiting for each one to finish.

1. Set up **Company A**, select its files, optionally review mappings, and click **Add to Queue**.
2. Change settings to **Company B**, select its respective files, and click **Add to Queue**.
3. Repeat for as many companies as needed.
4. Click **Run Queue** to run all jobs in order.

> 🗑️ **Note:** You can remove an accidental job from the queue by selecting it and clicking **Remove Selected**.

---

## 4. The Extract Headers Tab

This tab creates a catalogue of all column headers found within your files. **It does not extract actual data rows**—only the column names and their guessed categories.

1. Select a company and passwords (same process as the Scan tab).
2. Choose an extraction date.
3. Tick the files you want to inspect.
4. Click **Extract Headers**.

The resulting Excel file will be saved under `Output\HD_extract_<date>_<company>\`. Open this file to see every column header, its position, and its auto-assigned category. This is helpful when creating or updating your master company header dictionary.

---

## 5. Quick Tips

* **Drag & Drop:** You can drag Excel/CSV files directly from Windows File Explorer onto the app's file list area. They will automatically copy into the Watch Folder and refresh the view.
* **Auto‑refresh:** Tick the **Auto‑refresh (5s)** checkbox to have the file list update automatically every 5 seconds.
* **Preview Sheets:** When previewing multi-sheet Excel files, use the sheet selector directly above the preview grid to switch views.
* **Keyboard Shortcut:** In the company code box, type to filter and press `Enter` to select when only one choice remains.

---

## 6. Understanding the Output

After a successful scan, the output directory structure is organized like this within your configured **Output Folder**:

```text
Output Folder\
└── 20260627\
    └── ABC\
        ├── CSVs\                   (Intermediate processing files)
        ├── Archived\               (Original files moved here after scan)
        ├── Unzipped\               (Temporary files from zip archives)
        ├── Compiled\               (Final Excel files go here)
        │   └── OFAC_ABS_Log_20260627_ABC.xlsx
        └── Log_20260627.csv        (Detailed per‑file processing log)

```

### Compiled Excel Structure

The final compiled Excel file contains the following standardized columns:

| Column | Description |
| --- | --- |
| **SURNAME** | Last name |
| **FIRST_NAME** | First name |
| **COMPLETE_NAME** | Full name |
| **SEX** | M, F, or U (Unknown) |
| **DATE_OF_BIRTH** | MM/DD/YYYY format |
| **CMPY_NO** | Company code |
| **POLICY_NUMBER** | Policy number |
| **FILE_PATH** | Full path of the original source file |
| **SHEET** | Which sheet/table the row originated from |

> ⚠️ **Row Limits:** If the total number of rows exceeds 1,000,000, the compiler will automatically split the output into multiple Excel files (`_1.xlsx`, `_2.xlsx`, etc.).

---

## 7. Troubleshooting

| Problem | What to Do |
| --- | --- |
| **No companies appear in the dropdown** | Make sure the passwords CSV file has two columns: `Code` and `Password`. The column headers must match exactly (case-insensitive). |
| **Headers not detected** | Use the **Detect Headers** tool to see what was found. If no columns are assigned, your header dictionary may be missing keywords. Update the `header.xlsx` file in the Header Files Folder. |
| **“Add to Queue” button is greyed out** | You must select at least one file, a company, and at least one active password. |
| **Scan completes but no data extracted** | Check the log for error messages. The password may be incorrect, or the file may be in an unsupported format. |
| **Compiled file is missing data** | Use the **Detect Headers** popup to make sure the correct columns are mapped properly before executing the scan. |
| **Output folder not opening** | The folder is only created *after* a successful scan. If the scan fails, check the log window for errors. |
| **“No files found in watch folder”** | Verify that the watch folder path in Settings is accurate and that your files are not accidentally placed inside a subfolder. |
| **Want to stop scanning midway** | Click **Stop**. The current file will finish processing its active table, and the rest of the queue will safely clear. |

---

*For further help, contact your system administrator or IT support.*
