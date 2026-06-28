Here is the entire guide wrapped in a clean, copying-friendly raw block so you can paste it directly into your GitHub `README.md` or wiki page without losing any formatting:

```markdown
# OFAC Scanner – Complete Setup Guide (Windows)

No prior knowledge of Python or coding needed. Just follow these steps to get everything running cleanly using Visual Studio Code.

---

## 1. Install Python

1. Open a web browser and go to [python.org](https://python.org).
2. Click the yellow **Download Python** button (the latest version is fine).
3. Run the downloaded file.
4. > ⚠️ **CRITICAL STEP:** On the first screen, check the box that says **"Add Python to PATH"** at the bottom of the window. If you miss this, the rest of the guide will not work.
5. Click **Install Now**.
6. Wait for the installation to finish and close the window.

> ⚠️ **YOU CAN SKIP THIS PART IF YOU ALREADY HAVE PYTHON INSTALLED**

---

## 2. Install VS Code & Python Extension

1. Go to [code.visualstudio.com/download](https://code.visualstudio.com/download).
2. Click the **Windows** download button, run the installer, and accept the default settings.
3. Open **Visual Studio Code**.
4. Click on the **Extensions** icon on the far left sidebar (it looks like 4 blocks/squares).
5. In the search bar, type `Python`.
6. Find the official **Python** extension by Microsoft and click **Install**.

---

## 3. Open the Project Folder

1. In VS Code, click **File** → **Open Folder**.
2. Browse to the folder where your OFAC Scanner files are located (e.g., `C:\OFAC\Scanner`). 
   * *Note: If you received a ZIP file, extract it completely to this location first.*
3. Click **Select Folder**.
4. You should now see your project files (like `gui/`, `utils/`, `engine/`) listed on the left side explorer.

---

## 4. Set Up & Activate a Virtual Environment (`.venv`)

Using a virtual environment ensures the scanner's libraries don't conflict with other files on your computer.

### Step A: Create the Environment
1. Open a new terminal in VS Code: click **Terminal** → **New Terminal** from the top menu.
2. At the bottom of the screen, type the following command and press **Enter**:
```bash
python -m venv .venv

```

3. Wait a few seconds. You will notice a new folder named `.venv` appear in your left-hand file explorer.

### Step B: Activate the Environment

VS Code will usually pop up a notification in the bottom right asking: *"We noticed a new environment has been created. Do you want to select it for the workspace folder?"* Click **Yes**.

If that popup does not appear, or if your terminal line doesn't update, you must activate it manually. Look at the right side of your terminal bar to see if you are using **PowerShell** or **Command Prompt (cmd)**, then run the matching command below:

* **For PowerShell (Default in VS Code):**

```powershell
.\.venv\Scripts\Activate.ps1

```

*If you get an error saying "script execution is disabled on this system", paste this command first, hit Enter, and try the activation command again:*

```powershell
Set-ExecutionPolicy RemoteSigned -Scope Process

```

* **For Command Prompt (cmd):**

```cmd
.\.venv\Scripts\activate.bat

```

> 🌟 **How to know it worked:** You should now see **`(.venv)`** written at the very beginning of your terminal command line (e.g., `(.venv) PS C:\OFAC\Scanner>`). This means your isolated environment is active!

---

## 5. Install the Required Python Libraries

With `(.venv)` explicitly showing at the start of your terminal line, type the following command and press **Enter**:

```bash
pip install -r requirements.txt

```

Wait until you see a message saying "Successfully installed...".

---

## 6. Prepare Your Working Folders

The scanner needs a few folders to process files. Create these folders now using Windows File Explorer:

* **Watch Folder** – Where you drop the Excel/CSV files to be scanned.
* *Example:* `C:\OFAC\Incoming`


* **Output Folder** – Where final results will be saved.
* *Example:* `C:\OFAC\Output` *(or a network shared path like `\\server\share\Output`)*


* **Header Files Folder** – Contains the Excel dictionary files that map your column fields (Names, Dates of Birth, Policy Numbers).
* *Example:* `C:\OFAC\Headers`



Inside the Headers folder, you must have a base `header.xlsx` file. If you have company-specific mappings, place them in a subfolder named `by_company`:

```text
C:\OFAC\Headers\
├── header.xlsx
└── by_company\
    ├── header_ABC.xlsx
    └── header_XYZ.xlsx

```

---

## 7. Create the Company Passwords CSV

1. Open **Notepad** on your computer.
2. Type your company codes and passwords exactly like this (no spaces after commas):

```text
Code,Password
ABC,pass123
XYZ,secret99

```

3. Click **File** → **Save As**.
4. Change *Save as type* to **All Files (*.*)**.
5. Name the file `passwords.csv` and save it to a clean location (e.g., `C:\OFAC\passwords.csv`).

---

## 8. Run the Scanner & Configure Settings

1. In the VS Code left sidebar, expand the `gui` folder and click on `app.py` to open it.
2. Look at the top right corner of VS Code—you should see a **Play button** (Run Python File). Click it.
* *Alternative:* Type the following into your active terminal and hit Enter:



```bash
python gui/app.py

```

3. The scanner window will pop up. Go straight to the **Settings** tab.
4. Fill out the paths exactly as you set them up:

| Field | What to Enter |
| --- | --- |
| **Input (Watch) Folder** | `C:\OFAC\Incoming` |
| **Company Passwords CSV** | `C:\OFAC\passwords.csv` |
| **Output Folder** | `C:\OFAC\Output` (or your network path) |
| **Header Files Folder** | `C:\OFAC\Headers` |
| **Theme** | Leave as "flatly" or pick a custom style |

5. Click **Save Settings**. The app will remember these paths permanently.

---

## 9. Test the Scanner

1. Drop a sample Excel file containing dummy customer details (names, DOBs, policy numbers) into your **Watch Folder**.
2. Switch to the **Scan** tab in the application.
3. Click **Refresh**—your sample file should appear in the list.
4. Select your company code from the dropdown menu.
5. Check the boxes next to the password, the file, and click **Detect Headers** to confirm the scanner reads your columns accurately.
6. Click **Add to Queue**, then click **Run Queue**.
7. Once finished, click **Open Output** to view your results.

---

## 10. Desktop Shortcut for Daily Use (Optional)

To run the scanner without opening VS Code every morning:

1. Right-click on your desktop → **New** → **Shortcut**.
2. In the location box, paste the path to your virtual environment's Python launcher, followed directly by the script path:

```text
"C:\OFAC\Scanner\.venv\Scripts\pythonw.exe" "C:\OFAC\Scanner\gui\app.py"

```

*(Using `pythonw.exe` ensures the scanner opens cleanly without leaving a blank black command prompt window open behind it).* 3. Click **Next**, name it **OFAC Scanner**, and click **Finish**.

---

## Troubleshooting

| Problem | Fix |
| :--- | :--- |
| **"Terminal shows errors about Script Execution Policies"** | Run `Set-ExecutionPolicy RemoteSigned -Scope Process` in your VS Code terminal window. Alternatively, open Windows PowerShell as Administrator and run:<br><br>```powershell<br>Set-ExecutionPolicy RemoteSigned -Scope CurrentUser<br>```<br>then restart VS Code. |
| **"No module named..."** | Make sure `(.venv)` is explicitly visible at the start of your terminal line before running:<br><br>```bash<br>pip install -r requirements.txt<br>``` |
| **"Headers not detected"** | Double check that your base `header.xlsx` sheet names or cell keywords match the targets exactly (e.g., name, dob, policynum). |
