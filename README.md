# ACC Issue Exporter

This tool converts an Autodesk Construction Cloud (ACC) Issues report into a clean, filterable Excel workbook.

The ACC login and report generation are completed manually in your normal browser. The exporter then reads the downloaded report and creates the final workbook through a desktop GUI.

## What Gets Exported

The output includes issues that meet either condition:

- `Assigned to` contains the inputted name.
- The issue `Title` starts with `RBT EPMS`, if this field is selected as yes.

Duplicate issues are included only once.

The final workbook contains these columns:

1. Issue number
2. Title
3. Description
4. Location
5. Company
6. Created by
7. Status

The output is formatted as an Excel table with filter dropdowns, including a filter on `Status` for Open, Closed, Draft, and other statuses. The header row is frozen for easier scrolling.

## Requirements

- Windows
- Python 3.14 or another supported Python 3 version
- Access to the ACC (Autodesk Construction Cloud) project
- An ACC Issues report downloaded as `.xlsx`, `.xls`, or `.csv`

The required Python packages are listed in `requirements.txt`.

## One-Time Setup

Open PowerShell in the project directory:

```powershell
cd C:\Users\YOURROVISYSUSERNAME\Desktop\forma_issue_exporter
```

If the virtual environment does not exist, create it:

```powershell
python -m venv .venv
```

Install the dependencies:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

The script automatically switches to `.venv` when it is run with system Python, but using the explicit command above is the most reliable option. The `Gooey` package provides the desktop interface.

## Create the ACC Report

1. Open Autodesk Construction Cloud in your normal browser.
2. Navigate to the project **Issues** page.
3. Click **Export all**.
4. In the report setup dialog, select **Issue summary**.
5. Enter any report title; the title does not affect the conversion.
6. For **File format**, select **Excel**.
7. Complete the initial setup and click **Run report**.
8. Wait for the report to finish processing.
9. When the download link appears, click it and save the Excel file somewhere easy to find, such as your Downloads folder.

Do not rename the worksheet or edit the source report before running the converter. If the workbook contains an `Issues` worksheet, the script uses that worksheet automatically.

## Run the GUI Converter

From the project directory, run:

```powershell
.\.venv\Scripts\python.exe acc_issue_exporter.py
```

The **ACC Issue Exporter** window provides these controls:

- **File**: choose the downloaded ACC `.csv` or Excel report with the file picker.
- **Company**: enter the company name as it appears in the report's `Assigned to` column.
- **Include RBT EPMS**: select this to also include issues whose title starts with the configured RBT EPMS name.
- **RBT EPMS name**: optionally change the title prefix used by that filter.
- **Reset cache**: clear saved filter values before running.

Click **Start** to create the workbook. The GUI shows the job status while the report is being processed and returns to a completed state after the Excel file is written. The application no longer waits for an extra command-line confirmation after finishing.

The command-line options are also available when launching the script directly:

```powershell
python acc_issue_exporter.py --file "C:\Users\vivek.darji\Downloads\Issue summary-202609221901.xlsx" --company "Rovisys"
```

The script accepts `.xlsx`, `.xls`, and `.csv` input files. When using the GUI, use the file picker rather than pasting a quoted path into the field.

The answers for the company and RBT EPMS filters are stored in `input_cache.json` beside the script or executable. The `--reset-cache` option clears that file before the next run:

```powershell
python acc_issue_exporter.py --reset-cache --file "C:\Users\vivek.darji\Downloads\Issue summary-202609221901.xlsx" --company "Rovisys"
```

### Run the Windows Executable

A packaged Windows executable can be downloaded from the GitHub Releases page. The release workflow produces:

```text
acc_issue_exporter.exe
```

Copy the executable to the folder where you want the converter to run, then double-click it or launch it from PowerShell. It opens the same GUI and creates the `exports` folder beside itself. The executable does not require Python or the project virtual environment.

The executable also accepts the command-line options described above:

```powershell
.\acc_issue_exporter.exe --file "C:\Users\vivek.darji\Downloads\Issue summary-202609221901.xlsx" --company "Rovisys"
```

## Rebuild the Executable

Install the project dependencies and PyInstaller in the virtual environment, then run:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt pyinstaller
.\.venv\Scripts\pyinstaller.exe app.spec --clean --noconfirm
```

The rebuilt executable is written to `dist\app.exe` when using `app.spec`. The GitHub Actions release workflow builds `dist\acc_issue_exporter.exe` directly from `acc_issue_exporter.py` and publishes that file.

## Output

The finished workbook is saved in the project's `exports` folder:

```text
exports\rovisys_issues_YYYYMMDD_HHMMSS.xlsx
```

Open the workbook in Excel and use the filter arrow in the `Status` column to show only Open, Closed, or another status.

## Troubleshooting

### `ModuleNotFoundError: No module named 'xlsxwriter'`

Install dependencies into the project virtual environment and run the script with that interpreter:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
python acc_issue_exporter.py
```

### `Export file does not exist`

Use the GUI file picker and select the downloaded report. Include the file extension and make sure the report has finished downloading. When using the command line, surround paths containing spaces with double quotes.

### The GUI stays on `Running`

Use the current version of the script or executable. The exporter must be allowed to return after writing the workbook; it should not wait for an extra `Press ENTER` prompt. A successful run prints the output path and finishes automatically.

### The output contains zero issues

Check that the ACC report is the Issue summary for the intended project and that the source workbook contains an `Issues` worksheet. The converter includes rows assigned to the company name you entered or, when enabled, rows whose title begins with the configured RBT EPMS name.

### Excel reports that it repaired the workbook

Regenerate the file with the current version of the script and ensure `XlsxWriter` is installed from `requirements.txt`. The current writer creates the Excel table and filter metadata in one step.

## Security

The exporter does not need Autodesk credentials. Sign in through your normal browser and download the report manually. Do not store Autodesk passwords in `.env`, source files, or the repository.

# TLDR

- Download `ACCIssueExporter.exe` from the Releases page and copy it to any folder.
- Download an ACC Issues report as described in the Create the ACC Report section.
- Open the exporter, choose the report, enter the company name, and select the RBT EPMS option if needed.
- Click **Start** and wait for the GUI to show completion.
- Open the new `exports` folder beside the executable to find the Excel workbook.
- Done!
