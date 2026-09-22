# Forma Issue Exporter

This tool converts an Autodesk Construction Cloud (ACC) Issues report into a clean, filterable Excel workbook.

The ACC login and report generation are completed manually in your normal browser. The Python script then reads the downloaded report and creates the final workbook.

## What Gets Exported

The output includes issues that meet either condition:

- `Assigned to` contains `Rovisys`.
- The issue `Title` starts with `RBT EPMS`.

Duplicate issues are included only once.

The final workbook contains these columns:

1. Issue number
2. Title
3. Description
4. Location
5. Created by
6. Status

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

The script automatically switches to `.venv` when it is run with system Python, but using the explicit command above is the most reliable option.

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

## Run the Converter

From the project directory, run:

```powershell
.\.venv\Scripts\python.exe app.py
```

When prompted, enter the full path to the downloaded report. Quoted Windows paths are supported:

```text
"C:\Users\vivek.darji\Downloads\Issue summary-202609221901.xlsx"
```

You can also pass the file directly:

```powershell
.\.venv\Scripts\python.exe app.py --file "C:\Users\vivek.darji\Downloads\Issue summary-202609221901.xlsx"
```

The script accepts `.xlsx`, `.xls`, and `.csv` input files.

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
.\.venv\Scripts\python.exe app.py
```

### `Export file does not exist`

Check the path entered at the prompt. Include the file extension and make sure the report has finished downloading. For a path containing spaces, surrounding it with double quotes is supported.

### The output contains zero issues

Check that the ACC report is the Issue summary for the intended project and that the source workbook contains an `Issues` worksheet. The converter includes rows assigned to `Rovisys` or rows whose title begins with `RBT EPMS`.

### Excel reports that it repaired the workbook

Regenerate the file with the current version of the script and ensure `XlsxWriter` is installed from `requirements.txt`. The current writer creates the Excel table and filter metadata in one step.

## Security

The exporter does not need Autodesk credentials. Sign in through your normal browser and download the report manually. Do not store Autodesk passwords in `.env`, source files, or the repository.
