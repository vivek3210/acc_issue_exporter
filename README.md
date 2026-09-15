# Autodesk Forma Issue Scraper

A Python application that uses Playwright browser automation to extract issues from Autodesk Forma, filter issues assigned to **RoviSys**, and export the results to an Excel spreadsheet.

---

## Features

- Automated browser interaction using Playwright
- Autodesk Forma login support
- Session persistence using `auth.json`
- Filters issues assigned to **RoviSys**
- Exports results to Excel (`.xlsx`)
- Supports scheduled execution
- No Autodesk API access required

---

## Requirements

### Software

- Python 3.10+
- Google Chrome
- Internet access to Autodesk Forma

Verify Python version:

```bash
python --version
```

Example:

```text
Python 3.11.8
```

---

## Project Structure

```text
forma_scraper/
│
├── app.py
├── auth.json
├── .env
├── requirements.txt
├── exports/
│   └── rovisys_issues_YYYYMMDD.xlsx
│
└── README.md
```

---

## Installation

### Create Virtual Environment

```bash
python -m venv venv
```

Activate the environment:

**Windows**

```bash
venv\Scripts\activate
```

**Mac/Linux**

```bash
source venv/bin/activate
```

---

### Install Dependencies

```bash
pip install -r requirements.txt
```

---

## requirements.txt

Create a file named `requirements.txt`:

```text
playwright
pandas
openpyxl
python-dotenv
```

Install:

```bash
pip install -r requirements.txt
```

---

## Install Playwright Browser

Install Chromium:

```bash
python -m playwright install chromium
```

Or install all supported browsers:

```bash
python -m playwright install
```

---

## Corporate Environment / Browser Installation Issues

If Playwright cannot download Chromium due to proxy, firewall, or corporate restrictions, use your locally installed Chrome browser.

Replace:

```python
browser = p.chromium.launch(
    headless=False
)
```

With:

```python
browser = p.chromium.launch(
    channel="chrome",
    headless=False
)
```

---

## Environment Variables

Create a `.env` file in the project root.

```env
FORMA_PROJECT_URL=https://app.autodeskforma.com/projects/YOUR_PROJECT_ID
```

Replace `YOUR_PROJECT_ID` with your Autodesk Forma project URL.

---

## First-Time Authentication

Run:

```bash
python app.py
```

A browser window will open.

1. Sign in to Autodesk Forma.
2. Complete MFA if required.
3. Return to the terminal.
4. Press **ENTER**.

The session will be saved to:

```text
auth.json
```

Future executions will reuse this saved session.

---

## Running the Application

```bash
python app.py
```

Example output:

```text
Loading Issues page...
Found 42 issue rows

RoviSys Issues Found: 17

Excel file created:
exports/rovisys_issues_20260915_113000.xlsx
```

---

## Output Location

Generated Excel files are saved to:

```text
exports/
```

Example:

```text
exports/
└── rovisys_issues_20260915_113000.xlsx
```

---

## Updating HTML Selectors

Autodesk Forma is a dynamic web application and selectors may need to be updated.

Current placeholder selectors:

```python
[data-testid='issue-row']
.issue-title
.issue-status
.assignee
.company-name
```

To identify the correct selectors:

1. Open Autodesk Forma.
2. Press **F12**.
3. Navigate to the **Elements** tab.
4. Locate an issue row.
5. Copy the relevant CSS selectors.

Update the selectors in `app.py`.

---

## Scheduling Daily Exports

### Windows Task Scheduler

Program:

```text
python
```

Arguments:

```text
C:\Scripts\forma_scraper\app.py
```

Start In:

```text
C:\Scripts\forma_scraper
```

Example schedule:

```text
Daily
06:00 AM
```

---

## Troubleshooting

### BrowserType.launch Error

Error:

```text
BrowserType.launch:
Executable doesn't exist
```

Solution:

```bash
python -m playwright install chromium
```

Or use local Chrome:

```python
browser = p.chromium.launch(
    channel="chrome",
    headless=False
)
```

---

### Authentication Issues

Delete:

```text
auth.json
```

Run again:

```bash
python app.py
```

Log in manually and save a new session.

---

### Excel File Not Created

Verify:

- `exports` folder exists
- User has write permissions
- No file is currently open in Excel

---

## Dependencies

### Playwright

Browser automation framework used to interact with Autodesk Forma.

```bash
pip install playwright
```

### Pandas

Used to manipulate issue data and export to Excel.

```bash
pip install pandas
```

### OpenPyXL

Excel workbook creation and formatting.

```bash
pip install openpyxl
```

### Python Dotenv

Loads environment variables from a `.env` file.

```bash
pip install python-dotenv
```

---

## Future Enhancements

- Microsoft Teams notifications
- Email reports
- Historical issue tracking
- Auto-refresh issue exports
- Dashboard worksheet
- Screenshot capture
- Network-call extraction instead of HTML scraping
- Automated scheduled execution

---

## Quick Start

```bash
python -m venv venv

venv\Scripts\activate

pip install -r requirements.txt

python -m playwright install chromium

python app.py
```

The application will open Autodesk Forma, collect issues assigned to **RoviSys**, and export them to an Excel file in the `exports` folder.
