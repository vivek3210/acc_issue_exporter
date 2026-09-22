import argparse
import os
import sys
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PROJECT_PYTHON = BASE_DIR / ".venv" / "Scripts" / "python.exe"

if PROJECT_PYTHON.exists() and Path(sys.executable).resolve() != PROJECT_PYTHON.resolve():
    os.execv(str(PROJECT_PYTHON), [str(PROJECT_PYTHON), str(Path(__file__).resolve()), *sys.argv[1:]])

import pandas as pd
from dotenv import load_dotenv

TARGET_COMPANY = "Rovisys"
OUTPUT_COLUMNS = [
    "Issue number",
    "Title",
    "Description",
    "Location",
    "Created by",
    "Status",
]


def normalized_columns(frame):

    return {
        "".join(character for character in str(column).casefold() if character.isalnum()): column
        for column in frame.columns
    }


def select_column(frame, aliases):

    columns = normalized_columns(frame)
    for alias in aliases:
        column = columns.get(
            "".join(character for character in alias.casefold() if character.isalnum())
        )
        if column is not None:
            return frame[column]

    return pd.Series("", index=frame.index)


def read_export(source):

    if source.suffix.lower() == ".csv":
        for encoding in ("utf-8-sig", "utf-16", "cp1252"):
            try:
                return pd.read_csv(source, encoding=encoding)
            except UnicodeDecodeError:
                continue
        raise ValueError(f"Could not decode CSV export: {source}")

    workbook = pd.ExcelFile(source)
    sheet = "Issues" if "Issues" in workbook.sheet_names else workbook.sheet_names[0]
    return pd.read_excel(source, sheet_name=sheet)


def export_to_excel(source):

    issues = read_export(source).dropna(how="all")
    assigned_to = select_column(issues, ["Assigned to", "Assignee"])
    title = select_column(issues, ["Title", "Issue title"])
    assignment_mask = assigned_to.fillna("").astype(str).str.contains(
        TARGET_COMPANY,
        case=False,
        na=False,
    )
    title_mask = title.fillna("").astype(str).str.lstrip().str.startswith(
        "RBT EPMS",
        na=False,
    )
    issues = issues[assignment_mask | title_mask].copy()

    description = select_column(
        issues,
        ["Description", "Title", "Issue title"],
    )
    title = select_column(issues, ["Title", "Issue title"])
    description = description.fillna("")
    description = description.where(description.astype(str).str.strip().ne(""), title)

    created_by = select_column(
        issues,
        ["Created by", "Created By", "Creator", "Author"],
    )
    created_by = created_by.fillna("").astype(str).str.split("\n").str[0].str.strip()

    issues = pd.DataFrame(
        {
            "Issue number": select_column(
                issues,
                ["Issue number", "Issue ID", "ID", "Number"],
            ),
            "Title": select_column(
                issues,
                ["Title", "Issue title"],
            ),
            "Description": description,
            "Location": select_column(
                issues,
                ["Location", "Location description", "Area"],
            ),
            "Created by": created_by,
            "Status": select_column(
                issues,
                ["Status", "Issue status"],
            ),
        }
    )

    export_dir = BASE_DIR / "exports"
    export_dir.mkdir(exist_ok=True)
    destination = export_dir / (
        f"rovisys_issues_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    )
    with pd.ExcelWriter(destination, engine="xlsxwriter") as writer:
        issues.to_excel(writer, index=False, sheet_name="Issues")
        worksheet = writer.sheets["Issues"]
        row_count, column_count = issues.shape

        worksheet.freeze_panes(1, 0)
        worksheet.set_column("A:A", 16)
        worksheet.set_column("B:B", 55)
        worksheet.set_column("C:C", 70)
        worksheet.set_column("D:D", 45)
        worksheet.set_column("E:E", 28)
        worksheet.set_column("F:F", 16)

        if row_count:
            worksheet.add_table(
                0,
                0,
                row_count,
                column_count - 1,
                {
                    "name": "RovisysIssues",
                    "style": "Table Style Medium 2",
                    "columns": [
                        {"header": column} for column in issues.columns
                    ],
                },
            )

    print(f"\nRovisys-assigned Issues Found: {len(issues)}")
    print(f"Excel file created: {destination}")
    return destination


def main():

    load_dotenv(BASE_DIR / ".env")
    parser = argparse.ArgumentParser(
        description="Convert an Autodesk ACC Issues export to Excel."
    )
    parser.add_argument(
        "--file",
        type=Path,
        help="Existing ACC CSV/XLSX export to convert.",
    )
    args = parser.parse_args()

    if args.file:
        source = Path(str(args.file).strip().strip('"')).expanduser().resolve()
    else:
        entered_path = input("Enter the CSV or Excel filename/path: ").strip()
        entered_path = entered_path.strip('"')
        source = Path(entered_path).expanduser()
        if not source.is_absolute():
            source = BASE_DIR / source
        source = source.resolve()

    if not source.is_file():
        raise FileNotFoundError(f"Export file does not exist: {source}")

    export_to_excel(source)


if __name__ == "__main__":
    main()
