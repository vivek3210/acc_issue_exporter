import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from gooey import Gooey
from gooey import GooeyParser
from argparse import ArgumentParser

# FIX: Check if running as a compiled PyInstaller executable
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

PROJECT_PYTHON = BASE_DIR / ".venv" / "Scripts" / "python.exe"

# Only attempt virtual env handoff if running as a raw script
if not getattr(sys, 'frozen', False) and PROJECT_PYTHON.exists() and Path(sys.executable).resolve() != PROJECT_PYTHON.resolve():
    os.execv(str(PROJECT_PYTHON), [str(PROJECT_PYTHON), str(Path(__file__).resolve()), *sys.argv[1:]])


import pandas as pd
from dotenv import load_dotenv

CACHE_FILE = BASE_DIR / "input_cache.json"
DEFAULT_RBT_EPMS_NAME = "RBT EPMS"


def load_input_cache():

    try:
        with CACHE_FILE.open(encoding="utf-8") as cache_file:
            cache = json.load(cache_file)
    except (OSError, json.JSONDecodeError):
        return {}

    return cache if isinstance(cache, dict) else {}


def save_input_cache(target_company, include_rbt_epms, rbt_epms_name):

    cache = {
        "target_company": target_company,
        "include_rbt_epms": include_rbt_epms,
        "rbt_epms_name": rbt_epms_name,
    }
    try:
        with CACHE_FILE.open("w", encoding="utf-8") as cache_file:
            json.dump(cache, cache_file, indent=2)
    except OSError:
        pass


def prompt_for_filters(reset_cache=False, target_company=None,
                       include_rbt_epms=False, rbt_epms_name=None):

    cache = {} if reset_cache else load_input_cache()
    cached_company = str(cache.get("target_company", "")).strip()
    target_company = str(target_company or cached_company).strip()
    while not target_company:
        raise ValueError("A company name is required.")

    include_rbt_epms = bool(include_rbt_epms)

    rbt_epms_name = str(rbt_epms_name or DEFAULT_RBT_EPMS_NAME).strip()
    if not rbt_epms_name:
        rbt_epms_name = DEFAULT_RBT_EPMS_NAME
    if not include_rbt_epms:
        rbt_epms_name = DEFAULT_RBT_EPMS_NAME

    save_input_cache(target_company, include_rbt_epms, rbt_epms_name)
    return target_company, include_rbt_epms, rbt_epms_name

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
    title_mask = (
        title.fillna("").astype(str).str.lstrip().str.startswith(
            RBT_EPMS_NAME,
            na=False,
        )
        if INCLUDE_RBT_EPMS
        else pd.Series(False, index=issues.index)
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
            "Company": select_column(
                issues,
                ["Company", "Organization"],
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

@Gooey(
    program_name="ACC Issue Exporter",
    program_description="Filter an Autodesk ACC Issues export and create an Excel file.",
    default_size=(700, 550),
    clear_before_run=True,
    show_success_modal=False,
)
def main():

    load_dotenv(BASE_DIR / ".env")
    parser = GooeyParser(
        description="Convert an Autodesk ACC Issues export to Excel."
    )
    parser.add_argument(
        "--file",
        type=Path,
        help="Existing ACC CSV/XLSX export to convert.",
        widget="FileChooser",
        gooey_options={"wildcard": "ACC exports (*.csv;*.xls;*.xlsx)|*.csv;*.xls;*.xlsx"},
        required=True,
    )
    parser.add_argument(
        "--company",
        dest="target_company",
        help="Company name to match in ACC.",
        required=True,
    )
    parser.add_argument(
        "--include-rbt-epms",
        action="store_true",
        help="Include issues whose title starts with the RBT EPMS name.",
    )
    parser.add_argument(
        "--rbt-epms-name",
        default=DEFAULT_RBT_EPMS_NAME,
        help="Title prefix used for RBT EPMS issues.",
    )
    parser.add_argument(
        "--reset-cache",
        action="store_true",
        help="Forget cached company and RBT EPMS answers before prompting.",
    )
    args = parser.parse_args()

    global TARGET_COMPANY, INCLUDE_RBT_EPMS, RBT_EPMS_NAME
    TARGET_COMPANY, INCLUDE_RBT_EPMS, RBT_EPMS_NAME = prompt_for_filters(
        reset_cache=args.reset_cache,
        target_company=args.target_company,
        include_rbt_epms=args.include_rbt_epms,
        rbt_epms_name=args.rbt_epms_name,
    )

    source = Path(str(args.file).strip().strip('"')).expanduser().resolve()

    if not source.is_file():
        raise FileNotFoundError(f"Export file does not exist: {source}")

    export_to_excel(source)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ Script failed with error: {e}")
    finally:
        print("\n" + "-"*40)
        print("Process finished.")
