# app.py

import os
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

load_dotenv()

FORMA_PROJECT_URL = os.getenv("FORMA_PROJECT_URL")
AUTH_FILE = "auth.json"


def export_to_excel(issues):

    if not os.path.exists("exports"):
        os.makedirs("exports")

    filename = (
        f"exports/rovisys_issues_"
        f"{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    )

    df = pd.DataFrame(issues)

    df.to_excel(
        filename,
        index=False
    )

    print(f"\nExcel file created: {filename}")

    return filename


def scrape_issues():

    results = []

    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=False
        )

        # Reuse saved Autodesk login
        if os.path.exists(AUTH_FILE):

            context = browser.new_context(
                storage_state=AUTH_FILE
            )

        else:

            context = browser.new_context()

            print("\nNo saved login found.")
            print("Please log into Autodesk manually.")

        page = context.new_page()

        page.goto(
            FORMA_PROJECT_URL,
            wait_until="networkidle"
        )

        # First-run login workflow
        if not os.path.exists(AUTH_FILE):

            input(
                "\nAfter logging into Forma press ENTER..."
            )

            context.storage_state(
                path=AUTH_FILE
            )

            print("Login session saved.")

        print("Loading Issues page...")

        # OPTIONAL:
        # Replace with your actual Issues URL if needed
        # page.goto(
        #     f"{FORMA_PROJECT_URL}/issues",
        #     wait_until="networkidle"
        # )

        page.wait_for_timeout(5000)

        # ------------------------------------------------------------------
        # Update these selectors after inspecting the Forma issue table
        # ------------------------------------------------------------------

        issue_rows = page.locator(
            "[data-testid='issue-row']"
        )

        count = issue_rows.count()

        print(f"Found {count} issue rows")

        for i in range(count):

            try:

                row = issue_rows.nth(i)

                title = row.locator(
                    ".issue-title"
                ).inner_text()

                status = row.locator(
                    ".issue-status"
                ).inner_text()

                assignee = row.locator(
                    ".assignee"
                ).inner_text()

                company = row.locator(
                    ".company-name"
                ).inner_text()

                if company.strip().lower() != "rovisys":
                    continue

                issue = {
                    "Title": title,
                    "Status": status,
                    "Assignee": assignee,
                    "Company": company,
                }

                results.append(issue)

            except Exception as ex:

                print(
                    f"Row {i} skipped: {ex}"
                )

        browser.close()

    return results


def main():

    if not FORMA_PROJECT_URL:

        raise ValueError(
            "FORMA_PROJECT_URL is missing from .env"
        )

    issues = scrape_issues()

    print(
        f"\nRoviSys Issues Found: {len(issues)}"
    )

    export_to_excel(issues)


if __name__ == "__main__":
    main()