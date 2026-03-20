import time
import os
from threading import Thread

import pandas as pd
from get_for_ex_trans import (
    get_foreign_exchange_transactions,
    get_immediate_transactions,
)
from add_rows_to_csv import parse_amount
import shutil
from pathlib import Path
from typing import Optional

from playwright.sync_api import sync_playwright, Page, BrowserContext

from config import get_paths, load_env, validate_env, update_state
from utils import checkDate


headless = False


def toggleHeadless(headlessState: int):  # should be 0 (off) or 1 (on)
    global headless
    if headlessState == 1:
        headless = True
    else:
        headless = False


def getExcelFileThreaded(year, month):
    Thread(target=getExcelFile, args=(year, month), daemon=True).start()


def _login_max(page: Page) -> None:
    load_env()
    validate_env(require_max=True, require_leumi=False)
    max_username = os.environ.get("MAX_USERNAME")
    max_password = os.environ.get("MAX_PASSWORD")
    user_id = os.environ.get("ID")

    page.goto("https://www.max.co.il/login", wait_until="networkidle")
    page.wait_for_timeout(1000)
    # Click "login with password" button if it exists
    try:
        page.get_by_text("כניסה עם סיסמה").click()
    except Exception:
        pass

    page.fill("#user-name", max_username or "")
    page.fill("#password", max_password or "")
    page.keyboard.press("Enter")

    # Optional ID field
    try:
        page.wait_for_selector("#idInput input", timeout=3000)
        if user_id:
            page.fill("#idInput input", user_id)
            page.keyboard.press("Enter")
            page.wait_for_timeout(3000)
    except Exception:
        pass


def _go_to_max_transaction_details(page: Page, year: str, month: str) -> None:
    print("Going to transaction details")
    if month == "12":
        month_int = "1"
        year_int = str(int(year) + 1)
    else:
        month_int = str(int(month) + 1)
        year_int = year
    url = (
        "https://www.max.co.il/transaction-details/personal?filter=-1_-1_1_"
        + year_int
        + "-"
        + month_int
        + "-01_0_0_-1&sort=1a_1a_1a_1a_1a_1a"
    )
    page.goto(url, wait_until="networkidle")


def _download_max_excel(page: Page, year: str, month: str) -> Optional[Path]:
    paths = get_paths()
    target_file = (
        paths.max_exports_dir / f"transaction-details_export_{year}_{month}.xlsx"
    )

    with page.expect_download() as download_info:
        page.wait_for_timeout(1000)
        page.locator(".download-excel").click()
    download = download_info.value
    temp_path = Path(download.path())
    shutil.move(str(temp_path), target_file)
    print(f"Downloaded Max Excel to {target_file}")
    return target_file


def _capture_foreign_exchange_table(page: Page, *, foreign_html_path: Path, immediate_html_path: Path) -> None:
    try:
        # locator = page.locator(
        #     "css=app-table.ng-star-inserted:nth-child(6) > div:nth-child(1)"
        # )
        locator = page.locator("app-table").filter(has_text='עסקאות חו"ל ומט"ח')
        locator.wait_for(timeout=10000)
        html = locator.first.inner_html()
        page.wait_for_timeout(10000)
        foreign_html_path.parent.mkdir(parents=True, exist_ok=True)
        foreign_html_path.write_text(html, encoding="utf-8")
        print(f"Deal table HTML saved to {foreign_html_path}")
        # added immediate transactions table
        locator = page.locator("app-table").filter(has_text="עסקאות בחיוב מיידי")
        locator.wait_for(timeout=10000)
        html = locator.first.inner_html()
        page.wait_for_timeout(10000)
        immediate_html_path.parent.mkdir(parents=True, exist_ok=True)
        immediate_html_path.write_text(html, encoding="utf-8")
        print(f"Immediate transactions HTML saved to {immediate_html_path}")
    except Exception as e:
        print("Error capturing table:", e)


def getExcelFile(year, month):
    year, month = checkDate(year, month)
    print(f"date = {year}-{month}")

    paths = get_paths()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context: BrowserContext = browser.new_context(accept_downloads=True)
        page = context.new_page()

        _login_max(page)
        _go_to_max_transaction_details(page, year, month)
        excel_path = _download_max_excel(page, year, month)
        foreign_html_path = paths.data_dir / "foreign_exchange_transactions.html"
        immediate_html_path = paths.data_dir / "immediate_transactions.html"
        _capture_foreign_exchange_table(
            page,
            foreign_html_path=foreign_html_path,
            immediate_html_path=immediate_html_path,
        )

        time.sleep(5)

        context.close()
        browser.close()

    if excel_path is None:
        raise RuntimeError("Failed to download Max Excel file")

    combined_transactions_csv_path = paths.data_dir / "transactions.csv"
    foreign_exchange_transactions_csv_path = (
        paths.data_dir / "foreign_exchange_transactions.csv"
    )
    immediate_transactions_csv_path = paths.data_dir / "immediate_transactions.csv"

    # `get_foreign_exchange_transactions` and `get_immediate_transactions` each write a CSV.
    # If they target the same file, the second call overwrites the first.
    # Write them separately, then concatenate into a single `transactions.csv`.
    get_foreign_exchange_transactions(
        str(foreign_html_path),
        str(foreign_exchange_transactions_csv_path),
    )
    get_immediate_transactions(
        str(immediate_html_path),
        str(immediate_transactions_csv_path),
    )

    foreign_exchange_transactions_dataframe = pd.read_csv(
        foreign_exchange_transactions_csv_path,
        encoding="utf-8-sig",
    )
    immediate_transactions_dataframe = pd.read_csv(
        immediate_transactions_csv_path,
        encoding="utf-8-sig",
    )
    combined_transactions_dataframe = pd.concat(
        [
            foreign_exchange_transactions_dataframe,
            immediate_transactions_dataframe,
        ],
        ignore_index=True,
    )
    combined_transactions_dataframe.to_csv(
        combined_transactions_csv_path,
        index=False,
        encoding="utf-8-sig",
    )

    parse_amount(str(combined_transactions_csv_path), str(excel_path))

    # Remember last downloaded month in state for convenience
    update_state({"last_filled_year": year, "last_filled_month": month})

    return str(excel_path)
