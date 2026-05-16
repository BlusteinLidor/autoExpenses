import time
import os
from threading import Thread

import pandas as pd
from get_for_ex_trans import (
    get_foreign_exchange_transactions,
    get_immediate_transactions,
    write_empty_transactions_csv,
)
from add_rows_to_csv import parse_amount
import shutil
from pathlib import Path
from typing import Any, Optional

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
    try:
        page.get_by_text("כניסה עם סיסמה").click()
    except Exception:
        pass

    page.fill("#user-name", max_username or "")
    page.fill("#password", max_password or "")
    page.keyboard.press("Enter")

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
    page.wait_for_timeout(2000)


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


FOREIGN_TABLE_LABEL = 'עסקאות חו"ל ומט"ח'
IMMEDIATE_TABLE_LABEL = "עסקאות בחיוב מיידי"


def _label_variants(table_label: str) -> tuple[str, ...]:
    """Quote variants Max may use in section headings."""
    variants = {table_label}
    if '"' in table_label:
        variants.add(table_label.replace('"', "'"))
        variants.add(table_label.replace('"', "״"))
    return tuple(variants)


def _scroll_transaction_details_page(page: Page) -> None:
    """Scroll the full page so lazy-loaded Max sections can render."""
    try:
        page.evaluate(
            """async () => {
                const delay = (ms) => new Promise((r) => setTimeout(r, ms));
                const step = Math.max(280, Math.floor(window.innerHeight * 0.75));
                const maxY = Math.max(
                    document.body.scrollHeight,
                    document.documentElement.scrollHeight
                );
                window.scrollTo(0, 0);
                await delay(350);
                for (let y = 0; y <= maxY; y += step) {
                    window.scrollTo(0, y);
                    await delay(280);
                }
                window.scrollTo(0, maxY);
                await delay(450);
                window.scrollTo(0, 0);
                await delay(250);
            }"""
        )
    except Exception:
        for _ in range(10):
            page.mouse.wheel(0, 900)
            page.wait_for_timeout(350)
    page.wait_for_timeout(600)


def _section_label_on_page(page: Page, table_label: str) -> bool:
    for variant in _label_variants(table_label):
        try:
            if page.get_by_text(variant, exact=False).count() > 0:
                return True
        except Exception:
            continue
    return False


def _table_locator_for_label(page: Page, table_label: str):
    locator = page.locator("app-table").filter(has_text=table_label)
    if locator.count() > 0:
        return locator
    for variant in _label_variants(table_label):
        if variant == table_label:
            continue
        alt = page.locator("app-table").filter(has_text=variant)
        if alt.count() > 0:
            return alt
    return locator


def _is_table_section_absent(
    page: Page,
    table_label: str,
    *,
    page_pre_scrolled: bool = False,
) -> bool:
    """
    Return True only when the section heading is not on the page after a full scroll.
    """
    if not page_pre_scrolled:
        _scroll_transaction_details_page(page)
    if not _section_label_on_page(page, table_label):
        return True

    locator = _table_locator_for_label(page, table_label)
    if locator.count() == 0:
        return True

    try:
        locator.first.wait_for(state="visible", timeout=4000)
        return False
    except Exception:
        return True


def _scroll_table_into_view(page: Page, locator) -> None:
    try:
        locator.first.scroll_into_view_if_needed(timeout=5000)
    except Exception:
        pass
    try:
        page.evaluate(
            """(el) => {
                let node = el;
                while (node) {
                    if (node.scrollHeight > node.clientHeight + 20) {
                        node.scrollTop = node.scrollHeight;
                    }
                    node = node.parentElement;
                }
            }""",
            locator.first.element_handle(),
        )
    except Exception:
        pass
    page.wait_for_timeout(1500)


def _expand_table_rows(page: Page, locator) -> int:
    """Scroll and click 'show more' until row count stabilizes."""
    _scroll_table_into_view(page, locator)
    previous_count = -1
    stable_rounds = 0
    for _ in range(12):
        try:
            row_count = locator.locator("div.row.body, motion.div.row.body").count()
        except Exception:
            row_count = 0

        if row_count == previous_count:
            stable_rounds += 1
        else:
            stable_rounds = 0
        previous_count = row_count
        if stable_rounds >= 2:
            break

        clicked_more = False
        for label in ("הצג עוד", "טען עוד", "עוד", "Show more"):
            try:
                more_button = locator.get_by_text(label, exact=False)
                if more_button.count() > 0 and more_button.first.is_visible():
                    more_button.first.click()
                    clicked_more = True
                    page.wait_for_timeout(1200)
                    break
            except Exception:
                continue

        if not clicked_more:
            try:
                page.evaluate(
                    """(el) => {
                        let node = el;
                        while (node) {
                            if (node.scrollHeight > node.clientHeight + 20) {
                                node.scrollTop += Math.max(200, node.clientHeight * 0.8);
                            }
                            node = node.parentElement;
                        }
                    }""",
                    locator.first.element_handle(),
                )
            except Exception:
                pass
            page.wait_for_timeout(800)

    try:
        return locator.locator("motion.div.row.body, div.row.body").count()
    except Exception:
        return 0


def _capture_table_html(
    page: Page,
    *,
    table_label: str,
    html_path: Path,
    csv_path: Path,
    page_pre_scrolled: bool = False,
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "label": table_label,
        "html_path": str(html_path),
        "csv_path": str(csv_path),
        "captured": False,
        "absent": False,
        "row_count": 0,
        "parsed_count": 0,
        "error": None,
    }
    html_path.parent.mkdir(parents=True, exist_ok=True)

    if _is_table_section_absent(
        page, table_label, page_pre_scrolled=page_pre_scrolled
    ):
        result["absent"] = True
        write_empty_transactions_csv(str(csv_path))
        print(f"{table_label}: section not on Max page for this month; skipping.")
        return result

    try:
        locator = _table_locator_for_label(page, table_label)
        locator.first.wait_for(state="visible", timeout=20000)
        _scroll_table_into_view(page, locator)
        row_count = _expand_table_rows(page, locator)
        html = locator.first.inner_html()
        html_path.write_text(html, encoding="utf-8")
        result["captured"] = True
        result["row_count"] = row_count
        print(f"{table_label}: captured {row_count} visible row(s) -> {html_path}")
    except Exception as error:
        if _is_table_section_absent(page, table_label):
            result["absent"] = True
            write_empty_transactions_csv(str(csv_path))
            print(f"{table_label}: section not on Max page for this month; skipping.")
            return result
        result["error"] = str(error)
        print(f"Error capturing {table_label}: {error}")
    return result


def _capture_foreign_exchange_table(
    page: Page,
    *,
    foreign_html_path: Path,
    immediate_html_path: Path,
    foreign_csv_path: Path,
    immediate_csv_path: Path,
) -> dict[str, Any]:
    _scroll_transaction_details_page(page)
    foreign_result = _capture_table_html(
        page,
        table_label=FOREIGN_TABLE_LABEL,
        html_path=foreign_html_path,
        csv_path=foreign_csv_path,
        page_pre_scrolled=True,
    )
    immediate_result = _capture_table_html(
        page,
        table_label=IMMEDIATE_TABLE_LABEL,
        html_path=immediate_html_path,
        csv_path=immediate_csv_path,
        page_pre_scrolled=True,
    )
    return {
        "foreign": foreign_result,
        "immediate": immediate_result,
    }


def _parse_captured_table_html(
    capture_result: dict[str, Any],
    *,
    html_path: Path,
    csv_path: Path,
    parse_html,
) -> int:
    if capture_result.get("absent"):
        return int(capture_result.get("parsed_count") or 0)
    if not capture_result.get("captured") or not html_path.exists():
        write_empty_transactions_csv(str(csv_path))
        return 0
    return parse_html(str(html_path), str(csv_path))


def getExcelFile(year, month):
    year, month = checkDate(year, month)
    print(f"date = {year}-{month}")

    paths = get_paths()
    month_dir = paths.data_dir / f"{year}_{month}"
    month_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context: BrowserContext = browser.new_context(accept_downloads=True)
        page = context.new_page()

        _login_max(page)
        _go_to_max_transaction_details(page, year, month)
        excel_path = _download_max_excel(page, year, month)
        foreign_html_path = month_dir / "foreign_exchange_transactions.html"
        immediate_html_path = month_dir / "immediate_transactions.html"
        foreign_exchange_transactions_csv_path = (
            month_dir / "foreign_exchange_transactions.csv"
        )
        immediate_transactions_csv_path = month_dir / "immediate_transactions.csv"
        capture_report = _capture_foreign_exchange_table(
            page,
            foreign_html_path=foreign_html_path,
            immediate_html_path=immediate_html_path,
            foreign_csv_path=foreign_exchange_transactions_csv_path,
            immediate_csv_path=immediate_transactions_csv_path,
        )

        time.sleep(2)

        context.close()
        browser.close()

    if excel_path is None:
        raise RuntimeError("Failed to download Max Excel file")

    combined_transactions_csv_path = month_dir / "transactions.csv"
    foreign_exchange_transactions_csv_path = (
        month_dir / "foreign_exchange_transactions.csv"
    )
    immediate_transactions_csv_path = month_dir / "immediate_transactions.csv"

    foreign_count = _parse_captured_table_html(
        capture_report["foreign"],
        html_path=foreign_html_path,
        csv_path=foreign_exchange_transactions_csv_path,
        parse_html=get_foreign_exchange_transactions,
    )
    immediate_count = _parse_captured_table_html(
        capture_report["immediate"],
        html_path=immediate_html_path,
        csv_path=immediate_transactions_csv_path,
        parse_html=get_immediate_transactions,
    )
    capture_report["foreign"]["parsed_count"] = foreign_count
    capture_report["immediate"]["parsed_count"] = immediate_count

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

    appended_rows = parse_amount(
        str(combined_transactions_csv_path),
        str(excel_path),
    )

    capture_report["appended_rows"] = appended_rows
    update_state(
        {
            "last_filled_year": year,
            "last_filled_month": month,
            "last_max_capture": capture_report,
        }
    )

    return str(excel_path)
