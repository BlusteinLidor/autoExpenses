from datetime import datetime
import os
from pathlib import Path
from typing import Optional, Tuple

from playwright.sync_api import sync_playwright, Page, BrowserContext

from config import get_paths, load_env, validate_env, update_state
from utils import checkDate


def _debug_enabled() -> bool:
    return os.environ.get("LEUMI_DEBUG", "").strip() in (
        "1",
        "true",
        "True",
        "yes",
        "YES",
    )


def month_number_to_hebrew(month: str | int) -> tuple[str, str]:
    month_int = int(str(month).strip())
    months = {
        1: "ינואר",
        2: "פברואר",
        3: "מרץ",
        4: "אפריל",
        5: "מאי",
        6: "יוני",
        7: "יולי",
        8: "אוגוסט",
        9: "ספטמבר",
        10: "אוקטובר",
        11: "נובמבר",
        12: "דצמבר",
    }
    try:
        current = months[month_int]
        next_month_int = 1 if month_int == 12 else month_int + 1
        next_month = months[next_month_int]
        return current, next_month
    except KeyError as e:
        raise ValueError(f"month must be between 1-12, got {month!r}") from e


def _login_leumi(page: Page) -> None:
    load_env()
    validate_env(require_max=False, require_leumi=True)
    username = os.environ.get("LEUMI_USERNAME")
    password = os.environ.get("LEUMI_PASSWORD")

    page.goto(
        "https://hb2.bankleumi.co.il/staticcontent/gate-keeper/he/?trackingCode=c0278d68-d0b3-4dfe-6557-2227c641e379&sysNum=23&langNum=1#/hpsummary",
        wait_until="networkidle",
    )

    # Try to remove or hide the cookies overlay that intercepts clicks
    try:
        page.wait_for_timeout(500)
        page.evaluate(
            "const el = document.querySelector('.app-cookies-overlay'); if (el) el.remove();"
        )
    except Exception:
        # If this fails, we continue and rely on force clicks below.
        pass

    # username
    username_box = page.get_by_role("textbox", name="שם משתמש")
    username_box.click(force=True)
    username_box.fill(username or "")

    # password
    pw_box = page.get_by_role("textbox", name="סיסמה")
    pw_box.click(force=True)
    pw_box.fill(password or "")

    # submit
    page.get_by_role("button", name="כניסה לחשבון").click()


def _export_transactions_for_month(
    page: Page, year: str, month: str
) -> Optional[Tuple[Path, Path]]:
    """
    Placeholder implementation: this function should navigate from the summary view
    to the transactions list for the given month and trigger an export/download.

    Because Leumi's HTML/flow can change and is not fully specified here,
    the CSS/XPath selectors will likely need to be customized.
    """
    paths = get_paths()
    # This path is where we want the final CSV/Excel to live.
    target = paths.leumi_exports_dir / f"leumi_transactions_{year}_{month}.xls"

    # current_year = datetime.now().year
    # last_month = datetime.now().month - 1 if datetime.now().month > 1 else 12
    # last_month_in_hebrew, last_next_month_in_hebrew = month_number_to_hebrew(last_month)
    # month_in_hebrew, next_month_in_hebrew = month_number_to_hebrew(month)

    start_day = "10"
    end_day = "9"

    month_int = int(month)
    if len(year) == 4:
        year_int = int(year) % 2000
    else:
        year_int = int(year)
    next_month_int = month_int + 1 if month_int < 12 else 1
    if next_month_int == 1:
        next_year_int = year_int + 1
    else:
        next_year_int = year_int
    with page.expect_download() as download_info:
        # Go to the checking account ("עובר ושב")
        page.locator(
            "app-footer a[aria-label='עובר ושב'][href*='BusinessAccountTrx']"
        ).click()

        # Open advanced search
        page.get_by_text("חיפוש מתקדם", exact=True).click()

        # Set date range – these are your recorded clicks; customize as needed
        page.get_by_text("תקופה").click()
        # page.locator(".ts-btn.btn-default").first.click()
        page.get_by_placeholder("מתאריך").click()
        page.get_by_placeholder("מתאריך").press("ControlOrMeta+a")
        page.get_by_placeholder("מתאריך").fill(f"{start_day}.{month_int}.{year_int}")
        page.get_by_placeholder("מתאריך").click()
        page.wait_for_timeout(1000)
        page.get_by_placeholder("עד תאריך").click()
        page.get_by_placeholder("עד תאריך").press("ControlOrMeta+a")
        page.get_by_placeholder("עד תאריך").fill(
            f"{end_day}.{next_month_int}.{next_year_int}"
        )
        page.get_by_placeholder("עד תאריך").click()

        # Apply filter
        page.wait_for_timeout(1000)
        page.get_by_label("סנן").click()

        # Export to Excel and confirm
        page.wait_for_timeout(5000)
        page.get_by_title("יצוא לאקסל").click()
        page.get_by_text("המשך").click()

    download = download_info.value
    temp_path = Path(download.path())
    temp_path.replace(target)

    with page.expect_download() as download1_info:
        page.locator("app-nav-menu").get_by_text("דף הבית").click()
        page.locator("#center_hpsummary").get_by_label(
            "כרטיסי אשראי", exact=True
        ).click()
        page.locator("button").filter(has_text="אפריל").click()
        page.locator("li").filter(has_text="מרץ").click()
        page.get_by_title("יצוא לאקסל", exact=True).click()
        page.get_by_text("המשך").click()
    download1 = download1_info.value
    temp_path1 = Path(download1.path())
    cards_target = paths.leumi_exports_dir / f"leumi_credit_cards_{year}_{month}.xls"
    temp_path1.replace(cards_target)

    return target, cards_target


def _scrape_balance(page: Page) -> Optional[str]:
    """
    Scrape current balance from the Leumi summary page.
    Selectors are placeholders and must be adapted to the actual DOM.
    """
    try:
        # Example: find an element that contains the main account balance.
        # The real selector should be inspected via browser devtools.
        balance_el = page.locator(".item-divider-left").first
        text = (
            balance_el.inner_text()
            .strip()
            .replace("₪", "")
            .replace(",", "")
            .replace("\u200f", "")
            .strip()
        )
        return text
    except Exception:
        return None


def getLeumiData(year: str, month: str) -> Optional[Tuple[Path, Path]]:
    """
    Run Leumi flow:
    - login
    - (placeholder) navigate and export transactions for the given month
    - scrape current balance and store in state
    Returns the exported transactions file path (or None if not implemented yet).
    """
    year, month = checkDate(year, month)

    with sync_playwright() as p:
        # debug = _debug_enabled()
        debug = True
        paths = get_paths()
        debug_dir = paths.data_dir / "leumi_debug"
        debug_dir.mkdir(parents=True, exist_ok=True)
        trace_path = debug_dir / f"trace_leumi_{year}_{month}.zip"

        browser = p.chromium.launch(headless=(not debug))
        context: BrowserContext = browser.new_context(
            accept_downloads=True,
            record_video_dir=str(debug_dir) if debug else None,
            record_video_size={"width": 1280, "height": 720} if debug else None,
        )
        if debug:
            context.tracing.start(screenshots=True, snapshots=True, sources=True)
        page = context.new_page()

        _login_leumi(page)
        page.wait_for_timeout(3000)

        # Scrape balance from the summary view
        balance_text = _scrape_balance(page)
        if balance_text:
            update_state(
                {
                    "leumi_balance": balance_text,
                }
            )

        # Export transactions (will raise NotImplementedError until implemented)
        exported_path = _export_transactions_for_month(page, year, month)
        if exported_path is None:
            raise RuntimeError("Failed to export transactions")

        if debug:
            page.screenshot(
                path=str(debug_dir / f"final_{year}_{month}.png"), full_page=True
            )
            context.tracing.stop(path=str(trace_path))

        context.close()
        browser.close()

    return exported_path[0], exported_path[1]
