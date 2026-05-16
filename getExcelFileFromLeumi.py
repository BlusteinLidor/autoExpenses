from datetime import datetime
import os
import tempfile
from pathlib import Path
from typing import Optional, Tuple
import traceback

from playwright.sync_api import (
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
    Page,
    BrowserContext,
    Locator,
)

from config import get_paths, load_env, validate_env, update_state
from utils import checkDate


def _dismiss_intercepting_overlays(page: Page) -> None:
    # Remove common overlays/popups that can intercept export clicks.
    try:
        page.evaluate(
            """
            () => {
                const selectors = [
                    '.app-cookies-overlay',
                    '.modal-backdrop',
                    '.block-ui-wrapper',
                    '.loading-overlay',
                ];
                for (const selector of selectors) {
                    document.querySelectorAll(selector).forEach((el) => el.remove());
                }
            }
            """
        )
    except Exception:
        pass


def _ensure_clickable(page: Page, button_label: str, timeout_ms: int = 15000) -> None:
    _dismiss_intercepting_overlays(page)
    page.wait_for_load_state("networkidle")
    button = page.get_by_title(button_label, exact=True).first
    button.wait_for(state="visible", timeout=timeout_ms)
    button.scroll_into_view_if_needed()
    button.click(trial=True, timeout=timeout_ms)
    coverage_info = button.evaluate(
        """
        (el) => {
            const rect = el.getBoundingClientRect();
            const cx = rect.left + (rect.width / 2);
            const cy = rect.top + (rect.height / 2);
            const top = document.elementFromPoint(cx, cy);
            const uncovered = !top || top === el || el.contains(top);
            return {
                uncovered,
                blocker: uncovered ? null : ((top.outerHTML || '').slice(0, 220)),
            };
        }
        """
    )
    if not coverage_info.get("uncovered", False):
        blocker = coverage_info.get("blocker")
        raise RuntimeError(f'Export button "{button_label}" is blocked by: {blocker}')


def _export_modal(page: Page) -> Locator:
    return page.locator("ngb-modal-window").filter(has=page.locator("app-exporttol-modal"))


def _export_modal_is_open(page: Page) -> bool:
    try:
        return page.locator("ngb-modal-window app-exporttol-modal").is_visible()
    except Exception:
        return False


def _export_continue_button(page: Page, continue_text: str = "המשך") -> Locator:
    """Continue button in the Excel export modal (primary footer button)."""
    footer = _export_modal(page).locator(".modal-footer")
    return footer.locator("button.btn-primary").or_(
        footer.get_by_text(continue_text, exact=True)
    ).first


def _save_export_from_iframe(page: Page, timeout_ms: int) -> str:
    page.wait_for_function(
        """
        () => {
            const iframe = document.getElementById('frmExportData');
            if (!iframe || !iframe.contentDocument) return false;
            return !!iframe.contentDocument.querySelector('table.xlTable');
        }
        """,
        timeout=timeout_ms,
    )
    return page.frame_locator("#frmExportData").locator("html").inner_html()


def _open_export_modal(page: Page, export_button: Locator) -> None:
    if _export_modal_is_open(page):
        return
    export_button.click(timeout=15000)


def _confirm_export_modal(page: Page, continue_text: str = "המשך") -> None:
    _export_modal(page).wait_for(state="visible", timeout=15000)
    continue_button = _export_continue_button(page, continue_text=continue_text)
    continue_button.wait_for(state="visible", timeout=15000)
    continue_button.click(timeout=15000)
    page.wait_for_timeout(1500)


def _click_export_and_continue(page: Page, export_button: Locator) -> None:
    _open_export_modal(page, export_button)
    _confirm_export_modal(page)


def _download_from_export_dialog(
    page: Page,
    button_label: str,
    continue_text: str = "המשך",
    timeout_ms: int = 90000,
) -> Path:
    _ensure_clickable(page, button_label=button_label, timeout_ms=15000)
    export_button = page.get_by_title(button_label, exact=True).first

    download_timeout_ms = min(timeout_ms, 45000)
    try:
        with page.expect_download(timeout=download_timeout_ms) as download_info:
            _click_export_and_continue(page, export_button)
        return Path(download_info.value.path())
    except PlaywrightTimeoutError:
        # Export may already be confirmed; Leumi often loads HTML into #frmExportData only.
        iframe_timeout_ms = min(timeout_ms, 30000)
        try:
            return _write_iframe_export(page, timeout_ms=iframe_timeout_ms)
        except PlaywrightTimeoutError:
            if _export_modal_is_open(page):
                _confirm_export_modal(page, continue_text=continue_text)
            else:
                _click_export_and_continue(page, export_button)
            return _write_iframe_export(page, timeout_ms=timeout_ms)


def _write_iframe_export(page: Page, timeout_ms: int) -> Path:
    html = _save_export_from_iframe(page, timeout_ms=timeout_ms)
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".xls", delete=False, encoding="utf-8"
    ) as tmp:
        if html.lstrip().lower().startswith("<!doctype"):
            tmp.write(html)
        else:
            tmp.write(f"<!DOCTYPE html><html>{html}</html>")
        return Path(tmp.name)


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

    month_in_hebrew, _ = month_number_to_hebrew(month)
    _, next_current_month_in_hebrew = month_number_to_hebrew(datetime.now().month)

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
    # Go to the checking account ("עובר ושב")
    page.locator("app-footer a[aria-label='עובר ושב'][href*='BusinessAccountTrx']").click()

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
    page.get_by_placeholder("עד תאריך").fill(f"{end_day}.{next_month_int}.{next_year_int}")
    page.get_by_placeholder("עד תאריך").click()

    # Apply filter
    page.wait_for_timeout(1000)
    page.get_by_label("סנן").click()

    # Export to Excel and confirm
    page.wait_for_timeout(5000)
    temp_path = _download_from_export_dialog(page, button_label="יצוא לאקסל")
    temp_path.replace(target)

    page.locator("app-nav-menu").get_by_text("דף הבית").click()
    page.locator("#center_hpsummary").get_by_label("כרטיסי אשראי", exact=True).click()
    page.locator("button").filter(has_text=next_current_month_in_hebrew).click()
    page.locator("li").filter(has_text=month_in_hebrew).click()
    temp_path1 = _download_from_export_dialog(page, button_label="יצוא לאקסל")
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
        debug = _debug_enabled()
        paths = get_paths()
        debug_dir = paths.data_dir / "leumi_debug"
        debug_dir.mkdir(parents=True, exist_ok=True)
        trace_path = debug_dir / f"trace_leumi_{year}_{month}.zip"
        error_screenshot_path = debug_dir / f"error_{year}_{month}.png"
        error_html_path = debug_dir / f"error_{year}_{month}.html"

        browser = p.chromium.launch(headless=(not debug))
        context: BrowserContext = browser.new_context(
            accept_downloads=True,
            record_video_dir=str(debug_dir) if debug else None,
            record_video_size={"width": 1280, "height": 720} if debug else None,
        )
        if debug:
            context.tracing.start(screenshots=True, snapshots=True, sources=True)
        page = context.new_page()

        try:
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
        except Exception as e:
            print("[leumi][error] Leumi flow failed.")
            print(f"[leumi][error] {type(e).__name__}: {e}")
            print(traceback.format_exc())

            # Best-effort debug artifacts for troubleshooting.
            try:
                page.screenshot(path=str(error_screenshot_path), full_page=True)
                print(f"[leumi][debug] Saved error screenshot: {error_screenshot_path}")
            except Exception as capture_error:
                print(f"[leumi][debug] Failed to save error screenshot: {capture_error}")

            try:
                error_html_path.write_text(page.content(), encoding="utf-8")
                print(f"[leumi][debug] Saved error HTML: {error_html_path}")
            except Exception as capture_error:
                print(f"[leumi][debug] Failed to save error HTML: {capture_error}")

            raise
        finally:
            if debug:
                try:
                    context.tracing.stop(path=str(trace_path))
                    print(f"[leumi][debug] Saved Playwright trace: {trace_path}")
                except Exception as trace_error:
                    print(f"[leumi][debug] Failed to save Playwright trace: {trace_error}")

            context.close()
            browser.close()

    return exported_path[0], exported_path[1]
