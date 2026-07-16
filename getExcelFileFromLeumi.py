from datetime import datetime
import os
import re
import tempfile
import time
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


def _wait_for_credit_card_view(page: Page, timeout_ms: int = 60000) -> None:
    page.wait_for_load_state("networkidle")
    try:
        page.locator("bll-loader ngx-spinner, ngx-spinner").first.wait_for(
            state="hidden", timeout=timeout_ms
        )
    except PlaywrightTimeoutError:
        pass

    hebrew_months = [
        "ינואר",
        "פברואר",
        "מרץ",
        "אפריל",
        "מאי",
        "יוני",
        "יולי",
        "אוגוסט",
        "ספטמבר",
        "אוקטובר",
        "נובמבר",
        "דצמבר",
    ]
    for month_name in hebrew_months:
        month_button = page.locator("button").filter(has_text=month_name).first
        try:
            month_button.wait_for(state="visible", timeout=3000)
            return
        except PlaywrightTimeoutError:
            continue

    export_button = page.get_by_title("יצוא לאקסל", exact=True).first
    export_button.wait_for(state="visible", timeout=timeout_ms)


def _navigate_to_credit_card_export_view(page: Page) -> None:
    _close_leumi_header_overlays(page)
    page.locator("app-footer a[aria-label='דף הבית']").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)

    card_entry = page.locator(
        "creditcard-directive a[data-lb-key='SHEMESHPREMIUM.creditCard.TEST.Link']"
    ).first
    title_entry = page.locator(
        "creditcard-directive a[aria-label='כרטיסי אשראי']"
    ).first

    if card_entry.count() > 0:
        card_entry.click(force=True, timeout=15000)
    elif title_entry.count() > 0:
        title_entry.click(force=True, timeout=15000)
    else:
        page.locator(
            "app-footer a[aria-label='כרטיסי אשראי'][href*='CardsWorld']"
        ).click()

    _wait_for_credit_card_view(page)
    _close_leumi_header_overlays(page)


def _hebrew_month_labels() -> list[str]:
    return [
        "ינואר",
        "פברואר",
        "מרץ",
        "אפריל",
        "מאי",
        "יוני",
        "יולי",
        "אוגוסט",
        "ספטמבר",
        "אוקטובר",
        "נובמבר",
        "דצמבר",
    ]


def _full_year(year: str | int) -> str:
    text = str(year).strip()
    if len(text) == 4:
        return text
    return f"20{int(text):02d}"


def _find_visible_month_picker_button(page: Page):
    """
    Return the visible period dropdown button that currently shows a Hebrew month.

    Prefer the control near 'הצגת פירוט עסקאות לתקופה' so we don't match unrelated buttons.
    """
    labels = _hebrew_month_labels()
    scoped = page.locator("button").filter(
        has_text=re.compile("|".join(re.escape(label) for label in labels))
    )
    # Prefer a button whose text looks like "<month> <year>".
    try:
        count = scoped.count()
    except Exception:
        count = 0
    for i in range(count):
        btn = scoped.nth(i)
        try:
            if not btn.is_visible():
                continue
            text = re.sub(r"\s+", " ", (btn.inner_text() or "").strip())
            for label in labels:
                if text.startswith(label):
                    return btn, label
        except Exception:
            continue

    for label in labels:
        candidate = page.locator("button").filter(has_text=label)
        try:
            if candidate.count() > 0 and candidate.first.is_visible():
                return candidate.first, label
        except Exception:
            continue
    return None, None


def _picker_button_label(page: Page) -> str | None:
    _, label = _find_visible_month_picker_button(page)
    return label


def _click_credit_card_month_option(page: Page, *, month_label: str, year: str) -> bool:
    """Open the period dropdown and choose '<month> <year>' (fallback: month only)."""
    year_full = _full_year(year)
    exact_label = f"{month_label} {year_full}"
    picker, current_label = _find_visible_month_picker_button(page)
    if picker is None:
        return False
    if current_label == month_label:
        # Already selected; still verify year-ish text when possible.
        try:
            text = re.sub(r"\s+", " ", (picker.inner_text() or "").strip())
            if year_full in text or text.startswith(month_label):
                return True
        except Exception:
            return True

    picker.click(timeout=15000)
    page.wait_for_timeout(400)

    option_candidates = [
        page.locator("li").filter(has_text=exact_label),
        page.get_by_role("option", name=exact_label),
        page.locator("li").filter(has_text=re.compile(rf"^{re.escape(month_label)}\s+{re.escape(year_full)}$")),
        page.locator("li").filter(has_text=month_label),
    ]
    clicked = False
    for option in option_candidates:
        try:
            target = option.first
            target.wait_for(state="visible", timeout=8000)
            target.scroll_into_view_if_needed()
            target.click(timeout=8000)
            clicked = True
            break
        except Exception:
            continue
    if not clicked:
        page.keyboard.press("Escape")
        return False

    # Wait until the picker reflects the chosen month.
    deadline = time.monotonic() + 12.0
    while time.monotonic() < deadline:
        _close_leumi_header_overlays(page)
        selected = _picker_button_label(page)
        if selected == month_label:
            return True
        page.wait_for_timeout(300)
    return False


def _select_credit_card_month(
    page: Page,
    month_in_hebrew: str,
    next_current_month_in_hebrew: str,
    *,
    year: str,
    billing_month_in_hebrew: str | None = None,
    month_number: str | int | None = None,
) -> str:
    """
    Select the Leumi cards billing period.

    Tries, in order:
    1) run month (e.g. יוני 2026) — historical behavior that worked for May
    2) billing/charge month (usually run month + 1, e.g. יולי) — aligns with Max

    Returns the Hebrew month label that ended up selected.
    """
    del next_current_month_in_hebrew  # kept for call-site compatibility
    _close_leumi_header_overlays(page)

    year_full = _full_year(year)
    billing_year = year_full
    if month_number is not None and int(month_number) == 12:
        billing_year = str(int(year_full) + 1)

    candidates: list[tuple[str, str]] = []
    for label, label_year in (
        (month_in_hebrew, year_full),
        (billing_month_in_hebrew, billing_year),
    ):
        if label and (label, label_year) not in candidates:
            candidates.append((label, label_year))
    if not candidates:
        raise RuntimeError("No credit-card month candidates to select.")

    last_error = None
    for label, label_year in candidates:
        try:
            ok = _click_credit_card_month_option(
                page, month_label=label, year=label_year
            )
            if not ok:
                last_error = f"could not click/select {label!r} ({label_year})"
                continue
            _close_leumi_header_overlays(page)
            selected = _picker_button_label(page)
            if selected != label:
                last_error = f"after selecting {label!r}, UI shows {selected!r}"
                continue

            # Billing-period statement should be present for a historical/current charge month.
            settled = page.get_by_text("עסקאות בש\"ח במועד החיוב", exact=False)
            pending_only = page.get_by_text("עסקאות אחרונות שטרם נקלטו", exact=False)
            try:
                if settled.count() == 0 and pending_only.count() > 0:
                    last_error = (
                        f"after selecting {label}, only pending card transactions "
                        "are visible (שטרם נקלטו)"
                    )
                    continue
            except Exception:
                pass
            print(f"[leumi] Selected credit-card period month: {label} {label_year}")
            return label
        except Exception as e:
            last_error = str(e)
            continue

    raise RuntimeError(
        "Failed to select Leumi credit-card month from candidates "
        f"{candidates!r}. Last error: {last_error}"
    )


def _dismiss_intercepting_overlays(page: Page) -> None:
    # Remove common overlays/popups that can intercept export clicks.
    try:
        page.evaluate(
            """
            () => {
                const selectors = [
                    '.app-cookies-overlay',
                    '.modal-backdrop',
                    '.backdrop',
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


def _close_leumi_header_overlays(page: Page, timeout_ms: int = 4000) -> None:
    """
    Leumi sometimes leaves the app-header sub-nav/backdrop open, which intercepts
    pointer events and breaks clicks (e.g. on "יצוא לאקסל").
    We try non-destructive close actions first, and only then fall back to DOM cleanup.
    """
    backdrop = page.locator("app-header .backdrop").first
    sub_nav_open = page.locator("app-header .sub-nav.open").first

    try:
        has_open_menu = sub_nav_open.count() > 0 and sub_nav_open.is_visible()
        has_backdrop = backdrop.count() > 0 and backdrop.is_visible()
        if not has_open_menu and not has_backdrop:
            return
    except Exception:
        pass

    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(150)
    except Exception:
        pass

    try:
        if backdrop.count() > 0 and backdrop.is_visible():
            backdrop.click(force=True, timeout=1500)
            page.wait_for_timeout(150)
    except Exception:
        pass

    try:
        page.evaluate(
            """
            () => {
                document.querySelectorAll('app-header .sub-nav.open').forEach((el) => {
                    el.classList.remove('open');
                });
                document.querySelectorAll('app-header .backdrop').forEach((el) => el.remove());
            }
            """
        )
    except Exception:
        pass

    try:
        if backdrop.count() > 0:
            backdrop.wait_for(state="hidden", timeout=timeout_ms)
    except Exception:
        pass
    try:
        if sub_nav_open.count() > 0:
            sub_nav_open.wait_for(state="hidden", timeout=timeout_ms)
    except Exception:
        pass

    _dismiss_intercepting_overlays(page)


def _ensure_clickable(page: Page, button_label: str, timeout_ms: int = 15000) -> None:
    button = page.get_by_title(button_label, exact=True).first
    start = time.monotonic()
    last_err: Exception | None = None
    succeeded = False

    while (time.monotonic() - start) * 1000 < timeout_ms:
        remaining = timeout_ms - int((time.monotonic() - start) * 1000)
        attempt_timeout = max(1000, min(5000, remaining))

        _close_leumi_header_overlays(page)
        _dismiss_intercepting_overlays(page)
        try:
            page.wait_for_load_state("networkidle", timeout=min(3000, attempt_timeout))
        except Exception:
            pass

        button.wait_for(state="visible", timeout=attempt_timeout)
        button.scroll_into_view_if_needed()

        try:
            button.click(trial=True, timeout=attempt_timeout)
            succeeded = True
            break
        except PlaywrightTimeoutError as e:
            last_err = e
            page.wait_for_timeout(250)
        except Exception as e:
            last_err = e
            page.wait_for_timeout(250)

    if not succeeded:
        if last_err is not None:
            raise last_err
        raise PlaywrightTimeoutError(
            f'Export button "{button_label}" was not clickable within {timeout_ms}ms'
        )

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
    _close_leumi_header_overlays(page)
    export_button.click(timeout=15000, force=True)


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

    # submit and wait until the login form is gone / dashboard shell appears
    page.get_by_role("button", name="כניסה לחשבון").click()
    _wait_for_leumi_dashboard(page)


def _dismiss_leumi_blocking_banners(page: Page) -> None:
    """Close common post-login banners (popup-blocked notice, walkme, cookies)."""
    _dismiss_intercepting_overlays(page)
    _close_leumi_header_overlays(page)
    try:
        page.evaluate(
            """
            () => {
                const texts = ['חלון הצגה', 'Pop-up', 'נחסם'];
                for (const el of Array.from(document.querySelectorAll('button, a, span, div'))) {
                    const t = (el.innerText || el.textContent || '').trim();
                    if (!t || t.length > 80) continue;
                    if (t === 'x' || t === 'X' || t === 'סגור' || t === 'אישור' || t === 'המשך') {
                        const parentText = (el.closest('div') || el).innerText || '';
                        if (texts.some((s) => parentText.includes(s))) {
                            try { el.click(); } catch (e) {}
                        }
                    }
                }
            }
            """
        )
    except Exception:
        pass


def _wait_for_leumi_dashboard(page: Page, timeout_ms: int = 90000) -> None:
    """
    After login, Leumi's Angular shell can take a while (and may show a blank page
    briefly). Wait for a real dashboard control before continuing.
    """
    _dismiss_leumi_blocking_banners(page)
    deadline = time.monotonic() + (timeout_ms / 1000.0)
    last_error: Exception | None = None
    strong_selectors = [
        "app-footer a[aria-label='עובר ושב']",
        "app-footer a[href*='BusinessAccountTrx']",
        "a[aria-label='עובר ושב']",
        "a[href*='BusinessAccountTrx']",
        ".item-divider-left",
    ]
    while time.monotonic() < deadline:
        _dismiss_leumi_blocking_banners(page)
        # Login form still present => not logged in yet.
        try:
            login_btn = page.get_by_role("button", name="כניסה לחשבון")
            if login_btn.count() > 0 and login_btn.first.is_visible():
                page.wait_for_timeout(1000)
                continue
        except Exception as e:
            last_error = e

        for sel in strong_selectors:
            loc = page.locator(sel).first
            try:
                if loc.count() > 0 and loc.is_visible():
                    return
            except Exception as e:
                last_error = e

        # Accept visible footer shell as a weaker ready signal near timeout.
        try:
            footer = page.locator("app-footer").first
            if footer.count() > 0 and footer.is_visible():
                remaining_ms = (deadline - time.monotonic()) * 1000
                if remaining_ms < 15000:
                    return
        except Exception as e:
            last_error = e

        page.wait_for_timeout(1000)

    detail = f" Last error: {last_error}" if last_error else ""
    raise RuntimeError(
        "Leumi dashboard did not become ready after login "
        f"(waited {timeout_ms}ms). Page may be blank or a popup was blocked.{detail}"
    )


def _navigate_to_checking_account(page: Page) -> None:
    """Open checking-account transactions with resilient selectors + SPA fallback."""
    _dismiss_leumi_blocking_banners(page)
    candidates = [
        page.locator("app-footer a[aria-label='עובר ושב'][href*='BusinessAccountTrx']"),
        page.locator("app-footer a[aria-label='עובר ושב']"),
        page.locator("app-footer a[href*='BusinessAccountTrx']"),
        page.locator("a[aria-label='עובר ושב']"),
        page.locator("a[href*='BusinessAccountTrx']"),
        page.get_by_role("link", name="עובר ושב"),
        page.get_by_text("עובר ושב", exact=True),
    ]
    for loc in candidates:
        try:
            target = loc.first
            if target.count() == 0:
                continue
            if not target.is_visible():
                continue
            target.scroll_into_view_if_needed()
            target.click(timeout=15000)
            page.wait_for_timeout(1500)
            # Success heuristic: advanced search / period controls appear on trx page.
            if (
                page.get_by_text("חיפוש מתקדם", exact=True).count() > 0
                or page.get_by_placeholder("מתאריך").count() > 0
                or "BusinessAccountTrx" in (page.url or "")
            ):
                return
            # Click happened; give the SPA a bit more time then continue.
            page.wait_for_timeout(2000)
            if page.get_by_text("חיפוש מתקדם", exact=True).count() > 0:
                return
        except Exception:
            continue

    # SPA hash fallback used by Leumi NavState strings.
    try:
        page.evaluate(
            """
            () => {
                const hash = '#/ts/BusinessAccountTrx';
                if (location.hash !== hash) {
                    location.hash = hash;
                }
            }
            """
        )
        page.wait_for_timeout(2500)
        if page.get_by_text("חיפוש מתקדם", exact=True).count() > 0:
            return
    except Exception:
        pass

    raise RuntimeError(
        "Could not open Leumi checking account (עובר ושב / BusinessAccountTrx). "
        "Dashboard may still be loading or a popup was blocked."
    )


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

    month_in_hebrew, billing_month_in_hebrew = month_number_to_hebrew(month)
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
    _navigate_to_checking_account(page)

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
    _close_leumi_header_overlays(page)
    temp_path = _download_from_export_dialog(page, button_label="יצוא לאקסל")
    temp_path.replace(target)

    _navigate_to_credit_card_export_view(page)
    selected_card_month = _select_credit_card_month(
        page,
        month_in_hebrew,
        next_current_month_in_hebrew,
        year=year,
        billing_month_in_hebrew=billing_month_in_hebrew,
        month_number=month,
    )
    _close_leumi_header_overlays(page)
    # Prefer export controls near the settled billing table when both sections exist.
    try:
        settled_heading = page.get_by_text("עסקאות בש\"ח במועד החיוב", exact=False)
        if settled_heading.count() > 0:
            settled_heading.first.scroll_into_view_if_needed()
            page.wait_for_timeout(500)
    except Exception:
        pass
    temp_path1 = _download_from_export_dialog(page, button_label="יצוא לאקסל")
    cards_target = paths.leumi_exports_dir / f"leumi_credit_cards_{year}_{month}.xls"
    temp_path1.replace(cards_target)
    _validate_leumi_cards_export_file(
        cards_target, month_in_hebrew=selected_card_month
    )

    return target, cards_target


def _validate_leumi_cards_export_file(path: Path, *, month_in_hebrew: str) -> None:
    from merge_expenses import inspect_leumi_cards_export

    info = inspect_leumi_cards_export(path)
    title = str(info.get("title") or "")
    if info.get("is_pending"):
        raise RuntimeError(
            f"Leumi cards export for {month_in_hebrew} looks like pending transactions "
            f"({title!r}), not the billing-period statement. Refusing to continue."
        )
    if not info.get("is_settled"):
        raise RuntimeError(
            f"Leumi cards export for {month_in_hebrew} has unexpected title {title!r}. "
            "Expected a billing-period statement (במועד החיוב)."
        )
    if int(info.get("html_data_rows") or 0) == 0:
        print(
            f"[leumi] Warning: billing-period card export for {month_in_hebrew} "
            "has zero transaction rows."
        )


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

        browser = p.chromium.launch(
            headless=(not debug),
            args=["--disable-popup-blocking"],
        )
        context: BrowserContext = browser.new_context(
            accept_downloads=True,
            record_video_dir=str(debug_dir) if debug else None,
            record_video_size={"width": 1280, "height": 720} if debug else None,
            viewport={"width": 1440, "height": 900},
        )
        if debug:
            context.tracing.start(screenshots=True, snapshots=True, sources=True)
        page = context.new_page()

        try:
            _login_leumi(page)
            _dismiss_leumi_blocking_banners(page)

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
