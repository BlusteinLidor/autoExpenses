"""
Merge Max and Leumi exports into a single workbook for getExpenses.

- Max: xlsx with columns B=expense name, C=category, F=cost from row 5.
- Leumi: HTML exported as .xls (transactions + credit cards); we parse and normalize
  to the same (name, cost) shape.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl import load_workbook

from handleExcel import (
    resolve_transaction_is_income,
    card_settlement_kind,
    normalize_expense_source,
    normalize_expense_date,
    EXPENSE_SOURCE_MAX,
    EXPENSE_SOURCE_LEUMI_CHECKING,
    EXPENSE_SOURCE_LEUMI_CARD,
    EXPENSE_SOURCE_COL,
    EXPENSE_DATE_COL,
)

# name, cost, category, is_income, source, date_iso
ExpenseRow = Tuple[str, float, Optional[str], bool, str, Optional[str]]
LEUMI_CARDS_SETTLED_MARKERS = ("במועד החיוב",)
LEUMI_CARDS_PENDING_MARKERS = ("שטרם נקלטו",)


def _parse_number_hebrew(s: str) -> float:
    """Parse a number that may use Hebrew locale / bidi marks; preserves sign."""
    if s is None or (isinstance(s, float) and (s != s or s == 0)):
        return 0.0
    text = (
        str(s)
        .strip()
        .replace("\u200f", "")
        .replace("\u200e", "")
        .replace("₪", "")
        .replace(",", "")
        .replace(" ", "")
    )
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def _extract_text(cell) -> str:
    """Get single-line text from a table cell (BeautifulSoup td)."""
    if cell is None:
        return ""
    text = cell.get_text(separator=" ", strip=True)
    return " ".join(text.split()) if text else ""


def read_leumi_transactions_html(path: Path) -> List[Tuple[str, float, bool, Optional[str]]]:
    """
    Parse Leumi checking-account export (HTML saved as .xls).
    Columns: תאריך, תאריך ערך, תיאור, אסמכתא, בחובה (debit), בזכות (credit), ...
    Returns list of (description, amount, is_income, date_iso).
    """
    rows: List[Tuple[str, float, bool, Optional[str]]] = []
    raw = path.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(raw, "html.parser")
    table = soup.find("table", class_="xlTable")
    if not table:
        return rows
    trs = table.find_all("tr")
    if len(trs) < 3:
        return rows
    # Prefer header-based indices when available.
    header_cells = [_extract_text(td) for td in trs[1].find_all("td")]
    date_idx, desc_idx, debit_idx, credit_idx = 0, 2, 4, 5
    for i, header in enumerate(header_cells):
        if header == "תאריך" or (header.startswith("תאריך") and "ערך" not in header and "עסקה" not in header):
            date_idx = i
        elif "תיאור" in header:
            desc_idx = i
        elif "בחובה" in header:
            debit_idx = i
        elif "בזכות" in header:
            credit_idx = i

    for tr in trs[2:]:
        tds = tr.find_all("td")
        needed = max(date_idx, desc_idx, debit_idx, credit_idx) + 1
        if len(tds) < needed:
            continue
        desc = _extract_text(tds[desc_idx])
        debit = _parse_number_hebrew(_extract_text(tds[debit_idx]))
        credit = _parse_number_hebrew(_extract_text(tds[credit_idx]))
        if not desc:
            continue
        signed = debit - credit
        amount = abs(signed)
        date_label = normalize_expense_date(_extract_text(tds[date_idx]))
        if amount != 0:
            rows.append(
                (
                    desc,
                    amount,
                    resolve_transaction_is_income(desc, signed < 0),
                    date_label,
                )
            )
    return rows


def inspect_leumi_cards_export(path: Path) -> Dict[str, Any]:
    """Inspect Leumi card HTML export title/schema without fully parsing amounts."""
    info: Dict[str, Any] = {
        "title": "",
        "headers": [],
        "is_settled": False,
        "is_pending": False,
        "html_data_rows": 0,
        "path": str(path),
    }
    if not path.exists():
        return info
    raw = path.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(raw, "html.parser")
    table = soup.find("table", class_="xlTable")
    if not table:
        return info
    trs = table.find_all("tr")
    if not trs:
        return info
    title = _extract_text(trs[0])
    info["title"] = title
    info["is_settled"] = any(marker in title for marker in LEUMI_CARDS_SETTLED_MARKERS)
    info["is_pending"] = any(marker in title for marker in LEUMI_CARDS_PENDING_MARKERS)
    if len(trs) >= 2:
        info["headers"] = [_extract_text(td) for td in trs[1].find_all("td")]
    info["html_data_rows"] = max(0, len(trs) - 2)
    return info


def _leumi_cards_column_map(headers: List[str]) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    for i, header in enumerate(headers):
        if "שם בית העסק" in header or header == "שם בית העסק":
            mapping["name"] = i
        elif "סכום חיוב" in header:
            mapping["charge"] = i
        elif "סכום העסקה" in header:
            mapping["transaction_amount"] = i
        elif "תאריך העסקה" in header or header == "תאריך":
            mapping["date"] = i
        elif header in ("שעה",) or header.startswith("שעה"):
            mapping["hour"] = i
    return mapping


def read_leumi_credit_cards_html(
    path: Path,
) -> List[Tuple[str, float, bool, Optional[str]]]:
    """
    Parse Leumi credit-cards export (HTML saved as .xls).

    Supports:
    - Settled billing export: תאריך העסקה, שם בית העסק, ..., סכום חיוב
    - Pending export: תאריך העסקה, שעה, שם בית העסק, ..., סכום העסקה
      (still parsed, but reconciliation should reject pending exports)

    Returns list of (business_name, amount, is_income, date_iso).
    Negative amounts (refunds/credits) are stored as income.
    """
    rows: List[Tuple[str, float, bool, Optional[str]]] = []
    raw = path.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(raw, "html.parser")
    table = soup.find("table", class_="xlTable")
    if not table:
        return rows
    trs = table.find_all("tr")
    if len(trs) < 3:
        return rows

    headers = [_extract_text(td) for td in trs[1].find_all("td")]
    colmap = _leumi_cards_column_map(headers)
    name_idx = colmap.get("name")
    amount_idx = colmap.get("charge", colmap.get("transaction_amount"))
    date_idx = colmap.get("date", 0)

    # Legacy fallback only when headers look like the settled schema without שעה.
    if name_idx is None or amount_idx is None:
        if len(headers) >= 6 and "שעה" not in "".join(headers):
            name_idx = 1
            amount_idx = 5
            date_idx = 0
        elif len(headers) >= 6 and "שעה" in "".join(headers):
            name_idx = 2
            amount_idx = 5
            date_idx = 0
        else:
            return rows

    for tr in trs[2:]:
        tds = tr.find_all("td")
        needed = max(name_idx, amount_idx, date_idx) + 1
        if len(tds) < needed:
            continue
        name = _extract_text(tds[name_idx])
        # Guard against the old bug: hour mistaken for merchant name.
        if name_idx == colmap.get("hour") or (
            len(name) == 5 and name[2] == ":" and name.replace(":", "").isdigit()
        ):
            # Try adjacent merchant column when hour leaked in.
            if name_idx + 1 < len(tds):
                alt = _extract_text(tds[name_idx + 1])
                if alt and not (len(alt) == 5 and alt[2] == ":"):
                    name = alt
        signed = _parse_number_hebrew(_extract_text(tds[amount_idx]))
        if not name or signed == 0:
            continue
        amount = abs(signed)
        is_income = signed < 0
        date_label = normalize_expense_date(_extract_text(tds[date_idx]))
        rows.append((name, amount, is_income, date_label))
    return rows


def _read_max_sheet_rows(ws) -> List[ExpenseRow]:
    rows: List[ExpenseRow] = []
    row = 5
    while True:
        name_cell = ws[f"B{row}"].value
        if name_cell is None or (isinstance(name_cell, str) and not name_cell.strip()):
            break
        cost_cell = ws[f"F{row}"].value
        try:
            cost = abs(float(cost_cell)) if cost_cell is not None else 0.0
        except (TypeError, ValueError) as e:
            print(
                f"[merge_expenses] Row {row} non-numeric cost in F: {cost_cell!r} -> 0.0 ({e})"
            )
            cost = 0.0
        income_flag = ws[f"G{row}"].value
        is_income = resolve_transaction_is_income(
            str(name_cell).strip(),
            income_flag in (1, True, "1", "income", "yes"),
        )
        category_cell = ws[f"C{row}"].value
        category = str(category_cell).strip() if category_cell is not None else None
        source_cell = ws[f"{EXPENSE_SOURCE_COL}{row}"].value
        # Max native exports put amounts in H; always normalize to a known label.
        source = normalize_expense_source(source_cell, default=EXPENSE_SOURCE_MAX)
        date_label = normalize_expense_date(ws[f"{EXPENSE_DATE_COL}{row}"].value)
        rows.append(
            (
                str(name_cell).strip(),
                cost,
                category or None,
                is_income,
                source,
                date_label,
            )
        )
        row += 1
    return rows


def read_max_expense_rows(path: Path) -> List[ExpenseRow]:
    """
    Read expense name (B), main category (C), cost (F) and date (A) from row 5.

    Max exports often have separate sheets (billing vs immediate). Across sheets we keep
    the maximum per-sheet multiplicity for each (name, amount, is_income) so the same
    charge is not double-counted when it appears on both sheets. Per-row dates are
    preserved from the winning sheet.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    best_rows: Dict[Tuple[str, float, bool], List[ExpenseRow]] = {}

    for ws in wb.worksheets:
        sheet_rows: Dict[Tuple[str, float, bool], List[ExpenseRow]] = {}
        for row in _read_max_sheet_rows(ws):
            name, cost, _category, is_income, _source, _date = row
            key = (name, round(float(cost), 2), bool(is_income))
            sheet_rows.setdefault(key, []).append(row)
        for key, rows in sheet_rows.items():
            prev = best_rows.get(key)
            if prev is None or len(rows) > len(prev):
                best_rows[key] = rows
            elif len(rows) == len(prev):
                # Prefer rows that still carry category / date labels.
                prev_score = sum(1 for r in prev if r[2] or r[5])
                new_score = sum(1 for r in rows if r[2] or r[5])
                if new_score > prev_score:
                    best_rows[key] = rows

    wb.close()

    rows: List[ExpenseRow] = []
    for key_rows in best_rows.values():
        rows.extend(key_rows)
    return rows


def write_combined_workbook(rows: List[ExpenseRow], output_path: Path) -> None:
    """
    Write a workbook compatible with getExpenses:
    from row 5, A=date, B=name, C=category (main), F=cost, H=source.
    """
    wb = Workbook()
    ws = wb.active
    start_row = 5
    for i, (name, cost, category, is_income, source, date_label) in enumerate(rows):
        r = start_row + i
        if date_label:
            ws[f"{EXPENSE_DATE_COL}{r}"] = date_label
        ws[f"B{r}"] = name
        ws[f"C{r}"] = category
        ws[f"F{r}"] = cost
        ws[f"{EXPENSE_SOURCE_COL}{r}"] = normalize_expense_source(
            source, default=EXPENSE_SOURCE_MAX
        )
        if is_income:
            ws[f"G{r}"] = 1
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def drop_checking_card_settlements(
    rows: List[ExpenseRow],
    *,
    drop_max_settlements: bool,
    drop_leumi_card_settlements: bool,
) -> Tuple[List[ExpenseRow], Dict[str, Any]]:
    """
    Prefer merchant-level detail over checking-account card-payment lumps.

    - When Max detail rows exist, drop Max settlement lines from checking.
    - When Leumi card merchants exist (healthy export), drop Mastercard settlement.

    Non-settlement checking activity is always kept. Settlements are kept when the
    matching detail source is missing so expenses are not silently lost.
    """
    kept: List[ExpenseRow] = []
    dropped: List[Dict[str, Any]] = []

    for row in rows:
        name, amount, category, is_income, source, _date = row
        if source != EXPENSE_SOURCE_LEUMI_CHECKING:
            kept.append(row)
            continue
        kind = card_settlement_kind(name)
        should_drop = (kind == "max" and drop_max_settlements) or (
            kind == "leumi_mastercard" and drop_leumi_card_settlements
        )
        if should_drop:
            dropped.append(
                {
                    "name": name,
                    "amount": round(float(amount), 2),
                    "kind": kind,
                    "is_income": bool(is_income),
                }
            )
            continue
        kept.append(row)

    dropped_max = [d for d in dropped if d.get("kind") == "max"]
    dropped_mc = [d for d in dropped if d.get("kind") == "leumi_mastercard"]
    stats: Dict[str, Any] = {
        "dropped_count": len(dropped),
        "dropped_amount_sum": round(sum(float(d["amount"]) for d in dropped), 2),
        "dropped_max_count": len(dropped_max),
        "dropped_max_amount_sum": round(
            sum(float(d["amount"]) for d in dropped_max), 2
        ),
        "dropped_leumi_mastercard_count": len(dropped_mc),
        "dropped_leumi_mastercard_amount_sum": round(
            sum(float(d["amount"]) for d in dropped_mc), 2
        ),
        "dropped_rows": dropped,
        "drop_max_settlements": bool(drop_max_settlements),
        "drop_leumi_card_settlements": bool(drop_leumi_card_settlements),
    }
    return kept, stats


def merge_max_and_leumi(
    max_excel_path: Path,
    leumi_transactions_path: Path,
    leumi_cards_path: Path,
    output_path: Path,
) -> Tuple[Path, Dict[str, Any]]:
    """
    Combine Max export and Leumi (transactions + cards) into one workbook at output_path.

    Prefers merchant detail over checking card-settlement lumps when detail is present.
    Returns (output_path, dedupe_stats).
    Order after dedupe: Max rows, non-settlement checking, Leumi card merchants.
    """
    max_rows: List[ExpenseRow] = []
    checking_rows: List[ExpenseRow] = []
    card_rows: List[ExpenseRow] = []

    if max_excel_path.exists():
        max_rows = [
            (name, cost, category, is_income, EXPENSE_SOURCE_MAX, date_label)
            for name, cost, category, is_income, _source, date_label in read_max_expense_rows(
                max_excel_path
            )
        ]

    if leumi_transactions_path.exists():
        for name, amount, is_income, date_label in read_leumi_transactions_html(
            leumi_transactions_path
        ):
            checking_rows.append(
                (
                    name,
                    amount,
                    None,
                    is_income,
                    EXPENSE_SOURCE_LEUMI_CHECKING,
                    date_label,
                )
            )

    cards_info: Dict[str, Any] = {"healthy": False, "parsed_rows": 0}
    if leumi_cards_path.exists():
        cards_info = inspect_leumi_cards_export(leumi_cards_path)
        parsed_cards = read_leumi_credit_cards_html(leumi_cards_path)
        cards_info["parsed_rows"] = len(parsed_cards)
        cards_info["healthy"] = bool(
            cards_info.get("is_settled")
            and not cards_info.get("is_pending")
            and len(parsed_cards) > 0
        )
        for name, amount, is_income, date_label in parsed_cards:
            card_rows.append(
                (name, amount, None, is_income, EXPENSE_SOURCE_LEUMI_CARD, date_label)
            )

    drop_max = len(max_rows) > 0
    drop_leumi_mc = bool(cards_info.get("healthy"))

    all_rows = [*max_rows, *checking_rows, *card_rows]
    deduped_rows, dedupe_stats = drop_checking_card_settlements(
        all_rows,
        drop_max_settlements=drop_max,
        drop_leumi_card_settlements=drop_leumi_mc,
    )

    max_expense_sum = round(
        sum(r[1] for r in max_rows if not r[3]), 2
    )
    card_expense_sum = round(
        sum(r[1] for r in card_rows if not r[3]), 2
    )
    warnings: List[str] = []
    if drop_max and dedupe_stats["dropped_max_count"]:
        gap = round(
            dedupe_stats["dropped_max_amount_sum"] - max_expense_sum, 2
        )
        if abs(gap) > 1.0:
            warnings.append(
                f"Max settlement total ({dedupe_stats['dropped_max_amount_sum']}) "
                f"differs from Max merchant total ({max_expense_sum}) by {gap}. "
                "Merchant detail kept; settlement lumps dropped."
            )
    if drop_leumi_mc and dedupe_stats["dropped_leumi_mastercard_count"]:
        gap = round(
            dedupe_stats["dropped_leumi_mastercard_amount_sum"] - card_expense_sum,
            2,
        )
        if abs(gap) > 1.0:
            warnings.append(
                f"Mastercard settlement total "
                f"({dedupe_stats['dropped_leumi_mastercard_amount_sum']}) "
                f"differs from Leumi card merchant total ({card_expense_sum}) "
                f"by {gap}. Merchant detail kept; settlement lump dropped."
            )
    if not drop_max:
        warnings.append(
            "Max merchant export empty — Max settlement lines in checking were kept."
        )
    if not drop_leumi_mc:
        warnings.append(
            "Leumi card merchant export not healthy — Mastercard settlement "
            "line(s) in checking were kept."
        )

    dedupe_stats["warnings"] = warnings
    dedupe_stats["kept_rows"] = len(deduped_rows)
    dedupe_stats["source_rows_before"] = len(all_rows)
    dedupe_stats["max_merchant_rows"] = len(max_rows)
    dedupe_stats["max_merchant_amount_sum"] = max_expense_sum
    dedupe_stats["leumi_card_merchant_rows"] = len(card_rows)
    dedupe_stats["leumi_card_merchant_amount_sum"] = card_expense_sum
    dedupe_stats["card_export_healthy"] = bool(cards_info.get("healthy"))

    if dedupe_stats["dropped_count"]:
        print(
            f"[merge_expenses] Dropped {dedupe_stats['dropped_count']} checking "
            f"card-settlement row(s) totaling {dedupe_stats['dropped_amount_sum']} "
            "(preferring merchant detail)."
        )
    for warning in warnings:
        print(f"[merge_expenses] {warning}")

    write_combined_workbook(deduped_rows, output_path)
    return output_path, dedupe_stats
