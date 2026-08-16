"""
Post-download reconciliation: compare raw exports to parsed rows and flag gaps.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from bs4 import BeautifulSoup

from merge_expenses import (
    inspect_leumi_cards_export,
    read_leumi_credit_cards_html,
    read_leumi_transactions_html,
    read_max_expense_rows,
)


def _html_table_data_row_count(path: Path) -> int:
    if not path.exists():
        return 0
    raw = path.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(raw, "html.parser")
    table = soup.find("table", class_="xlTable")
    if not table:
        return 0
    trs = table.find_all("tr")
    return max(0, len(trs) - 2)


def build_reconciliation_report(
    *,
    year: str,
    month: str,
    max_excel_path: Path,
    leumi_transactions_path: Optional[Path] = None,
    leumi_cards_path: Optional[Path] = None,
    combined_path: Optional[Path] = None,
    max_capture: Optional[Dict[str, Any]] = None,
    merge_dedupe_stats: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    warnings: List[str] = []
    blocking_errors: List[str] = []

    max_rows = read_max_expense_rows(max_excel_path) if max_excel_path.exists() else []
    max_sum = round(sum(r[1] for r in max_rows), 2)

    # (description, amount, is_income, date_iso) — same shape as merge_expenses readers
    leumi_checking: List[Tuple[str, float, bool, Optional[str]]] = []
    leumi_cards: List[Tuple[str, float, bool, Optional[str]]] = []
    checking_html_rows = 0
    cards_info: Dict[str, Any] = {
        "title": "",
        "is_settled": False,
        "is_pending": False,
        "html_data_rows": 0,
        "parsed_rows": 0,
        "healthy": False,
    }

    if leumi_transactions_path and leumi_transactions_path.exists():
        checking_html_rows = _html_table_data_row_count(leumi_transactions_path)
        leumi_checking = read_leumi_transactions_html(leumi_transactions_path)
        if checking_html_rows > 0 and len(leumi_checking) < checking_html_rows:
            warnings.append(
                f"Leumi checking: parsed {len(leumi_checking)} of {checking_html_rows} "
                "HTML data row(s)."
            )
        if checking_html_rows > 0 and len(leumi_checking) == 0:
            blocking_errors.append(
                "Leumi checking export has HTML rows but none were parsed."
            )

    if leumi_cards_path and leumi_cards_path.exists():
        cards_info = inspect_leumi_cards_export(leumi_cards_path)
        leumi_cards = read_leumi_credit_cards_html(leumi_cards_path)
        cards_info["parsed_rows"] = len(leumi_cards)
        if cards_info.get("is_pending"):
            blocking_errors.append(
                "Leumi credit-card export is pending transactions "
                "('עסקאות אחרונות שטרם נקלטו'), not the billing-period statement. "
                "Month selection or export target is wrong."
            )
        elif not cards_info.get("is_settled"):
            blocking_errors.append(
                "Leumi credit-card export title was unexpected: "
                f"{cards_info.get('title')!r}. Expected billing-period statement."
            )
        elif cards_info.get("html_data_rows", 0) == 0:
            warnings.append("Leumi credit-card billing export has zero transactions.")
        elif len(leumi_cards) == 0:
            blocking_errors.append(
                "Leumi credit-card export has HTML rows but none were parsed "
                "(column layout mismatch?)."
            )
        elif len(leumi_cards) < int(cards_info.get("html_data_rows") or 0):
            warnings.append(
                f"Leumi cards: parsed {len(leumi_cards)} of "
                f"{cards_info.get('html_data_rows')} HTML data row(s)."
            )

        cards_info["healthy"] = bool(
            cards_info.get("is_settled")
            and not cards_info.get("is_pending")
            and len(leumi_cards) > 0
        )
    elif leumi_cards_path is not None:
        warnings.append("Leumi credit-card export file is missing.")
        cards_info["healthy"] = False

    if max_capture and isinstance(max_capture, dict):
        immediate = max_capture.get("immediate") or {}
        foreign = max_capture.get("foreign") or {}
        immediate_parsed = int(immediate.get("parsed_count") or 0)
        immediate_visible = int(immediate.get("row_count") or 0)
        if immediate.get("error"):
            warnings.append(
                "Max immediate-charge scrape error: " + str(immediate.get("error"))
            )
        elif (
            not immediate.get("absent")
            and immediate_visible > 0
            and immediate_parsed < immediate_visible
        ):
            warnings.append(
                f"Max immediate: parsed {immediate_parsed} of {immediate_visible} "
                "visible row(s)."
            )
        if foreign.get("error"):
            warnings.append(
                "Max foreign-exchange scrape error: " + str(foreign.get("error"))
            )

    dropped_settlements = int((merge_dedupe_stats or {}).get("dropped_count") or 0)
    expected_combined = (
        len(max_rows) + len(leumi_checking) + len(leumi_cards) - dropped_settlements
    )
    combined_count = None
    if combined_path and combined_path.exists():
        combined_count = len(read_max_expense_rows(combined_path))
        if combined_count != expected_combined:
            warnings.append(
                f"Combined workbook has {combined_count} row(s) but expected "
                f"{expected_combined} after settlement dedupe "
                f"(sources={len(max_rows)+len(leumi_checking)+len(leumi_cards)}, "
                f"dropped_settlements={dropped_settlements})."
            )

    if merge_dedupe_stats:
        for w in merge_dedupe_stats.get("warnings") or []:
            if str(w) not in warnings:
                warnings.append(str(w))

    report: Dict[str, Any] = {
        "year": year,
        "month": str(month),
        "max": {"parsed_rows": len(max_rows), "amount_sum": max_sum},
        "leumi_checking": {
            "html_data_rows": checking_html_rows,
            "parsed_rows": len(leumi_checking),
            "amount_sum": round(sum(row[1] for row in leumi_checking), 2),
        },
        "leumi_cards": cards_info,
        "combined": {
            "expected_rows": expected_combined,
            "parsed_rows": combined_count,
            "dropped_settlements": dropped_settlements,
        },
        "settlement_dedupe": merge_dedupe_stats or {},
        "card_export_healthy": bool(cards_info.get("healthy")),
        "warnings": warnings,
        "blocking_errors": blocking_errors,
        "ok": len(blocking_errors) == 0,
    }
    return report


def write_reconciliation_report(report: Dict[str, Any], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return path
