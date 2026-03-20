"""
Merge Max and Leumi exports into a single workbook for getExpenses.

- Max: xlsx with columns B=expense name, C=category, F=cost from row 5.
- Leumi: HTML exported as .xls (transactions + credit cards); we parse and normalize
  to the same (name, cost) shape.
"""

from pathlib import Path
from typing import List, Tuple, Optional

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl import load_workbook


def _parse_number_hebrew(s: str) -> float:
    """Parse a number that may use Hebrew locale (e.g. 1,234.56 or 1.234,56)."""
    if s is None or (isinstance(s, float) and (s != s or s == 0)):
        return 0.0
    s = str(s).strip().replace("\u200f", "").replace(",", "").replace(" ", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _extract_text(cell) -> str:
    """Get single-line text from a table cell (BeautifulSoup td)."""
    if cell is None:
        return ""
    text = cell.get_text(separator=" ", strip=True)
    return " ".join(text.split()) if text else ""


def read_leumi_transactions_html(path: Path) -> List[Tuple[str, float]]:
    """
    Parse Leumi checking-account export (HTML saved as .xls).
    Columns: תאריך, תאריך ערך, תיאור, אסמכתא, בחובה (debit), בזכות (credit), ...
    Returns list of (description, amount) where amount is debit - credit.
    """
    rows: List[Tuple[str, float]] = []
    raw = path.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(raw, "html.parser")
    table = soup.find("table", class_="xlTable")
    if not table:
        return rows
    trs = table.find_all("tr")
    if len(trs) < 3:
        return rows
    # Skip title row and header row
    for tr in trs[2:]:
        tds = tr.find_all("td")
        if len(tds) < 6:
            continue
        # תיאור = index 2, בחובה = 4, בזכות = 5
        desc = _extract_text(tds[2])
        debit = _parse_number_hebrew(_extract_text(tds[4]))
        credit = _parse_number_hebrew(_extract_text(tds[5]))
        if not desc:
            continue
        # Outgoing = debit; incoming = credit. Expense = debit - credit.
        # Normalize to abs: expenses in our output workbook should be non-negative.
        amount = abs(debit - credit)
        if amount != 0:
            rows.append((desc, amount))
    return rows


def read_leumi_credit_cards_html(path: Path) -> List[Tuple[str, float]]:
    """
    Parse Leumi credit-cards export (HTML saved as .xls).
    Columns: תאריך העסקה, שם בית העסק, סכום העסקה, סוג העסקה, פרטים, סכום חיוב
    Returns list of (business_name, charge_amount).
    """
    rows: List[Tuple[str, float]] = []
    raw = path.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(raw, "html.parser")
    table = soup.find("table", class_="xlTable")
    if not table:
        return rows
    trs = table.find_all("tr")
    if len(trs) < 3:
        return rows
    for tr in trs[2:]:
        tds = tr.find_all("td")
        if len(tds) < 6:
            continue
        # שם בית העסק = 1, סכום חיוב = 5
        name = _extract_text(tds[1])
        amount = abs(_parse_number_hebrew(_extract_text(tds[5])))
        if name and amount != 0:
            rows.append((name, amount))
    return rows


def read_max_expense_rows(path: Path) -> List[Tuple[str, float, Optional[str]]]:
    """
    Read expense name (B), main category (C) and cost (F) from row 5 until first empty B.
    We preserve main-category so OpenAI has guidance when mapping to sub-categories.
    """
    rows: List[Tuple[str, float, Optional[str]]] = []
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    row = 5
    while True:
        name_cell = ws[f"B{row}"].value
        if name_cell is None or (isinstance(name_cell, str) and not name_cell.strip()):
            break
        cost_cell = ws[f"F{row}"].value
        try:
            cost = abs(float(cost_cell)) if cost_cell is not None else 0.0
        except (TypeError, ValueError) as e:
            print(f"[merge_expenses] Row {row} non-numeric cost in F: {cost_cell!r} -> 0.0 ({e})")
            cost = 0.0
        category_cell = ws[f"C{row}"].value
        category = str(category_cell).strip() if category_cell is not None else None
        rows.append((str(name_cell).strip(), cost, category or None))
        row += 1
    wb.close()
    return rows


def write_combined_workbook(rows: List[Tuple[str, float, Optional[str]]], output_path: Path) -> None:
    """
    Write a workbook compatible with getExpenses: from row 5, B=name, C=category (main), F=cost.
    """
    wb = Workbook()
    ws = wb.active
    start_row = 5
    for i, (name, cost, category) in enumerate(rows):
        r = start_row + i
        ws[f"B{r}"] = name
        ws[f"C{r}"] = category
        ws[f"F{r}"] = cost
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def merge_max_and_leumi(
    max_excel_path: Path,
    leumi_transactions_path: Path,
    leumi_cards_path: Path,
    output_path: Path,
) -> Path:
    """
    Combine Max export and Leumi (transactions + cards) into one workbook at output_path.
    Returns output_path. Order: Max rows first, then Leumi transactions, then Leumi cards.
    """
    all_rows: List[Tuple[str, float, Optional[str]]] = []

    if max_excel_path.exists():
        all_rows.extend(read_max_expense_rows(max_excel_path))

    if leumi_transactions_path.exists():
        for name, amount in read_leumi_transactions_html(leumi_transactions_path):
            all_rows.append((name, amount, None))

    if leumi_cards_path.exists():
        for name, amount in read_leumi_credit_cards_html(leumi_cards_path):
            all_rows.append((name, amount, None))

    write_combined_workbook(all_rows, output_path)
    return output_path
