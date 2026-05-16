from bs4 import BeautifulSoup
import pandas as pd


def _extract_transaction_rows(soup: BeautifulSoup) -> list:
    """Extract transaction rows from Max table HTML with selector fallbacks."""
    row_selectors = (
        "motion.div.row.body",
        "motion-table-row",
        "motion-table-row.motion-table-row",
        "motion-table-row.motion-table-row.motion-table-row",
        "div.row.body",
        ".row.body",
    )
    for selector in row_selectors:
        rows = soup.select(selector)
        if rows:
            return rows
    return []


def _row_to_record(row) -> dict | None:
    merchant = row.select_one(".cell.name .text") or row.select_one(".cell.name")
    category = row.select_one(".cell.category")
    amount = row.select_one(".cell.sum .ltr-sum") or row.select_one(".cell.sum")

    merchant_text = merchant.get_text(strip=True) if merchant else ""
    category_text = category.get_text(strip=True) if category else ""
    amount_text = amount.get_text(strip=True) if amount else ""

    if not merchant_text and not amount_text:
        return None

    if merchant_text in ("", 'סה"כ', "סהכ"):
        return None

    return {
        "שם בית העסק": merchant_text,
        "קטגוריה": category_text,
        "סכום חיוב": amount_text,
    }


def _parse_transactions_from_html(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    records: list[dict] = []
    seen: set[tuple[str, str, str]] = set()
    for row in _extract_transaction_rows(soup):
        record = _row_to_record(row)
        if record is None:
            continue
        key = (
            record["שם בית העסק"],
            record["קטגוריה"],
            record["סכום חיוב"],
        )
        if key in seen:
            continue
        seen.add(key)
        records.append(record)
    return records


TRANSACTION_CSV_COLUMNS = ["שם בית העסק", "קטגוריה", "סכום חיוב"]


def _write_transactions_csv(records: list[dict], output_csv_path: str) -> int:
    df = pd.DataFrame(records, columns=TRANSACTION_CSV_COLUMNS)
    df.to_csv(output_csv_path, index=False, encoding="utf-8-sig")
    return len(records)


def write_empty_transactions_csv(output_csv_path: str) -> int:
    return _write_transactions_csv([], output_csv_path)


def get_immediate_transactions(
    htmlFilePath="immediate_transactions.html", outputCsvPath="transactions.csv"
) -> int:
    with open(htmlFilePath, encoding="utf-8") as f:
        html = f.read()
    records = _parse_transactions_from_html(html)
    return _write_transactions_csv(records, outputCsvPath)


def get_foreign_exchange_transactions(
    htmlFilePath="foreign_exchange_transactions.html", outputCsvPath="transactions.csv"
) -> int:
    with open(htmlFilePath, encoding="utf-8") as f:
        html = f.read()
    records = _parse_transactions_from_html(html)
    return _write_transactions_csv(records, outputCsvPath)
