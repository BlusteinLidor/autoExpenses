import csv
from openpyxl import load_workbook

from handleExcel import resolve_transaction_is_income


def parse_csv_charge_amount(raw: str) -> tuple[float, bool]:
    """
    Parse Max CSV 'סכום חיוב' values.
    Expenses use a minus sign (e.g. ₪-6.20); income/credits are positive (e.g. ₪500.00).
    Returns (absolute_amount, is_income).
    """
    if not raw:
        return 0.0, False
    s = (
        str(raw)
        .replace("₪", "")
        .replace(",", "")
        .replace("\u200f", "")
        .replace("\u200e", "")
        .strip()
    )
    if not s:
        return 0.0, False
    try:
        signed = float(s)
    except ValueError:
        return 0.0, False
    return abs(signed), signed > 0


def parse_amount(
    transactions_csv_path="transactions.csv", target_excel_path="target_file.xlsx"
) -> int:
    # Load the source rows from CSV
    with open(transactions_csv_path, newline="", encoding="utf-8") as src_file:
        reader = csv.reader(src_file)
        source_rows = list(reader)[1:]
    if not source_rows:
        return 0

    # Load the target Excel file
    wb = load_workbook(target_excel_path)
    ws = wb.active  # or wb['SheetName'] if you want a specific sheet

    # Find the first empty row based on the first column (A)
    first_empty_row = None
    total_amount_row = None
    added_amount = 0
    i = 0
    for row in range(
        1, ws.max_row + 2
    ):  # +2 to handle the case where the sheet is completely full
        if ws.cell(row=row, column=1).value in (None, ""):
            first_empty_row = row
            total_amount_row = first_empty_row + 2
            break

    # Insert rows (from bottom to top so they stay in order)
    for i, row_data in enumerate(source_rows):
        business_name = str(row_data[0]).strip() if row_data else ""
        ws.insert_rows(first_empty_row + i)  # shift down, make space
        for col, value in enumerate(row_data, start=1):
            # ws.cell(row=first_empty_row + i, column=col, value=value)
            if "₪" in value:
                amount, is_income = parse_csv_charge_amount(value)
                is_income = resolve_transaction_is_income(business_name, is_income)
                print(f"Parsed value: {amount} (income={is_income})")
                added_amount += amount
                ws.cell(row=first_empty_row + i, column=6, value=amount)
                if is_income:
                    ws.cell(row=first_empty_row + i, column=7, value=1)
            else:
                ws.cell(row=first_empty_row + i, column=col + 1, value=value)

    totalAmountCell = ws[f"A{total_amount_row + i + 1}"]
    totalAmountValue = totalAmountCell.value
    totalAmountParsed = totalAmountValue.replace("₪", "").replace(",", "").strip()
    floatTotalAmount = abs(float(totalAmountParsed)) + added_amount
    print(f"Total amount parsed: {floatTotalAmount}")
    print(f"Added amount: {added_amount}")
    print(f"New total amount: {floatTotalAmount}")
    totalAmountCell.value = f"₪{floatTotalAmount:.2f}"

    # Save back
    wb.save(target_excel_path)
    return len(source_rows)
