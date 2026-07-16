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


def _existing_expense_keys(ws) -> set[tuple[str, float, bool]]:
    """Collect (name, amount, is_income) already present from row 5 onward."""
    keys: set[tuple[str, float, bool]] = set()
    row = 5
    while True:
        name_cell = ws[f"B{row}"].value
        if name_cell is None or (isinstance(name_cell, str) and not name_cell.strip()):
            break
        name = str(name_cell).strip()
        cost_cell = ws[f"F{row}"].value
        try:
            cost = abs(float(cost_cell)) if cost_cell is not None else 0.0
        except (TypeError, ValueError):
            cost = 0.0
        income_flag = ws[f"G{row}"].value
        is_income = resolve_transaction_is_income(
            name,
            income_flag in (1, True, "1", "income", "yes"),
        )
        keys.add((name, round(cost, 2), bool(is_income)))
        row += 1
    return keys


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

    existing_keys = _existing_expense_keys(ws)
    # Also track keys within this CSV batch so CSV self-duplicates are skipped.
    pending_keys = set(existing_keys)

    # Find the first empty row based on the first column (A)
    first_empty_row = None
    total_amount_row = None
    added_amount = 0
    inserted = 0
    skipped = 0
    for row in range(
        1, ws.max_row + 2
    ):  # +2 to handle the case where the sheet is completely full
        if ws.cell(row=row, column=1).value in (None, ""):
            first_empty_row = row
            total_amount_row = first_empty_row + 2
            break

    # Insert rows (from bottom to top so they stay in order)
    for row_data in source_rows:
        business_name = str(row_data[0]).strip() if row_data else ""
        amount = 0.0
        is_income = False
        has_amount = False
        for value in row_data:
            if "₪" in str(value):
                amount, is_income = parse_csv_charge_amount(value)
                is_income = resolve_transaction_is_income(business_name, is_income)
                has_amount = True
                break

        if business_name and has_amount:
            key = (business_name, round(amount, 2), bool(is_income))
            if key in pending_keys:
                skipped += 1
                print(
                    f"Skipping duplicate already in workbook/CSV: "
                    f"{business_name} {amount} (income={is_income})"
                )
                continue
            pending_keys.add(key)

        ws.insert_rows(first_empty_row + inserted)  # shift down, make space
        for col, value in enumerate(row_data, start=1):
            if "₪" in str(value):
                print(f"Parsed value: {amount} (income={is_income})")
                added_amount += amount
                ws.cell(row=first_empty_row + inserted, column=6, value=amount)
                if is_income:
                    ws.cell(row=first_empty_row + inserted, column=7, value=1)
            else:
                ws.cell(row=first_empty_row + inserted, column=col + 1, value=value)
        inserted += 1

    if inserted == 0:
        print(f"No new rows to append (skipped {skipped} duplicate(s)).")
        wb.close()
        return 0

    totalAmountCell = ws[f"A{total_amount_row + inserted}"]
    totalAmountValue = totalAmountCell.value
    totalAmountParsed = str(totalAmountValue).replace("₪", "").replace(",", "").strip()
    floatTotalAmount = abs(float(totalAmountParsed)) + added_amount
    print(f"Total amount parsed: {floatTotalAmount}")
    print(f"Added amount: {added_amount}")
    print(f"New total amount: {floatTotalAmount}")
    if skipped:
        print(f"Skipped {skipped} duplicate row(s) already present in the workbook.")
    totalAmountCell.value = f"₪{floatTotalAmount:.2f}"

    # Save back
    wb.save(target_excel_path)
    return inserted
