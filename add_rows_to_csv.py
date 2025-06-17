import csv
from openpyxl import load_workbook


def parse_amount(
    transactions_csv_path="transactions.csv", target_excel_path="target_file.xlsx"
):
    # Load the source rows from CSV
    with open(transactions_csv_path, newline="", encoding="utf-8") as src_file:
        reader = csv.reader(src_file)
        source_rows = list(reader)[1:]

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
        ws.insert_rows(first_empty_row + i)  # shift down, make space
        for col, value in enumerate(row_data, start=1):
            # ws.cell(row=first_empty_row + i, column=col, value=value)
            if "₪" in value:
                value = (
                    value.replace("₪", "")
                    .replace(",", "")
                    .replace("\u200f", "")
                    .strip()
                )
                value = float(value) if value else 0.0
                print(f"Parsed value: {value}")
                added_amount += float(value)
                ws.cell(row=first_empty_row + i, column=6, value=value)
            else:
                ws.cell(row=first_empty_row + i, column=col + 1, value=value)

    totalAmountCell = ws[f"A{total_amount_row + i + 1}"]
    totalAmountValue = totalAmountCell.value
    totalAmountParsed = totalAmountValue.replace("₪", "").replace(",", "").strip()
    floatTotalAmount = float(totalAmountParsed) + added_amount
    print(f"Total amount parsed: {floatTotalAmount}")
    print(f"Added amount: {added_amount}")
    print(f"New total amount: {floatTotalAmount}")
    totalAmountCell.value = f"₪{floatTotalAmount:.2f}"

    # Save back
    wb.save(target_excel_path)
