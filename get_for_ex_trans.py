from bs4 import BeautifulSoup
import pandas as pd


def get_foreign_exchange_transactions(
    htmlFilePath="foreign_exchange_transactions.html", outputCsvPath="transactions.csv"
):
    # Load the HTML content (you can also read from a file)
    with open(htmlFilePath, encoding="utf-8") as f:
        html = f.read()

    soup = BeautifulSoup(html, "html.parser")

    # Extract all transaction rows
    rows = soup.select("div.row.body")

    # Prepare list for storing extracted data
    data = []

    for row in rows:
        # date = row.select_one(".cell.date .text")
        merchant = row.select_one(".cell.name .text")
        category = row.select_one(".cell.category")
        # card = row.select_one(".cell.card-number")
        # type_ = row.select_one(".cell.type span")
        amount = row.select_one(".cell.sum .ltr-sum")

        data.append(
            {
                # "Date": date.text.strip() if date else "",
                "שם בית העסק": merchant.text.strip() if merchant else "",
                "קטגוריה": category.text.strip() if category else "",
                # "Card": card.text.strip() if card else "",
                # "Type": type_.text.strip() if type_ else "",
                "סכום חיוב": amount.text.strip() if amount else "",
            }
        )

    # Convert to pandas DataFrame
    df = pd.DataFrame(data)

    # Display or save
    # print(df)
    df.to_csv(outputCsvPath, index=False, encoding="utf-8-sig")
