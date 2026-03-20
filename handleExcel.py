from openpyxl import load_workbook
from openpyxl.comments import Comment
from openAI import sortExpensesAI
from data import months, monthToExpenseColDict, fullDict
from config import get_paths
from rag_categories import get_category_corrections
from difflib import SequenceMatcher
import re

_CARD_STATEMENT_DUPLICATE_PATTERNS = [
    "מקס איט פיננ~י",
    "ל.מאסטרקרד(יש)",
]


def _normalize_for_fuzzy_name(s: str) -> str:
    return re.sub(r"[^א-תa-zA-Z0-9]", "", str(s or "").lower())


def is_card_statement_duplicate_expense(expense_name: str, threshold: float = 0.66) -> bool:
    """
    Detect card-statement duplicate lines (that already appear in bank statement).
    Uses a little fuzziness to catch small variations/typos.
    """
    normalized = _normalize_for_fuzzy_name(expense_name)
    if not normalized:
        return False
    for pattern in _CARD_STATEMENT_DUPLICATE_PATTERNS:
        score = SequenceMatcher(None, normalized, _normalize_for_fuzzy_name(pattern)).ratio()
        if score >= threshold:
            return True
    return False


def _safe_float_from_string(s: str, context: str = ""):
    """Parse a string to float; return (value, None) or (0.0, error_msg). Used for AI output and debug."""
    if s is None or (isinstance(s, str) and not s.strip()):
        return 0.0, None
    s = str(s).strip().replace("\u200f", "").replace("\u200e", "").replace("₪", "").replace(",", "").strip()
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1].strip()
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
    if not m:
        return 0.0, f"no number found in {s!r}" + (f" ({context})" if context else "")
    try:
        val = float(m.group(0))
        # Expenses in the output workbook should never be negative; remove any leading minus.
        return abs(-val if neg else val), None
    except ValueError as e:
        return 0.0, str(e) + (f" ({context})" if context else "")


def get_allowed_categories(outputWorkbookPath: str) -> list[str]:
    """Read allowed (expense) sub-category labels from template/output workbook (column B, rows 6..12 and 17..105)."""
    wb = load_workbook(outputWorkbookPath, data_only=True)
    ws = wb.active
    categories: list[str] = []
    category_rows = list(range(6, 13)) + list(range(17, 106))
    seen = set()
    for r in category_rows:
        v = ws[f"B{r}"].value
        if v is None:
            continue
        s = str(v).strip()
        if s and s not in seen:
            categories.append(s)
            seen.add(s)
    wb.close()
    return categories


def parse_ai_output_lines(openAIOutput: str) -> tuple[list[dict], list[str]]:
    """
    Parse AI output lines in format: name - cost - category
    Returns (parsed_items, parse_errors).
    """
    parsed: list[dict] = []
    errors: list[str] = []
    eachLineList = [line.strip() for line in openAIOutput.split("\n") if line.strip()]
    for i, val in enumerate(eachLineList):
        line_norm = (
            val.replace("–", "-").replace("—", "-").replace("−", "-").replace("\u00a0", " ")
        )
        # Preferred strict parsing: exact " - " separators.
        parts = val.rsplit(" - ", 2)
        name = cost_str = category = None

        if len(parts) == 3:
            name, cost_str, category = parts
        else:
            # Fallback: parse with a regex to tolerate whitespace around the separators
            # while still respecting the required structure: name - amount - category.
            #
            # Note: expense names may contain '-' in real data, but OpenAI is instructed
            # to replace '-' with '~' in names, so separators are the main '-' tokens.
            m = re.match(
                r"^(?P<name>.+?)\s*-\s*(?P<amount>\(?[-+]?\d+(?:\.\d+)?\)?)\s*-\s*(?P<cat>.+?)\s*$",
                line_norm,
            )
            if not m:
                errors.append(f"Line {i+1} unparseable: {val}")
                continue
            name = m.group("name")
            cost_str = m.group("amount")
            category = m.group("cat")

        cost_value, parse_err = _safe_float_from_string(
            cost_str, context=f"line {i+1}: {val!r}"
        )
        if parse_err is not None:
            # If strict splitting mis-parsed (e.g. because category text contains ' - '),
            # retry regex parsing anchored by the numeric amount.
            m = re.match(
                r"^(?P<name>.+?)\s*-\s*(?P<amount>\(?[-+]?\d+(?:\.\d+)?\)?)\s*-\s*(?P<cat>.+?)\s*$",
                line_norm,
            )
            if m:
                name = m.group("name")
                cost_str = m.group("amount")
                category = m.group("cat")
                cost_value, parse_err = _safe_float_from_string(
                    cost_str, context=f"line {i+1}: {val!r}"
                )

            if parse_err is not None:
                errors.append(f"Line {i+1} invalid cost: {val} ({parse_err})")
                continue
        parsed.append(
            {
                # Normalize a couple of trailing "quote / replacement" artifacts that the model may add.
                # This is intentionally conservative: we only trim a small set of characters at the end.
                "name": re.sub(r"[\uFFFD\"״״׳׳“””']+$", "", name.strip()).strip(),
                "cost": float(cost_value),
                "category": str(category).strip(),
            }
        )
    return parsed, errors


def build_ai_output_from_items(items: list[dict]) -> str:
    """Build normalized AI-output-like text from reviewed items."""
    lines = []
    for item in items:
        lines.append(f"{item['name']} - {item['cost']} - {item['category']}")
    return "\n".join(lines)


def getExpenses(workbook_path):
    print(f"[getExpenses] Loading workbook: {workbook_path}")
    def _to_float(v) -> float:
        if v is None:
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip()
        if not s:
            return 0.0
        # Remove common currency / bidi marks and normalize separators
        s = (
            s.replace("\u200f", "")
            .replace("\u200e", "")
            .replace("₪", "")
            .replace(",", "")
            .strip()
        )
        # Handle parentheses as negative amounts: (123.45)
        neg = False
        if s.startswith("(") and s.endswith(")"):
            neg = True
            s = s[1:-1].strip()
        # Extract first number-looking token
        m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
        if not m:
            return 0.0
        try:
            val = float(m.group(0))
            # Expenses in our workbook should never be negative; remove leading minus signs.
            return abs(-val if neg else val)
        except ValueError:
            return 0.0

    # load the excel workbook
    wb = load_workbook(workbook_path)
    # get the active (default) worksheet
    ws = wb.active
    # expense name column
    expenseNameCol = "B"
    # category column
    categoryCol = "C"
    # expense cost column
    expenseCol = "F"
    # first expense row
    categoryRow = 6
    expensesDict = {}
    totalCost = 0.0
    # loop through the rows and get the expense name and the category, separated by a dash
    while ws[expenseNameCol + str(categoryRow)].value is not None:
        expenseNameCell = ws[expenseNameCol + str(categoryRow)]
        expenseCostCell = ws[expenseCol + str(categoryRow)]
        expenseCategoryCell = ws[categoryCol + str(categoryRow)]
        raw_cost = expenseCostCell.value
        cost_value = _to_float(raw_cost)
        if raw_cost is not None and not isinstance(raw_cost, (int, float)):
            print(f"[getExpenses] Row {categoryRow}: B={expenseNameCell.value!r} F(raw)={raw_cost!r} -> cost={cost_value}")
        # if the expense name is already in the dict, add the cost of the expense to the last cost

        # if the expense name is already in the dict, add the cost of the expense to the last cost
        if expenseNameCell.value in expensesDict:
            expensesDict[expenseNameCell.value][0] += cost_value
            totalCost += cost_value
        # else, add the expense name, it's cost and it's category
        else:
            expensesDict.update(
                {
                    expenseNameCell.value: [
                        cost_value,
                        expenseCategoryCell.value if expenseCategoryCell.value is not None else "",
                    ]
                }
            )
            totalCost += cost_value
        # go to the next row
        categoryRow += 1

    expensesSorted = sortExpensesAI(expensesDict)
    paths = get_paths()
    out_path = paths.data_dir / "expensesDict.txt"
    with out_path.open("w", encoding="utf-8") as file:
        file.write(
            str(expensesDict)
            + "\n"
            + expensesSorted
            + "\n Total cost: "
            + str(totalCost)
        )

    return expensesSorted


# @TODO add the expenses to the final excel file - go through each line in chat's response, for each line, check the name of the expense and it's cost, add the cost to a
# dict with the total cost for each category. then, add the category's cost to the final excel sheet


def fillCells(outputWorkbookPath, openAIOutput, month):
    # load the excel workbook
    wb = load_workbook(outputWorkbookPath)
    # get the active (default) worksheet
    ws = wb.active
    # category column
    categoryCol = "B"
    # expense column
    if month not in months:
        expenseCol = monthToExpenseColDict[12]
    else:
        expenseCol = monthToExpenseColDict[int(month)]

    errorString = ""
    # Rows 6..12 and 17..105 are expense sub-categories we allow matching into.
    category_rows = list(range(6, 13)) + list(range(17, 106))

    # Build allowed categories from the template itself (column B in category_rows)
    allowed_categories: list[str] = []
    seen = set()
    for r in category_rows:
        v = ws[f"{categoryCol}{r}"].value
        if v is None:
            continue
        s = str(v).strip()
        if s and s not in seen:
            allowed_categories.append(s)
            seen.add(s)

    def _normalize(s: str) -> str:
        return (
            str(s)
            .replace("\u200f", "")
            .replace("\u200e", "")
            .replace("׳", "'")
            .replace("״", '"')
            .strip()
        )

    def _normalize_expense_name_for_corrections(s: str) -> str:
        """
        Normalize expense names so `category_corrections.json` matches OpenAI output.
        OpenAI instruction: replace '-' with '~' in expense names.
        """
        s = str(s or "").strip()
        s = s.replace("-", "~")
        s = re.sub(r"\s+", " ", s).strip()
        return s

    # Map normalized fullDict main-category -> list of allowed template sub-categories
    allowed_norm_to_label = {_normalize(cat): cat for cat in allowed_categories}
    main_to_subcats: dict[str, list[str]] = {}
    for main_cat, subcats in fullDict.items():
        if isinstance(subcats, (list, tuple)):
            subs = [allowed_norm_to_label.get(_normalize(s)) for s in subcats]
        else:
            subs = [allowed_norm_to_label.get(_normalize(subcats))]
        subs = [s for s in subs if s]
        if subs:
            main_to_subcats[_normalize(main_cat)] = subs

    def _best_category_match(raw_category: str) -> str:
        """
        Map an AI category to the closest existing sheet *expense sub-category* label.
        If the model returns a main-category by mistake, use it as a hint and restrict matching.
        """
        raw = _normalize(raw_category)
        if not raw:
            return raw_category

        # Candidate list: restrict to the sub-categories under the returned main-category (when possible).
        candidates = main_to_subcats.get(raw, allowed_categories)

        # Exact (normalized) match first
        for cat in candidates:
            if _normalize(cat) == raw:
                return cat

        # Best fuzzy match among candidates
        best_cat = ""
        best_score = 0.0
        for cat in candidates:
            score = SequenceMatcher(None, raw, _normalize(cat)).ratio()
            if score > best_score:
                best_score = score
                best_cat = cat

        # Keep a reasonable threshold; otherwise fall back to an "אחר/Other" bucket if present.
        if best_cat and best_score >= 0.74:
            return best_cat
        for cat in candidates:
            if "אחר" in cat or "Other" in cat:
                return cat
        return best_cat or raw_category

    # initiate a list of the expenses names and their associated category
    expenseNameCostCategoryDict = {}
    # User corrections override AI: expense name -> exact template category
    user_corrections = get_category_corrections()
    if user_corrections:
        print(f"[fillCells] Applying {len(user_corrections)} user category correction(s)")
    user_corrections_norm = {
        _normalize_expense_name_for_corrections(k): v for k, v in (user_corrections or {}).items()
    }

    # split OpenAI's output to separate lines
    eachLineList = [line.strip() for line in openAIOutput.split("\n") if line.strip()]
    print(f"[fillCells] Parsing {len(eachLineList)} lines from AI output; month={month}, outputWorkbookPath={outputWorkbookPath}")
    # for each line
    for i, val in enumerate(eachLineList):
        line_no = i + 1
        line_norm = (
            val.replace("–", "-").replace("—", "-").replace("−", "-").replace("\u00a0", " ")
        )

        # robustly split between the expense, cost and category
        parts = val.rsplit(" - ", 2)
        name = cost_str = category = None
        cost_value = 0.0
        parse_err = None

        if len(parts) == 3:
            name, cost_str, category = parts
            cost_value, parse_err = _safe_float_from_string(
                cost_str, context=f"line {line_no}: {val!r}"
            )

        # If strict split mis-parses (often because category contains ' - '),
        # retry regex anchored by the numeric amount token.
        if parse_err is not None or name is None or category is None:
            m = re.match(
                r"^(?P<name>.+?)\s*-\s*(?P<amount>\(?[-+]?\d+(?:\.\d+)?\)?)\s*-\s*(?P<cat>.+?)\s*$",
                line_norm,
            )
            if not m:
                print(
                    f"[fillCells] Line {line_no}: unparseable (expected 'name - cost - category'): {val!r}"
                )
                errorString += f"Unparseable line: {val}\n"
                continue
            name = m.group("name")
            cost_str = m.group("amount")
            category = m.group("cat")
            cost_value, parse_err = _safe_float_from_string(
                cost_str, context=f"line {line_no}: {val!r}"
            )
            if parse_err is not None:
                print(f"[fillCells] Line {line_no}: cost parse issue: {parse_err}")
                errorString += f"Invalid cost value: {val}\n"
                continue
        # If user specified a correction for this expense, use it; otherwise use AI category
        normalized_name_key = _normalize_expense_name_for_corrections(name)
        if normalized_name_key in user_corrections_norm:
            category = user_corrections_norm[normalized_name_key]
        expenseNameCostCategoryDict[name] = [cost_value, _best_category_match(category)]

    for val in expenseNameCostCategoryDict:
        found = False
        target_category = expenseNameCostCategoryDict.get(val)[1]
        add_amount = expenseNameCostCategoryDict.get(val)[0]
        for categoryRow in category_rows:
            cell = ws[f"{categoryCol}{categoryRow}"]
            if cell.value is None:
                continue
            if cell.value == target_category:
                existing_val = ws[f"{expenseCol}{categoryRow}"].value
                if existing_val is None:
                    ws[f"{expenseCol}{categoryRow}"].value = add_amount
                else:
                    existing_float, _ = _safe_float_from_string(
                        str(existing_val), context=f"cell {expenseCol}{categoryRow}"
                    )
                    ws[f"{expenseCol}{categoryRow}"].value = existing_float + add_amount

                # Add comment with expense origin
                if ws[f"{expenseCol}{categoryRow}"].comment is None:
                    ws[f"{expenseCol}{categoryRow}"].comment = Comment("", "Automated")
                comment = Comment(
                    val + " - " + str(add_amount),
                    "Automated",
                )
                if comment.text != "":
                    ws[f"{expenseCol}{categoryRow}"].comment.text += comment.text + "\n"
                found = True
        if not found:
            errorString += (
                str(val)
                + " "
                + str(expenseNameCostCategoryDict.get(val)[0])
                + str(expenseNameCostCategoryDict.get(val)[1])
                + "\n"
            )

    paths = get_paths()
    out_path = paths.data_dir / "errors.txt"
    with out_path.open("w", encoding="utf-8") as file:
        file.write(errorString)

    wb.save(outputWorkbookPath)


# getExpenses(r"C:\Users\liido\Downloads\transaction-details_export_1728243409088.xlsx")

# string = """TRES JOLIE - 20 - רהיטים, כלי בית וגן
# מאפיתת אורן משי באר שבע - 44 - מצרכי מזון
# סופר פארם מצדה ב"ש - 46.8 - תרופות, בדיקות וטיפולים רפואיים
# 7 בעיר הבלוק - 39.8 - מצרכי מזון
# סהרה אלקטרוניקה בע"מ - 184 - מתנות
# חברת החשמל לישראל בע"מ - 368.73 - חשבון חשמל
# פז אפליקציית יילו - 683.15 - חשבון גז
# פרפל יבוא ושיווק - 77 - תכשיטים וקוסמטיקה
# סופרמרקט הגשר - 15 - מצרכי מזון
# אייבורי מחשבים - 195 - גאדג'טים ואלקטרוניקה
# אפריל פארק הקרח - 127 - מוצרי היגיינה
# חברת פרטנר תקשורת בע"מ (ה - 119 - סלולארי
# מחסני השוק אינטרנט - 1350.04 - מצרכי מזון
# היי קוקי - 122.55 - אוכל בחוץ (כולל מסעדות, בתי קפה, משלוחים)
# כביש 6 - 76.38 - כביש 6 וחוצה צפון
# WOLT - 320 - אוכל בחוץ (כולל מסעדות, בתי קפה, משלוחים)
# סופר קופיקס - 63.4 - מצרכי מזון
# פלאפון חשבון תקופתי - 41.84 - סלולארי
# מ. התחבורה ר.רכב - 906 - דו"חות וקנסות
# שוהם סטוק בע"מ - 58.74 - מצרכי מזון
# AIG ביטוח רכב - 2855 - ביטוח רכב
# המילטון חשמל ואלקטרוניקה - 450 - גאדג'טים ואלקטרוניקה
# סיטי שופ - 79.28 - ספרים ומוזיקה
# הסטוק גרנד קניון ב"ש~צמרת - 15.9 - מתנות
# תכשיטי ישראל - 30 - מתנות
# אושר עד באר שבע - 140.7 - מצרכי מזון
# פנגו חשבונית חודשית - 9.32 - הוצאות תחבורה ציבורית ואופניים
# מי שבע בע"מ - 114.02 - דו"חות וקנסות
# רוזה & פולה - 40 - אוכל בחוץ (כולל מסעדות, בתי קפה, משלוחים)
# דמי כרטיס - 0.0 - משיכת כספים (כספומט)
# עיריית באר שבע - 390.58 - דו"חות וקנסות
# BIT - 100 - העברת כספים
# יעקב כהן - 60 - תספורת
# מס הכנסה עצמאים וחברות - 500 - דו"חות וקנסות"""
#
# fillCells("C:/Users/liido/Downloads/expenses_2024.xlsx", string, 12)
