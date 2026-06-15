from openpyxl import load_workbook
from openpyxl.comments import Comment
from openAI import sortExpensesAI
from data import months, monthToExpenseColDict, fullDict
from config import (
    TEMPLATE_INCOME_CATEGORY_ROW_RANGES,
    TEMPLATE_INVESTMENT_CATEGORY_ROW_RANGES,
    TEMPLATE_LUXURY_EXPENSE_CATEGORY_ROW_RANGES,
    TEMPLATE_NECESSARY_EXPENSE_CATEGORY_ROW_RANGES,
    get_paths,
    template_category_rows,
)
from rag_categories import get_category_corrections
from difflib import SequenceMatcher
import re

_CARD_STATEMENT_DUPLICATE_PATTERNS = [
    "מקס איט פיננ~י",
    "ל.מאסטרקרד(יש)",
]

EXPENSE_SOURCE_MAX = "Max"
EXPENSE_SOURCE_LEUMI_CHECKING = 'לאומי עו"ש'
EXPENSE_SOURCE_LEUMI_CARD = "לאומי כרטיס"
EXPENSE_SOURCE_COL = "H"
DEFAULT_EXPENSE_SOURCE = EXPENSE_SOURCE_MAX

_DEFAULT_INCOME_FALLBACK_CATEGORY = "הכנסה אחרת / חד פעמית"
# Explicit income-source overrides: when an expense name matches one of these
# patterns, force it into the matching income category (rows 6..12 in template).
_INCOME_NAME_CATEGORY_OVERRIDES = {
    "עירית מגדל ה~י": "שכר עבודה אלה - נטו",
    "נאנומושן בע\"~י": "שכר עבודה לידור - נטו",
}

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


def _normalize_income_name_key(s: str) -> str:
    """
    Normalize source names for income override matching.
    Handles model quirks such as replacing '-' with '~'.
    """
    normalized = str(s or "").strip().lower()
    normalized = normalized.replace("-", "~")
    normalized = normalized.replace("״", '"').replace("׳", "'")
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized


# Max credits that fund investments (not income). Positive card amounts are otherwise
# treated as income; these names must stay in the investment section of the template.
_INVESTMENT_TRANSFER_NAME_KEYS = frozenset(
    {
        _normalize_income_name_key("צובר ושב"),
    }
)
_INVESTMENT_NAME_CATEGORY_OVERRIDES = {
    "צובר ושב": "הפקדות לפוליסת חיסכון פיננסי",
}


def is_investment_transfer_name(name: str) -> bool:
    return _normalize_income_name_key(name) in _INVESTMENT_TRANSFER_NAME_KEYS


def resolve_transaction_is_income(name: str, raw_is_income: bool) -> bool:
    if is_investment_transfer_name(name):
        return False
    return bool(raw_is_income)


def _normalize_expense_name_for_corrections(s: str) -> str:
    """
    Normalize expense names so `category_corrections.json` matches OpenAI output.
    OpenAI instruction: replace '-' with '~' in expense names.
    """
    normalized = str(s or "").strip()
    normalized = normalized.replace("-", "~")
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


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


_CATEGORY_GROUP_SPECS: tuple[tuple[str, str, tuple[tuple[int, int], ...]], ...] = (
    ("income", "Income", TEMPLATE_INCOME_CATEGORY_ROW_RANGES),
    ("necessary_expenses", "Necessary Expenses", TEMPLATE_NECESSARY_EXPENSE_CATEGORY_ROW_RANGES),
    ("luxury_expenses", "Luxury Expenses", TEMPLATE_LUXURY_EXPENSE_CATEGORY_ROW_RANGES),
    ("investments", "Investments", TEMPLATE_INVESTMENT_CATEGORY_ROW_RANGES),
)


def _categories_from_row_ranges(
    ws,
    row_ranges: tuple[tuple[int, int], ...],
    seen: set[str] | None = None,
) -> list[str]:
    categories: list[str] = []
    local_seen = seen if seen is not None else set()
    for start, end_exclusive in row_ranges:
        for r in range(start, end_exclusive):
            v = ws[f"B{r}"].value
            if v is None:
                continue
            s = str(v).strip()
            if s and s not in local_seen:
                categories.append(s)
                local_seen.add(s)
    return categories


def get_allowed_category_groups(outputWorkbookPath: str) -> list[dict]:
    """Read allowed sub-categories grouped by template section (income, expenses, investments)."""
    wb = load_workbook(outputWorkbookPath, data_only=True)
    ws = wb.active
    seen: set[str] = set()
    groups: list[dict] = []
    for group_id, label, row_ranges in _CATEGORY_GROUP_SPECS:
        categories = _categories_from_row_ranges(ws, row_ranges, seen)
        if categories:
            groups.append({"id": group_id, "label": label, "categories": categories})
    wb.close()
    return groups


def get_allowed_categories(outputWorkbookPath: str) -> list[str]:
    """Read allowed sub-category labels from template/output workbook (column B, category rows only)."""
    return [
        category
        for group in get_allowed_category_groups(outputWorkbookPath)
        for category in group["categories"]
    ]


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


def _normalize_review_name_key(name: str) -> str:
    return _normalize_expense_name_for_corrections(name)


def _expense_dict_is_income(expense_values) -> bool:
    if not expense_values or len(expense_values) < 3:
        return False
    return bool(expense_values[2])


def _merge_expense_sources(existing_source: str | None, new_source: str) -> str:
    sources: list[str] = []
    for part in str(existing_source or "").split(","):
        label = part.strip()
        if label and label not in sources:
            sources.append(label)
    if new_source and new_source not in sources:
        sources.append(new_source)
    return ", ".join(sources) if sources else DEFAULT_EXPENSE_SOURCE


def _lookup_expense_source_from_dict(expenses_dict: dict | None, name: str) -> str:
    if not expenses_dict:
        return DEFAULT_EXPENSE_SOURCE
    direct = expenses_dict.get(name)
    if direct is not None and len(direct) > 3 and direct[3]:
        return str(direct[3])
    normalized = _normalize_review_name_key(name)
    for expense_name, expense_values in expenses_dict.items():
        if _normalize_review_name_key(str(expense_name)) == normalized:
            if len(expense_values) > 3 and expense_values[3]:
                return str(expense_values[3])
    return DEFAULT_EXPENSE_SOURCE


def _lookup_is_income_from_dict(expenses_dict: dict | None, name: str) -> bool:
    if is_investment_transfer_name(name):
        return False
    if not expenses_dict:
        return False
    direct = expenses_dict.get(name)
    if direct is not None:
        return _expense_dict_is_income(direct)
    normalized = _normalize_review_name_key(name)
    for expense_name, expense_values in expenses_dict.items():
        if _normalize_review_name_key(str(expense_name)) == normalized:
            return _expense_dict_is_income(expense_values)
    return False


def _workbook_row_is_income(cell_value) -> bool:
    return cell_value in (1, True, "1", "income", "yes")


def parse_errors_to_review_items(parse_errors: list[str]) -> list[dict]:
    """Turn AI parse failures into review rows the user can categorize manually."""
    review_items: list[dict] = []
    line_pattern = re.compile(
        r"^(?P<name>.+?)\s*-\s*(?P<amount>\(?[-+]?\d+(?:\.\d+)?\)?)\s*-\s*(?P<cat>.+?)\s*$"
    )
    for err in parse_errors:
        raw_line = err
        match = re.search(r":\s*(.+)$", err)
        if match:
            raw_line = match.group(1).strip()
        line_norm = (
            raw_line.replace("–", "-")
            .replace("—", "-")
            .replace("−", "-")
            .replace("\u00a0", " ")
        )
        parsed_name = ""
        parsed_cost = 0.0
        parsed_category = ""
        line_match = line_pattern.match(line_norm)
        if line_match:
            parsed_name = line_match.group("name").strip()
            parsed_category = line_match.group("cat").strip()
            parsed_cost, _ = _safe_float_from_string(line_match.group("amount"))
        review_items.append(
            {
                "name": parsed_name or raw_line[:120] or "Unknown expense",
                "cost": float(parsed_cost),
                "category": parsed_category,
                "is_possible_duplicate": bool(
                    parsed_name and is_card_statement_duplicate_expense(parsed_name)
                ),
                "needs_manual_review": True,
                "error_reason": err,
            }
        )
    return review_items


def find_expenses_missing_from_parsed(
    expenses_dict: dict,
    parsed_items: list[dict],
) -> list[dict]:
    """Find source workbook expenses that never made it into parsed AI output."""
    parsed_keys: list[str] = []
    for item in parsed_items:
        key = _normalize_review_name_key(str(item.get("name", "")))
        if key:
            parsed_keys.append(key)

    missing_items: list[dict] = []
    for expense_name, expense_values in expenses_dict.items():
        source_key = _normalize_review_name_key(str(expense_name))
        if not source_key:
            continue
        matched = False
        for parsed_key in parsed_keys:
            if parsed_key == source_key:
                matched = True
                break
            if SequenceMatcher(None, parsed_key, source_key).ratio() >= 0.9:
                matched = True
                break
        if matched:
            continue
        cost_value = float(expense_values[0]) if expense_values else 0.0
        source_category = str(expense_values[1]) if len(expense_values) > 1 else ""
        missing_items.append(
            {
                "name": str(expense_name),
                "cost": cost_value,
                "category": source_category,
                "is_income": _expense_dict_is_income(expense_values),
                "is_possible_duplicate": is_card_statement_duplicate_expense(str(expense_name)),
                "needs_manual_review": True,
                "error_reason": "AI did not return a categorized line for this expense.",
            }
        )
    return missing_items


def build_review_items(
    parsed_items: list[dict],
    parse_errors: list[str],
    expenses_dict: dict | None = None,
) -> tuple[list[dict], list[str]]:
    """
    Merge successfully parsed AI items with rows that need manual review.
    Returns (review_items, warnings).
    """
    warnings: list[str] = []
    review_items: list[dict] = []
    seen_keys: set[tuple[str, float]] = set()

    def _append_item(item: dict) -> None:
        name = str(item.get("name", "")).strip()
        cost = abs(float(item.get("cost", 0.0)))
        is_income = resolve_transaction_is_income(
            name,
            bool(item.get("is_income", False))
            or _lookup_is_income_from_dict(expenses_dict, name),
        )
        category = str(item.get("category", "")).strip()
        investment_category = {
            _normalize_income_name_key(k): v
            for k, v in _INVESTMENT_NAME_CATEGORY_OVERRIDES.items()
        }.get(_normalize_income_name_key(name))
        if investment_category:
            category = investment_category
        key = (_normalize_review_name_key(name), round(cost, 2), is_income)
        if key in seen_keys:
            return
        seen_keys.add(key)
        review_items.append(
            {
                "name": name,
                "cost": cost,
                "category": category,
                "is_income": is_income,
                "is_possible_duplicate": bool(item.get("is_possible_duplicate", False)),
                "needs_manual_review": bool(item.get("needs_manual_review", False)),
                "error_reason": (
                    None
                    if item.get("error_reason") is None
                    else str(item.get("error_reason", "")).strip() or None
                ),
            }
        )

    for item in parsed_items:
        _append_item(
            {
                **item,
                "is_possible_duplicate": is_card_statement_duplicate_expense(
                    str(item.get("name", ""))
                ),
                "needs_manual_review": False,
                "error_reason": None,
            }
        )

    for item in parse_errors_to_review_items(parse_errors):
        _append_item(item)

    if expenses_dict:
        for item in find_expenses_missing_from_parsed(expenses_dict, parsed_items):
            _append_item(item)

    manual_count = sum(1 for item in review_items if item.get("needs_manual_review"))
    income_count = sum(1 for item in review_items if item.get("is_income"))
    if manual_count:
        warnings.append(
            f"{manual_count} transaction(s) need manual categorization before finalize."
        )
    if income_count:
        warnings.append(
            f"{income_count} income transaction(s) detected — use income categories when reviewing."
        )
    return review_items, warnings


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
        incomeFlagCell = ws["G" + str(categoryRow)]
        sourceCell = ws[EXPENSE_SOURCE_COL + str(categoryRow)]
        raw_cost = expenseCostCell.value
        cost_value = _to_float(raw_cost)
        is_income = resolve_transaction_is_income(
            str(expenseNameCell.value or ""),
            _workbook_row_is_income(incomeFlagCell.value),
        )
        source_label = (
            str(sourceCell.value).strip()
            if sourceCell.value is not None and str(sourceCell.value).strip()
            else DEFAULT_EXPENSE_SOURCE
        )
        if raw_cost is not None and not isinstance(raw_cost, (int, float)):
            print(f"[getExpenses] Row {categoryRow}: B={expenseNameCell.value!r} F(raw)={raw_cost!r} -> cost={cost_value}")
        # if the expense name is already in the dict, add the cost of the expense to the last cost

        # if the expense name is already in the dict, add the cost of the expense to the last cost
        if expenseNameCell.value in expensesDict:
            existing = expensesDict[expenseNameCell.value]
            existing[0] += cost_value
            if len(existing) < 3:
                existing.append(is_income)
            else:
                existing[2] = existing[2] or is_income
            if len(existing) < 4:
                existing.append(source_label)
            else:
                existing[3] = _merge_expense_sources(existing[3], source_label)
            totalCost += cost_value
        # else, add the expense name, it's cost and it's category
        else:
            expensesDict.update(
                {
                    expenseNameCell.value: [
                        cost_value,
                        expenseCategoryCell.value if expenseCategoryCell.value is not None else "",
                        is_income,
                        source_label,
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

    return expensesSorted, expensesDict


# @TODO add the expenses to the final excel file - go through each line in chat's response, for each line, check the name of the expense and it's cost, add the cost to a
# dict with the total cost for each category. then, add the category's cost to the final excel sheet


def fillCells(
    outputWorkbookPath,
    openAIOutput,
    month,
    expenses_dict=None,
    trust_provided_categories: bool = False,
):
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
    category_rows = template_category_rows()

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
    income_overrides_norm = {
        _normalize_income_name_key(k): v
        for k, v in _INCOME_NAME_CATEGORY_OVERRIDES.items()
    }
    investment_overrides_norm = {
        _normalize_income_name_key(k): v
        for k, v in _INVESTMENT_NAME_CATEGORY_OVERRIDES.items()
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
        if not trust_provided_categories:
            # Income/investment hard overrides apply to AI output only; reviewed choices win.
            normalized_income_key = _normalize_income_name_key(name)
            forced_income_category = income_overrides_norm.get(normalized_income_key)
            if forced_income_category:
                if forced_income_category not in allowed_categories:
                    forced_income_category = (
                        _DEFAULT_INCOME_FALLBACK_CATEGORY
                        if _DEFAULT_INCOME_FALLBACK_CATEGORY in allowed_categories
                        else _best_category_match(_DEFAULT_INCOME_FALLBACK_CATEGORY)
                    )
                category = forced_income_category

            forced_investment_category = investment_overrides_norm.get(normalized_income_key)
            if forced_investment_category:
                category = forced_investment_category

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

                source_label = _lookup_expense_source_from_dict(expenses_dict, val)
                comment_line = f"{val} - {add_amount} ({source_label})"
                if ws[f"{expenseCol}{categoryRow}"].comment is None:
                    ws[f"{expenseCol}{categoryRow}"].comment = Comment("", "Automated")
                comment = Comment(
                    comment_line,
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
