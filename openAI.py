import os
from typing import Dict, Any, List

from openpyxl import load_workbook

from openai import OpenAI

from config import get_paths, load_env, template_category_rows, validate_env
from data import fullDict
from rag_categories import get_examples_for_prompt


def _load_allowed_subcategories_from_template() -> List[str]:
    """
    Read expense sub-category labels from the template workbook (column B, category rows only).
    This keeps OpenAI output constrained to the exact labels that will later be matched in Excel.
    """
    paths = get_paths()
    if not paths.template_expenses.exists():
        return []

    wb = load_workbook(str(paths.template_expenses), data_only=True)
    ws = wb.active
    labels: List[str] = []
    seen = set()
    for r in template_category_rows():
        v = ws[f"B{r}"].value
        if v is None:
            continue
        s = str(v).strip()
        if not s or s in seen:
            continue
        seen.add(s)
        labels.append(s)
    wb.close()
    return labels


def _build_system_prompt() -> str:
    allowed_subcategories = _load_allowed_subcategories_from_template()
    allowed_subcategories_text = (
        "\n".join(f"- {s}" for s in allowed_subcategories) if allowed_subcategories else "(לא נמצא קובץ תבנית - אין רשימת תתי-קטגוריות קשיחה)"
    )

    content_system = """
    אני רוצה לסווג את ההוצאות שלי לפי תת-קטגוריות. אני אוסיף רשימה, שהערכים הראשונים שלה יהיו הקטגוריות הראשיות, והערכים השניים שלה יהיו תת-קטגוריות.
    לאחר מכן, אוסיף רשימה נוספת שהערכים הראשונים שלה יהיו השמות של ההוצאות, והערכים הבאים שלו יהיו זוגות, שהערך הראשון בזוג הוא סכום ההוצאה, והערך השני בזוג הוא הקטגוריה הראשית של ההוצאה.
    'אני רוצה שתחזיר לי רשימה שתראה כך: 'שם ההוצאה - סכום ההוצאה - תת-הקטגוריה המתאימה לפי סיווג.
    תחזיר בבקשה רק את הרשימה, ללא מלל נוסף. בנוסף, אם באחד שמות ההוצאות יש מקף '-', אנא החלף אותו ב '~'.
    אצרף כאן את הרשימה הראשונה - הרשימה של הקטגוריות לתת-קטגוריות :
    """ + str(fullDict)

    content_system = (
        """
    אני רוצה לסווג את ההוצאות שלי לפי תת-קטגוריות. אני אוסיף רשימה, שהערכים הראשונים שלה יהיו הקטגוריות הראשיות, והערכים השניים שלה יהיו תת-קטגוריות.
    לאחר מכן, אוסיף רשימה נוספת שהערכים הראשונים שלה יהיו השמות של ההוצאות, והערכים הבאים שלו יהיו זוגות, שהערך הראשון בזוג הוא סכום ההוצאה, והערך השני בזוג הוא הקטגוריה הראשית של ההוצאה.
    'אני רוצה שתחזיר לי רשימה שתראה כך: 'שם ההוצאה - סכום ההוצאה - תת-הקטגוריה המתאימה לפי סיווג.
    תחזיר בבקשה רק את הרשימה, ללא מלל נוסף. בנוסף, אם באחד שמות ההוצאות יש מקף '-', אנא החלף אותו ב '~'.
    דבר נוסף מאוד חשוב - וודא שאין שגיאות כתיב בשמות ההוצאות, תת הקטגוריות והקטגוריות. כמו כן, וודא שכל הוצאה מקוטלגת בקטגוריה המתאימה לה, כמה שיותר אחוז התאמה. אין גישה לאינטרנט—הסתמך רק על המידע שסופק בהודעה זו.
    אצרף כאן את הרשימה הראשונה - הרשימה של הקטגוריות לתת-קטגוריות : """
        + str(fullDict)
        + """
    דרישה קריטית לפלט:
    - השדה השלישי (תת-הקטגוריה) חייב להיות EXACTLY אחד מהרשימה הבאה (כפי שהיא מופיעה, כולל סימני פיסוק).
    - אסור להחזיר קבוצה של קטגוריות ראשיות (כמו למשל "מזון וצריכה", "שונות" וכו') - תמיד להחזיר תת-קטגוריה בלבד.
    - תת-הקטגוריה חייבת להיות אחת מאלה:
    """
        + allowed_subcategories_text
        + """
    כמו כן, מצרף דוגמאות:
קטגוריות לתת-קטגוריות (זו הרשימה שהבאתי לך, ממנה אתה לוקח את החלוקה של הקטגוריות):
"אופנה": ("ביגוד, תכשיטים וקוסמטיקה")
"דלק, חשמל וגז": ("חשבון חשמל", "חשבון גז", "חשבון תמי 4", "דלק לרכב"),

רשימת הוצאות (זו הרשימה בה יש שם הוצאה, הסכום, והתת קטגוריה שלה):
'שרות בוש/סימנס/קונסטרוקטה': [509.31000000000006, 'חשמל ומחשבים']
'ליבי': [558, 'מזון וצריכה']

פלט נדרש:
שרות בוש/סימנס/קונסטרוקטה - 509.31000000000006 - גאדג'טים ואלקטרוניקה'
ליבי - 558 - אוכל בחוץ (כולל מסעדות, בתי קפה, משלוחים)

דגשים נוספים:
- יש להחזיר שורה אחת לכל הוצאה מהרשימה הנתונה (לא לדלג על אף הוצאה).
- הפרדת השדות חייבת להיות בדיוק עם ' - ' (רווח-מקף-רווח).
- לגבי שם ההוצאה: להעתיק את השם בדיוק כפי שהוא מופיע אצלך (המפתח ברשימה), ולהחליף רק מקף '-' ב-'~'. אסור לשנות תווים אחרים (כולל גרשיים/ציטוטים/תווים מיוחדים).
"""
    )
    return content_system


def sortExpensesAI(expensesDict: Dict[str, Any], batch_size: int = 25) -> str:
    # Load environment and validate OpenAI API key
    load_env()
    validate_env(require_max=False, require_leumi=False)

    api_key = os.environ.get("OPENAI_API_KEY")
    client = OpenAI(api_key=api_key)

    items = list(expensesDict.items())
    if not items:
        return ""

    outputs: list[str] = []
    for start in range(0, len(items), batch_size):
        batch_items = items[start : start + batch_size]
        batch_dict = dict(batch_items)
        expected_lines = len(batch_dict)

        # Optionally augment the user message with RAG examples for this batch
        examples_text = get_examples_for_prompt(batch_dict)
        user_content = "רשימת ההוצאות: " + str(batch_dict)
        if examples_text:
            user_content += examples_text
        user_content += (
            "\n\nהערה: אם שם הוצאה מסתיים ב-' ##N' (מספר), זה רק מזהה ייחודי לשורה כפולה — "
            "החזר את שם ההוצאה ללא הסיומת ##N, ושמור על שורה נפרדת לכל פריט."
        )

        last_output = ""

        def _count_item_lines(s: str) -> int:
            # Only count lines that look like: name - amount - category
            # The exact separators must be ' - ' (spaces around dash).
            lines = [l.strip() for l in s.splitlines() if l.strip()]
            return sum(1 for l in lines if l.count(" - ") == 2)

        for attempt in range(2):
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                temperature=0,
                max_tokens=2000,
                messages=[
                    {"role": "system", "content": _build_system_prompt()},
                    {"role": "user", "content": user_content},
                ],
            )
            output = response.choices[0].message.content or ""
            output_norm = output.replace("\r\n", "\n").replace("\r", "\n").strip()
            last_output = output_norm

            item_lines_count = _count_item_lines(output_norm)
            if item_lines_count == expected_lines:
                break

            # One retry with stricter instruction; avoids silent omission.
            user_content_retry = (
                user_content
                + f"\n\nהערה קריטית: החזרת {item_lines_count} שורות בפורמט 'שם - סכום - תת-קטגוריה' אבל נדרשות {expected_lines}."
                + " החזר בדיוק שורה אחת לכל הוצאה מהרשימה. אל תדלג על אף הוצאה."
                + " החזר רק את הרשימה בפורמט זה, ללא שום מלל נוסף."
            )
            user_content = user_content_retry

        if last_output:
            outputs.append(last_output)

    return "\n".join([o for o in outputs if o]).strip()

