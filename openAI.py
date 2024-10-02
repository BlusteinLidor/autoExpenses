import os
from dotenv import load_dotenv
from openai import OpenAI
from data import fullDict

def sortExpensesAI(expensesDict):

    content_system = """
    אני רוצה לסווג את ההוצאות שלי לפי תת-קטגוריות. אני אוסיף רשימה, שהערכים הראשונים שלה יהיו הקטגוריות הראשיות, והערכים השניים שלה יהיו תת-קטגוריות.
    לאחר מכן, אוסיף רשימה נוספת שהערכים הראשונים שלה יהיו השמות של ההוצאות, והערכים הבאים שלו יהיו זוגות, שהערך הראשון בזוג הוא סכום ההוצאה, והערך השני בזוג הוא הקטגוריה הראשית של ההוצאה.
    'אני רוצה שתחזיר לי רשימה שתראה כך: 'שם ההוצאה - סכום ההוצאה - תת-הקטגוריה המתאימה לפי סיווג.
    תחזיר בבקשה רק את הרשימה, ללא מלל נוסף. בנוסף, אם באחד שמות ההוצאות יש מקף '-', אנא החלף אותו ב '~'.
    אצרף כאן את הרשימה הראשונה - הרשימה של הקטגוריות לתת-קטגוריות :
    """ + str(fullDict)


    # Load environment variables from .env file
    load_dotenv()

    # Retrieve API key from the .env file
    client = OpenAI(
        api_key=os.environ.get("OPENAI_API_KEY")
    )

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": content_system},
            {"role": "user", "content": "רשימת ההוצאות: " + str(expensesDict)},
        ]
    )
    output = response.choices[0].message.content
    return output



