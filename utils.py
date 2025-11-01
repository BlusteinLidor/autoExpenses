from data import END_MONTH, START_MONTH, START_YEAR, defaultMonth, defaultYear
from datetime import datetime

def checkDate(year, month):
    """Check if the given year and month are within the allowed range."""
    current_year = datetime.now().year
    if not (START_YEAR <= int(year) <= current_year):
        year = defaultYear
    if not (START_MONTH <= int(month) <= END_MONTH):
        month = defaultMonth
    return year, month