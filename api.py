from pathlib import Path
from typing import Dict, Any

from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from config import get_paths, get_state, update_state
from pipeline import run_for_month
from data import monthToExpenseColDict
from openpyxl import load_workbook


app = FastAPI(title="AutoExpenses API")


BASE_DIR = Path(__file__).resolve().parent
FRONTEND_DIR = BASE_DIR / "frontend"
STATIC_DIR = FRONTEND_DIR / "static"

# Serve the web UI (index.html) + its static assets from /static/*
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/")
def index() -> FileResponse:
    return FileResponse(str(FRONTEND_DIR / "index.html"))


@app.post("/run-month")
def run_month(payload: Dict[str, Any]) -> Dict[str, Any]:
    year = str(payload.get("year"))
    month = str(payload.get("month"))
    include_leumi = bool(payload.get("include_leumi", False))

    src, out = run_for_month(year, month, include_leumi=include_leumi)
    return {
        "source_excel": str(src),
        "output_excel": str(out),
        "state": get_state(),
    }


@app.get("/state")
def read_state() -> Dict[str, Any]:
    return get_state()


@app.post("/assets")
def write_assets(payload: Dict[str, Any]) -> Dict[str, Any]:
    cars = payload.get("cars", [])
    investments = payload.get("investments", [])
    update_state({"cars": cars, "investments": investments})
    return get_state()


@app.get("/assets")
def read_assets() -> Dict[str, Any]:
    state = get_state()
    return {
        "cars": state.get("cars", []),
        "investments": state.get("investments", []),
        "leumi_balance": state.get("leumi_balance"),
    }


@app.get("/expenses/summary")
def expenses_summary(year: int, month: int) -> Dict[str, float]:
    """
    Summarize expenses by category for a given year/month.
    """
    paths = get_paths()

    month_padded = str(month).zfill(2)
    # pipeline writes per-month workbooks like: expenses_output_{year}_{MM}.xlsx
    candidate_path = paths.data_dir / f"expenses_output_{year}_{month_padded}.xlsx"
    if not candidate_path.exists():
        # Fallback for older flows / manual runs.
        candidate_path = (
            paths.current_expenses if paths.current_expenses.exists() else None
        )
    if not candidate_path:
        return {}

    wb = load_workbook(str(candidate_path), data_only=True)
    ws = wb.active

    if month not in monthToExpenseColDict:
        return {}
    col_letter = monthToExpenseColDict[month]

    first_row = 17
    last_row = 105
    category_col = "B"

    summary: Dict[str, float] = {}
    for row in range(first_row, last_row):
        category = ws[f"{category_col}{row}"].value
        if not category:
            continue
        cell = ws[f"{col_letter}{row}"]
        if cell.value:
            summary[str(category)] = summary.get(str(category), 0.0) + float(
                cell.value
            )

    return summary


