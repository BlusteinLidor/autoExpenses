from pathlib import Path
from typing import Dict, Any
from numbers import Number
from uuid import uuid4
from datetime import datetime, timezone

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from config import get_paths, get_state, update_state
from pipeline import run_for_month, prepare_for_month, finalize_for_month
from data import monthToExpenseColDict
from openpyxl import load_workbook
from handleExcel import (
    parse_ai_output_lines,
    get_allowed_categories,
    build_ai_output_from_items,
    is_card_statement_duplicate_expense,
)


app = FastAPI(title="AutoExpenses API")
_PENDING_REVIEWS: Dict[str, Dict[str, Any]] = {}
_PIPELINE_PROGRESS: Dict[str, Dict[str, Any]] = {}


BASE_DIR = Path(__file__).resolve().parent
FRONTEND_DIR = BASE_DIR / "frontend"
STATIC_DIR = FRONTEND_DIR / "static"

# Serve the web UI (index.html) + its static assets from /static/*
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


def _user_friendly_pipeline_error_message(error: Exception) -> str:
    """Map internal pipeline failures to short user-facing messages."""
    raw_message = str(error)
    lower_message = raw_message.lower()

    if "step 1 (download max)" in lower_message:
        return (
            "Failed to download the Max Excel file. "
            "Please check that the Max site is available and try again."
        )
    if "step 2 (leumi download/merge)" in lower_message:
        return (
            "Failed while downloading or merging Leumi data. "
            "Please check your Leumi access and try again."
        )
    if "step 3 (getexpenses/categorize)" in lower_message:
        return (
            "Failed while categorizing expenses. "
            "Please verify your API configuration and try again."
        )
    if "step 5 (fillcells)" in lower_message:
        return (
            "Failed while writing results to the workbook. "
            "Please ensure the file is not open and try again."
        )
    return "The monthly run failed. Please try again."


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception) -> JSONResponse:
    # Ensure the frontend always gets JSON, even for unexpected server errors.
    return JSONResponse(
        status_code=500,
        content={
            "detail": "Something went wrong on the server. Please try again.",
            "error_code": "UNEXPECTED_SERVER_ERROR",
        },
    )


@app.get("/")
def index() -> FileResponse:
    return FileResponse(str(FRONTEND_DIR / "index.html"))


@app.post("/run-month")
def run_month(payload: Dict[str, Any]) -> Dict[str, Any]:
    year = str(payload.get("year"))
    month = str(payload.get("month"))
    include_leumi = bool(payload.get("include_leumi", False))

    if not year.isdigit() or not month.isdigit():
        raise HTTPException(
            status_code=400,
            detail="Invalid year/month. Please choose valid numeric values.",
        )

    month_int = int(month)
    if month_int < 1 or month_int > 12:
        raise HTTPException(
            status_code=400,
            detail="Invalid month. Please choose a month between 1 and 12.",
        )

    try:
        src, out = run_for_month(year, month, include_leumi=include_leumi)
    except Exception as error:
        raise HTTPException(
            status_code=500,
            detail=_user_friendly_pipeline_error_message(error),
        ) from error

    return {
        "source_excel": str(src),
        "output_excel": str(out),
        "state": get_state(),
    }


@app.post("/run-month/prepare")
def run_month_prepare(payload: Dict[str, Any]) -> Dict[str, Any]:
    year = str(payload.get("year"))
    month = str(payload.get("month"))
    include_leumi = bool(payload.get("include_leumi", False))

    if not year.isdigit() or not month.isdigit():
        raise HTTPException(
            status_code=400,
            detail="Invalid year/month. Please choose valid numeric values.",
        )

    month_int = int(month)
    if month_int < 1 or month_int > 12:
        raise HTTPException(
            status_code=400,
            detail="Invalid month. Please choose a month between 1 and 12.",
        )

    run_token = str(payload.get("run_token", "")).strip() or str(uuid4())

    def _update_progress(step_text: str) -> None:
        _PIPELINE_PROGRESS[run_token] = {
            "step": step_text,
            "updated_at": datetime.now(timezone.utc).isoformat(),
            "done": False,
        }

    _update_progress("Starting monthly run...")

    try:
        src, out, ai_output = prepare_for_month(
            year=year,
            month=month,
            include_leumi=include_leumi,
            progress_callback=_update_progress,
        )
    except Exception as error:
        _PIPELINE_PROGRESS[run_token] = {
            "step": "Run failed.",
            "updated_at": datetime.now(timezone.utc).isoformat(),
            "done": True,
        }
        raise HTTPException(
            status_code=500,
            detail=_user_friendly_pipeline_error_message(error),
        ) from error

    parsed_items, parse_errors = parse_ai_output_lines(ai_output)
    allowed_categories = get_allowed_categories(str(out))
    review_items = []
    for item in parsed_items:
        review_items.append(
            {
                "name": item.get("name", ""),
                "cost": float(item.get("cost", 0.0)),
                "category": str(item.get("category", "")),
                "is_possible_duplicate": bool(
                    is_card_statement_duplicate_expense(str(item.get("name", "")))
                ),
            }
        )

    review_token = str(uuid4())
    _PENDING_REVIEWS[review_token] = {
        "year": year,
        "month": month,
        "output_workbook_path": str(out),
    }

    return {
        "run_token": run_token,
        "review_token": review_token,
        "source_excel": str(src),
        "output_excel": str(out),
        "allowed_categories": allowed_categories,
        "items": review_items,
        "parse_errors": parse_errors,
    }


@app.post("/run-month/finalize")
def run_month_finalize(payload: Dict[str, Any]) -> Dict[str, Any]:
    review_token = str(payload.get("review_token", "")).strip()
    reviewed_items = payload.get("reviewed_items", [])
    if not review_token:
        raise HTTPException(status_code=400, detail="Missing review token.")

    review_state = _PENDING_REVIEWS.get(review_token)
    if review_state is None:
        raise HTTPException(
            status_code=400,
            detail="Review session not found or expired. Please run prepare again.",
        )

    if not isinstance(reviewed_items, list):
        raise HTTPException(status_code=400, detail="reviewed_items must be a list.")

    categorized_output = build_ai_output_from_items(reviewed_items)
    run_token = str(payload.get("run_token", "")).strip()

    def _update_progress(step_text: str) -> None:
        if not run_token:
            return
        _PIPELINE_PROGRESS[run_token] = {
            "step": step_text,
            "updated_at": datetime.now(timezone.utc).isoformat(),
            "done": False,
        }

    if run_token:
        _update_progress("Step 5/5: Filling workbook...")

    try:
        final_path = finalize_for_month(
            year=review_state["year"],
            month=review_state["month"],
            output_workbook_path=Path(review_state["output_workbook_path"]),
            categorized_output=categorized_output,
            progress_callback=_update_progress if run_token else None,
        )
    except Exception as error:
        if run_token:
            _PIPELINE_PROGRESS[run_token] = {
                "step": "Run failed.",
                "updated_at": datetime.now(timezone.utc).isoformat(),
                "done": True,
            }
        raise HTTPException(
            status_code=500,
            detail=_user_friendly_pipeline_error_message(error),
        ) from error
    finally:
        _PENDING_REVIEWS.pop(review_token, None)

    if run_token:
        _PIPELINE_PROGRESS[run_token] = {
            "step": "Done.",
            "updated_at": datetime.now(timezone.utc).isoformat(),
            "done": True,
        }

    return {
        "output_excel": str(final_path),
        "state": get_state(),
    }


@app.get("/run-month/progress/{run_token}")
def run_month_progress(run_token: str) -> Dict[str, Any]:
    progress = _PIPELINE_PROGRESS.get(run_token)
    if progress is None:
        return {"step": "", "updated_at": None, "done": False}
    return progress


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
        raw_value = cell.value
        if raw_value is None or raw_value == "":
            continue

        amount: float | None = None
        if isinstance(raw_value, Number):
            amount = float(raw_value)
        elif isinstance(raw_value, str):
            # Some template rows in month columns may contain labels (e.g., month names).
            normalized = raw_value.replace(",", "").strip()
            try:
                amount = float(normalized)
            except ValueError:
                continue

        if amount is not None:
            summary[str(category)] = summary.get(str(category), 0.0) + amount

    return summary


