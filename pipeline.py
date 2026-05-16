from pathlib import Path
from typing import Any, Callable, Dict, Optional, Tuple
import traceback

from config import get_paths, get_state, update_state, load_env, validate_env
from getExcelFileFromMax import getExcelFile
from getExcelFileFromLeumi import getLeumiData
from handleExcel import getExpenses, fillCells
from merge_expenses import merge_max_and_leumi
from yearly_totals import sync_month_to_year_total


def _emit_progress(progress_callback: Optional[Callable[[str], None]], step_text: str) -> None:
    if progress_callback is None:
        return
    progress_callback(step_text)


def prepare_for_month(
    year: str,
    month: str,
    include_leumi: bool = False,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> Tuple[Path, Path, str, Dict[str, Any], list[str]]:
    """
    End-to-end pipeline for a given year/month:
    - Download Max Excel (and optionally Leumi when include_leumi=True)
    - Merge both into a single workbook when include_leumi=True
    - Run getExpenses on the merged or Max-only file
    - Use OpenAI (+ RAG) to categorize
    - Fill cells in the current output workbook
    Returns (source_excel_path, output_workbook_path).
    """
    load_env()
    validate_env(require_max=True, require_leumi=include_leumi)

    paths = get_paths()
    month_padded = str(month).zfill(2)
    year_dir = paths.data_dir / str(year)
    year_dir.mkdir(parents=True, exist_ok=True)

    # 1. Download from Max; function already appends foreign exchange rows.
    try:
        _emit_progress(progress_callback, "Step 1/5: Getting data from Max...")
        print("[pipeline] Step 1: Download Max Excel")
        source_excel_path = Path(getExcelFile(year, month))
    except Exception as e:
        raise RuntimeError(f"Pipeline failed at step 1 (download Max): {e}") from e

    if include_leumi:
        try:
            _emit_progress(
                progress_callback, "Step 2/5: Getting data from Leumi and merging..."
            )
            print("[pipeline] Step 2: Download Leumi + merge")
            leumi_trans_path, leumi_cards_path = getLeumiData(year, month)
            combined_path = year_dir / f"combined_expenses_{year}_{month}.xlsx"
            merge_max_and_leumi(
                source_excel_path,
                Path(leumi_trans_path),
                Path(leumi_cards_path),
                combined_path,
            )
            source_excel_path = combined_path
        except Exception as e:
            print("[pipeline][error] Step 2 failed: Download Leumi + merge")
            print(f"[pipeline][error] {type(e).__name__}: {e}")
            print(traceback.format_exc())
            raise RuntimeError(f"Pipeline failed at step 2 (Leumi download/merge): {e}") from e

    # 3. Categorize expenses from the (merged or Max-only) file
    try:
        _emit_progress(progress_callback, "Step 3/5: Categorizing expenses with AI...")
        print("[pipeline] Step 3: getExpenses (read + AI categorize)")
        sorted_expenses, source_expenses_dict = getExpenses(str(source_excel_path))
    except Exception as e:
        raise RuntimeError(f"Pipeline failed at step 3 (getExpenses/categorize): {e}") from e

    # 4. Ensure we have a current output workbook to write into
    _emit_progress(progress_callback, "Step 4/5: Preparing output workbook...")
    print("[pipeline] Step 4: Prepare output workbook")
    output_workbook_path = year_dir / f"expenses_output_{year}_{month_padded}.xlsx"
    if not output_workbook_path.exists():
        if not paths.template_expenses.exists():
            raise FileNotFoundError(
                f"Template workbook not found at {paths.template_expenses}. "
                "Place your empty standard template there."
            )
        output_workbook_path.write_bytes(paths.template_expenses.read_bytes())

    capture_warnings: list[str] = []
    state = get_state()
    capture_report = state.get("last_max_capture")
    if isinstance(capture_report, dict):
        immediate = capture_report.get("immediate") or {}
        foreign = capture_report.get("foreign") or {}
        immediate_parsed = int(immediate.get("parsed_count") or 0)
        immediate_visible = int(immediate.get("row_count") or 0)
        if immediate.get("error"):
            capture_warnings.append(
                "Could not scrape immediate-charge table from Max: "
                + str(immediate.get("error"))
            )
        elif not immediate.get("absent"):
            if immediate_visible > 0 and immediate_parsed < immediate_visible:
                capture_warnings.append(
                    f"Only {immediate_parsed} of {immediate_visible} visible immediate "
                    "transaction(s) were parsed. Some may be missing."
                )
            elif immediate_visible == 0 and immediate_parsed == 0:
                capture_warnings.append(
                    "No immediate-charge transactions were found on the Max page "
                    "for this month."
                )
        if foreign.get("error"):
            capture_warnings.append(
                "Could not scrape foreign-exchange table from Max: "
                + str(foreign.get("error"))
            )

    return (
        source_excel_path,
        output_workbook_path,
        sorted_expenses,
        source_expenses_dict,
        capture_warnings,
    )


def finalize_for_month(
    year: str,
    month: str,
    output_workbook_path: Path,
    categorized_output: str,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> Tuple[Path, Path]:
    """Fill workbook, sync yearly total workbook and update state."""
    try:
        _emit_progress(progress_callback, "Step 5/5: Filling workbook...")
        print(f"[pipeline] Step 5: fillCells into {output_workbook_path}")
        fillCells(str(output_workbook_path), categorized_output, month=month)
        total_workbook_path = sync_month_to_year_total(year=year, month=month)
        print(
            f"[pipeline] Yearly total synced for {year}-{str(month).zfill(2)} "
            f"into {total_workbook_path}"
        )
    except Exception as e:
        raise RuntimeError(f"Pipeline failed at step 5 (fillCells): {e}") from e

    update_state({"last_filled_year": year, "last_filled_month": month})
    return output_workbook_path, total_workbook_path


def run_for_month(
    year: str,
    month: str,
    include_leumi: bool = False,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> Tuple[Path, Path]:
    source_excel_path, output_workbook_path, sorted_expenses, _, _ = prepare_for_month(
        year=year,
        month=month,
        include_leumi=include_leumi,
        progress_callback=progress_callback,
    )
    final_monthly, _ = finalize_for_month(
        year=year,
        month=month,
        output_workbook_path=output_workbook_path,
        categorized_output=sorted_expenses,
        progress_callback=progress_callback,
    )
    return source_excel_path, final_monthly
