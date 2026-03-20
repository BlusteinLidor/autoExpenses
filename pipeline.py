from pathlib import Path
from typing import Tuple

from config import get_paths, update_state, load_env, validate_env
from getExcelFileFromMax import getExcelFile
from getExcelFileFromLeumi import getLeumiData
from handleExcel import getExpenses, fillCells
from merge_expenses import merge_max_and_leumi


def prepare_for_month(
    year: str, month: str, include_leumi: bool = False
) -> Tuple[Path, Path, str]:
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

    # 1. Download from Max; function already appends foreign exchange rows.
    try:
        print("[pipeline] Step 1: Download Max Excel")
        source_excel_path = Path(getExcelFile(year, month))
    except Exception as e:
        raise RuntimeError(f"Pipeline failed at step 1 (download Max): {e}") from e

    if include_leumi:
        try:
            print("[pipeline] Step 2: Download Leumi + merge")
            leumi_trans_path, leumi_cards_path = getLeumiData(year, month)
            combined_path = paths.data_dir / f"combined_expenses_{year}_{month}.xlsx"
            merge_max_and_leumi(
                source_excel_path,
                Path(leumi_trans_path),
                Path(leumi_cards_path),
                combined_path,
            )
            source_excel_path = combined_path
        except Exception as e:
            raise RuntimeError(f"Pipeline failed at step 2 (Leumi download/merge): {e}") from e

    # 3. Categorize expenses from the (merged or Max-only) file
    try:
        print("[pipeline] Step 3: getExpenses (read + AI categorize)")
        sorted_expenses = getExpenses(str(source_excel_path))
    except Exception as e:
        raise RuntimeError(f"Pipeline failed at step 3 (getExpenses/categorize): {e}") from e

    # 4. Ensure we have a current output workbook to write into
    print("[pipeline] Step 4: Prepare output workbook")
    output_workbook_path = paths.data_dir / f"expenses_output_{year}_{month_padded}.xlsx"
    if not output_workbook_path.exists():
        if not paths.template_expenses.exists():
            raise FileNotFoundError(
                f"Template workbook not found at {paths.template_expenses}. "
                "Place your empty standard template there."
            )
        output_workbook_path.write_bytes(paths.template_expenses.read_bytes())

    return source_excel_path, output_workbook_path, sorted_expenses


def finalize_for_month(
    year: str, month: str, output_workbook_path: Path, categorized_output: str
) -> Path:
    """Fill workbook with reviewed categories and update state."""
    try:
        print(f"[pipeline] Step 5: fillCells into {output_workbook_path}")
        fillCells(str(output_workbook_path), categorized_output, month=month)
    except Exception as e:
        raise RuntimeError(f"Pipeline failed at step 5 (fillCells): {e}") from e

    update_state({"last_filled_year": year, "last_filled_month": month})
    return output_workbook_path


def run_for_month(
    year: str, month: str, include_leumi: bool = False
) -> Tuple[Path, Path]:
    source_excel_path, output_workbook_path, sorted_expenses = prepare_for_month(
        year=year,
        month=month,
        include_leumi=include_leumi,
    )
    finalize_for_month(
        year=year,
        month=month,
        output_workbook_path=output_workbook_path,
        categorized_output=sorted_expenses,
    )
    return source_excel_path, output_workbook_path
