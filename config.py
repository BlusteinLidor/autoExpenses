import os
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, Dict, Any

from dotenv import load_dotenv

# Column B rows in template_expenses.xlsx that hold category labels.
# Excludes non-category rows: 13-16, 45-48, 91-94, and 105+.
TEMPLATE_INCOME_CATEGORY_ROW_RANGES: tuple[tuple[int, int], ...] = ((6, 13),)
TEMPLATE_NECESSARY_EXPENSE_CATEGORY_ROW_RANGES: tuple[tuple[int, int], ...] = ((17, 45),)
TEMPLATE_LUXURY_EXPENSE_CATEGORY_ROW_RANGES: tuple[tuple[int, int], ...] = ((49, 91),)
TEMPLATE_EXPENSE_CATEGORY_ROW_RANGES: tuple[tuple[int, int], ...] = (
    *TEMPLATE_NECESSARY_EXPENSE_CATEGORY_ROW_RANGES,
    *TEMPLATE_LUXURY_EXPENSE_CATEGORY_ROW_RANGES,
)
TEMPLATE_INVESTMENT_CATEGORY_ROW_RANGES: tuple[tuple[int, int], ...] = ((95, 105),)
TEMPLATE_CATEGORY_ROW_RANGES: tuple[tuple[int, int], ...] = (
    *TEMPLATE_INCOME_CATEGORY_ROW_RANGES,
    *TEMPLATE_EXPENSE_CATEGORY_ROW_RANGES,
    *TEMPLATE_INVESTMENT_CATEGORY_ROW_RANGES,
)


def template_category_rows() -> list[int]:
    rows: list[int] = []
    for start, end_exclusive in TEMPLATE_CATEGORY_ROW_RANGES:
        rows.extend(range(start, end_exclusive))
    return rows


@dataclass
class Paths:
    base_dir: Path
    data_dir: Path
    max_exports_dir: Path
    leumi_exports_dir: Path
    template_expenses: Path
    current_expenses: Path
    state_file: Path
    assets_file: Path
    category_corrections_file: Path
    drive_state_file: Path


def get_paths() -> Paths:
    base_dir = Path(__file__).resolve().parent
    data_dir = base_dir / "data"
    max_exports_dir = data_dir / "max_exports"
    leumi_exports_dir = data_dir / "leumi_exports"
    template_expenses = data_dir / "template_expenses.xlsx"
    current_expenses = data_dir / "expenses_output.xlsx"
    state_file = data_dir / "state.json"
    assets_file = data_dir / "assets.json"
    category_corrections_file = data_dir / "category_corrections.json"
    drive_state_file = data_dir / "drive_state.json"

    # Ensure required directories exist
    for d in (data_dir, max_exports_dir, leumi_exports_dir):
        d.mkdir(parents=True, exist_ok=True)

    return Paths(
        base_dir=base_dir,
        data_dir=data_dir,
        max_exports_dir=max_exports_dir,
        leumi_exports_dir=leumi_exports_dir,
        template_expenses=template_expenses,
        current_expenses=current_expenses,
        state_file=state_file,
        assets_file=assets_file,
        category_corrections_file=category_corrections_file,
        drive_state_file=drive_state_file,
    )


def load_env() -> None:
    # Load .env from project root if present
    base_dir = Path(__file__).resolve().parent
    env_path = base_dir / ".env"
    load_dotenv(env_path, override=True)


def validate_env(require_max: bool = True, require_leumi: bool = False) -> None:
    """
    Validate that required environment variables exist.
    Raise RuntimeError with a clear message if something is missing.
    """
    missing = []

    if require_max:
        for key in ("MAX_USERNAME", "MAX_PASSWORD"):
            if not os.environ.get(key):
                missing.append(key)

    # OpenAI is required whenever categorization runs
    if not os.environ.get("OPENAI_API_KEY"):
        missing.append("OPENAI_API_KEY")

    if require_leumi:
        for key in ("LEUMI_USERNAME", "LEUMI_PASSWORD"):
            if not os.environ.get(key):
                missing.append(key)

    if missing:
        raise RuntimeError(
            "Missing required environment variables: " + ", ".join(sorted(missing))
        )


def read_json(path: Path) -> Optional[Dict[str, Any]]:
    if not path.exists():
        return None
    import json

    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def write_json(path: Path, data: Dict[str, Any]) -> None:
    import json

    if not path.parent.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_state() -> Dict[str, Any]:
    paths = get_paths()
    state = read_json(paths.state_file) or {}
    return state


def update_state(updates: Dict[str, Any]) -> Dict[str, Any]:
    paths = get_paths()
    state = get_state()
    state.update(updates)
    write_json(paths.state_file, state)
    return state
