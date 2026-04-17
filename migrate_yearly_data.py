import argparse
import re
from pathlib import Path

from config import get_paths


YEAR_PATTERNS = [
    re.compile(r"^expenses_output_(?P<year>\d{4})_(?P<month>\d{2})\.xlsx$"),
    re.compile(r"^expenses_output_(?P<year>\d{4})_total\.xlsx$"),
    re.compile(r"^combined_expenses_(?P<year>\d{4})_(?P<month>\d{1,2})\.xlsx$"),
]


def _extract_year(file_name: str) -> str | None:
    for pattern in YEAR_PATTERNS:
        match = pattern.match(file_name)
        if match:
            return match.group("year")
    return None


def migrate_data_files(*, dry_run: bool = False) -> int:
    """
    One-time migration:
    Move year-tagged Excel files from data/ into data/<year>/.
    """
    paths = get_paths()
    data_dir: Path = paths.data_dir
    moved_count = 0

    for file_path in sorted(data_dir.glob("*.xlsx")):
        year = _extract_year(file_path.name)
        if year is None:
            continue

        target_dir = data_dir / year
        target_path = target_dir / file_path.name

        if target_path.exists():
            print(f"[skip] Destination already exists: {target_path}")
            continue

        print(f"[move] {file_path} -> {target_path}")
        if not dry_run:
            target_dir.mkdir(parents=True, exist_ok=True)
            file_path.replace(target_path)
            moved_count += 1

    if dry_run:
        print("[done] Dry run complete (no files were moved).")
    else:
        print(f"[done] Migration complete. Moved {moved_count} file(s).")
    return moved_count


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Move year-tagged Excel files from data/ into data/<year>/."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be moved without making changes.",
    )
    args = parser.parse_args()
    migrate_data_files(dry_run=args.dry_run)


if __name__ == "__main__":
    main()

