import os
import subprocess
from difflib import SequenceMatcher
from datetime import datetime
import re
from tkinter import *
from tkinter import ttk, messagebox
from getExcelFileFromMax import toggleHeadless
from pipeline import prepare_for_month, finalize_for_month
from config import get_paths
from handleExcel import (
    parse_ai_output_lines,
    get_allowed_categories,
    build_ai_output_from_items,
    is_card_statement_duplicate_expense,
)


def runPipelineThreaded():
    # Keep UI operations in the main thread to avoid Tkinter thread-safety issues.
    runPipeline()


def runPipeline():
    include_leumi = bool(includeLeumiVar.get())
    loading("Loading... Running pipeline (download + categorize). Please wait...")
    try:
        year = yearComboBox.get()
        month = monthComboBox.get()

        def _progress(step_text: str):
            loading(step_text)

        _, output_path, ai_output, _, _ = prepare_for_month(
            year=year,
            month=month,
            include_leumi=include_leumi,
            progress_callback=_progress,
        )
        show_corrections_dialog(
            year=year,
            month=month,
            output_path=output_path,
            ai_output=ai_output,
        )
    except Exception as e:
        messagebox.showerror("Failure", f"Error while running pipeline:\n{e}")
    finally:
        loading("")


def loading(message):
    global loadingText
    loadingText.set(message)
    window.update_idletasks()  # This forces the UI to refresh immediately


def _closest_category(raw_category: str, allowed_categories: list[str]) -> str:
    if not allowed_categories:
        return raw_category
    raw = str(raw_category or "").strip()
    if not raw:
        return allowed_categories[0]
    for cat in allowed_categories:
        if cat.strip() == raw:
            return cat
    best = allowed_categories[0]
    best_score = 0.0
    for cat in allowed_categories:
        score = SequenceMatcher(None, raw, cat.strip()).ratio()
        if score > best_score:
            best_score = score
            best = cat
    return best


def _create_category_combobox(
    parent,
    text_var: StringVar,
    allowed_categories: list[str],
    width: int = 45,
):
    """
    Create a category combobox that allows typing with prefix autocomplete,
    while keeping choices constrained to the provided categories.
    """
    combo = ttk.Combobox(
        parent,
        values=allowed_categories,
        textvariable=text_var,
        width=width,
        state="normal",
    )

    def _apply_canonical_case_if_exact():
        typed = str(text_var.get() or "").strip()
        for cat in allowed_categories:
            if typed == cat:
                return
            if typed.casefold() == cat.casefold():
                text_var.set(cat)
                return

    def _autocomplete_on_keyrelease(event):
        typed_raw = str(text_var.get() or "")
        typed = typed_raw.strip()
        if not typed:
            combo.configure(values=allowed_categories)
            return

        # Narrow dropdown options while typing.
        narrowed = [
            cat for cat in allowed_categories if typed.casefold() in cat.casefold()
        ]
        combo.configure(values=narrowed or allowed_categories)

        # Prefix autocomplete to the first match.
        prefix_matches = [
            cat for cat in allowed_categories if cat.casefold().startswith(typed.casefold())
        ]
        if prefix_matches:
            best = prefix_matches[0]
            if best.casefold() != typed.casefold():
                text_var.set(best)
                combo.icursor(len(typed))
                combo.selection_range(len(typed), END)
            else:
                _apply_canonical_case_if_exact()

    def _restore_options_on_focus_in(_event):
        combo.configure(values=allowed_categories)

    combo.bind("<KeyRelease>", _autocomplete_on_keyrelease)
    combo.bind("<FocusIn>", _restore_options_on_focus_in)
    combo.bind("<<ComboboxSelected>>", lambda _e: _apply_canonical_case_if_exact())
    return combo


def show_corrections_dialog(year: str, month: str, output_path, ai_output: str):
    parsed_items, parse_errors = parse_ai_output_lines(ai_output)
    allowed_categories = get_allowed_categories(str(output_path))

    if not parsed_items:
        err = "\n".join(parse_errors[:10]) if parse_errors else "No items parsed."
        messagebox.showerror("Failure", f"Could not parse categorized output:\n{err}")
        return

    dialog = Toplevel(window)
    dialog.title("Review and Correct Categories")
    dialog.geometry("950x650")
    dialog.transient(window)
    dialog.grab_set()

    header = Label(
        dialog,
        text="Review categories before filling the workbook",
        font=("Arial", 11, "bold"),
    )
    header.pack(pady=(10, 6))

    if parse_errors:
        Label(
            dialog,
            text=f"Skipped {len(parse_errors)} malformed line(s). Check data/errors.txt if needed.",
            foreground="orange",
            wraplength=900,
            justify=LEFT,
        ).pack(pady=(0, 6))

    container = Frame(dialog)
    container.pack(fill=BOTH, expand=True, padx=8, pady=8)

    canvas = Canvas(container)
    scrollbar = ttk.Scrollbar(container, orient=VERTICAL, command=canvas.yview)
    scrollable = Frame(canvas)
    scrollable.bind(
        "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=scrollable, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side=LEFT, fill=BOTH, expand=True)
    scrollbar.pack(side=RIGHT, fill=Y)

    regular_items = []
    duplicate_items = []
    for item in parsed_items:
        if is_card_statement_duplicate_expense(item["name"]):
            duplicate_items.append(item)
        else:
            regular_items.append(item)

    Label(scrollable, text="Expense Name", font=("Arial", 10, "bold")).grid(
        row=0, column=0, sticky=W, padx=6, pady=4
    )
    Label(scrollable, text="Amount", font=("Arial", 10, "bold")).grid(
        row=0, column=1, sticky=W, padx=6, pady=4
    )
    Label(scrollable, text="Category", font=("Arial", 10, "bold")).grid(
        row=0, column=2, sticky=W, padx=6, pady=4
    )
    Label(scrollable, text="Include?", font=("Arial", 10, "bold")).grid(
        row=0, column=3, sticky=W, padx=6, pady=4
    )

    SPLITTABLE_EXPENSE_NAME = "העברה דיגיטל"

    def _normalize_ws(s: str) -> str:
        return re.sub(
            r"\s+", " ", str(s or "").replace("\u200f", "").replace("\u200e", "")
        ).strip()

    def _is_split_expense(name: str) -> bool:
        return _normalize_ws(name) == SPLITTABLE_EXPENSE_NAME

    # For each regular item: either None (normal row) or split config (for UI splitting).
    split_ui_data = [None] * len(regular_items)

    category_vars = []
    include_vars = []

    for i, item in enumerate(regular_items, start=1):
        idx = i - 1

        if _is_split_expense(item["name"]):
            add_var = BooleanVar(value=True)
            include_vars.append(add_var)
            # Dummy var for alignment; ignored for split items.
            category_vars.append(StringVar(value=""))

            frame = Frame(scrollable, relief="groove", borderwidth=1)
            frame.grid(row=i, column=0, columnspan=4, sticky="we", padx=6, pady=4)

            Label(frame, text=item["name"], font=("Arial", 10, "bold")).grid(
                row=0, column=0, sticky=W, padx=6, pady=3, columnspan=2
            )
            Label(frame, text=str(item["cost"])).grid(
                row=0, column=2, sticky=W, padx=6, pady=3
            )
            ttk.Checkbutton(frame, variable=add_var).grid(
                row=0, column=3, sticky=W, padx=6, pady=3
            )

            default_category = _closest_category(item["category"], allowed_categories)
            part_cost_vars = [
                StringVar(value=str(item["cost"])),
                StringVar(value="0"),
                StringVar(value="0"),
            ]
            part_cat_vars = [
                StringVar(value=default_category),
                StringVar(value=default_category),
                StringVar(value=default_category),
            ]

            for part_no in (1, 2, 3):
                Label(frame, text=f"Part {part_no}").grid(
                    row=part_no, column=0, sticky=W, padx=10, pady=2
                )
                Entry(frame, textvariable=part_cost_vars[part_no - 1], width=12).grid(
                    row=part_no, column=1, sticky=W, padx=6, pady=2
                )
                combo = _create_category_combobox(
                    parent=frame,
                    text_var=part_cat_vars[part_no - 1],
                    allowed_categories=allowed_categories,
                    width=45,
                )
                combo.grid(row=part_no, column=2, sticky=W, padx=6, pady=2)

            split_ui_data[idx] = {
                "cost_vars": part_cost_vars,
                "cat_vars": part_cat_vars,
            }
        else:
            Label(
                scrollable, text=item["name"], anchor="w", justify=LEFT, wraplength=420
            ).grid(row=i, column=0, sticky=W, padx=6, pady=3)
            Label(scrollable, text=str(item["cost"])).grid(
                row=i, column=1, sticky=W, padx=6, pady=3
            )

            default_category = _closest_category(item["category"], allowed_categories)
            var = StringVar(value=default_category)
            combo = _create_category_combobox(
                parent=scrollable,
                text_var=var,
                allowed_categories=allowed_categories,
                width=45,
            )
            combo.grid(row=i, column=2, sticky=W, padx=6, pady=3)
            category_vars.append(var)

            add_var = BooleanVar(value=True)
            chk = ttk.Checkbutton(scrollable, variable=add_var)
            chk.grid(row=i, column=3, sticky=W, padx=6, pady=3)
            include_vars.append(add_var)

    duplicate_start_row = len(regular_items) + 2
    duplicate_category_vars = []
    duplicate_add_vars = []
    if duplicate_items:
        Label(
            scrollable,
            text="Possible duplicate card-statement charges (excluded by default)",
            font=("Arial", 10, "bold"),
            foreground="orange",
        ).grid(
            row=duplicate_start_row,
            column=0,
            columnspan=4,
            sticky=W,
            padx=6,
            pady=(16, 6),
        )
        Label(scrollable, text="Add?", font=("Arial", 10, "bold")).grid(
            row=duplicate_start_row + 1, column=3, sticky=W, padx=6, pady=4
        )
        for j, item in enumerate(duplicate_items, start=duplicate_start_row + 2):
            Label(
                scrollable,
                text=item["name"],
                anchor="w",
                justify=LEFT,
                wraplength=420,
                foreground="gray30",
            ).grid(row=j, column=0, sticky=W, padx=6, pady=3)
            Label(scrollable, text=str(item["cost"]), foreground="gray30").grid(
                row=j, column=1, sticky=W, padx=6, pady=3
            )

            default_category = _closest_category(item["category"], allowed_categories)
            var = StringVar(value=default_category)
            combo = _create_category_combobox(
                parent=scrollable,
                text_var=var,
                allowed_categories=allowed_categories,
                width=45,
            )
            combo.grid(row=j, column=2, sticky=W, padx=6, pady=3)
            duplicate_category_vars.append(var)

            add_var = BooleanVar(value=False)
            chk = ttk.Checkbutton(scrollable, variable=add_var)
            chk.grid(row=j, column=3, sticky=W, padx=6, pady=3)
            duplicate_add_vars.append(add_var)

    def apply_and_fill():
        reviewed_items = []
        excluded_regular = 0
        allowed_set = set(allowed_categories)

        def _parse_cost_entry(s: str) -> float:
            try:
                return abs(float(str(s or "").strip()))
            except ValueError:
                return 0.0

        for idx, item in enumerate(regular_items):
            if include_vars[idx].get():
                split_cfg = split_ui_data[idx]
                if split_cfg is not None:
                    cost_vars = split_cfg.get("cost_vars") or []
                    cat_vars = split_cfg.get("cat_vars") or []

                    any_added = False
                    for part_no in (1, 2, 3):
                        cv = (
                            cost_vars[part_no - 1]
                            if part_no - 1 < len(cost_vars)
                            else None
                        )
                        catv = (
                            cat_vars[part_no - 1]
                            if part_no - 1 < len(cat_vars)
                            else None
                        )
                        part_cost = _parse_cost_entry(
                            cv.get() if cv is not None else "0"
                        )
                        part_cat = str(catv.get() if catv is not None else "").strip()
                        if not part_cat and allowed_categories:
                            part_cat = allowed_categories[0]
                        if part_cat not in allowed_set:
                            messagebox.showerror(
                                "Failure",
                                f"Invalid category for '{item['name']}' part {part_no}: '{part_cat}'.\n"
                                "Please choose a category from the list.",
                            )
                            return

                        if part_cost > 0:
                            any_added = True
                            reviewed_items.append(
                                {
                                    "name": f"{item['name']} חלק {part_no}",
                                    "cost": part_cost,
                                    "category": part_cat,
                                }
                            )

                    if not any_added:
                        messagebox.showerror(
                            "Failure",
                            f"Please enter at least one positive split cost for '{item['name']}'.",
                        )
                        return
                else:
                    picked_category = str(category_vars[idx].get() or "").strip()
                    if picked_category not in allowed_set:
                        messagebox.showerror(
                            "Failure",
                            f"Invalid category for '{item['name']}': '{picked_category}'.\n"
                            "Please choose a category from the list.",
                        )
                        return
                    reviewed_items.append(
                        {
                            "name": item["name"],
                            "cost": item["cost"],
                            "category": picked_category,
                        }
                    )
            else:
                excluded_regular += 1

        included_duplicates = 0
        for idx, item in enumerate(duplicate_items):
            if duplicate_add_vars[idx].get():
                picked_category = str(duplicate_category_vars[idx].get() or "").strip()
                if picked_category not in allowed_set:
                    messagebox.showerror(
                        "Failure",
                        f"Invalid category for '{item['name']}': '{picked_category}'.\n"
                        "Please choose a category from the list.",
                    )
                    return
                included_duplicates += 1
                reviewed_items.append(
                    {
                        "name": item["name"],
                        "cost": item["cost"],
                        "category": picked_category,
                    }
                )

        reviewed_output = build_ai_output_from_items(reviewed_items)
        try:
            loading("Loading... Filling workbook with reviewed categories...")
            finalize_for_month(
                year=year,
                month=month,
                output_workbook_path=output_path,
                categorized_output=reviewed_output,
                progress_callback=loading,
            )
            dialog.destroy()
            messagebox.showinfo(
                "Success",
                f"The expenses are filled.\n"
                f"Auto-excluded duplicates: {len(duplicate_items) - included_duplicates}\n"
                f"Manually included duplicates: {included_duplicates}\n\n"
                f"Excluded regular items: {excluded_regular}\n\n"
                f"You can now open the output file at:\n{output_path}",
            )
        except Exception as e:
            messagebox.showerror(
                "Failure", f"Error while filling reviewed categories:\n{e}"
            )
        finally:
            loading("")

    bottom = Frame(dialog)
    bottom.pack(fill=X, padx=8, pady=8)
    ttk.Button(bottom, text="Apply and Fill", command=apply_and_fill).pack(
        side=RIGHT, padx=4
    )
    ttk.Button(bottom, text="Cancel", command=dialog.destroy).pack(side=RIGHT, padx=4)


def open_category_corrections_file():
    """Open data/category_corrections.json for editing; create with example if missing."""
    paths = get_paths()
    p = paths.category_corrections_file
    if not p.exists():
        p.write_text("{\n}\n", encoding="utf-8")
        messagebox.showinfo(
            "Category corrections",
            'Created empty file. Add lines like:\n"שם הוצאה": "תת-קטגוריה מדויקת"\n\nUse exact category names from your template (column B).',
        )
    try:
        if os.name == "nt":
            os.startfile(str(p))
        else:
            subprocess.run(["xdg-open", str(p)], check=False)
    except Exception as e:
        messagebox.showerror(
            "Error", f"Could not open file:\n{e}\n\nOpen manually: {p}"
        )


##################### UI ###################

# ui window
window = Tk()
window.title("Expense Sort Automation")
window.geometry("400x300")
window.resizable(False, False)  # Disable resizing
# @TODO add an icon to the window
menuBar = Menu(window)
settingsMenu = Menu(menuBar, tearoff=0)
# settingsMenu.add_command(label="Headless")
headlessCheckbuttonState = IntVar()
settingsMenu.add_checkbutton(
    label="Headless",
    command=lambda: toggleHeadless(headlessCheckbuttonState.get()),
    variable=headlessCheckbuttonState,
)
settingsMenu.add_command(
    label="Edit category corrections...", command=open_category_corrections_file
)
menuBar.add_cascade(label="Settings", menu=settingsMenu)
window.config(menu=menuBar)

# Styling
label_font = ("Arial", 10)
button_font = ("Arial", 10)

# months list
months = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
# year list
current_year = datetime.now().year
# last 3 years
years = [str(current_year - i) for i in range(3)]
# default year value
defaultYear = years[0]
# default month value
defaultMonth = months[0]

fileFrame = Frame(window)
fileFrame.pack(pady=10)
# Loading Label
loadingText = StringVar()
loadingLabel = Label(window, textvariable=loadingText, font=label_font)
loadingLabel.pack(pady=5)
# Year and Month Selection Frame
selectFrame = Frame(window)
selectFrame.pack(pady=10)

Label(selectFrame, text="Year:", font=label_font).grid(row=0, column=0, sticky=E)
yearComboBox = ttk.Combobox(selectFrame, values=years)
yearComboBox.set(defaultYear)
yearComboBox.grid(row=0, column=1, padx=5)

Label(selectFrame, text="Month:", font=label_font).grid(row=1, column=0, sticky=E)
monthComboBox = ttk.Combobox(selectFrame, values=months)
monthComboBox.set(defaultMonth)
monthComboBox.grid(row=1, column=1, padx=5)

# Run with Leumi checkbox
includeLeumiVar = IntVar(value=True)
includeLeumiCheck = ttk.Checkbutton(
    fileFrame,
    text="Run with Leumi (include Leumi export + merge)",
    variable=includeLeumiVar,
)
includeLeumiCheck.grid(row=1, column=0, columnspan=2, pady=5, sticky=W)

# Action Buttons Frame
actionFrame = Frame(window)
actionFrame.pack(pady=10)

global month
getExcelFileButton = ttk.Button(
    fileFrame,
    text="Run Pipeline",
    command=runPipelineThreaded,
)
getExcelFileButton.grid(row=0, column=0, padx=5)

closeButton = ttk.Button(actionFrame, text="Close", command=window.quit)
closeButton.grid(row=0, column=0, padx=5)

# Center everything on the window
window.update_idletasks()
x = (window.winfo_screenwidth() - window.winfo_reqwidth()) // 2
y = (window.winfo_screenheight() - window.winfo_reqheight()) // 2
window.geometry(f"+{x}+{y}")

# window loop
window.mainloop()
