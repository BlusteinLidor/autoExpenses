from tkinter import *
from tkinter import filedialog, ttk, messagebox
from threading import Thread
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

expensesSorted = ""
category_totals = {}
num_transactions = 0

# Run a callback on the main UI thread so messageboxes and label updates aren't blocked
def run_on_main(callback):
    try:
        window.after(0, callback)
    except Exception:
        pass


def chooseFileThreaded():
    Thread(target=chooseFile, daemon=True).start()


def chooseFile():
    workbook_path = filedialog.askopenfilename(
        title="Choose file",
        filetypes=(("Excel files", "*.xlsx"), ("All Files", "*.*")),
        initialdir=r"%userprofile%\downloads",
    )
    fileFormat = (workbook_path[-1:-6:-1])[::-1] if workbook_path else ""
    if fileFormat == ".xlsx":
        run_on_main(lambda: loading("Sorting expenses… Please wait."))
        try:
            from handleExcel import getExpenses
            global expensesSorted, category_totals, num_transactions
            expensesSorted, category_totals, num_transactions = getExpenses(workbook_path)
            run_on_main(lambda: on_sort_success())
        except Exception as e:
            run_on_main(lambda: on_sort_error(str(e)))
    else:
        if workbook_path:
            run_on_main(lambda: (loading(""), messagebox.showerror("Error", "The file format is not .xlsx")))


def chooseTargetFileThreaded():
    Thread(target=chooseTargetFile, daemon=True).start()


def chooseTargetFile():
    workbook_path = filedialog.askopenfilename(
        title="Choose file",
        filetypes=(("Excel files", "*.xlsx"), ("All Files", "*.*")),
        initialdir=r"%userprofile%\downloads",
    )
    fileFormat = (workbook_path[-1:-6:-1])[::-1] if workbook_path else ""
    if fileFormat == ".xlsx":
        run_on_main(lambda: loading("Filling expenses… Please wait."))
        try:
            from handleExcel import fillCells
            fillCells(workbook_path, expensesSorted, month=monthComboBox.get())
            run_on_main(lambda: messagebox.showinfo("Success", "Expenses are filled. You can open the output file."))
        except Exception as e:
            run_on_main(lambda: messagebox.showerror("Error", "Filling failed: " + str(e)))
        finally:
            run_on_main(lambda: loading(""))
    else:
        if workbook_path:
            run_on_main(lambda: (loading(""), messagebox.showerror("Error", "The file format is not .xlsx")))


def loading(message):
    loadingText.set(message)
    if message:
        statusText.set("")
        progressBar.start(10)
        getExcelFileButton.config(state=DISABLED)
        chooseFileButton.config(state=DISABLED)
        chooseTargetFileButton.config(state=DISABLED)
    else:
        progressBar.stop()
        getExcelFileButton.config(state=NORMAL)
        chooseFileButton.config(state=NORMAL)
        chooseTargetFileButton.config(state=NORMAL)
    window.update_idletasks()


def on_sort_success():
    loading("")
    statusText.set("Success. Expenses sorted.")
    statusLabel.config(fg=SUCCESS_COLOR)
    if callable(getattr(window, "update_summary_and_chart", None)):
        window.update_summary_and_chart()
    messagebox.showinfo("Success", "Expenses are sorted. Choose a target file.")


def on_sort_error(err_msg):
    loading("")
    statusText.set("Sorting failed: " + err_msg)
    statusLabel.config(fg=ERROR_COLOR)
    messagebox.showerror("Error", "Sorting failed: " + err_msg)


# def saveFile():
#    output_workbook_path = filedialog.asksaveasfilename(title="Save file", filetypes=(("Excel files", "*.xlsx")), initialdir=r"%userprofile%\downloads")


##################### UI ###################

# Colors and theme (works on Windows with ttk)
BG = "#f0f2f5"
CARD_BG = "#ffffff"
ACCENT = "#1a73e8"
ACCENT_HOVER = "#1557b0"
TEXT = "#202124"
TEXT_MUTED = "#5f6368"
SUCCESS_COLOR = "#0f9d58"
ERROR_COLOR = "#c5221f"

window = Tk()
window.title("Expense Sort Automation")
window.minsize(520, 700)
window.geometry("520x700")
window.resizable(True, True)
window.configure(bg=BG)

def on_headless_toggle():
    from getExcelFileFromMax import toggleHeadless
    toggleHeadless(headlessCheckbuttonState.get())

def on_get_excel_file():
    y, m = yearComboBox.get(), monthComboBox.get()
    loading("Getting Excel file…")
    def run_get_excel():
        try:
            from getExcelFileFromMax import getExcelFile
            getExcelFile(y, m)
        finally:
            run_on_main(lambda: loading(""))
    Thread(target=run_get_excel, daemon=True).start()

menuBar = Menu(window)
settingsMenu = Menu(menuBar, tearoff=0)
headlessCheckbuttonState = IntVar()
settingsMenu.add_checkbutton(
    label="Headless",
    command=on_headless_toggle,
    variable=headlessCheckbuttonState,
)
menuBar.add_cascade(label="Settings", menu=settingsMenu)
window.config(menu=menuBar)

# ttk style
style = ttk.Style()
style.theme_use("clam")
style.configure("TFrame", background=BG)
style.configure(
    "TButton",
    font=("Segoe UI", 10),
    padding=(12, 8),
)
style.configure("TLabel", font=("Segoe UI", 10), background=BG, foreground=TEXT)
style.configure("TLabelframe", font=("Segoe UI", 10), background=BG)
style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), background=BG, foreground=TEXT)

label_font = ("Segoe UI", 10)

# Main content with padding
main = Frame(window, bg=BG, padx=24, pady=20)
main.pack(fill=BOTH, expand=True)

# Card: File actions
fileCard = LabelFrame(main, text="  Files  ", bg=CARD_BG, fg=TEXT, font=("Segoe UI", 10, "bold"), padx=16, pady=14)
fileCard.pack(fill=X, pady=(0, 14))

fileFrame = Frame(fileCard, bg=CARD_BG)
fileFrame.pack(fill=X)

months = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
years = ["2022", "2023", "2024", "2025", "2026"]
defaultYear = years[-1]
defaultMonth = months[0]

getExcelFileButton = ttk.Button(
    fileFrame,
    text="Get Excel File",
    command=on_get_excel_file,
)
getExcelFileButton.pack(side=LEFT, padx=(0, 8))

chooseFileButton = ttk.Button(fileFrame, text="Choose File", command=chooseFileThreaded)
chooseFileButton.pack(side=LEFT, padx=4)

chooseTargetFileButton = ttk.Button(fileFrame, text="Choose Target File", command=chooseTargetFileThreaded)
chooseTargetFileButton.pack(side=LEFT, padx=4)

# Status (messages not blocked by tasks)
loadingText = StringVar()
statusText = StringVar()
statusFrame = Frame(main, bg=BG)
statusFrame.pack(fill=X, pady=(0, 14))
loadingLabel = Label(
    statusFrame,
    textvariable=loadingText,
    font=("Segoe UI", 9),
    fg=TEXT_MUTED,
    bg=BG,
    wraplength=400,
    justify=LEFT,
)
loadingLabel.pack(anchor=W)
progressBar = ttk.Progressbar(statusFrame, mode="indeterminate")
progressBar.pack(fill=X, pady=(4, 0))
statusLabel = Label(
    statusFrame,
    textvariable=statusText,
    font=("Segoe UI", 9),
    fg=TEXT_MUTED,
    bg=BG,
    wraplength=400,
    justify=LEFT,
)
statusLabel.pack(anchor=W, pady=(4, 0))

# Card: Year & Month
selectCard = LabelFrame(main, text="  Period  ", bg=CARD_BG, fg=TEXT, font=("Segoe UI", 10, "bold"), padx=16, pady=12)
selectCard.pack(fill=X, pady=(0, 14))

selectFrame = Frame(selectCard, bg=CARD_BG)
selectFrame.pack(fill=X)

Label(selectFrame, text="Year:", font=label_font, bg=CARD_BG, fg=TEXT).grid(row=0, column=0, sticky=E, padx=(0, 8), pady=4)
yearComboBox = ttk.Combobox(selectFrame, values=years, width=10)
yearComboBox.set(defaultYear)
yearComboBox.grid(row=0, column=1, padx=(0, 16), pady=4)

Label(selectFrame, text="Month:", font=label_font, bg=CARD_BG, fg=TEXT).grid(row=1, column=0, sticky=E, padx=(0, 8), pady=4)
monthComboBox = ttk.Combobox(selectFrame, values=months, width=10)
monthComboBox.set(defaultMonth)
monthComboBox.grid(row=1, column=1, padx=(0, 16), pady=4)

# Card: Summary & charts
summaryCard = LabelFrame(main, text="  Summary & charts  ", bg=CARD_BG, fg=TEXT, font=("Segoe UI", 10, "bold"), padx=16, pady=12)
summaryCard.pack(fill=BOTH, expand=True, pady=(0, 14))

summaryText = StringVar(value="Choose a file and sort to see breakdown.")
summaryLabel = Label(
    summaryCard,
    textvariable=summaryText,
    font=("Segoe UI", 10),
    fg=TEXT,
    bg=CARD_BG,
    wraplength=500,
)
summaryLabel.pack(anchor=W, pady=(0, 8))

chartFrame = Frame(summaryCard, bg=CARD_BG)
chartFrame.pack(fill=BOTH, expand=True)

# Category table (Treeview)
tableFrame = Frame(summaryCard, bg=CARD_BG)
tableFrame.pack(fill=BOTH, expand=True, pady=(8, 0))
treeScroll = ttk.Scrollbar(tableFrame)
treeScroll.pack(side=RIGHT, fill=Y)
categoryTree = ttk.Treeview(tableFrame, columns=("category", "amount", "pct"), show="headings", height=6, yscrollcommand=treeScroll.set)
categoryTree.heading("category", text="Category")
categoryTree.heading("amount", text="Amount (₪)")
categoryTree.heading("pct", text="% of total")
categoryTree.column("category", width=200)
categoryTree.column("amount", width=100)
categoryTree.column("pct", width=80)
categoryTree.pack(side=LEFT, fill=BOTH, expand=True)
treeScroll.config(command=categoryTree.yview)


def update_summary_and_chart():
    global category_totals, num_transactions
    for w in chartFrame.winfo_children():
        w.destroy()
    for row in categoryTree.get_children():
        categoryTree.delete(row)
    if not category_totals:
        summaryText.set("Choose a file and sort to see breakdown.")
        return
    total = sum(category_totals.values())
    n_cat = len(category_totals)
    summaryText.set("Total: ₪{:,.2f} · {} categories · {} transactions".format(total, n_cat, num_transactions))
    # Sort by amount descending for chart and table
    sorted_items = sorted(category_totals.items(), key=lambda x: -x[1])
    categories = [x[0] for x in sorted_items]
    amounts = [x[1] for x in sorted_items]
    for cat, amt in sorted_items:
        pct = (amt / total * 100) if total else 0
        categoryTree.insert("", END, values=(cat, "{:,.2f}".format(amt), "{:.1f}%".format(pct)))
    fig = Figure(figsize=(5, 4), dpi=100, facecolor=CARD_BG)
    ax = fig.add_subplot(111)
    ax.set_facecolor(CARD_BG)
    try:
        ax.tick_params(colors=TEXT)
        ax.xaxis.label.set_color(TEXT)
        ax.yaxis.label.set_color(TEXT)
    except Exception:
        pass
    # Horizontal bar chart (easier to read Hebrew labels)
    y_pos = range(len(categories))
    ax.barh(y_pos, amounts, color=ACCENT, height=0.7)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(categories, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel("Amount (₪)")
    fig.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=chartFrame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=BOTH, expand=True)


window.update_summary_and_chart = update_summary_and_chart

# Bottom: Close
actionFrame = Frame(main, bg=BG)
actionFrame.pack(fill=X)

def on_close():
    window.destroy()

closeButton = ttk.Button(actionFrame, text="Close", command=on_close)
closeButton.pack(side=RIGHT)

# Center on screen
window.update_idletasks()
x = (window.winfo_screenwidth() - window.winfo_reqwidth()) // 2
y = (window.winfo_screenheight() - window.winfo_reqheight()) // 2
window.geometry(f"+{x}+{y}")

# Clean exit when closing the window (X or Close)
window.protocol("WM_DELETE_WINDOW", on_close)
window.mainloop()
