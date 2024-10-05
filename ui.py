from tkinter import *
from tkinter import filedialog, ttk, messagebox
from getExcelFileFromMax import getExcelFile
from handleExcel import getExpenses, fillCells

expensesSorted = ""

# ui choose file function 
def chooseFile():
    # workbook path
    workbook_path = filedialog.askopenfilename(title="Choose file", filetypes=(("Excel files", "*.xlsx"), ("All Files", "*.*")), initialdir=r"%userprofile%\downloads")
    # save the 5 last characters of the file to check if its an xlsx excel file
    fileFormat = (workbook_path[-1:-6:-1])[::-1]
    # if it's an excel file, return the workbook path
    if fileFormat == ".xlsx":
        loading("Loading... Please wait while we sort the expenses, you'll get another message when we are done :)")
        global expensesSorted
        expensesSorted = getExpenses(workbook_path)
        messagebox.showinfo("Success", "The expenses are sorted, please choose a target file")
    # otherwise, print an error and return None
    else:
        messagebox.showerror("Failure", "Error: the file's fromat is not xlsx")

def chooseTargetFile():
    # workbook path
    workbook_path = filedialog.askopenfilename(
        title="Choose file", 
        filetypes=(("Excel files", "*.xlsx"), ("All Files", "*.*")), 
        initialdir=r"%userprofile%\downloads"
        )
    # save the 5 last characters of the file to check if its an xlsx excel file
    fileFormat = (workbook_path[-1:-6:-1])[::-1]
    # if it's an excel file, return the workbook path
    if fileFormat == ".xlsx":
        loading("Loading... Please wait while we fill the expenses, you'll get another message when we are done :)")
        # send workbook path to fillCells function
        fillCells(workbook_path, expensesSorted)
        messagebox.showinfo("Success", "The expenses are filled, you can now see the output file")
    # otherwise, print an error and return None
    else:
        messagebox.showerror("Failure", "Error: the file's fromat is not xlsx")

def loading(message):
    global loadingText
    loadingText.set(message)
    window.update_idletasks() # This forces the UI to refresh immediately

# def saveFile():
#    output_workbook_path = filedialog.asksaveasfilename(title="Save file", filetypes=(("Excel files", "*.xlsx")), initialdir=r"%userprofile%\downloads")


##################### UI ###################

# ui window
window = Tk()
window.title("Expense Sort Automation")
window.geometry("400x300")
window.resizable(False, False)  # Disable resizing
# @TODO add an icon to the window

# Styling
label_font = ("Arial", 10)
button_font = ("Arial", 10)

# File Selection Frame
fileFrame = Frame(window)
fileFrame.pack(pady=10)

# months list
months=["1","2","3","4","5","6","7","8","9","10","11","12"]
# year list
years=["2022","2023","2024"]
# default year value
defaultYear = years[-1]
# default month value
defaultMonth = months[0]

# choose file button
chooseFileButton = ttk.Button(fileFrame, text="Choose File", command=chooseFile)
chooseFileButton.grid(row=0, column=1, padx=5)
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

# Action Buttons Frame
actionFrame = Frame(window)
actionFrame.pack(pady=10)

getExcelFileButton = ttk.Button(fileFrame, text="Get Excel File", command=lambda: getExcelFile(year=yearComboBox.get(), month=monthComboBox.get()))
getExcelFileButton.grid(row=0, column=0, padx=5)

closeButton = ttk.Button(actionFrame, text="Close", command=window.quit)
closeButton.grid(row=0, column=0, padx=5)

# choose target file button
chooseTargetFileButton = ttk.Button(fileFrame, text="Choose Target File", command=chooseTargetFile)
chooseTargetFileButton.grid(row=0, column=2, padx=5)

# Center everything on the window
window.update_idletasks()
x = (window.winfo_screenwidth() - window.winfo_reqwidth()) // 2
y = (window.winfo_screenheight() - window.winfo_reqheight()) // 2
window.geometry(f"+{x}+{y}")

# window loop
window.mainloop()




