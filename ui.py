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
        loading()
        global expensesSorted
        expensesSorted = getExpenses(workbook_path)
        messagebox.showinfo("Success", "The expenses are sorted, please choose a target file")
    # otherwise, print an error and return None
    else:
        messagebox.showerror("Failure", "Error: the file's fromat is not xlsx")

def chooseTargetFile():
    # workbook path
    workbook_path = filedialog.askopenfilename(title="Choose file", filetypes=(("Excel files", "*.xlsx"), ("All Files", "*.*")), initialdir=r"%userprofile%\downloads")
    # save the 5 last characters of the file to check if its an xlsx excel file
    fileFormat = (workbook_path[-1:-6:-1])[::-1]
    # if it's an excel file, return the workbook path
    if fileFormat == ".xlsx":
        loadingText = "Loading... Please wait while we sort the expenses, you'll get another message when we are done :)"
        Label(window, text=loadingText).pack()
        # send workbook path to fillCells function
        fillCells(workbook_path, expensesSorted)
        messagebox.showinfo("Success", "The expenses are filled, you can now see the output file")
    # otherwise, print an error and return None
    else:
        messagebox.showerror("Failure", "Error: the file's fromat is not xlsx")

def loading():
    global loadingText
    loadingText.set("Loading... Please wait while we sort the expenses, you'll get another message when we are done :)")

# def saveFile():
#    output_workbook_path = filedialog.asksaveasfilename(title="Save file", filetypes=(("Excel files", "*.xlsx")), initialdir=r"%userprofile%\downloads")


##################### UI ###################

# ui window
window = Tk()

# bool to indicate if a date was chosen already
#dateChosen = False
# months list
months=["1","2","3","4","5","6","7","8","9","10","11","12"]
# year list
years=["2022","2023","2024"]
# default year value
defaultYear = years[-1]
# default month value
defaultMonth = months[0]

# choose file button
chooseFileButton = Button(text="Choose File", command=chooseFile)
# loading text
loadingText = StringVar()
# loading label
loadingLabel = Label(window, text=loadingText)
loadingText.set("")
# year combo box
yearComboBox = ttk.Combobox(window, values=years)
# year combo box default value
yearComboBox.set(defaultYear)
# month combo box
monthComboBox = ttk.Combobox(window, values=months)
# month combo box default value
monthComboBox.set(defaultMonth)
# accept year and month button
#acceptYearMonthButton = Button(text="Accept", command=getYearMonth)
# close window button
closeButton = Button(text="Close", command=quit)
# get excel file button
getExcelFileButton = Button(text="Get Excel File", command=lambda: getExcelFile(year = yearComboBox.get(), month = monthComboBox.get()))
# choose target file button
chooseTargetFileButton = Button(text="Choose Target File", command=chooseTargetFile)

# add buttons and input boxes to ui window
chooseFileButton.pack()
loadingLabel.pack()
yearComboBox.pack()
monthComboBox.pack()
#acceptYearMonthButton.pack()
getExcelFileButton.pack()
chooseTargetFileButton.pack()
closeButton.pack()
# window loop
window.mainloop()




