import tkinter
from tkinter import filedialog as fd

#returns the file path of a file on someones desktop
def getFilePath(title: str):
    tkinter.Tk().withdraw()
    return fd.askopenfile(title=title).name

    tkinter.Tk().withdraw()
    return fd.askopenfile().name
#returns true if the filePath is a excel file type; else returns false.
def isExcelFile(filePath: str):
    return (filePath.lower().endswith('.xlsx') or filePath.lower().endswith('.xls'))