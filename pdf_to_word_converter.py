# Convert Word Document to PDF
# adapted from https://medium.com/@prasanthrao/convert-pdf-files-to-editable-word-documents-using-python-44c6114a66b2
# !pip install pypiwin32  (this is a pre installed library if not found install it)

import win32com.client
import tkinter as tk
from tkinter import filedialog

# Access MS Word application to read the file
word = win32com.client.Dispatch("Word.Application")
word.visible = 0

# Show file selection dialog box
root = tk.Tk()
root.withdraw()
paths = filedialog.askopenfilenames()
root.update()
print(paths)

# File Paths
for path in paths:
    pdfdoc = r"\\".join(path.split('/'))
# pdfdoc = r"C:\\Users\stefan.lim\Desktop\\NTW\\Q1 2021\\Tui Ora 2021-22 Q1 Nga Tini Whetu.pdf"
# wordDoc = r"C:\\Users\stefan.lim\Desktop\\NTW"

# open pdf file and write it in word document
    wordObj = word.Documents.Open(pdfdoc)
    wordObj.SaveAs(wordObj, FileFormat=16) # file format for docx

# for more file formats refer the link "https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat"

    wordObj.Close()
    word.Quit()