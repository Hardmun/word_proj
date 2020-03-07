from docx import Document
from os import getcwd
from os import path

def getpath(folder = ''):
    return path.join(getcwd(), folder)

def splitWordFile():
    word = Document(getpath("files/1.docx"))
    cells = word.tables[0]._cells
    for cell in cells:
        if not cell.text.strip() == "":
            print(cell.text)
        # word.save(getpath("files/1_.docx"))

splitWordFile()