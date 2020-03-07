from docx import Document
from os import getcwd
from os import path

def getpath(folder = ''):
    return path.join(getcwd(), folder)

def splitWordFile():
    word1 = Document(getpath("files/1.docx"))
    word = word1
    rows = word.tables[0].rows
    # rows[13].delete()

    tbl = word.tables[0]._tbl
    tr = rows[13]._tr
    tbl.remove(tr)

    for row in rows:
        pass
    # cells = word.tables[0]._cells
    # for cell in cells:
    #     if not cell.text.strip() == "":
    #         print(cell.text)
    word1.save(getpath("files/1_.docx"))
    word.save(getpath("files/1__.docx"))

splitWordFile()