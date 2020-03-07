from common import full_path
import docx
import subprocess
import os
from win32com import client

for fl in os.listdir(os.getcwd()):
    if fl.endswith('.doc'):
        print(fl)
        wrd = client.Dispatch("Word.Application")
        wrd.visible = 0
        doc = wrd.Documents.Open(full_path("22.doc"))
        doc.SaveAs(full_path("testconvert"), FileFormat=12)
        doc.Close()
        wrd.Quit()
