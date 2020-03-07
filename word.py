from common import full_path
import docx
import subprocess
import os
import unoconv

for fl in os.listdir(os.getcwd()):
    if fl.endswith('.doc'):
        print(fl)
        subprocess.call(['d:\My documents\Python\word_proj\unoconv', '-d', 'document', '--format=docx', fl], shell=True)

