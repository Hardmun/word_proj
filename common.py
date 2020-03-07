from os import path
from os import getcwd
import sys

def projectdir(ctlg, usetempdir=False):
    """if executable file -  have to change the default path"""
    if getattr(sys, 'frozen', False) and usetempdir:
        exe_path = os.path.dirname(sys.executable)
        dirPath = os.path.join(getattr(sys, "_MEIPASS", exe_path), ctlg)
    else:
        dirPath = ctlg
    return dirPath

def full_path(path_to_file):
    return path.join(getcwd(), path_to_file)

