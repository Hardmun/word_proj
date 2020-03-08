from docx import Document
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os
import time

from configparser import ConfigParser

config = ConfigParser()
configList = config.read("settings.ini")
projectDir = os.getcwd()

if configList.__len__() == 0:
    config["Paths"] = {"Path": projectDir}

    with open("settings.ini", "w") as configfile:
        config.write(configfile)

def splitWordFile(filePath):
    word = Document(filePath)
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
    word.save(os.path.join(os.path.dirname(filePath), "mod.docx"))

# for fl in os.listdir(os.getcwd()):
#     if fl.endswith('.doc'):
#         print(fl)
#         wrd = client.Dispatch("Word.Application")
#         wrd.visible = 0
#         doc = wrd.Documents.Open(full_path("22.doc"))
#         doc.SaveAs(full_path("testconvert"), FileFormat=12)
#         doc.Close()
#         wrd.Quit()

class WordHandler(FileSystemEventHandler):
    def on_any_event(self, event):
        """path to file"""
        filePath = event.src_path
        fileDir = os.path.dirname(filePath)
        for file in os.listdir(fileDir):
            """converting *.doc into *.docx"""
            if file.endswith(".doc"):
                asdf = 0
            elif file.endswith(".docx"):
                splitWordFile(os.path.normpath(os.path.join(fileDir, file)))

class IniHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self.obs = None

    def on_modified(self, event):
        if event.src_path.find("settings.ini") != -1:
            config.read("settings.ini")
            observer = self.obs
            observer.schedule(WordHandler(), path=os.path.normpath(config.get("Paths", "Path")))

if __name__ == '__main__':
    """directory for observing"""
    observer = Observer()
    observer.schedule(WordHandler(), path=os.path.normpath(config.get("Paths", "Path")))
    observer.start()
    """if settings.ini was changed"""
    observerINI = Observer()
    IniHandler = IniHandler()
    IniHandler.obs = observer
    observerINI.schedule(IniHandler, path=projectDir)
    observerINI.start()

    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        observer.stop()
        observerINI.stop()

    observer.join()
    observerINI.join()
