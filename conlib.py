import os
from win32com import client
from tkinter import *

class converter:
    def __init__(self):
        pass
    def XLSXtoPDF(inputPath:str="input/example.xlsx",outputPath:str="output"):
        """Converts XLSX into PDF"""
        app = client.DispatchEx('Excel.Application')
        app.Interactive = False
        app.Visible = False
        workbook = app.Workbooks.open(inputPath)
        output = os.path.splitext(output)[0]
        try:
            workbook.ActiveSheet.ExportAsFixedFormat(0, output)
        except:
            print("Couldn't save, close the file.")
        workbook.Close()


class data:
    def searchFolder(path:str="example",type:str=".any") -> dict:
        """Searches only for files in a given folder."""
        all = {
            'files': [],
            'folders': [],
        }
        try:
            for item in os.listdir(path):
                itemPath = os.path.join(path, item)
                if os.path.isfile(itemPath):
                    if type != '.any':
                        if item[-(len(type)):] == type:
                            all['files'].append(item)
                        else:
                            print(f"Item abg: {item[(len(type))]}")
                    else:
                        all['files'].append(item)
                elif os.path.isdir(itemPath):
                    all['folders'].append(item)
        except FileNotFoundError:
            print(f"Couldnt find path:'{path}'")

        return all
    def cut_float(eingabe:float=0.0) -> float:
        eingabe = eingabe + 0.005
        eingabe = (int)(eingabe*100)
        eingabe = eingabe/100
        return eingabe



class window:
    """tkinter window settings"""
    def __init__(self, size:int=1920, title:str=None, sz_lmt:int=0):
        self.obj = Tk()
        self.wth = int(size)
        self.hth = int(self.wth*0.5625)
        self.tlt = str(title)

        self.obj.title(self.tlt)
        self.obj.geometry(f"{self.wth}x{self.hth}")

        if sz_lmt == 0:
            self.obj.resizable(width=False,height=False)
        else:
            self.obj.minsize(width=(self.wth-sz_lmt), height=(self.hth-sz_lmt))
            self.obj.maxsize(width=(self.wth+sz_lmt), height=(self.hth+sz_lmt))


class srf:
    def __init__(self):
        pass






if not __name__ in "__main__":
    for i in range(0,100):print(" ")
