"""THIS PROGRAMM OFFERS THE USER TO EDIT VALUES IN AN EXCEL-
FILE (needs to be in the same folder as the programm) WHILE 
ALSO BEING ABLE TO SAVE/UPDATE THE FILE EVEN IF THE EXCEL-
FILE IS OPENED. IF YOU WANT TO CONVERT IT INTO A PDF-FILE,
MAKE SURE TO SAVE THE WORKBOOK BEFORE DOING THAT.

YOU CAN CHANGE SETTINGS IN THE MAIN FOLDER USING THE 'config.json' 
FILE. FOR EXAMPLE: BY DEFAULT THE PROGRAMM IS LIMITED TO 
EDIT/SAVE/CONVERT FILES ONLY IF THE FILE IS IN THE SAME 
FOLDER AS THE MAIN PROGRAMM BUT I DO NOT RECOMMEND TO CHANGE
ANYTHING. THE SETTING ARE NOT MEANT TO BE CHANGED.

IF YOU FIND ANY BUGS OR ANYTHING. PLEASE LET ME KNOW"""




import os
try:
    import tkinter
    from tkinter import ttk
    from win32com import client
    import pywintypes
    import json
    import time
except ModuleNotFoundError:
    os.system("cmd /c pip install pywin32")
    os.system("cmd /c pip install tk")
    print("Restart the programm.")
for i in range(0,100):
    print(" ")


class DataCtrl:
    """Edit, Read and Save things like variables, files and so on."""
    class Json:
        """Create/Update/Delet/Read json files."""
        def read_json(path:str="config.json",decoding:str="utf-8") -> dict:
            """Returns the content of a json file as a dict."""
            return json.loads(open(str(path),encoding=decoding).read())
            # First, the file is opend, after that it is converted into a string
            # and at last the string is turned into a dict.


class XlsxEditing:
    """Editing and converting XLSX files."""
    def __init__(self, name: str = "example") -> None:
        self.name = name
        # The name of the currently used excel file (without file extension).

        # Only the name of the file is needed because
        # by default, the only files in the folder 'input' are used.

        self.app = client.DispatchEx("Excel.Application")
        # Setup the programm to open excel files.

        self.app.Interactive = True
        # I guess this indicates if the excel file is editable.

        self.app.Visible = False
        # Open the Microsoft Excel Window when opend.

        self.insert = f"{DataCtrl.Json.read_json()['convert']['input_path']}{self.name}"
        # The default path to the xlsx file.

        self.output = f"{DataCtrl.Json.read_json()['convert']['output_path']}{self.name}"
        # The default path to the pdf file, if converted to pdf.

        self.workbook = self.app.Workbooks.Open(self.insert)
        # Opens the workbook to work with.

        self.status = True
        # The current status of the workbook.

        self.worksheet = self.workbook.Worksheets('Tabelle1')
        # The opend sheet to work with.


    def open_close(self,boot:bool=True,fully_close:bool=False) -> None:
        """Open or close this excel-file. By default, the 
        workbook is opened when defined.
        True = Open; False = Close"""
        if boot:
            self.workbook = self.app.Workbooks.Open(self.insert)
            self.status = True
            # Opens the workbook to work with.

        elif boot is False and fully_close is False:
            self.workbook.Close()
            self.status = False
            # Closes the workbook.

        elif fully_close:
            self.workbook.Close()
            self.app.Quit()
            self.status = False


    def to_pdf(self) -> bool:
        """First a the workbook is open, after that the file 
        is exported as a pdf and at last the workbook is closed.
        You should save the file first before converting."""
        if not self.status:
            print("Open Workbook/Excel Application first.")
            return False


        print("Converting into PDF, Please wait...")

        try:
            print(f"Exporting: '{self.output}'")
            self.workbook.ActiveSheet.ExportAsFixedFormat(0, self.output)
            # Saves the workbook as a given format (like PDF).

            print(f"Converted {self.name}.xlsx into {self.name}.pdf")
            return True

        except pywintypes.com_error:
            # Unfortunately, I have no idea what the argument is called for this error.

            print("File could not be found.")
            print(f"Input > {self.insert}")
            print(f"Output > {self.output}")
            return False


    def edit_value(self,new_value:str="Example",
                   hztl_index:int=1,
                   vrtc_index:int=1,
                   save_after:bool=True) -> None:
        """Simply edits the value in a given cell.
        hztl_index = Horizontal Index/vrtc_index = vertical"""
        if not self.status:
            print("Open Workbook/Excel Application first.")
            return False

        self.worksheet.Cells(vrtc_index,hztl_index).Value = str(new_value)

        if save_after:
            self.workbook.Save()





if __name__ in "__main__":
    win = tkinter.Tk("1")
    win.title("ZA CHESTA")
    win.geometry("800x400")
    win.resizable(False,False)

    lable1 = ttk.Label(win)
    lable1.pack()

    lable1.configure(text="ZA CHESTA!")


    for i in range(0,100):
        win.update()
        time.sleep(0.1)
        lable1.configure(text=str(i))
