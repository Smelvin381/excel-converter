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

THIS ONLY WORKS WITH THE WINDOWS OS AND YOU MIGHT NEED TO
FIRST INSTALL some . I HAVE CREATED A BATCH FILE WHICH
WILL DO  

PLEASE PROVIDE ME WITH FEEDBACK AND ERROR REPORTS!
THANK YOU FOR USING MY PROGRAMM AND BEING AWSOME!"""


try:
    from win32com import client
    import pywintypes
    import tkinter
    import time
    import json
except ModuleNotFoundError:
    print("You are missing some modules.")
    print("Please run the batch file in the main folder.")




class DataCtrl:
    """Edit, Read and Save things like variables, files and so on."""
    class Clean:
        """Cleanses the confusing text from the terminal for
        better overview. You are welcome..."""
        def __init__(self) -> None:
            for i in range(0,100):
                print(" ")


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

        self.worksheet = self.workbook.Worksheets('Sheet1')
        # The opend sheet to work with.


    def open_close(self,boot:bool=True) -> None:
        """Open or close this excel-file. By default, the 
        workbook is opened when defined.
        True = Open; False = Close"""
        if boot:
            self.workbook = self.app.Workbooks.Open(self.insert)
            self.status = True
            # Opens the workbook to work with.

        else:
            self.workbook.Close()
            self.status = False
            # Closes the workbook.


    def save(self) -> None:
        """Save the current changes."""


    def to_pdf(self) -> bool:
        """First a the workbook is open, after that the file 
        is exported as a pdf and at last the workbook is closed.
        You should save the file first before converting."""


        print("Converting into PDF, Please wait...")

        try:
            print(f"Exporting: '{self.output}'")
            self.workbook.ActiveSheet.ExportAsFixedFormat(0, self.output)
            # Saves the workbook as a given format (like PDF).

            print(f"Converted {self.name}.xlsx into {self.name}.pdf")
            return True

        except pywintypes.com_error:
            # Unfortunately, I have no idea what the argument is called for this error.

            print("File could not be found. Make sure the file-path is correct.")
            print(f"Input > {self.insert}")
            print(f"Output > {self.output}")
            return False


    def edit_value(self) -> None:
        """Simply edits the value in a given cell."""
        self.worksheet.Cells(2,2).Value = "test"

        self.workbook.Save()




if __name__ in "__main__":
    DataCtrl.Clean()
    base = XlsxEditing("kopf")
    base.edit_value()
    base.to_pdf()
