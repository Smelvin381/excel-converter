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
        def __init__(self,path:str="config.json",decoding:str="utf-8") -> None:
            self.path = path
            # The path to the settings/configuration.

            self.decoding = decoding
            # The encoding of strings in a json file (like utf-8, ascii).

        def read_json(self) -> dict:
            """Returns the content of a json file as a dict."""
            return json.loads(open(str(self.path),encoding=self.decoding).read())
            # First, the file is opend, after that it is converted into a string
            # and at last the string is turned into a dict.


class XlsxEditing:
    """Editing and converting XLSX files."""
    settings = DataCtrl.Json()
    settings.read_json()["convert"]
    def __init__(self, name: str = "example") -> None:
        self.name = name
        # The name of the currently used excel file (without file extension).

        # Only the name of the file is needed because
        # by default, the only files in the folder 'input' are used.

        self.app = client.DispatchEx("Excel.Application")
        # Setup the programm to open excel files.

        self.app.Interactive = False
        # I guess this indicates if the excel file is editable.

        self.app.Visible = False
        # Honestly i have no idea.
    def to_pdf(self) -> bool:
        """First a the workbook is open, after that the file 
        is exported as a pdf and at last the workbook is closed.
        You should save the file first before converting."""

        insert = f"C:\\Users\\VW6F8P7\\Documents\\excel converter\\input\\{self.name}"
        # The default path to the xlsx file.

        output = f"C:\\Users\\VW6F8P7\\Documents\\excel converter\\output\\{self.name}"
        # The default path to the pdf file.

        print("Converting into PDF, Please wait...")

        try:
            print(f"Opening workbook: '{insert}'")
            workbook = self.app.Workbooks.Open(insert)
            # Opens the workbook to work with.

            print(f"Exporting: '{output}'")
            workbook.ActiveSheet.ExportAsFixedFormat(0, output)
            # Saves the workbook as a given format (like PDF).

            print(f"Closing workbook: {self.name}")
            workbook.Close()
            # Closes the workbook.


            print(f"Converted {self.name}.xlsx into {self.name}.pdf")
            return True
        except:  # Unfortunately, I have no idea what the argument is called for this type of error.
            return False




if __name__ in "__main__":
    DataCtrl.Clean()
    XlsxEditing("kopf")
    print(DataCtrl.Json())
