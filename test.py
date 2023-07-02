from win32com import client

excelApp = client.Dispatch("Excel.Application")
excelApp.Visible = False  # Set Excel to be visible (optional)

book = excelApp.Workbooks.Open("C:\\Users\\VW6F8P7\\Documents\\excel converter\\input\\kopf")  # Use the correct file extension (.xlsx)

workSheet = book.Worksheets('Sheet1')
workSheet.Cells(1, 1).Value = "test"

book.Save()
book.Close()

excelApp.Quit()  # Close the Excel application
