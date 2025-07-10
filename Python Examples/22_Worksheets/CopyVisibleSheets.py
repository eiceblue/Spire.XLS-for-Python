from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/CopyVisibleSheets.xlsx"
outputFile = "CopyVisibleSheets.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Create a new workbook
workbookNew = Workbook()
workbookNew.Version = ExcelVersion.Version2013
workbookNew.Worksheets.Clear()
#Loop through the worksheets
for sheet in workbook.Worksheets:
    #Judge if the worksheet is visible
    if sheet.Visibility == WorksheetVisibility.Visible:
        #Copy the sheet to new workbook
        name = sheet.Name
        workbookNew.Worksheets.AddCopy(sheet)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

