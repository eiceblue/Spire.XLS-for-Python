from spire.xls import *
from spire.common import *


outputFile = "MergeExcelFiles.xlsx"
files = []
files.append("./Demos/Data/MergeExcelFiles-1.xlsx" )
files.append("./Demos/Data/MergeExcelFiles-2.xls")
files.append("./Demos/Data/MergeExcelFiles-3.xlsx")

newbook = Workbook()
newbook.Version = ExcelVersion.Version2013
#Clear all worksheets
newbook.Worksheets.Clear()
#Create a workbook
tempbook = Workbook()
for file in files:
    #Load the file
    tempbook.LoadFromFile(file)
    for sheet in tempbook.Worksheets:
        #Copy every sheet in a workbook
        newbook.Worksheets.AddCopy(sheet, WorksheetCopyType.CopyAll)
#Save the file
newbook.SaveToFile(outputFile, ExcelVersion.Version2010)
newbook.Dispose()
tempbook.Dispose()

