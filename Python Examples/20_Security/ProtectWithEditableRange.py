from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ProtectWithEditableRange.xlsx"
outputFile = "ProtectWithEditableRange.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Define the specified ranges to allow users to edit while sheet is protected
sheet.AddAllowEditRange("EditableRanges", sheet.Range["B4:E12"])
#Protect worksheet with a password.
sheet.Protect("TestPassword", SheetProtectionType.All)
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

