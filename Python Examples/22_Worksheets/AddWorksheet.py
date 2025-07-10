from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/AddWorksheet.xlsx"
outputFile = "AddWorksheet.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Add a new worksheet named AddedSheet
sheet = workbook.Worksheets.Add("AddedSheet")
sheet.Range["C5"].Text = "This is a new sheet."
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

