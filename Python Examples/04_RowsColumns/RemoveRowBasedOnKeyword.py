from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorkbookToHTML.xlsx"
outputFile = "RemoveRowBasedOnKeyword.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Find the string
cr = sheet.FindString("Address", False, False)
#Delete the row which includes the string
sheet.DeleteRow(cr.Row)
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

