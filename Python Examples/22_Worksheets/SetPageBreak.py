from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorksheetSample1.xlsx"
outputFile = "SetPageBreak.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set Excel Page Break Horizontally
sheet.HPageBreaks.Add(sheet.Range["A8"])
sheet.HPageBreaks.Add(sheet.Range["A14"])
#Set Excel Page Break Vertically
#sheet.VPageBreaks.Add(sheet.Range["B1"])
#sheet.VPageBreaks.Add(sheet.Range["C1"])
#Set view mode to Preview mode
workbook.Worksheets[0].ViewMode = ViewMode.Preview
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

