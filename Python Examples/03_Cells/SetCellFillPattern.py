from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/CommonTemplate.xlsx"
outputFile = "SetCellFillPattern.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
worksheet = workbook.Worksheets[0]
#Set cell color
worksheet.Range["B7:F7"].Style.Color = Color.get_Yellow()
#Set cell fill pattern
worksheet.Range["B8:F8"].Style.FillPattern = ExcelPatternType.Percent125Gray
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


