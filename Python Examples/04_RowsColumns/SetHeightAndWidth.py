from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SetHeightAndWidth.xls"
outputFile = "SetHeightAndWidth.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
worksheet = workbook.Worksheets[0]
# Setting the width to 30
worksheet.SetColumnWidth(4, 30)
# Setting the height to 30
worksheet.SetRowHeight(4, 30)
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
