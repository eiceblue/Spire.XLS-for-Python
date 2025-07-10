from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_2.xlsx"
outputFile = "UnfreezeExcelPanes.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Unfreeze the panes.
sheet.RemovePanes()
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
