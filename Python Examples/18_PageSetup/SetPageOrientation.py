from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "SetPageOrientation.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Set the page orientation to Landscape. 
sheet.PageSetup.Orientation = PageOrientationType.Landscape
workbook.SaveToFile(outputFile,ExcelVersion.Version2010)
workbook.Dispose()

