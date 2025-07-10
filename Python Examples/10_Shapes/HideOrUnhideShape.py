from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_5.xlsx"
outputFile = "HideOrUnhideShape.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Hide the second shape in the worksheet
sheet.PrstGeomShapes[1].Visible = False
#Show the second shape in the worksheet
#sheet.PrstGeomShapes[1].Visible = true
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
