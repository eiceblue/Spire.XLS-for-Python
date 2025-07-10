from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_5.xlsx"
outputFile = "DeleteAllShapes.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Delete all shapes in the worksheet
for i in range(sheet.PrstGeomShapes.Count - 1, -1, -1):
    sheet.PrstGeomShapes[i].Remove()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

