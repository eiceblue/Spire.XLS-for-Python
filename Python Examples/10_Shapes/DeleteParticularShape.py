from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_5.xlsx"
outputFile = "DeleteParticularShape.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Delete the first shape in the worksheet
sheet.PrstGeomShapes[0].Remove()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

