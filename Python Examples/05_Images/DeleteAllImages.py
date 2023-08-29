from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_5.xlsx"
outputFile = "DeleteAllImages.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Delete all images from the worksheet.
for i in range(sheet.Pictures.Count - 1, -1, -1):
    sheet.Pictures[i].Remove()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


