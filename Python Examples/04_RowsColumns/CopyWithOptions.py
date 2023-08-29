from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Sample.xlsx"
outputFile = "CopyWithOptions.xlsx"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet1 = workbook.Worksheets[0]
#Add a new worksheet as destination sheet
destinationSheet = workbook.Worksheets.Add("DestSheet")
#Specify a copy range of original sheet
cellRange = sheet1.Range["B2:D4"]
#Copy the specified range to added worksheet and keep original styles and update reference
workbook.Worksheets[0].Copy(cellRange, workbook.Worksheets[1], 2, 1, True, True)

#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

