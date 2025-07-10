from spire.xls.common import *
from spire.xls import *



inputFile = "./Demos/Data/CreateTable.xlsx"
outputFile = "CopyCellsRange.xlsx"

#Create a workbook
workbook = Workbook()

#Load the Excel document from disk
workbook.LoadFromFile(inputFile)

#Get the first worksheet
sheet1 = workbook.Worksheets[0]

#Specify a destination range 
cells = sheet1.Range["G1:H19"]

#Copy the selected range to destination range 
sheet1.Range["B1:C19"].Copy(cells)

#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
