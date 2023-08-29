from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/RemoveDataValidation.xlsx"
outputFile = "RemoveDataValidation.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Create an array of rectangles, which is used to locate the ranges in worksheet.
rectangles = []
#Assign value to the first element of the array. This rectangle specifies the cells from A1 to B3.
rectangles.append(Rectangle.FromLTRB(0, 0, 1, 2))
#Remove validations in the ranges represented by rectangles.
workbook.Worksheets[0].DVTable.Remove(rectangles)
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()



