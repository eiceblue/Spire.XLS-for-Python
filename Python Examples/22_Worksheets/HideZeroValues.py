from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SampleB_2.xlsx"
outputFile = "HideZeroValues.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Set false to hide the zero values in sheet
sheet.IsDisplayZeros = False
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
