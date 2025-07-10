from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/CreateTable.xlsx"
outputFile = "SetExcelCalculationMode.xlsx"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Set excel calculation mode as Manual
workbook.CalculationMode = ExcelCalculationMode.Manual
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
