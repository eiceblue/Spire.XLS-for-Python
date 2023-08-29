from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/DataValidation.xlsx"
outputFile = "ListDataValidation.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set text for cells 
sheet.Range["A7"].Text = "Beijing"
sheet.Range["A8"].Text = "New York"
sheet.Range["A9"].Text = "Denver"
sheet.Range["A10"].Text = "Paris"
#Set data validation for cell
range = sheet.Range["D10"]
range.DataValidation.ShowError = True
range.DataValidation.AlertStyle = AlertStyleType.Stop
range.DataValidation.ErrorTitle = "Error"
range.DataValidation.ErrorMessage = "Please select a city from the list"
range.DataValidation.DataRange = sheet.Range["A7:A10"]
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
