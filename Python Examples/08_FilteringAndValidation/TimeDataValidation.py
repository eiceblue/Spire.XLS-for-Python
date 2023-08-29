from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/DataValidation.xlsx"
outputFile = "TimeDataValidation.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
sheet.Range["C12"].Text = "Please enter time between 09:00 and 18:00:"
sheet.Range["C12"].AutoFitColumns()
#Set Time data validation for cell "D12"
range = sheet.Range["D12"]
range.DataValidation.AllowType = CellDataType.Time
range.DataValidation.CompareOperator = ValidationComparisonOperator.Between
range.DataValidation.Formula1 = "09:00"
range.DataValidation.Formula2 = "18:00"
range.DataValidation.AlertStyle = AlertStyleType.Info
range.DataValidation.ShowError = True
range.DataValidation.ErrorTitle = "Time Error"
range.DataValidation.ErrorMessage = "Please enter a valid time"
range.DataValidation.InputMessage = "Time Validation Type"
range.DataValidation.IgnoreBlank = True
range.DataValidation.ShowInput = True
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
