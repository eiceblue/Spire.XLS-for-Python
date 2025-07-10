from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/DataValidation.xlsx"
outputFile = "WholeNumberDataValidation.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
sheet.Range["C12"].Text = "Please enter number between 10 and 100:"
sheet.Range["C12"].AutoFitColumns()
#Set Whole Number data validation for cell "D12"
range = sheet.Range["D12"]
range.DataValidation.AllowType = CellDataType.Integer
range.DataValidation.CompareOperator = ValidationComparisonOperator.Between
range.DataValidation.Formula1 = "10"
range.DataValidation.Formula2 = "100"
range.DataValidation.AlertStyle = AlertStyleType.Info
range.DataValidation.ShowError = True
range.DataValidation.ErrorTitle = "Error"
range.DataValidation.ErrorMessage = "Please enter a valid number"
range.DataValidation.InputMessage = "Whole Number Validation Type"
range.DataValidation.IgnoreBlank = True
range.DataValidation.ShowInput = True
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

