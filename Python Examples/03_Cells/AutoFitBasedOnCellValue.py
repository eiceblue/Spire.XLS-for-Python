from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ReadImages.xlsx"
outputFile = "AutoFitBasedOnCellValue.xlsx"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Set value for B8
cell = worksheet.Range["B8"]
cell.Text = "Welcome to Spire.XLS!"
#Set the cell style
style = cell.Style
style.Font.Size = 10
style.Font.IsBold = True
#Autofit column width and row height based on cell value
cell.AutoFitColumns()
cell.AutoFitRows()
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

