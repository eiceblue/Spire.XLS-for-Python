from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ExcelSample_N1.xlsx"
outputFile = "AddSpinnerControl.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set text for range C11
sheet.Range["C11"].Text = "Value:"
sheet.Range["C11"].Style.Font.IsBold = True
#Set value for range B10
sheet.Range["C12"].Value2 = Int32(0)
#Add spinner control
spinner = sheet.SpinnerShapes.AddSpinner(12, 4, 20, 20)
spinner.LinkedCell = sheet.Range["C12"]
spinner.Min = 0
spinner.Max = 100
spinner.IncrementalChange = 5
spinner.Display3DShading = True
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

