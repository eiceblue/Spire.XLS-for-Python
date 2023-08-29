from spire.xls import *
from spire.common import *


outputFile = "AddVariableArray.xlsx"

#Create a workbook
workbook = Workbook()
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set marker designer field in cell A1
sheet.Range["A1"].Value = "&=Array"
#Fill Array
workbook.MarkerDesigner.AddArray("Array", [String("Spire.Xls"), String("Spire.Doc"), String("Spire.PDF"), String("Spire.Presentation"), String("Spire.Email")])
workbook.MarkerDesigner.Apply()
workbook.CalculateAllValue()
#AutoFit
sheet.AllocatedRange.AutoFitRows()
sheet.AllocatedRange.AutoFitColumns()
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

