from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/FormulasSample.xlsx"
outputFile = "HideFormulas.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Hide the formulas in the used range
sheet.AllocatedRange.IsFormulaHidden = True
#Protect the worksheet with password
sheet.Protect("e-iceblue")
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

