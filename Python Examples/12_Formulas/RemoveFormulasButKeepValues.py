from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/RemoveFormulasButKeepValues.xlsx"
outputFile = "RemoveFormulasButKeepValues.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Loop through worksheets.
for sheet in workbook.Worksheets:
    #Loop through cells.
    for cell in sheet.Range:
        #If the cell contain formula, get the formula value, clear cell content, and then fill the formula value into the cell.
        if cell.HasFormula:
            value = cell.FormulaValue
            cell.Clear(ExcelClearOptions.ClearContent)
            cell.Value2 = value
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

