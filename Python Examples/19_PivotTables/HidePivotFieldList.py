from spire.xls import *

inputFile = "./Demos/Data/PivotTableExample.xlsx"
outputFile = "HidePivotFieldList.xlsx"

#Create a workbook
workbook = Workbook()

#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)

# Hide the PivotTable field list panel
workbook.HidePivotFieldList =True

#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()