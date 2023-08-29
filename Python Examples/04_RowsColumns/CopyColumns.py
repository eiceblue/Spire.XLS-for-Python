from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Copying.xls"
outputFile = "CopyColumns.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
sheet1 = workbook.Worksheets[0]
sheet2 = workbook.Worksheets[1]
#Copy the first column to the third column in the same sheet
sheet1.Copy(sheet1.Columns[0], sheet1.Columns[2], True, True, True)
#Copy the first column to the second column in the different sheet
sheet1.Copy(sheet1.Columns[0], sheet2.Columns[1], True, True, True)

#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


