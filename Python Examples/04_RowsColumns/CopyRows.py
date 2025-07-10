from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Copying.xls"
outputFile = "CopyRows.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
sheet1 = workbook.Worksheets[1]
sheet2 = workbook.Worksheets[0]
#Copy the first row to the third row in the same sheet
sheet1.Copy(sheet1.Rows[0], sheet1.Rows[2], True, True, True)
#Copy the first row to the second row in the different sheet
sheet1.Copy(sheet1.Rows[0], sheet2.Rows[1], True, True, True)
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)

workbook.Dispose()

