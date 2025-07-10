from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "MergeCells.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Merge the seventh column in Excel file.
workbook.Worksheets[0].Columns[6].Merge()
#Merge the particular range in Excel file.
workbook.Worksheets[0].Range["A14:D14"].Merge()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
