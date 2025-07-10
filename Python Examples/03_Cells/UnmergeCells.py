from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "UnmergeCells.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Unmerge the cells.
sheet.Range["F2"].UnMerge()
#Unmerge the cells.
sheet.Range["F7"].UnMerge()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


