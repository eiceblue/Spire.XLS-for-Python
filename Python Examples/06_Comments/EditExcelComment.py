from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_8.xlsx"
outputFile = "EditExcelComment.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the first comment.
comment = sheet.Comments[0]
#Edit the comment.
comment.Text = "This comment has been edited by Spire.XLS."
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

