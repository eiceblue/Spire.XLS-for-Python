from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/CommentSample.xlsx"
outputFile = "HideOrShowComment.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Hide the second comment
sheet.Comments[1].IsVisible = False
#Show the third comment
sheet.Comments[2].IsVisible = True
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


