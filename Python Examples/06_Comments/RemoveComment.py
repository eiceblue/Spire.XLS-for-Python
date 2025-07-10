from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/CommentSample.xlsx"
outputFile = "RemoveComment.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get all comments of the first sheet
comments = workbook.Worksheets[0].Comments
#Change the content of the first comment
comments[0].Text = "This comment has been changed."
#Remove the second comment
comments[1].Remove()
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

