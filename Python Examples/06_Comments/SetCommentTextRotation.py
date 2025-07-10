from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/CellValues.xlsx"
outputFile = "SetCommentTextRotation.xlsx"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get the default first  worksheet
sheet = workbook.Worksheets[0]
#Create Excel font
font = workbook.CreateFont()
font.FontName = "Arial"
font.Size = 11
font.KnownColor = ExcelColors.Orange
#Add the comment
range = sheet.Range["E1"]
range.Comment.Text = "This is a comment"
range.Comment.RichText.SetFont(0, (len(range.Comment.Text) - 1), font)
# Set its vertical and horizontal alignment 
range.Comment.VAlignment = CommentVAlignType.Center
range.Comment.HAlignment = CommentHAlignType.Right
#Set the comment text rotation
range.Comment.TextRotation = TextRotationType.LeftToRight
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
