from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorksheetSample1.xlsx"
outputFile = "AddCommentWithAuthor.xlsx"
     
#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the range that will add comment
range = sheet.Range["C1"]
#Set the author and comment content
author = "E-iceblue"
text = "This is demo to show how to add a comment with editable Author property."
#Add comment to the range and set properties
comment = range.AddComment()
comment.Width = 200
comment.Visible = True
comment.Text = author + ":\n" + text
#Set the font of the author
font = workbook.CreateFont()
font.FontName = "Tahoma"
font.KnownColor = ExcelColors.Black
font.IsBold = True
comment.RichText.SetFont(0, len(author), font)
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
