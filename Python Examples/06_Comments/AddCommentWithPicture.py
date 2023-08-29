from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Logo.png"
outputFile = "AddCommentWithPicture.xlsx"

#Create a workbook
workbook = Workbook()
sheet = workbook.Worksheets[0]
sheet.Range["C6"].Text = "E-iceblue"
#Add the comment
comment = sheet.Range["C6"].AddComment()
#Load the image file
image = Image.FromFile(inputFile)
comment.Fill.CustomPicture(image, "logo.png")
#Set the height and width of comment
comment.Height = image.Height
comment.Width = image.Width
comment.Visible = True
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
