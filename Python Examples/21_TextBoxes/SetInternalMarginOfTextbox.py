from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "SetInternalMarginOfTextbox.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Add a textbox to the sheet and set its position and size.
textbox = sheet.TextBoxes.AddTextBox(4, 2, 100, 300)
#Set the text on the textbox.
textbox.Text = "Insert TextBox in Excel and set the margin for the text"
textbox.HAlignment = CommentHAlignType.Center
textbox.VAlignment = CommentVAlignType.Center
#Set the inner margins of the contents.
textbox.InnerLeftMargin = 1
textbox.InnerRightMargin = 3
textbox.InnerTopMargin = 1
textbox.InnerBottomMargin = 1
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

