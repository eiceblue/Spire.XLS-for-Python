from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ManipulateTextBoxControl.xlsx"
outputFile = "ManipulateTextBox.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the first textbox
tb = sheet.TextBoxes[0]
#Change the text of textbox
tb.Text = "Spire.XLS for Python"
#Set the alignment of textbox as center
tb.HAlignment = CommentHAlignType.Center
tb.VAlignment = CommentVAlignType.Center
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

