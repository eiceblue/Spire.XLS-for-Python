from spire.xls import *
from spire.xls.common import *


outputFile = "CopyShapes.xlsx"

workbook = Workbook()
sheet = workbook.Worksheets[0]
#Create line shape
line = sheet.TypedLines.AddLine()
line.Top = 50
line.Left = 30
line.Width = 30
line.Height = 50
line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowDiamond
line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
CopyShapes = workbook.Worksheets[1]
#Copy the line into other sheet
CopyShapes.TypedLines.AddCopy(line)
#Create a button and then copy into other sheet
button = sheet.TypedRadioButtons.Add(5, 5, 20, 20)
CopyShapes.TypedRadioButtons.AddCopy(button)
#Create a textbox and then copy into other sheet
textbox = sheet.TypedTextBoxes.AddTextBox(5, 7, 50, 100)
CopyShapes.TypedTextBoxes.AddCopy(textbox)
#Create a checkbox and then copy into other sheet
checkbox = sheet.TypedCheckBoxes.AddCheckBox(10, 1, 20, 20)
CopyShapes.TypedCheckBoxes.AddCopy(checkbox)
#Create a comboboxes and then copy into other sheet
sheet.Range["A14"].Value = "1"
sheet.Range["A15"].Value = "2"
ComboBoxes = sheet.TypedComboBoxes.AddComboBox(10, 5, 30, 30)
ComboBoxes.ListFillRange = sheet.Range["A14:A15"]
CopyShapes.TypedComboBoxes.AddCopy(ComboBoxes)
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

