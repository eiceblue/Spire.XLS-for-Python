from spire.common import *
from spire.xls import *


inputFile = "./Demos/Data/InsertControls.xlsx"
outputFile = "InsertControls.xlsx"

wb = Workbook()
wb.LoadFromFile(inputFile)
ws = wb.Worksheets[0]

#Add a textbox 
textbox = ws.TextBoxes.AddTextBox(9, 2, 25, 100)
textbox.Text = "Hello World"
#Add a checkbox 
cb = ws.CheckBoxes.AddCheckBox(11, 2, 15, 100)
cb.CheckState = CheckState.Checked
cb.Text = "Check Box 1"
#Add a RadioButton 
rb = ws.RadioButtons.Add(13, 2, 15, 100)
rb.Text = "Option 1"

#Add a combox
cbx = ws.ComboBoxes.AddComboBox(15, 2, 15, 100) if isinstance(ws.ComboBoxes.AddComboBox(15, 2, 15, 100), IComboBoxShape) else None
cbx.ListFillRange = ws.Range["A41:A47"]
wb.SaveToFile(outputFile, ExcelVersion.Version2010)
wb.Dispose()
