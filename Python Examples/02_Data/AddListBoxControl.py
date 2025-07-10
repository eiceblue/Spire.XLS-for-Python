from spire.xls.common import *
from spire.xls import *


outputFile = "AddListBoxControl.xlsx"

workbook = Workbook()
sheet = workbook.Worksheets[0]
sheet.Range["A7"].Text = "Beijing"
sheet.Range["A8"].Text = "New York"
sheet.Range["A9"].Text = "ChengDu"
sheet.Range["A10"].Text = "Paris"
sheet.Range["A11"].Text = "Boston"
sheet.Range["A12"].Text = "London"
sheet.Range["C13"].Text = "City :"
sheet.Range["C13"].Style.Font.IsBold = True

listBox = sheet.ListBoxes.AddListBox(13, 4, 100, 80)
listBox.SelectionType = SelectionType.Single
listBox.SelectedIndex = 2
listBox.Display3DShading = True
listBox.ListFillRange = sheet.Range["A7:A12"]

workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()