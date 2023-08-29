from spire.xls import *
from spire.common import *


outputFile = "RemoveBorderlineOfTextbox.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.Version = ExcelVersion.Version2013
#Create a new worksheet named "Remove Borderline" and add a chart to the worksheet.
sheet = workbook.Worksheets[0]
sheet.Name = "Remove Borderline"
chart = sheet.Charts.Add()
#Create textbox1 in the chart and input text information.
textbox1 = chart.TextBoxes.AddTextBox(50, 50, 100, 600)
textbox1.Text = "The solution with borderline"
#Create textbox2 in the chart, input text information and remove borderline.
textbox2 = chart.TextBoxes.AddTextBox(1000, 50, 100, 600)
textbox2.Text = "The solution without borderline"
textbox2.Line.Weight = 0
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

