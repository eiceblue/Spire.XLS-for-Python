from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/AddTextBox.xlsx"
outputFile = "AddTextBox.xlsx"

#Create a Workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first chart
chart = sheet.Charts[0]
#Add a Textbox
textbox = chart.Shapes.AddTextBox()
textbox.Width = 1200
textbox.Height = 320
textbox.Left = 1000
textbox.Top = 480
textbox.Text = "This is a textbox"
#Save and Launch
workbook.SaveToFile(outputFile, FileFormat.Version2010)
workbook.Dispose()

