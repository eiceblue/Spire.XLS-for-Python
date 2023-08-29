from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WorksheetSample1.xlsx"
outputFile = "SplitWorksheetIntoPanes.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Vertical and horizontal split the worksheet into four panes
sheet.FirstVisibleColumn = 2
sheet.FirstVisibleRow = 5
sheet.VerticalSplit = 4000
sheet.HorizontalSplit = 5000
#Set the active pane
sheet.ActivePane = 1
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

