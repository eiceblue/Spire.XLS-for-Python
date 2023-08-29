from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WorksheetSample1.xlsx"
outputFile = "ApplyStyleToWorksheet.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Create a cell style
style = workbook.Styles.Add("newStyle")
style.Color = Color.get_LightBlue()
style.Font.Color = Color.get_White()
style.Font.Size = 15
style.Font.IsBold = True
#Apply the style to the first worksheet
sheet.ApplyStyle(style)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

