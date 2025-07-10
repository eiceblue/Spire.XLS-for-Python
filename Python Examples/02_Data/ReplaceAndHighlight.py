from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ReplaceAndHighlight.xlsx"
outputFile = "ReplaceAndHighlight.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
worksheet = workbook.Worksheets[0]
ranges = worksheet.FindAllString("Total", True, True)

for range in ranges:
    #reset the text, in other words, replace the text
    range.Text = "Sum"
    #set the color
    range.Style.Color = Color.get_Yellow()

workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

