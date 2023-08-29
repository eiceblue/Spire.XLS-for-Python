from spire.xls import *
from spire.common import *


outputFile = "SetDefaultRowAndColumnStyle.xlsx"

workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
#Create a cell style and set the color
style = workbook.Styles.Add("Mystyle")
style.Color = Color.get_Yellow()
#Set the default style for the first row and column 
sheet.SetDefaultRowStyle(1, style)
sheet.SetDefaultColumnStyle(1, style)
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()




