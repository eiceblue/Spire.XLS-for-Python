from spire.common import *
from spire.xls import *


outputFile = "CopyDataWithStyle.xlsx"

#Create a workbook
workbook = Workbook()
#Get the default first worksheet
worksheet = workbook.Worksheets[0]

#Set the values for some cells.
cells = worksheet.Range["A1:J50"]
for i in range(1, 11):
    for j in range(1, 9):
       text = str(i - 1) + "," + str(j - 1)
       cells[i,j].Text = text

#Get a source range (A1:D3).
srcRange = worksheet.Range["A1:D3"]

#Create a style object.
style = workbook.Styles.Add("style")

#Specify the font attribute.
style.Font.FontName = "Calibri"

#Specify the shading color.
style.Font.Color = Color.get_Red()

#Specify the border attributes.
style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin
style.Borders[BordersLineType.EdgeTop].Color = Color.get_Blue()
style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin
style.Borders[BordersLineType.EdgeBottom].Color = Color.get_Blue()
style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin
style.Borders[BordersLineType.EdgeTop].Color = Color.get_Blue()
style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin
style.Borders[BordersLineType.EdgeRight].Color = Color.get_Blue()
srcRange.CellStyleName = style.Name

#Set the destination range
destRange = worksheet.Range["A12:D14"]

#Copy the range data with style
srcRange.Copy(destRange, True, True)

#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
