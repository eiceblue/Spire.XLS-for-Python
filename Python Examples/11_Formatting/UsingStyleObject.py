from spire.xls import *
from spire.common import *


outputFile = "UsingStyleObject.xlsx"

#Create a workbook
workbook = Workbook()
#Add a new worksheet to the Excel object
sheet = workbook.Worksheets.Add("new sheet")
#Access the "B1" cell from the worksheet
cell = sheet.Range["B1"]
#Add some value to the "B1" cell
cell.Text = "Hello Spire!"
#Create a new style
style = workbook.Styles.Add("newStyle")
#Set the vertical alignment of the text in the "B1" cell
style.VerticalAlignment = VerticalAlignType.Center
#Set the horizontal alignment of the text in the "B1" cell
style.HorizontalAlignment = HorizontalAlignType.Center
#Set the font color of the text in the "B1" cell
style.Font.Color = Color.get_Blue()
#Shrink the text to fit in the cell
style.ShrinkToFit = True
#Set the bottom border color of the cell to GreenYellow
style.Borders[BordersLineType.EdgeBottom].Color = Color.get_GreenYellow()
#Set the bottom border type of the cell to Medium
style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Medium
#Assign the Style object to the "B1" cell
cell.Style = style
#Apply the same style to some other cells
sheet.Range["B4"].Style = style
sheet.Range["B4"].Text = "Test"
sheet.Range["C3"].CellStyleName = style.Name
sheet.Range["C3"].Text = "Welcome to use Spire.XLS"
sheet.Range["D4"].Style = style
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


