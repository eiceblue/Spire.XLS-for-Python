from spire.xls import *
from spire.common import *



outputFile = "ForegroundAndBackground.xlsx"


#Create a workbook
workbook = Workbook()
workbook.Version = ExcelVersion.Version2010
#Get the first sheet
sheet = workbook.Worksheets[0]
#Create a new style
style = workbook.Styles.Add("newStyle1")
#Set filling pattern type
style.Interior.FillPattern = ExcelPatternType.Gradient
#Set filling Background color
style.Interior.Gradient.BackKnownColor = ExcelColors.Green
#Set filling Foreground color
style.Interior.Gradient.ForeKnownColor = ExcelColors.Yellow
#set gradient style
style.Interior.Gradient.GradientStyle = GradientStyleType.From_Center
#Apply the style to  "B2" cell
sheet.Range["B2"].CellStyleName = style.Name
sheet.Range["B2"].Text = "Test"
sheet.Range["B2"].RowHeight = 30
sheet.Range["B2"].ColumnWidth = 50
#Create a new style
style = workbook.Styles.Add("newStyle2")
#Set filling pattern type
style.Interior.FillPattern = ExcelPatternType.Gradient
#Set filling Foreground color
style.Interior.Gradient.ForeKnownColor = ExcelColors.Red
#Apply the style to  "B4" cell
sheet.Range["B4"].CellStyleName = style.Name
sheet.Range["B4"].RowHeight = 30
sheet.Range["B4"].ColumnWidth = 60
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


