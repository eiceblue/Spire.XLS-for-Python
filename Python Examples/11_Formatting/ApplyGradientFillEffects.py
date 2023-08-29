from spire.xls import *
from spire.common import *


outputFile = "ApplyGradientFillEffects.xlsx"

#Create a workbook
workbook = Workbook()
workbook.Version = ExcelVersion.Version2010
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get "B5" cell
range = sheet.Range["B5"]
#Set row height and column width
range.RowHeight = 50
range.ColumnWidth = 30
range.Text = "Hello"
#Set alignment style
range.Style.HorizontalAlignment = HorizontalAlignType.Center
#Set gradient filling effects
range.Style.Interior.FillPattern = ExcelPatternType.Gradient
range.Style.Interior.Gradient.ForeColor = Color.FromRgb(255, 255, 255)
range.Style.Interior.Gradient.BackColor = Color.FromRgb(79, 129, 189)
range.Style.Interior.Gradient.TwoColorGradient(GradientStyleType.Horizontal, GradientVariantsType.ShadingVariants1)
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

