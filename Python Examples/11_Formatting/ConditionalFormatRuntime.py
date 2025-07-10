from spire.xls import *
from spire.xls.common import *


def AddComparisonRule1( sheet):
    #Create conditional formatting rule
    xcfs1 = sheet.ConditionalFormats.Add()
    xcfs1.AddRange(sheet.Range["A1:D1"])
    cf1 = xcfs1.AddCondition()
    cf1.FormatType = ConditionalFormatType.CellValue
    cf1.FirstFormula = "150"
    cf1.Operator = ComparisonOperatorType.Greater
    cf1.FontColor = Color.get_Red()
    cf1.BackColor = Color.get_LightBlue()
def AddComparisonRule2( sheet):
    xcfs2 = sheet.ConditionalFormats.Add()
    xcfs2.AddRange(sheet.Range["A2:D2"])
    cf2 = xcfs2.AddCondition()
    cf2.FormatType = ConditionalFormatType.CellValue
    cf2.FirstFormula = "500"
    cf2.Operator = ComparisonOperatorType.Less
    #Set border color
    cf2.LeftBorderColor = Color.get_Pink()
    cf2.RightBorderColor = Color.get_Pink()
    cf2.TopBorderColor = Color.get_DeepSkyBlue()
    cf2.BottomBorderColor = Color.get_DeepSkyBlue()
    cf2.LeftBorderStyle = LineStyleType.Medium
    cf2.RightBorderStyle = LineStyleType.Thick
    cf2.TopBorderStyle = LineStyleType.Double
    cf2.BottomBorderStyle = LineStyleType.Double
def AddComparisonRule3( sheet):
    #Create conditional formatting rule
    xcfs1 = sheet.ConditionalFormats.Add()
    xcfs1.AddRange(sheet.Range["A3:D3"])
    cf1 = xcfs1.AddCondition()
    cf1.FormatType = ConditionalFormatType.CellValue
    cf1.FirstFormula = "300"
    cf1.SecondFormula = "500"
    cf1.Operator = ComparisonOperatorType.Between
    cf1.BackColor = Color.get_Yellow()
def AddComparisonRule4( sheet):
    #Create conditional formatting rule
    xcfs1 = sheet.ConditionalFormats.Add()
    xcfs1.AddRange(sheet.Range["A4:D4"])
    cf1 = xcfs1.AddCondition()
    cf1.FormatType = ConditionalFormatType.CellValue
    cf1.FirstFormula = "100"
    cf1.SecondFormula = "200"
    cf1.Operator = ComparisonOperatorType.NotBetween
    #Set fill pattern type
    cf1.FillPattern = ExcelPatternType.ReverseDiagonalStripe
    #Set foreground color
    cf1.Color = Color.FromRgb(255, 255, 0)
    #Set background color
    cf1.BackColor = Color.FromRgb(0, 255, 255)

inputFile = "./Demos/Data/ConditionalFormatRuntime.xlsx"
outputFile = "ConditionalFormatRuntime.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
AddComparisonRule1(sheet)
AddComparisonRule2(sheet)
AddComparisonRule3(sheet)
AddComparisonRule4(sheet)
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

