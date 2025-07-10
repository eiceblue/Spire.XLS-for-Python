from spire.xls import *
from spire.xls.common import *


def AddConditionalFormattingForExistingSheet( sheet):
    sheet.AllocatedRange.RowHeight = 15
    sheet.AllocatedRange.ColumnWidth = 16
    #Create conditional formatting rule
    xcfs1 = sheet.ConditionalFormats.Add()
    xcfs1.AddRange(sheet.Range["A1:D1"])
    cf1 = xcfs1.AddCondition()
    cf1.FormatType = ConditionalFormatType.CellValue
    cf1.FirstFormula = "150"
    cf1.Operator = ComparisonOperatorType.Greater
    cf1.FontColor = Color.get_Red()
    cf1.BackColor = Color.get_LightBlue()
    xcfs2 = sheet.ConditionalFormats.Add()
    xcfs2.AddRange(sheet.Range["A2:D2"])
    cf2 = xcfs2.AddCondition()
    cf2.FormatType = ConditionalFormatType.CellValue
    cf2.FirstFormula = "300"
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
    #Add data bars
    xcfs3 = sheet.ConditionalFormats.Add()
    xcfs3.AddRange(sheet.Range["A3:D3"])
    cf3 = xcfs3.AddCondition()
    cf3.FormatType = ConditionalFormatType.DataBar
    cf3.DataBar.BarColor = Color.get_CadetBlue()
    #Add icon sets
    xcfs4 = sheet.ConditionalFormats.Add()
    xcfs4.AddRange(sheet.Range["A4:D4"])
    cf4 = xcfs4.AddCondition()
    cf4.FormatType = ConditionalFormatType.IconSet
    cf4.IconSet.IconSetType = IconSetType.ThreeTrafficLights1
    #Add color scales
    xcfs5 = sheet.ConditionalFormats.Add()
    xcfs5.AddRange(sheet.Range["A5:D5"])
    cf5 = xcfs5.AddCondition()
    cf5.FormatType = ConditionalFormatType.ColorScale
    #Highlight duplicate values in range "A6:D6" with BurlyWood color
    xcfs6 = sheet.ConditionalFormats.Add()
    xcfs6.AddRange(sheet.Range["A6:D6"])
    cf6 = xcfs6.AddCondition()
    cf6.FormatType = ConditionalFormatType.DuplicateValues
    cf6.BackColor = Color.get_BurlyWood()


inputFile = "./Demos/Data/ConditionalFormatting.xlsx"
outputFile = "SimpleConditionalFormatting.xlsx"

#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
oldSheet = workbook.Worksheets[0]
AddConditionalFormattingForExistingSheet(oldSheet)
result = "SimpleConditionalFormatting_result.xlsx"
#Save and Launch
workbook.SaveToFile(result, ExcelVersion.Version2010)
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


