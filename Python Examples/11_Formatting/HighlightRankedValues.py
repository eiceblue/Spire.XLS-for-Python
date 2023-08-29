from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_6.xlsx"
outputFile = "HighlightRankedValues.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Apply conditional formatting to range ��D2:D10�� to highlight the top 2 values.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(sheet.Range["D2:D10"])
format1 = xcfs.AddTopBottomCondition(TopBottomType.Top, 2)
format1.FormatType = ConditionalFormatType.TopBottom
format1.BackColor = Color.get_Red()
#Apply conditional formatting to range ��E2:E10�� to highlight the bottom 2 values.
xcfs1 = sheet.ConditionalFormats.Add()
xcfs1.AddRange(sheet.Range["E2:E10"])
format2 = xcfs1.AddTopBottomCondition(TopBottomType.Bottom, 2)
format2.FormatType = ConditionalFormatType.TopBottom
format2.BackColor = Color.get_ForestGreen()
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
