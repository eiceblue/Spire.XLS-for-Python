from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_6.xlsx"
outputFile = "HighlightDuplicateUniqueValues.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Use conditional formatting to highlight duplicate values in range "C2:C10" with IndianRed color.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(sheet.Range["C2:C10"])
format1 = xcfs.AddCondition()
format1.FormatType = ConditionalFormatType.DuplicateValues
format1.BackColor = Color.get_IndianRed()
#Use conditional formatting to highlight unique values in range "C2:C10" with Yellow color.
xcfs1 = sheet.ConditionalFormats.Add()
xcfs1.AddRange(sheet.Range["C2:C10"])
format2 = xcfs.AddCondition()
format2.FormatType = ConditionalFormatType.UniqueValues
format2.BackColor = Color.get_Yellow()
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

