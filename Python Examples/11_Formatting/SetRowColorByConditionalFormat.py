from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "SetRowColorByConditionalFormat.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Select the range that you want to format.
dataRange = sheet.AllocatedRange
#Set conditional formatting.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(dataRange)
format1 = xcfs.AddCondition()
#Determines the cells to format.
format1.FirstFormula = "=MOD(ROW(),2)=0"
#Set conditional formatting type
format1.FormatType = ConditionalFormatType.Formula
#Set the color.
format1.BackColor = Color.get_LightSeaGreen()
#Set the backcolor of the odd rows as Yellow.
xcfs1 = sheet.ConditionalFormats.Add()
xcfs1.AddRange(dataRange)
format2 = xcfs.AddCondition()
format2.FirstFormula = "=MOD(ROW(),2)=1"
format2.FormatType = ConditionalFormatType.Formula
format2.BackColor = Color.get_Yellow()
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


