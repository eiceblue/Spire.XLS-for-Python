from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_6.xlsx"
outputFile = "HighlightAverageValues.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Add conditional format.
format1 = sheet.ConditionalFormats.Add()
#Set the cell range to apply the formatting.
format1.AddRange(sheet.Range["E2:E10"])
#Add below average condition.
cf1 = format1.AddAverageCondition(AverageType.Below)
#Highlight cells below average values.
cf1.BackColor = Color.get_SkyBlue()
#Add conditional format.
format2 = sheet.ConditionalFormats.Add()
#Set the cell range to apply the formatting.
format2.AddRange(sheet.Range["E2:E10"])
#Add above average condition.
cf2 = format1.AddAverageCondition(AverageType.Above)
#Highlight cells above average values.
cf2.BackColor = Color.get_Orange()
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

