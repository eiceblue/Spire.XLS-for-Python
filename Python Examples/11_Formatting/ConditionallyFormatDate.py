from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_6.xlsx"
outputFile = "ConditionallyFormatDate.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Highlight cells that contain a date occurring in the last 7 days.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(sheet.AllocatedRange)
conditionalFormat = xcfs.AddTimePeriodCondition(TimePeriodType.Last7Days)
conditionalFormat.BackColor = Color.get_Orange()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

