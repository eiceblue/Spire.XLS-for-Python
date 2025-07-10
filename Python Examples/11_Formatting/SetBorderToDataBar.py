from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SetBorderToDataBar.xlsx"
outputFile = "SetBorderToDataBar_out.xlsx"

#create a workbook
workbook = Workbook()

#load an excel document
workbook.LoadFromFile(inputFile)

#get the first worksheet
sheet=workbook.Worksheets[0]

#get the databar format 
xcfs = sheet.ConditionalFormats[0]
cf = xcfs[0]
dataBar1 = cf.DataBar
dataBar1.BarBorder.Type = DataBarBorderType.DataBarBorderSolid
dataBar1.BarBorder.Color = Color.get_Red()

#set to new data bar
sheet["E1"].NumberValue = 200
xcfs2 = sheet.ConditionalFormats.Add()
xcfs2.AddRange(sheet.Range["E1"])
cf2 = xcfs2.AddCondition()
cf2.FormatType = ConditionalFormatType.DataBar
cf2.DataBar.BarBorder.Type = DataBarBorderType.DataBarBorderSolid
cf2.DataBar.BarBorder.Color = Color.get_Red()
cf2.DataBar.BarColor = Color.get_GreenYellow()

#save the file
workbook.SaveToFile(outputFile, FileFormat.Version2013)
workbook.Dispose()


