from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SparkLine.xlsx"
outputFile = "SparkLine.xlsx"

#Load a Workbook from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Add sparkline
sparklineGroup = sheet.SparklineGroups.AddGroup(SparklineType.Line)
sparklines = sparklineGroup.Add()
sparklines.Add(sheet["A2:D2"], sheet["E2"])
sparklines.Add(sheet["A3:D3"], sheet["E3"])
sparklines.Add(sheet["A4:D4"], sheet["E4"])
sparklines.Add(sheet["A5:D5"], sheet["E5"])
sparklines.Add(sheet["A6:D6"], sheet["E6"])
sparklines.Add(sheet["A7:D7"], sheet["E7"])
sparklines.Add(sheet["A8:D8"], sheet["E8"])
sparklines.Add(sheet["A9:D9"], sheet["E9"])
sparklines.Add(sheet["A10:D10"], sheet["E10"])
sparklines.Add(sheet["A11:D11"], sheet["E11"])
#Save and Launch
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

