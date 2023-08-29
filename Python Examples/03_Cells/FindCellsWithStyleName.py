from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SampleB_2.xlsx"
outputFile = "FindCellsWithStyleName.xlsx"

#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the cell style name
styleName = sheet.Range["A1"].CellStyleName
ranges = sheet.AllocatedRange
for cc in ranges.Cells:
    #Find the cells which have the same style name
    if cc.CellStyleName == styleName:
        #Set value
        cc.Value = "Same style"
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

