from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SampleB_2.xlsx"
outputFile = "CutCellsToOtherPosition.xlsx"

#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
Ori = sheet.Range["A1:C5"]
Dest = sheet.Range["A26:C30"]
#Copy the range to other position
sheet.Copy(Ori, Dest, True, True, True)
#Remove all content in original cells
for cr in Ori.Cells:
    cr.ClearAll()
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


