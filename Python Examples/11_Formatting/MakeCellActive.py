from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/templateAz.xlsx"
outputFile = "MakeCellActive.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the 2nd sheet
sheet = workbook.Worksheets[1]
#Set the 2nd sheet as an active sheet.
sheet.Activate()
#Set B2 cell as an active cell in the worksheet.
sheet.SetActiveCell(sheet.Range["B2"])
#Set the B column as the first visible column in the worksheet.
sheet.FirstVisibleColumn = 1
#Set the 2nd row as the first visible row in the worksheet.
sheet.FirstVisibleRow = 1
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

