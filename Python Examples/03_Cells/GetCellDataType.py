from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_2.xlsx"
outputFile = "GetCellDataType.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the cell types of the cells in range "C13:F13”
for range in sheet.Range["H2:H7"].Cells:
    cellType = sheet.GetCellType(range.Row, range.Column, False)
    sheet[range.Row,range.Column + 1].Text = str(cellType)
    sheet[range.Row,range.Column + 1].Style.Font.Color = Color.get_Red()
    sheet[range.Row,range.Column + 1].Style.Font.IsBold = True
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


