from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SampleB_2.xlsx"
outputFile = "ToPdfWithChangePageSize.xlsx"

#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
for sheet in workbook.Worksheets:
    #Change the page size
    sheet.PageSetup.PaperSize = PaperSizeType.PaperA3
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


