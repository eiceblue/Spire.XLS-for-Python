from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SampleB_2.xlsx"
outputFile = "FitWidthWhenConvertToPDF.pdf"

#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
for sheet in workbook.Worksheets:
    #Auto fit page height
    sheet.PageSetup.FitToPagesTall = 0
    #Fit one page width
    sheet.PageSetup.FitToPagesWide = 1
workbook.SaveToFile(outputFile, FileFormat.PDF)
workbook.Dispose()


