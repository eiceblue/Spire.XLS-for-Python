from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/EachWorksheetToDifferentPDFSample.xlsx"

#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
for sheet in workbook.Worksheets:
    FileName =  sheet.Name + ".pdf"
    #Save the sheet to PDF
    sheet.SaveToPdf(FileName)
workbook.Dispose()

