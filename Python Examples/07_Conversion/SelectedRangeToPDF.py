from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ConversionSample1.xlsx"
outputFile = "SelectedRangeToPDF.pdf"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Add a new sheet to workbook
workbook.Worksheets.Add("newsheet")
#Copy your area to new sheet.
workbook.Worksheets[0].Range["A9:E15"].Copy(workbook.Worksheets[1].Range["A9:E15"], False, True)
#Auto fit column width
workbook.Worksheets[1].Range["A9:E15"].AutoFitColumns()
workbook.Worksheets[1].SaveToPdf(outputFile)
workbook.Dispose()

