from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ToPDF.xlsx"
outputFile = "ToPDF.pdf"

#create a workbook
workbook = Workbook()
#load a excel document
workbook.LoadFromFile(inputFile)
workbook.ConverterSetting.SheetFitToPage = True
#convert to PDF file
workbook.SaveToFile(outputFile, FileFormat.PDF)
workbook.Dispose()

