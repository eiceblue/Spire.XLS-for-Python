from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ToPDF_A1BExample.xlsx"
outputFile = "ToPDFA1B.pdf"

#Create a workbook
workbook = Workbook()
#Load an excel file
workbook.LoadFromFile(inputFile)
#Convert excel to PDFA/1-B
workbook.ConverterSetting.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B
workbook.SaveToFile(outputFile, FileFormat.PDF)
workbook.Dispose()

