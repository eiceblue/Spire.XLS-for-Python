from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ToPDF.xlsx"
outputFile = "ToPdfSimply.pdf"

#Create a workbook
workbook = Workbook()
#Load a excel document
workbook.LoadFromFile(inputFile)
#Convert excel to pdf
workbook.SaveToFile(outputFile, FileFormat.PDF)
workbook.Dispose()

