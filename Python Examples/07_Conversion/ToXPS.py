from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ToXPS.xlsx"
outputFile = "ToXPS.xps"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#convert to XPS file
workbook.SaveToFile(outputFile, FileFormat.XPS)
workbook.Dispose()
