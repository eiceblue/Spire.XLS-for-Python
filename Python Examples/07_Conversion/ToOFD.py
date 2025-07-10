from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorksheetSample1.xlsx"
outputFile = "ToOFD.ofd"

#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
#convert to OFD file
workbook.SaveToFile(outputFile, FileFormat.OFD)
workbook.Dispose()
