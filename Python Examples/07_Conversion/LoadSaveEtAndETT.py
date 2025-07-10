from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorksheetSample1.xlsx"
outputFile = "ToET.et"

#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
#convert to ET file
workbook.SaveToFile(outputFile, FileFormat.ET)
workbook.Dispose()
