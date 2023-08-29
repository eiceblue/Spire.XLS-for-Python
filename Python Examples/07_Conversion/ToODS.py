from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ToODS.xlsx"
outputFile = "ToODS.ods"

#create a workbook
workbook = Workbook()
#load a excel document
workbook.LoadFromFile(inputFile)
#convert to ODS file
workbook.SaveToFile(outputFile, FileFormat.ODS)
workbook.Dispose()


