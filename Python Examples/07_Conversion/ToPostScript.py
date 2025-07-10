from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ToPostScript.xlsx"
outputFile = "ToPostScript.ps"

#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
workbook.SaveToFile(outputFile, FileFormat.PostScript)
workbook.Dispose()


