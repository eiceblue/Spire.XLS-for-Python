from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SaveStream.xls"
outputFile = "SaveStream.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Save an excel workbook to stream
fileStream = Stream(outputFile)
workbook.SaveToStream(fileStream, FileFormat.Version2010)
fileStream.Close()
workbook.Dispose()

