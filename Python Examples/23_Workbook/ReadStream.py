from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ReadStream.xlsx"
outputFile = "ReadStream.xlsx"

workbook = Workbook()
#Open excel from a stream
fileStream = Stream(inputFile)
workbook.LoadFromStream(fileStream)
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
