from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/OfficeOpenXMLToExcel.Xml"
outputFile = "OfficeOpenXMLToExcel.xlsx"

workbook = Workbook()
#Initialize worksheet
fileStream = Stream(inputFile)
workbook.LoadFromXml(fileStream)
fileStream.Close()
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()



