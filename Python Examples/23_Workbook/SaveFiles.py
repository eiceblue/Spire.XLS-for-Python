from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ExcelSample_N1.xlsx"
outputFile_xls = "SaveFiles.xls"
outputFile_xlsx = "SaveFiles.xlsx"
outputFile_xlsb = "SaveFiles.xlsb"
outputFile_ods = "SaveFiles.ods"
outputFile_pdf = "SaveFiles.pdf"
outputFile_xml = "SaveFiles.xml"
outputFile_xps = "SaveFiles.xps"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
# Save in Excel 97-2003 format
workbook.SaveToFile(outputFile_xls, ExcelVersion.Version97to2003)
workbook.Dispose()


#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
# Save in Excel2010 xlsx format
workbook.SaveToFile(outputFile_xlsx, ExcelVersion.Version2010)
workbook.Dispose()


#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
# Save in XLSB format
workbook.SaveToFile(outputFile_xlsb, ExcelVersion.Xlsb2010)
workbook.Dispose()


#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
# Save in ODS format
workbook.SaveToFile(outputFile_ods, ExcelVersion.ODS)
workbook.Dispose()


#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
# Save in PDF format
workbook.SaveToFile(outputFile_pdf, FileFormat.PDF)
workbook.Dispose()


#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
# Save in XML format
workbook.SaveToFile(outputFile_xml, FileFormat.XML)
workbook.Dispose()


#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
# Save in XPS format
workbook.SaveToFile(outputFile_xps, FileFormat.XPS)
workbook.Dispose()

