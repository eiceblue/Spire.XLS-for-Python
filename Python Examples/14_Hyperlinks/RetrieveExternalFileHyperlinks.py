import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/RetrieveExternalFileHyperlinks.xlsx"
outputFile = "RetrieveExternalFileHyperlinks.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
content = []
#Retrieve external file hyperlinks.
for item in sheet.HyperLinks:
    address = item.Address
    sheetName = item.Range.WorksheetName
    range = item.Range
    content.append("Cell[{0},{1}] in sheet \"" + sheetName + "\" contains File URL: {2}".format(range.Row, range.Column, address))
File.AppendAllText(outputFile, content)
#Save to file
#workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

