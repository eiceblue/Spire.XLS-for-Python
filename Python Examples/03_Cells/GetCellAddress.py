import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ExcelSample_N1.xlsx"
outputFile = "GetCellAddress.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
builder = []
#Get a cell range
range = sheet.Range["A1:B5"]
#Get address of range
address = range.RangeAddressLocal
builder.append("Address of range: " + address)
#Get the cell count of range
count = range.CellsCount
builder.append("Cell count of range: " + str(count))
#Get the address of the entire column of range
entireColAddress = range.EntireColumn.RangeAddressLocal
builder.append("Address of entire column of the range: " + entireColAddress)
#Get the address of the entire row of range
entireRowAddress = range.EntireRow.RangeAddressLocal
builder.append("Address of entire row of the range " + entireRowAddress)
File.AppendAllText(outputFile, builder)


