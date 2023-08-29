import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ExcelSample_N1.xlsx"
outputFile = "GetExcelVersion.txt"

builder = []
#Create a workbook
workbook = Workbook()
#Load the document
workbook.LoadFromFile(inputFile)
#Get the version
version = workbook.Version
builder.append(str(version))
#Save to file
File.AppendAllText(outputFile, builder)
workbook.Dispose()

