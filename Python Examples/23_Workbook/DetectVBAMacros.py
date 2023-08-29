import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *

inputFile = "./Demos/Data/MacroSample.xls"
outputFile = "DetectVBAMacros.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Detect if the Excel file contains VBA macros
value = ""
hasMacros = workbook.HasMacros
if hasMacros:
    value = "Yes"
else:
    value = "No"
File.AppendAllText(outputFile, value)

