import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ProtectedWorkbook.xlsx"
outputFile = "DetectProtection.txt"

value = Workbook.IsPasswordProtected(inputFile)
boolvalue = ""
if value:
    boolvalue = "Yes"
else:
    boolvalue = "No"
File.AppendText(outputFile, boolvalue)

