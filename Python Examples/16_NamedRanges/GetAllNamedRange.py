import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/AllNamedRanges.xlsx"
outputFile = "GetAllNamedRange.txt"

sb = []
#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get all named range
ranges = workbook.NameRanges
for nameRange in ranges:
    sb.append(nameRange.Name )
#Save and launch result file
File.AppendAllText(outputFile, sb)
workbook.Dispose()

