import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/AllNamedRanges.xlsx"
outputFile = "GetNamedRangeAddress.txt"

sb = []
#Create a workbook and load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get specific named range by index
NamedRange = workbook.NameRanges[0]
#Get the address of the named range
address = NamedRange.RefersToRange.RangeAddress
sb.append("The address of the named range " + NamedRange.Name + " is " + address)
#Save and launch result file
File.AppendAllText(outputFile, sb)
workbook.Dispose()

