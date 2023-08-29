import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/AllNamedRanges.xlsx"
outputFile = "GetSpecificNamedRange.txt"

sb = []
#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get specific named range by index
name1 = workbook.NameRanges[1].Name
sb.append("Get the specific named range " + name1 + " by index")
#Get specific named range by name
name2 = workbook.NameRanges["NameRange3"].Name
sb.append("Get the specific named range " + name2 + " by name")
#Save and launch result file
File.AppendAllText(outputFile, sb)
workbook.Dispose()

