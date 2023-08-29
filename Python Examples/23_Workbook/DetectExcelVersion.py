import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


outputFile = "DetectExcelVersion.txt"

#Files
files = ["./Demos/Data/ExcelSample97_N.xls", "./Demos/Data/ExcelSample_N1.xlsx", "./Demos/Data/ExcelSample_N.xlsb"]
builder = []
for file in files:
    #Create a workbook
    workbook = Workbook()
    #Load the document
    workbook.LoadFromFile(file)
    #Get the version
    version = workbook.Version
    builder.append(str(version))
#Save to txt file
File.AppendAllText(outputFile, builder)


