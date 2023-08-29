import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ReadHyperlinks.xlsx"
outputFile = "ReadHyperlinks.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
address1 = sheet.HyperLinks[0].Address
address2 = sheet.HyperLinks[1].Address
File.AppendText(outputFile, address1 + "\r\n" + address2)
workbook.Dispose()

