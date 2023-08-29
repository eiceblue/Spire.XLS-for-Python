import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/HyperlinksSample2.xlsx"
outputFile = "GetHyperLinkType.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Iterate all hyperlinks
sb = []
for item in sheet.HyperLinks:
    #Get hyperlink address
    address = item.Address
    #Get hyperlink type
    type = item.Type
    sb.append("Link address: " + address)
    sb.append("Link type: " + str(type))
    sb.append("")
File.AppendAllText(outputFile, sb)
workbook.Dispose()

