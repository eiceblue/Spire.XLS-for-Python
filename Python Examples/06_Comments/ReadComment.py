import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ReadComment.xls"
outputFile = "ReadComment.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
builder = []
builder.append(sheet.Range["A1"].Comment.Text+"\n\t")
builder.append(str(sheet.Range["A2"].Comment.RichText.RtfText))
File.AppendAllText(outputFile, builder)
workbook.Dispose()



