import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ConversionSample2.xlsx"
outputFile = "ExceltoTxt.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet in excel workbook
sheet = workbook.Worksheets[0]
sheet.SaveToFile(outputFile, " ", Encoding.get_UTF8())
workbook.Dispose()

