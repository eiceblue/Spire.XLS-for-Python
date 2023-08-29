import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/AccessCell.xlsx"
outputFile = "AccessCell.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
builder = []
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Access cell by its name
range1 = sheet.Range["A1"]
builder.append("Value of range1: " + range1.Text)
#Access cell by index of row and column
range2 = sheet.Range[2,1]
builder.append("Value of range2: " + range2.Text)
#Access cell in cell collection
range3 = sheet.Cells[2]
builder.append("Value of range3: " + range3.Text)
File.AppendAllText(outputFile, builder)

