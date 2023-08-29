import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *



inputFile = "./Demos/Data/FindCellsSample.xlsx"
outputFile =  "FindStringAndNumber.txt"

#Create a workbook
workbook = Workbook()

#Load the document from disk
workbook.LoadFromFile(inputFile)

#Get the first worksheet
sheet = workbook.Worksheets[0]

#Find cells with the input string
textRanges = sheet.FindAllString("E-iceblue", False, False)

#Create a string builder
builder = []

#Append the address of found cells in builder
if len(textRanges) != 0:
    for range in textRanges:
        address = range.RangeAddress
        builder.append("address of found text cell is: " + address)
else:
    builder.append("No cells that contain the text")

#Find cells with the input integer or double
numberRanges = sheet.FindAllNumber(100, True)

#Append the address of found cells in builder
if len(numberRanges) != 0:
    for range in numberRanges:
        address = range.RangeAddress
        builder.append("The address of found number cell is: " + address)
else:
    builder.append("No cells that contain the number")
    
File.AppendAllText(outputFile, builder)
