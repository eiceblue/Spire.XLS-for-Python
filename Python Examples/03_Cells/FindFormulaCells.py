from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/FindCellsSample.xlsx"
outputFile = "FindFormulaCells.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Find the cells that contain formula "=SUM(A11,A12)"
ranges = sheet.FindAll("=SUM(A11,A12)", FindType.Formula, ExcelFindOptions.none)
#Create a string builder
builder = []
#Append the address of found cells to builder
if len(ranges) != 0:
    for range in ranges:
        address = range.RangeAddress
        builder.append("The address of found cell is: " + address)
else:
    builder.append("No cell contain the formula")
AppendAllText(outputFile, builder)

