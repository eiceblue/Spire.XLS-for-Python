from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

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
AppendAllText(outputFile, builder)

