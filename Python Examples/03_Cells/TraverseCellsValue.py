from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/CellValues.xlsx"
outputFile = "TraverseCellsValue.txt"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Get the cell range collection 
cellRangeCollection = worksheet.Cells
#Create StringBuilder to save 
content = []
content.append("Values of the first sheet:")
#Traverse cells value
for cellRange in cellRangeCollection:
    #Set string format for displaying
    result = "Cell: " + cellRange.RangeAddress + "   Value: " + cellRange.Value
    #Add result string to StringBuilder
    content.append(result)
#Save them to a txt file
AppendAllText(outputFile, content)

