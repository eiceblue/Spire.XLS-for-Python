from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/GetFreezePaneRange.xlsx"
outputFile = "GetFreezePaneRange.txt"

#Create a workbook and load a file
wb = Workbook()
wb.LoadFromFile(inputFile)
sheet = wb.Worksheets[0]
rowIndex = None
colIndex = None
#The row and column index of the frozen pane is passed through the out parameter. 
#If it returns to 0, it means that it is not frozen
indexs = sheet.GetFreezePanes()
colIndex = indexs[1]
rowIndex = indexs[0]
r = "Row index: " + str(rowIndex) + ", column index: " + str(colIndex)
#Save the document and launch it
AppendAllText(outputFile, r)
wb.Dispose()

