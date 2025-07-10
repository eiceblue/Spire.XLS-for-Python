from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "CountNumberOfCells.txt"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
content = []
#Get the number of cells.
content.append("Number of Cells: " + str(sheet.Cells.Length))
#Save to file.
AppendAllText(outputFile, content)

