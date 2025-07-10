from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "GetCellValueByCellName.txt"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Specify a cell by its name.
cell = sheet.Range["A2"]
content = []
#Get vaule of cell "A2".
content.append("The vaule of cell A2 is: " + cell.Value) #Save to file.
AppendAllText(outputFile, content)

