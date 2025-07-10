from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()
	
inputFile = "./Demos/Data/WorksheetSample3.xlsx"
outputFile = "OutputGetWorksheetNames.txt"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the names of all worksheets
sb = []
for sheet in workbook.Worksheets:
    sb.append(sheet.Name)
#Save the documen
AppendAllText(outputFile, sb)
workbook.Dispose()

