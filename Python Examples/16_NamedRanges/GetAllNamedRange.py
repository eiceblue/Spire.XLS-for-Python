from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/AllNamedRanges.xlsx"
outputFile = "GetAllNamedRange.txt"

sb = []
#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get all named range
ranges = workbook.NameRanges
for nameRange in ranges:
    sb.append(nameRange.Name )
#Save and launch result file
AppendAllText(outputFile, sb)
workbook.Dispose()

