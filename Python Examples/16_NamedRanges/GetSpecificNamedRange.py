from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/AllNamedRanges.xlsx"
outputFile = "GetSpecificNamedRange.txt"

sb = []
#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get specific named range by index
name1 = workbook.NameRanges[1].Name
sb.append("Get the specific named range " + name1 + " by index")
#Get specific named range by name
name2 = workbook.NameRanges["NameRange3"].Name
sb.append("Get the specific named range " + name2 + " by name")
#Save and launch result file
AppendAllText(outputFile, sb)
workbook.Dispose()

