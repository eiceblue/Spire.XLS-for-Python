from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "GetIntersectionOfTwoRanges.txt"


#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the two ranges.
range = sheet.Range["A2:D7"].Intersect(sheet.Range["B2:E8"])
content = []
content.append("The intersection of the two ranges \"A2:D7\" and \"B2:E8\" is:")
#Get the intersection of the two ranges.
for r in range.Cells:
    content.append(str(r.Value))
#Save to file.
AppendAllText(outputFile, content)


