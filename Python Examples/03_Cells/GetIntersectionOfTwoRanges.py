import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


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
File.AppendAllText(outputFile, content)


