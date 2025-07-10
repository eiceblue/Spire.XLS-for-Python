from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/ReadImages.xlsx"
outputFile = "DetectEmptyWorksheet.txt"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
worksheet1 = workbook.Worksheets[0]
#Detect the first worksheet is empty or not
detect1 = worksheet1.IsEmpty
#Get the second worksheet
worksheet2 = workbook.Worksheets[1]
#Detect the second worksheet is empty or not
detect2 = worksheet2.IsEmpty
#Create StringBuilder to save 
content = []
#Set string format for displaying
result = "The first worksheet is empty or not: " + str(detect1) + "\r\nThe second worksheet is empty or not: " + str(detect2)
#Add result string to StringBuilder
content.append(result)
#Save the document
AppendAllText(outputFile, content)
workbook.Dispose()
