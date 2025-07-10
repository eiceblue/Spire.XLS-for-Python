from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()
	

inputFile = "./Demos/Data/ProtectedWorksheet.xlsx"
outputFile = "VerifyProtectedWorksheet.txt"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
worksheet = workbook.Worksheets[0]
#Verify the first worksheet 
detect = worksheet.IsPasswordProtected
#Create StringBuilder to save 
content = []
#Set string format for displaying
result = "The first worksheet is password protected or not: " + str(detect)
#Add result string to StringBuilder
content.append(result)
#Save the document
AppendAllText(outputFile, content)
workbook.Dispose()
