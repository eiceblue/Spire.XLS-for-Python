from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/Sample.xlsx"
outputFile = "GetSettingsOfDataValidation.txt"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Cell B4 has the Decimal Validation
cell = worksheet.Range["B4"]
#Get the valditation of this cell
validation = cell.DataValidation
#Get the settings
allowType = str(validation.AllowType)
data = str(validation.CompareOperator)
minimum = str(validation.Formula1)
maximum = str(validation.Formula2)
ignoreBlank = str(validation.IgnoreBlank)
#Create StringBuilder to save 
content = []
#Set string format for displaying
result = "Settings of Validation: \r\nAllow Type: " + allowType + "\r\nData: " + data + "\r\nMinimum: " + minimum + "\r\nMaximum: " + maximum + "\r\nIgnoreBlank: " + ignoreBlank
#Add result string to StringBuilder
content.append(result)
#Save them to a txt file
AppendAllText(outputFile, content)

