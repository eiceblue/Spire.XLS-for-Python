from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/Sample.xlsx"
outputFile = "VerifyDataByValidation.txt"

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
#Get the specified data range
minimum = Double.Parse(validation.Formula1)
maximum = Double.Parse(validation.Formula2)
#Create StringBuilder to save 
content = []
#Set different numbers for the cell
for i in range(5, 100, 40):
    cell.NumberValue = i
    result = None
    #Verify 
    if cell.NumberValue < minimum or cell.NumberValue > maximum:
        #Set string format for displaying
        result = "Is input " + str(i) + " a valid value for this Cell: false"
    else:
        #Set string format for displaying
        result = "Is input " + str(i) + " a valid value for this Cell: true"
    #Add result string to StringBuilder
    content.append(result)
#Save them to a txt file
AppendAllText(outputFile, content)
