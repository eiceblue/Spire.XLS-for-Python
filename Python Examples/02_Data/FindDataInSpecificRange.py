from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

def findTextFromRange(range, builder):
    #Find string from this range
    textRanges = range.FindAllString("E-iceblue", False, False)

    #Append the address of found cells in builder
    if len(textRanges) != 0:
        for r in textRanges:
            address = r.RangeAddress
            builder.append("The address of found text cell is: " + address)
    else:
        builder.append("No cell contain the text")
        

def findNumberFromRange(range, builder):
    #Find number from this range
    numberRanges = range.FindAllNumber(100, True)

    #Append the address of found cells in builder
    if len(numberRanges) != 0:
        for r in numberRanges:
            address = r.RangeAddress
            builder.append("The address of found number cell is: " + address)
    else:
        builder.append("No cell contain the number")



inputFile = "./Demos/Data/FindCellsSample.xlsx"
outputFile = "FindDataInSpecificRange.txt"


#Create a workbook
workbook = Workbook()

#Load the document from disk
workbook.LoadFromFile(inputFile)

#Get the first worksheet
sheet = workbook.Worksheets[0]

#Specify a range
range = sheet.Range[1,1,12,8]

#Create a string builder
builder = []


#Find text from this range
findTextFromRange(range, builder)

#Find number from this range
findNumberFromRange(range, builder)
AppendAllText(outputFile, builder)


