from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()
	

outputFile = "DetectExcelVersion.txt"

#Files
files = ["./Demos/Data/ExcelSample97_N.xls", "./Demos/Data/ExcelSample_N1.xlsx", "./Demos/Data/ExcelSample_N.xlsb"]
builder = []
for file in files:
    #Create a workbook
    workbook = Workbook()
    #Load the document
    workbook.LoadFromFile(file)
    #Get the version
    version = workbook.Version
    builder.append(str(version))
#Save to txt file
AppendAllText(outputFile, builder)


