from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()
	

inputFile = "./Demos/Data/ExcelSample_N1.xlsx"
inputFile_97 = "./Demos/Data/ExcelSample97_N.xls"
inputFile_xml ="./Demos/Data/OfficeOpenXML_N.xml"
inputFile_csv = "./Demos/Data/CSVSample_N.csv"
outputFile = "OpenFiles.txt"

#Create string builder
builder = []
#1. Load file by file path
#Create a workbook
workbook1 = Workbook()
#Load the document from disk
workbook1.LoadFromFile(inputFile)
builder.append("Workbook opened using file path successfully!")
#2. Load file by file stream
stream = Stream(inputFile)
#Create a workbook
workbook2 = Workbook()
#Load the document from disk
workbook2.LoadFromStream(stream)
builder.append("Workbook opened using file stream successfully!")
stream.Dispose()
#3. Open Microsoft Excel 97 - 2003 file
wbExcel97 = Workbook()
wbExcel97.LoadFromFile(inputFile_97, ExcelVersion.Version97to2003)
builder.append("Microsoft Excel 97 - 2003 workbook opened successfully!")
#4. Open xml file
wbXML = Workbook()
wbXML.LoadFromXml(inputFile_xml)
builder.append("XML file opened successfully!")
#5. Open csv file
wbCSV = Workbook()
wbCSV.LoadFromFile(inputFile_csv, ",", 1, 1)
builder.append("CSV file opened successfully!")
#Save to txt file
AppendAllText(outputFile, builder)
