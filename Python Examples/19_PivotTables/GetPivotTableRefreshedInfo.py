from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/PivotTable.xlsx"
outputFile = "GetPivotTableRefreshedInfo.txt"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Get the first pivot table
pivotTable = worksheet.PivotTables[0]
#Get the refreshed information
dateTime = pivotTable.Cache.RefreshDate
refreshedBy = pivotTable.Cache.RefreshedBy
#Create StringBuilder to save 
sb = []
#Set string format for displaying
result = "Pivot table refreshed by:  " + refreshedBy + "\r\nPivot table refreshed date: " + str(dateTime)
#Add result string to StringBuilder
sb.append(result)
AppendAllText(outputFile, sb)
workbook.Dispose()
