from spire.xls.common import *
from spire.xls import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "AddTableWithFilter.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
 #Create a List Object named in Table.
sheet.ListObjects.Create("Table", sheet.Range[1,1,sheet.LastRow,sheet.LastColumn])
#Set the BuiltInTableStyle for List object.
sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleLight9
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()