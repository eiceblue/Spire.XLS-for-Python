from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/AddCustomProperties.xlsx"
outputFile = "AddCustomProperties.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Add a custom property ro make the document as final
workbook.CustomDocumentProperties.Add("_MarkAsFinal", True)
#Add other custom properties to the workbook
workbook.CustomDocumentProperties.Add("The Editor", "E-iceblue")
workbook.CustomDocumentProperties.Add("Phone number", 81705109)
workbook.CustomDocumentProperties.Add("Revision number", 7.12)
workbook.CustomDocumentProperties.Add("Revision date", DateTime.get_Now())
#Save the document and launch it
workbook.SaveToFile(outputFile, FileFormat.Version2013)
workbook.Dispose()

