from spire.xls import *
from spire.common import *



def ReplaceTextInTextBox(sheet, sFind, sReplace):
    for tb in sheet.TextBoxes:
        if tb.Text != "":
            if tb.Text.__contains__(sFind):
                tb.Text = tb.Text.replace(sFind, sReplace)

inputFile = "./Demos/Data/ReplaceTextInTextBox.xlsx"
outputFile = "ReplaceTextInTextBox.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
tag = "TAG_1$TAG_2"
replace = "Spire.XLS for .NET$Spire.XLS for JAVA"
i = 0
while i < len(tag.split('$')):
    #Replace text in textbox
    ReplaceTextInTextBox(sheet, "<" + tag.split('$')[i] + ">", replace.split('$')[i])
    i += 1
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

    
    
    
    
    

