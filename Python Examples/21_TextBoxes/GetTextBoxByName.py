import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *

outputFile = "GetTextBoxByName.txt"

#Create a workbook
workbook = Workbook()
#Get the default first  worksheet
sheet = workbook.Worksheets[0]
#Insert a TextBox
sheet.Range["A2"].Text = "Name："
textBox = sheet.TextBoxes.AddTextBox(2, 2, 18, 65)
#Set the name 
textBox.Name = "FirstTextBox"
#Set string text for TextBox 
textBox.Text = "Spire.XLS for Python  is a professional Excel Python API that can be used to create, read, write and convert Excel files in any type of python application. Spire.XLS for Python offers object model Excel API for speeding up Excel programming in python platform - create new Excel documents from template, edit existing Excel documents and convert Excel files."
#Get the TextBox by the name
FindTextBox = sheet.TextBoxes["FirstTextBox"]
#Get the TextBox text 
text = FindTextBox.Text
#Create StringBuilder to save 
content = []
#Set string format for displaying
result = "The text of \"" + textBox.Name + "\" is :" + text
#Add result string to StringBuilder
content.append(result)
File.AppendAllText(outputFile, content)

