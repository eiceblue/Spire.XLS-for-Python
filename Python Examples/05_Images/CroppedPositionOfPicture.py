from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/ReadImages.xlsx"
outputFile = "CroppedPositionOfPicture.txt"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet1 = workbook.Worksheets[0]
#Get the image from the first sheet
picture = sheet1.Pictures[0]
#Get the cropped position
left = picture.Left
top = picture.Top
width = picture.Width
height = picture.Height
#Create StringBuilder to save 
content = []
#Set string format for displaying
displayString = "Crop position: Left " + str(left) + "\r\nCrop position: Top " + str(top) + "\r\nCrop position: Width " + str(width) + "\r\nCrop position: Height " + str(height)
#Add result string to StringBuilder
content.append(displayString)
#Save them to a txt file
AppendAllText(outputFile, content)

