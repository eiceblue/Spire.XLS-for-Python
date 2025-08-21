# Python代码核心功能提取

# Spire.XLS Python Create Excel with Multiple Sheets
## Create an Excel workbook with five sheets
```python
workbook = Workbook()
workbook.CreateEmptySheets(5)
for i in range(0, 5):
    sheet = workbook.Worksheets[i]
    sheet.Name = "Sheet" + str(i)
```

---

# Create Excel with One Sheet
## Create a workbook with one sheet and fill data with row and column numbers
```python
workbook = Workbook()
workbook.CreateEmptySheets(1)
sheet = workbook.Worksheets[0]
for row in range(1, 100):
    for col in range(1, 31):
        sheet.Range[row,col].Text = str(row) + "," + str(col)
```

---

# spire.xls python batch file creation
## create multiple Excel files with sample data
```python
for n in range(0, 50):
    workbook = Workbook()
    workbook.CreateEmptySheets(5)
    for i in range(0, 5):
        sheet = workbook.Worksheets[i]
        sheet.Name = "Sheet" + str(i)
        for row in range(1, 15):
            for col in range(1, 5):
                sheet.Range[row,col].Text = "row" + str(row) + " col" + str(col)

    workbook.SaveToFile("Workbook" + str(n) + ".xlsx", ExcelVersion.Version2010)
    workbook.Dispose()
```

---

# spire.xls python hello world
## create a simple excel file with hello world text
```python
workbook = Workbook()
sheet = workbook.Worksheets.Add("MySheet")
sheet.Range["A1"].Text = "Hello World"
sheet.Range["A1"].AutoFitColumns()
```

---

# spire.xls python open existing file
## open an existing Excel file, add a new worksheet and save
```python
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets.Add("MySheet")
sheet.Range["A1"].Text = "Hello World"
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
```

---

# spire.xls python label control
## add label control to Excel worksheet
```python
# Create a workbook
workbook = Workbook()
# Get the first worksheet
sheet = workbook.Worksheets[0]
# Add a label control to the worksheet
label = sheet.LabelShapes.AddLabel(10, 2, 30, 200)
# Set the text of the label
label.Text = "This is a Label Control"
```

---

# spire.xls python listbox
## add listbox control to Excel worksheet
```python
listBox = sheet.ListBoxes.AddListBox(13, 4, 100, 80)
listBox.SelectionType = SelectionType.Single
listBox.SelectedIndex = 2
listBox.Display3DShading = True
listBox.ListFillRange = sheet.Range["A7:A12"]
```

---

# Spire.XLS Python ScrollBar Control
## Add a scroll bar control to an Excel worksheet
```python
# Add scroll bar control
scrollBar = sheet.ScrollBarShapes.AddScrollBar(10, 3, 150, 20)
scrollBar.LinkedCell = sheet.Range["B10"]
scrollBar.Min = 1
scrollBar.Max = 150
scrollBar.IncrementalChange = 1
scrollBar.Display3DShading = True
```

---

# spire.xls python table
## create table with filter in excel
```python
#Create a List Object named in Table.
sheet.ListObjects.Create("Table", sheet.Range[1,1,sheet.LastRow,sheet.LastColumn])
#Set the BuiltInTableStyle for List object.
sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleLight9
```

---

# spire.xls python table
## add total row to table
```python
#Create a table with the data from the specific cell range.
table = sheet.ListObjects.Create("Table", sheet.Range["A1:D4"])
#Display total row.
table.DisplayTotalRow = True
#Add a total row.
cols = table.Columns
cols[0].TotalsRowLabel = "Total"
cols[1].TotalsCalculation = ExcelTotalsCalculation.Sum
cols[2].TotalsCalculation = ExcelTotalsCalculation.Sum
cols[3].TotalsCalculation = ExcelTotalsCalculation.Sum
```

---

# Spire.XLS Python Subscript and Superscript
## Apply subscript and superscript formatting to Excel cell text
```python
#Set the rtf value of "B3" to "R100-0.06".
range = sheet.Range["B3"]
range.RichText.Text = "R100-0.06"

#Create a font. Set the IsSubscript property of the font to "true".
font = workbook.CreateFont()
font.IsSubscript = True
font.Color = Color.get_Green()

#Set font for specified range of the text in "B3".
range.RichText.SetFont(4, 8, font)

#Set the rtf value of "D3" to "a2 + b2 = c2".
range = sheet.Range["D3"]
range.RichText.Text = "a2 + b2 = c2"

#Create a font. Set the IsSuperscript property of the font to "true".
font = workbook.CreateFont()
font.IsSuperscript = True

#Set font for specified range of the text in "D3".
range.RichText.SetFont(1, 1, font)
range.RichText.SetFont(6, 6, font)
range.RichText.SetFont(11, 11, font)

sheet.AllocatedRange.AutoFitColumns()
```

---

# spire.xls python style cloning
## clone Excel font style and modify cloned style
```python
#Create a style with font properties
style = workbook.Styles.Add("style")
style.Font.FontName = "Calibri"
style.Font.Color = Color.get_Red()
style.Font.Size = 12
style.Font.IsBold = True
style.Font.IsItalic = True

#Apply style to a cell
sheet.Range["A1"].CellStyleName = style.Name

#Clone the style and apply to another cell
csOrieign = style.clone()
sheet.Range["B2"].CellStyleName = csOrieign.Name

#Clone the style, modify font color, and apply to a third cell
csGreen = style.clone()
csGreen.Font.Color = Color.get_Green()
sheet.Range["C3"].CellStyleName = csGreen.Name
```

---

# spire.xls python copy cells range
## copy a range of cells to another range in Excel worksheet
```python
#Get the first worksheet
sheet1 = workbook.Worksheets[0]

#Specify a destination range 
cells = sheet1.Range["G1:H19"]

#Copy the selected range to destination range 
sheet1.Range["B1:C19"].Copy(cells)
```

---

# Spire.XLS Python Copy Data with Style
## Copy cell range data with formatting from source to destination
```python
#Get a source range (A1:D3).
srcRange = worksheet.Range["A1:D3"]

#Create a style object.
style = workbook.Styles.Add("style")
cellStyle = (CellStyle)(style)

#Specify the font attribute.
cellStyle.Font.FontName = "Calibri"

#Specify the shading color.
cellStyle.Font.Color = Color.get_Red()

#Specify the border attributes.
cellStyle.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin
cellStyle.Borders[BordersLineType.EdgeTop].Color = Color.get_Blue()
cellStyle.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin
cellStyle.Borders[BordersLineType.EdgeBottom].Color = Color.get_Blue()
cellStyle.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin
cellStyle.Borders[BordersLineType.EdgeTop].Color = Color.get_Blue()
cellStyle.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin
cellStyle.Borders[BordersLineType.EdgeRight].Color = Color.get_Blue()
srcRange.CellStyleName = style.Name

#Set the destination range
destRange = worksheet.Range["A12:D14"]

#Copy the range data with style
srcRange.Copy(destRange, True, True)
```

---

# spire.xls python copy formula values
## copy only formula values from one range to another in Excel
```python
#Set the copy option
copyOptions = CopyRangeOptions.OnlyCopyFormulaValue

#Copy range
sheet.Copy(sheet.Range["A2:C2"], sheet.Range["A5:C5"], copyOptions)
```

---

# Excel Nested Group Creation
## Create nested groups in Excel worksheet using Spire.XLS
```python
#Set the summary rows appear above detail rows.
sheet.PageSetup.IsSummaryRowBelow = False

#Group the rows that you want to group.
sheet.GroupByRows(2, 9, False)
sheet.GroupByRows(4, 5, False)
sheet.GroupByRows(8, 9, False)
```

---

# spire.xls python table creation
## create table in excel worksheet with style
```python
# Add a new List Object to the worksheet
sheet.ListObjects.Create("table", sheet.Range[1,1,19,5])
# Add Default Style to the table
sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleLight9
```

---

# Data Sorting in Excel
## Sort Excel data by multiple columns
```python
# Add sort columns (column 2 and 3, ascending order)
workbook.DataSorter.SortColumns.Add(2, OrderBy.Ascending)
workbook.DataSorter.SortColumns.Add(3, OrderBy.Ascending)
# Apply sorting to the specified range
workbook.DataSorter.Sort(worksheet["A1:E19"])
```

---

# Excel Data Validation
## Implementing different types of data validation in Excel cells
```python
#Decimal DataValidation
sheet.Range["B11"].Text = "Input Number(3-6):"
rangeNumber = sheet.Range["B12"]
#Set the operator for the data validation.
rangeNumber.DataValidation.CompareOperator = ValidationComparisonOperator.Between
#Set the value or expression associated with the data validation.
rangeNumber.DataValidation.Formula1 = "3"
#The value or expression associated with the second part of the data validation.
rangeNumber.DataValidation.Formula2 = "6"
#Set the data validation type.
rangeNumber.DataValidation.AllowType = CellDataType.Decimal
#Set the data validation error message.
rangeNumber.DataValidation.ErrorMessage = "Please input correct number!"
#Enable the error.
rangeNumber.DataValidation.ShowError = True
rangeNumber.Style.KnownColor = ExcelColors.Gray25Percent

#Date DataValidation
sheet.Range["B14"].Text = "Input Date:"
rangeDate = sheet.Range["B15"]
rangeDate.DataValidation.AllowType = CellDataType.Date
rangeDate.DataValidation.CompareOperator = ValidationComparisonOperator.Between
rangeDate.DataValidation.Formula1 = "1/1/1970"
rangeDate.DataValidation.Formula2 = "12/31/1970"
rangeDate.DataValidation.ErrorMessage = "Please input correct date!"
rangeDate.DataValidation.ShowError = True
rangeDate.DataValidation.AlertStyle = AlertStyleType.Warning
rangeDate.Style.KnownColor = ExcelColors.Gray25Percent

#TextLength DataValidation
sheet.Range["B17"].Text = "Input Text:"
rangeTextLength = sheet.Range["B18"]
rangeTextLength.DataValidation.AllowType = CellDataType.TextLength
rangeTextLength.DataValidation.CompareOperator = ValidationComparisonOperator.LessOrEqual
rangeTextLength.DataValidation.Formula1 = "5"
rangeTextLength.DataValidation.ErrorMessage = "Enter a Valid String!"
rangeTextLength.DataValidation.ShowError = True
rangeTextLength.DataValidation.AlertStyle = AlertStyleType.Stop
rangeTextLength.Style.KnownColor = ExcelColors.Gray25Percent
```

---

# Excel Group Management
## Expand and collapse row groups in Excel worksheets
```python
#Expand the grouped rows with ExpandCollapseFlags set to expand parent
sheet.Range["A16:G19"].ExpandGroup(GroupByType.ByRows, ExpandCollapseFlags.ExpandParent)

#Collapse the grouped rows
sheet.Range["A10:G12"].CollapseGroup(GroupByType.ByRows)
```

---

# Excel Find and Replace Data
## Find and replace text in Excel cells with highlighting
```python
#Get the first worksheet
worksheet = workbook.Worksheets[0]

#Find the string
ranges = worksheet.FindAllString("Area", False, False)
#Traverse the found ranges
for range in ranges:
    #Replace it with new text
    range.Text = "Area Code"
    #Highlight the color
    range.Style.Color = Color.get_Yellow()
```

---

# Find Data in Excel Range
## Core functionality for finding text and numbers in a specific Excel range

```python
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
```

---

# Spire.XLS for Python - Find String and Number
## This code demonstrates how to find specific strings and numbers in an Excel worksheet
```python
#Find cells with the input string
textRanges = sheet.FindAllString("E-iceblue", False, False)

#Process found text cells
if len(textRanges) != 0:
    for range in textRanges:
        address = range.RangeAddress
        # Do something with the found cell address
        print("Found text cell at: " + address)
else:
    print("No cells contain the specified text")

#Find cells with the input integer or double
numberRanges = sheet.FindAllNumber(100, True)

#Process found number cells
if len(numberRanges) != 0:
    for range in numberRanges:
        address = range.RangeAddress
        # Do something with the found cell address
        print("Found number cell at: " + address)
else:
    print("No cells contain the specified number")
```

---

# spire.xls python table formatting
## format Excel table with built-in styles and total row
```python
#Add Default Style to the table
sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium9
#Show Total
sheet.ListObjects[0].DisplayTotalRow = True
#Set calculation type
sheet.ListObjects[0].Columns[0].TotalsRowLabel = "Total"
sheet.ListObjects[0].Columns[1].TotalsCalculation = ExcelTotalsCalculation.none
sheet.ListObjects[0].Columns[2].TotalsCalculation = ExcelTotalsCalculation.none
sheet.ListObjects[0].Columns[3].TotalsCalculation = ExcelTotalsCalculation.Sum
sheet.ListObjects[0].Columns[4].TotalsCalculation = ExcelTotalsCalculation.Sum

sheet.ListObjects[0].ShowTableStyleRowStripes = True

sheet.ListObjects[0].ShowTableStyleColumnStripes = True
```

---

# spire.xls python controls
## insert various controls into Excel worksheet
```python
#Add a textbox 
textbox = ws.TextBoxes.AddTextBox(9, 2, 25, 100)
textbox.Text = "Hello World"
#Add a checkbox 
cb = ws.CheckBoxes.AddCheckBox(11, 2, 15, 100)
cb.CheckState = CheckState.Checked
cb.Text = "Check Box 1"
#Add a RadioButton 
rb = ws.RadioButtons.Add(13, 2, 15, 100)
rb.Text = "Option 1"

#Add a combobox
cbx = ws.ComboBoxes.AddComboBox(15, 2, 15, 100)
cbx.ListFillRange = ws.Range["A41:A47"]
```

---

# Insert HTML string into Excel cell
## Demonstrates how to insert HTML content into a cell in Excel
```python
# Insert Html String in range A1
htmlCode = "<div>first line<strong>second line</strong>third line</div>"
range = sheet["A1"]
range.HtmlString = htmlCode
```

---

# Spire.XLS Python Named Ranges
## Creating and setting named ranges in Excel
```python
workbook = Workbook()
sheet = workbook.Worksheets[0]
#Creating a named range
NamedRange = workbook.NameRanges.Add("NewNamedRange")
#Setting the range of the named range
NamedRange.RefersToRange = sheet.Range["A8:E12"]
```

---

# Excel Text Replacement and Highlighting
## Replace specific text and highlight cells in Excel
```python
# Find all cells containing "Total"
ranges = worksheet.FindAllString("Total", True, True)

for range in ranges:
    # Replace the text with "Sum"
    range.Text = "Sum"
    # Set the highlight color to yellow
    range.Style.Color = Color.get_Yellow()
```

---

# Excel Data Retrieval and Extraction
## Extract rows containing specific text from one Excel sheet to another
```python
# Retrieve data and extract it from source sheet to destination sheet
i = 1
columnCount = len(sheet.Columns)
cells = sheet.Columns[0].Cells
for range in cells:
            if range.Text == "teacher":
                sourceRange = sheet.Range[range.Row,1,range.Row,columnCount]
                destRange = newSheet.Range[i,1,i,columnCount]
                sheet.Copy(sourceRange, destRange, True)
                i += 1
```

---

# spire.xls data validation across sheets
## Setting up data validation in Excel that references a range in a separate worksheet
```python
# Get the first worksheet
sheet1 = workbook.Worksheets[0]
sheet1.Range["B10"].Text = "Here is a dataValidation example."
# Get the second worksheet
sheet2 = workbook.Worksheets[1]
# Enable data validation to reference cells in different sheets
workbook.Allow3DRangesInDataValidation = True
# Set up data validation on cell B11 in sheet1 referencing range A1:A7 in sheet2
sheet1.Range["B11"].DataValidation.DataRange = sheet2.Range["A1:A7"]
```

---

# Split Excel Data into Multiple Columns
## This code demonstrates how to split data in an Excel worksheet into multiple columns based on a delimiter (space).
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]

#Split data into separate columns by the delimited characters – space.
splitText = None
text = None
i = 1
while i < sheet.LastRow:
    text = sheet.Range[i + 1,1].Text
    splitText = text.split(' ')
    j = 0
    while j < len(splitText):
        sheet.Range[i + 1,1 + j + 1].Text = splitText[j]
        j += 1
    i += 1
```

---

# Excel Subtotal Creation
## Create subtotals in Excel spreadsheet using Spire.XLS
```python
#Select data range
range = sheet.Range["A1:B18"]
#Subtotal selected data
sheet.Subtotal(range, 0, [1], SubtotalTypes.Sum, True, False, True)
```

---

# spire.xls python richtext
## write rich text with different font styles to excel cell
```python
workbook = Workbook()
sheet = workbook.Worksheets[0]

fontBold = workbook.CreateFont()
fontBold.IsBold = True

fontUnderline = workbook.CreateFont()
fontUnderline.Underline = FontUnderlineType.Single

fontItalic = workbook.CreateFont()
fontItalic.IsItalic = True

fontColor = workbook.CreateFont()
fontColor.KnownColor = ExcelColors.Green

richText = sheet.Range["B11"].RichText
richText.Text = "Bold and underlined and italic and colored text."
richText.SetFont(0, 3, fontBold)
richText.SetFont(9, 18, fontUnderline)
richText.SetFont(24, 29, fontItalic)
richText.SetFont(35, 41, fontColor)
```

---

# Access Cells in Excel Worksheet
## Demonstrates different ways to access cells in an Excel worksheet
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Access cell by its name
range1 = sheet.Range["A1"]
#Access cell by index of row and column
range2 = sheet.Range[2,1]
#Access cell in cell collection
range3 = sheet.Cells[2]
```

---

# Spire.XLS Python Multiple Fonts in Single Cell
## Apply different fonts to different parts of text within a single Excel cell
```python
#Create a font object in workbook, setting the font color, size and type.
font1 = workbook.CreateFont()
font1.KnownColor = ExcelColors.LightBlue
font1.IsBold = True
font1.Size = 10
#Create another font object specifying its properties.
font2 = workbook.CreateFont()
font2.KnownColor = ExcelColors.Red
font2.IsBold = True
font2.IsItalic = True
font2.FontName = "Times New Roman"
font2.Size = 11
#Write a RichText string to the cell 'H5', and set the font for it.
richText = sheet.Range["H5"].RichText
richText.Text = "This document was created with Spire.XLS for python."
richText.SetFont(0, 29, font1)
richText.SetFont(31, 48, font2)
```

---

# AutoFit Cells Based on Cell Value
## This code demonstrates how to auto-fit column width and row height based on cell value in Excel using Spire.XLS for Python
```python
#Set value for B8
cell = worksheet.Range["B8"]
cell.Text = "Welcome to Spire.XLS!"
#Set the cell style
style = cell.Style
style.Font.Size = 10
style.Font.IsBold = True
#Autofit column width and row height based on cell value
cell.AutoFitColumns()
cell.AutoFitRows()
```

---

# Cell Style Name Processing
## Find cells with the same style name and set their values
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the cell style name
styleName = sheet.Range["A1"].CellStyleName
ranges = sheet.AllocatedRange
for cc in ranges.Cells:
    #Find the cells which have the same style name
    if cc.CellStyleName == styleName:
        #Set value
        cc.Value = "Same style"
```

---

# spire.xls python text to number conversion
## convert text format to number format in Excel cells
```python
#Convert text string format to number format
worksheet.Range["D2:D8"].ConvertToNumber()
```

---

# Spire.XLS Python Cell Format Copying
## Copy cell format from one column to another
```python
#Copy the cell format from column 2 and apply to cells of column 5.
count = sheet.Rows.Length
i = 1
while i < count + 1:
    sheet.Range["E{0}".format(i)].Style = sheet.Range["B{0}".format(i)].Style
    i += 1
```

---

# Spire.XLS Python Cell Count
## Count the number of cells in a worksheet
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the number of cells.
cellCount = sheet.Cells.Length
```

---

# spire.xls python cut cells
## Cut cells from one position to another position in worksheet
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Define source and destination ranges
Ori = sheet.Range["A1:C5"]
Dest = sheet.Range["A26:C30"]
#Copy the range to other position
sheet.Copy(Ori, Dest, True, True, True)
#Remove all content in original cells
for cr in Ori.Cells:
    cr.ClearAll()
```

---

# Detecting and unmerging cells in Excel
## This code detects merged cells in the first worksheet and unmerges them
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the merged cell ranges in the first worksheet and put them into a CellRange array.
range = sheet.MergedCells
#Traverse through the array and unmerge the merged cells.
for cell in range:
    cell.UnMerge()
```

---

# Spire.XLS Python Duplicate Cell Range
## Demonstrates how to duplicate a cell range in Excel using Spire.XLS for Python
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Copy data from source range to destination range and maintain the format.
sheet.Copy(sheet.Range["A6:F6"], sheet.Range["A16:F16"], True)
```

---

# Excel Cell Emptying Methods
## Different ways to clear or empty cells in Excel
```python
#Set the value as null to remove the original content from the Excel Cell.
sheet.Range["C6"].Value = ""
#Clear the contents to remove the original content from the Excel Cell.
sheet.Range["B6"].ClearContents()
#Remove the contents with format from the Excel cell.
sheet.Range["D6"].ClearAll()
```

---

# Spire.XLS Python Cell Filtering
## Filter cells by cell color
```python
#Create an auto filter in the sheet and specify the range to be filtered
sheet.AutoFilters.Range = sheet.Range["G1:G19"]
#Get the column to be filtered
filtercolumn = sheet.AutoFilters[0]
#Add a color filter to filter the column based on cell color
sheet.AutoFilters.AddFillColorFilter(filtercolumn, Color.get_Red())
#Filter the data
sheet.AutoFilters.Filter()
```

---

# Find cells with style name
## Find all cells that have the same style name as a specific cell and mark them

```python
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the cell style name
styleName = sheet.Range["A1"].CellStyleName
ranges = sheet.AllocatedRange
for cc in ranges.Cells:
    #Find the cells which have the same style name
    if cc.CellStyleName == styleName:
        #Set value
        cc.Value = "Same style"
```

---

# Find Formula Cells in Excel
## Locate cells containing specific formulas in Excel worksheets
```python
#Assume workbook is a loaded Excel workbook
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Find the cells that contain formula "=SUM(A11,A12)"
ranges = sheet.FindAll("=SUM(A11,A12)", FindType.Formula, ExcelFindOptions.none)
#Create a string builder
builder = []
#Append the address of found cells to builder
if len(ranges) != 0:
    for range in ranges:
        address = range.RangeAddress
        builder.append("The address of found cell is: " + address)
else:
    builder.append("No cell contain the formula")
```

---

# Get Cell Address in Excel
## Demonstrates how to get various address information from cell ranges in Excel
```python
# Assuming workbook is already initialized
#Get the first worksheet
sheet = workbook.Worksheets[0]
# Get a cell range
range = sheet.Range["A1:B5"]
# Get address of range
address = range.RangeAddressLocal
# Get the cell count of range
count = range.CellsCount
# Get the address of the entire column of range
entireColAddress = range.EntireColumn.RangeAddressLocal
# Get the address of the entire row of range
entireRowAddress = range.EntireRow.RangeAddressLocal
```

---

# Spire.XLS Python Get Cell Data Type
## Get and display cell data types in Excel
```python
# Get the cell types of the cells in range "H2:H7"
for range in sheet.Range["H2:H7"].Cells:
    cellType = sheet.GetCellType(range.Row, range.Column, False)
    sheet[range.Row,range.Column + 1].Text = str(cellType)
    sheet[range.Row,range.Column + 1].Style.Font.Color = Color.get_Red()
    sheet[range.Row,range.Column + 1].Style.Font.IsBold = True
```

---

# Get Cell Displayed Text
## Retrieve the displayed text of a cell in Excel, considering its formatting
```python
#Get a cell from worksheet
cell = worksheet.Range["B8"]
#Set value and format for the cell
cell.NumberValue = 0.012345
style = cell.Style
style.NumberFormat = "0.00"
#Get the displayed text of the cell
displayedText = cell.DisplayedText
```

---

# Get Cell Value by Name
## Demonstrates how to get a cell's value using its name in Excel
```python
# Assuming a workbook is already loaded
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Specify a cell by its name.
cell = sheet.Range["A2"]
#Get value of cell "A2".
cell_value = cell.Value
```

---

# Spire.XLS Python Range Intersection
## Get the intersection of two ranges in an Excel worksheet
```python
#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the two ranges.
range = sheet.Range["A2:D7"].Intersect(sheet.Range["B2:E8"])
#Iterate through cells in the intersection
for r in range.Cells:
    # Access cell value
    cell_value = str(r.Value)
```

---

# Hide Cell Content in Excel
## Demonstrates how to hide cell content by setting the number format
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Hide the area by setting the number format as ";;;".
sheet.Range["C5:D6"].NumberFormat = ";;;"
```

---

# spire.xls python merge cells
## merge cells in excel worksheet
```python
#Create a workbook.
workbook = Workbook()
#Merge the seventh column in Excel file.
workbook.Worksheets[0].Columns[6].Merge()
#Merge the particular range in Excel file.
workbook.Worksheets[0].Range["A14:D14"].Merge()
```

---

# spire.xls python copy formula values
## copy only formula values from Excel cells
```python
#Set the copy option
copyOptions = CopyRangeOptions.OnlyCopyFormulaValue
sourceRange = sheet.Range["A6:E6"]
sheet.Copy(sourceRange, sheet.Range["A8:E8"], copyOptions)
sourceRange.Copy(sheet.Range["A10:E10"], copyOptions)
```

---

# spire.xls python cell formatting
## Set cell fill patterns and colors
```python
#Set cell color
worksheet.Range["B7:F7"].Style.Color = Color.get_Yellow()
#Set cell fill pattern
worksheet.Range["B8:F8"].Style.FillPattern = ExcelPatternType.Percent125Gray
```

---

# spire.xls python DB number formatting
## Set DB number format for Excel cells
```python
#Get the cell range
range = sheet.Range["A1:A3"]
#Set the DB num format
range.NumberFormat = "[DBNum2][$-804]General"
```

---

# Spire.XLS Python Shrink Text to Fit
## Enable ShrinkToFit property for a cell in Excel
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#The cell range to shrink text.
cell = sheet.Range["B13:C13"]
#Enable ShrinkToFit.
style = cell.Style
style.ShrinkToFit = True
```

---

# Traverse Excel Cell Values
## This code demonstrates how to traverse through all cells in an Excel worksheet and retrieve their values
```python
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Get the cell range collection 
cellRangeCollection = worksheet.Cells
#Traverse cells value
for cellRange in cellRangeCollection:
    #Set string format for displaying
    result = "Cell: " + cellRange.RangeAddress + "   Value: " + cellRange.Value
    #Add result string to list
    content.append(result)
```

---

# Spire.XLS Python Ungroup Cells
## Demonstrates how to ungroup rows in an Excel worksheet
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Ungroup the row 10 to 12.
sheet.UngroupByRows(10, 12)
#Ungroup the row 16 to 19.
sheet.UngroupByRows(16, 19)
```

---

# spire.xls python unmerge cells
## Unmerge specific cells in Excel worksheet
```python
#Unmerge the cells.
sheet.Range["F2"].UnMerge()
#Unmerge the cells.
sheet.Range["F7"].UnMerge()
```

---

# Excel Cell Line Breaks
## How to use explicit line breaks in Excel cells
```python
#Specify a cell range
c5 = sheet1.Range["C5"]
#Set the cell width for specified range
sheet1.SetColumnWidth(c5.Column, 70)
#Put the string value with explicit line breaks
c5.Value = "Spire.XLS for .NET is a professional Excel .NET API\n that can be used to create, read, \nwrite, convert and print Excel files in any type \nof .NET(C#, VB.NET, ASP.NET, .NET Core) application. \nSpire.XLS for .NET offers object model\n Excel API for speeding up Excel programming in .NET platform -\n create new Excel documents from template, edit existing \nExcel documents and \nconvert Excel files."
#Set Text wrap
c5.IsWrapText = True
```

---

# Spire.XLS Python Text Wrapping
## Wrap or unwrap text in Excel cells
```python
#Wrap the excel text
sheet.Range["C1"].Text = "e-iceblue is in facebook and welcome to like us"
sheet.Range["C1"].Style.WrapText = True
sheet.Range["D1"].Text = "e-iceblue is in twitter and welcome to follow us"
sheet.Range["D1"].Style.WrapText = True
#Unwrap the excel text
sheet.Range["C2"].Text = "http://www.facebook.com/pages/e-iceblue/139657096082266"
sheet.Range["C2"].Style.WrapText = False
sheet.Range["D2"].Text = "https://twitter.com/eiceblue"
sheet.Range["D2"].Style.WrapText = False
```

---

# Excel AutoFit Column in Range
## Automatically adjust column width within a specified range in Excel
```python
#Create a workbook
workbook = Workbook()
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Autofit the Column of the worksheet
sheet.AutoFitColumn(2, 2, 5)
```

---

# Spire.XLS Python AutoFit Row
## Automatically adjust the height of a specific row in Excel
```python
#Create a workbook
workbook = Workbook()
#Get the first worksheet
sheet = workbook.Worksheets[0]
# Autofit the second row of the worksheet
sheet.AutoFitRow(2, 1, 2, False)
```

---

# Check AutoFit Row or Column in Excel
## Check whether rows or columns in Excel have auto-fit settings
```python
# Get the first worksheet
sheet = workbook.Worksheets[0]
# Check whether the cell has an adaptive row height set
isRowAutofit = sheet.GetRowIsAutoFit(2)
# Check whether the cell has an adaptive column width set
isColAutofit = sheet.GetColumnIsAutoFit(2)
```

---

# Spire.XLS for Python - Check Hidden Rows or Columns
## This code demonstrates how to check if a specific row or column is hidden in an Excel worksheet

```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Check whether the row is hidden
isRowHide = sheet.GetRowIsHide(2)
#Check whether the column is hidden
isColHide = sheet.GetColumnIsHide(2)
```

---

# Spire.XLS Python Column Copying
## Copy columns within and between worksheets
```python
# Get worksheets
sheet1 = workbook.Worksheets[0]
sheet2 = workbook.Worksheets[1]
# Copy the first column to the third column in the same sheet
sheet1.Copy(sheet1.Columns[0], sheet1.Columns[2], True, True, True)
# Copy the first column to the second column in the different sheet
sheet1.Copy(sheet1.Columns[0], sheet2.Columns[1], True, True, True)
```

---

# spire.xls python row copying
## copy rows within and between worksheets
```python
#Copy the first row to the third row in the same sheet
sheet1.Copy(sheet1.Rows[0], sheet1.Rows[2], True, True, True)
#Copy the first row to the second row in the different sheet
sheet1.Copy(sheet1.Rows[0], sheet2.Rows[1], True, True, True)
```

---

# spire.xls python copy column and row
## copy single column and row in excel worksheet
```python
#Get the first worksheet
sheet1 = workbook.Worksheets[0]
#Specify a destination range to copy one column 
columnCells = sheet1.Range["G1:G19"]
#Copy the second column to destination range 
sheet1.Columns[1].Copy(columnCells)
#Specify a destination range to copy one row 
rowCells = sheet1.Range["A21:E21"]
#Copy the first row to destination range 
sheet1.Rows[0].Copy(rowCells)
```

---

# Spire.XLS Python Copy Range with Options
## Copy a range of cells from one worksheet to another with options to keep styles and update references
```python
#Get the first worksheet
sheet1 = workbook.Worksheets[0]
#Add a new worksheet as destination sheet
destinationSheet = workbook.Worksheets.Add("DestSheet")
#Specify a copy range of original sheet
cellRange = sheet1.Range["B2:D4"]
#Copy the specified range to added worksheet and keep original styles and update reference
workbook.Worksheets[0].Copy(cellRange, workbook.Worksheets[1], 2, 1, True, True)
```

---

# Delete blank rows and columns in Excel
## This code demonstrates how to delete blank rows and columns from an Excel worksheet
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Delete blank rows from the worksheet.
for i in range(sheet.Rows.Length - 1, -1, -1):
    if sheet.Rows[i].IsBlank:
        sheet.DeleteRow(i + 1)
#Delete blank columns from the worksheet.
for j in range(sheet.Columns.Length - 1, -1, -1):
    if sheet.Columns[j].IsBlank:
        sheet.DeleteColumn(j + 1)
```

---

# Spire.XLS Python Row and Column Operations
## Delete multiple rows and columns from an Excel worksheet
```python
#Delete 4 rows from the fifth row
sheet.DeleteRow(5, 4)
#Delete 2 columns from the second column
sheet.DeleteColumn(2, 2)
```

---

# spire.xls python core functionality
## get default row and column count of Excel worksheet
```python
#Create a workbook
workbook = Workbook()
#Clear all worksheets
workbook.Worksheets.Clear()
#Create a new worksheet
sheet = workbook.CreateEmptySheet()
#Get row and column count
rowCount = sheet.Rows.Length
columnCount = sheet.Columns.Length
```

---

# Spire.XLS Python Row and Column Grouping
## Group rows and columns in Excel using Spire.XLS for Python
```python
workbook = Workbook()
sheet = workbook.Worksheets[0]
# Grouping rows
sheet.GroupByRows(1, 5, False)
# Grouping columns
sheet.GroupByColumns(1, 3, False)
```

---

# spire.xls python hide/show headers
## Hide or show row and column headers in Excel worksheet
```python
#Create a workbook
workbook = Workbook()
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Hide the headers of rows and columns
sheet.RowColumnHeadersVisible = False
#Show the headers of rows and columns
#sheet.RowColumnHeadersVisible = true
```

---

# Spire.XLS Python Hide Rows and Columns
## Demonstrates how to hide specific rows and columns in an Excel worksheet
```python
# Hiding the column of the worksheet
worksheet.HideColumn(2)
# Hiding the row of the worksheet
worksheet.HideRow(4)
```

---

# Insert Rows and Columns in Excel
## Demonstrate how to insert single and multiple rows and columns in an Excel worksheet
```python
#Inserting a row into the worksheet 
worksheet.InsertRow(2)
#Inserting a column into the worksheet 
worksheet.InsertColumn(2)
#Inserting multiple rows into the worksheet
worksheet.InsertRow(5, 2)
#Inserting multiple columns into the worksheet
worksheet.InsertColumn(5, 2)
```

---

# Excel Row Removal Based on Keyword
## Removes a row from an Excel worksheet that contains a specific keyword
```python
# Create a workbook
workbook = Workbook()
sheet = workbook.Worksheets[0]
# Find the string "Address" in the sheet
cr = sheet.FindString("Address", False, False)
# Delete the row that contains the found string
sheet.DeleteRow(cr.Row)
```

---

# spire.xls python column width
## Set column width in pixels
```python
#Set the width of the third column to 400 pixels
sheet.SetColumnWidthInPixels(3, 400)
```

---

# Spire.XLS Python Default Column Width
## Set default column width for Excel worksheet
```python
# Get the first worksheet
sheet = workbook.Worksheets[0]
# Set default column width
sheet.DefaultColumnWidth = 25
```

---

# spire.xls python row and column styling
## set default style for rows and columns
```python
workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
#Create a cell style and set the color
style = workbook.Styles.Add("Mystyle")
style.Color = Color.get_Yellow()
#Set the default style for the first row and column 
sheet.SetDefaultRowStyle(1, style)
sheet.SetDefaultColumnStyle(1, style)
```

---

# spire.xls python row height
## set default row height for excel worksheet
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
# Set default row height
sheet.DefaultRowHeight = 30
```

---

# spire.xls python row and column sizing
## set column width and row height in Excel worksheet
```python
# Setting the width to 30
worksheet.SetColumnWidth(4, 30)
# Setting the height to 30
worksheet.SetRowHeight(4, 30)
```

---

# spire.xls python summary column direction
## Set summary column direction in Excel worksheet
```python
sheet = workbook.Worksheets[0]
#Group Columns
sheet.GroupByColumns(1, 4, True)
#Set summary columns to right of details
sheet.PageSetup.IsSummaryRowBelow = True
```

---

# Excel Summary Row Direction
## Set the direction of summary rows in Excel (above or below detail rows)
```python
#Create a workbook
workbook = Workbook()
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Group rows
sheet.GroupByRows(1, 4, True)
#Set summary rows above details
sheet.PageSetup.IsSummaryRowBelow = False
```

---

# spire.xls python rows and columns
## Unhide rows and columns in Excel worksheet
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Unhide the row
sheet.ShowRow(15)
#Unhide the column
sheet.ShowColumn(4)
```

---

# Excel Picture Alignment
## Align picture within a cell in Excel
```python
#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
sheet.Range["A1"].Text = "Align Picture Within A Cell:"
sheet.Range["A1"].Style.VerticalAlignment = VerticalAlignType.Top
#Add a picture to the worksheet
picture = sheet.Pictures.Add(1, 1, "image_path")
#Adjust the column width and row height so that the cell can contain the picture.
sheet.Columns[0].ColumnWidth = 40
sheet.Rows[0].RowHeight = 200
#Vertically and horizontally align the image.
picture.LeftColumnOffset = 100
picture.TopRowOffset = 25
```

---

# Spire.XLS Python Picture Compression
## Compress pictures in Excel worksheets
```python
# Iterate through all worksheets
for sheet in workbook.Worksheets:
    # Iterate through all pictures in each worksheet
    for picture in sheet.Pictures:
        # Compress each picture with 50% compression ratio
        picture.Compress(50)
```

---

# spire.xls python image handling
## copy picture between worksheets
```python
#Get the first worksheet
sheet1 = workbook.Worksheets[0]
#Add a new worksheet as destination sheet
destinationSheet = workbook.Worksheets.Add("DestSheet")
#Get the first picture from the first worksheet
sourcePicture = sheet1.Pictures[0]
#Get the image
image = sourcePicture.Picture
#Add the image into the added worksheet 
destinationSheet.Pictures.Add(2, 2, image)
```

---

# spire.xls python get picture position
## Get the cropped position of a picture in an Excel worksheet
```python
#Get the first worksheet
sheet1 = workbook.Worksheets[0]
#Get the image from the first sheet
picture = sheet1.Pictures[0]
#Get the cropped position
left = picture.Left
top = picture.Top
width = picture.Width
height = picture.Height
```

---

# Spire.XLS Python Delete Images
## Delete all images from an Excel worksheet
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Delete all images from the worksheet.
for i in range(sheet.Pictures.Count - 1, -1, -1):
    sheet.Pictures[i].Remove()
```

---

# Get Image Crop Position in Excel
## Extract the position and dimensions of a cropped image in an Excel worksheet
```python
#Get the first worksheet
sheet1 = workbook.Worksheets[0]
#Get the image from the first sheet
picture = sheet1.Pictures[0]
#Get the cropped position
left = picture.Left
top = picture.Top
width = picture.Width
height = picture.Height
```

---

# Insert Excel Background Image
## Set background image for Excel worksheet
```python
#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Open an image. 
image = Stream("image_path")
#Set the image to be background image of the worksheet.
sheet.PageSetup.BackgoundImage = image
```

---

# spire.xls python image manipulation
## locate and adjust image position in Excel
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
pic = sheet.Pictures[0]
pic.LeftColumnOffset = 300
pic.TopRowOffset = 300
```

---

# Spire.XLS Python Picture Offset
## Set picture offset in Excel worksheet
```python
#Insert a picture
pic = sheet.Pictures.Add(2, 2, inputFile)
#Set left offset and top offset from the current range
pic.LeftColumnOffset = 200
pic.TopRowOffset = 100
```

---

# spire.xls python picture reference range
## set reference range for a picture in Excel worksheet
```python
#Get the first picture in worksheet
picture = sheet.Pictures[0]
#Set the reference range of the picture to A1:B3
picture.RefRange = "A1:B3"
```

---

# spire.xls python read images
## extract and save images from excel files
```python
#Create a Workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first image
pic = sheet.Pictures[0]
#save
pic.Picture.Save(outputFile)
workbook.Dispose()
```

---

# Spire.XLS Python Remove Picture Border
## Remove border from a picture in Excel worksheet
```python
# Get the first worksheet
sheet1 = workbook.Worksheets[0]
# Get the first picture from the first worksheet
picture = sheet1.Pictures[0]
# Remove the picture border
# Method-1:
picture.Line.Visible = False
# Method-2:
# picture.Line.Weight = 0
```

---

# spire.xls python image manipulation
## reset size and position for image in Excel
```python
#Add a picture to the first worksheet
picture = sheet.Pictures.Add(1, 1, inputFile)
#Set the size for the picture
picture.Width = 200
picture.Height = 200
#Set the position for the picture
picture.Left = 200
picture.Top = 100
```

---

# spire.xls python chart image offset
## set image offset for chart background
```python
#Add chart
chart1 = sheet1.Charts.Add(ExcelChartType.ColumnClustered)
chart1.DataRange = sheet.Range["D1:E8"]
chart1.SeriesDataFromRange = False
#Set chart position
chart1.LeftColumn = 1
chart1.TopRow = 11
chart1.RightColumn = 8
chart1.BottomRow = 33
#Add picture as background
chart1.ChartArea.Fill.CustomPicture(Stream("image_path"), "None")
chart1.ChartArea.Fill.Tile = False
#Set the image offset  
chart1.ChartArea.Fill.PicStretch.Left = 20
chart1.ChartArea.Fill.PicStretch.Top = 20
chart1.ChartArea.Fill.PicStretch.Right = 5
chart1.ChartArea.Fill.PicStretch.Bottom = 5
```

---

# spire.xls python image
## write image to Excel worksheet
```python
#Add an image to the specific cell
sheet.Pictures.Add(14, 5, "image_path.png")
```

---

# Add Comment with Author in Excel
## This code demonstrates how to add a comment with author information to an Excel cell
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the range that will add comment
range = sheet.Range["C1"]
#Set the author and comment content
author = "E-iceblue"
text = "This is demo to show how to add a comment with editable Author property."
#Add comment to the range and set properties
comment = range.AddComment()
comment.Width = 200
comment.Visible = True
comment.Text = author + ":\n" + text
#Set the font of the author
font = workbook.CreateFont()
font.FontName = "Tahoma"
font.KnownColor = ExcelColors.Black
font.IsBold = True
comment.RichText.SetFont(0, len(author), font)
```

---

# spire.xls python comment with picture
## add comment with image to excel cell
```python
#Assuming sheet is already defined
#Add comment to cell
comment = sheet.Range["C6"].AddComment()
#Load image and set it for comment
image = Stream("path_to_image.png")
comment.Fill.CustomPicture(image, "logo.png")
comment.Visible = True
```

---

# spire.xls python comment
## edit excel comment
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the first comment.
comment = sheet.Comments[0]
#Edit the comment.
comment.Text = "This comment has been edited by Spire.XLS."
```

---

# Excel Comment Visibility Control
## Hide or show comments in Excel worksheets
```python
# Get the first worksheet
sheet = workbook.Worksheets[0]
# Hide the second comment
sheet.Comments[1].IsVisible = False
# Show the third comment
sheet.Comments[2].IsVisible = True
```

---

# spire.xls python comment reading
## Read comments from Excel cells
```python
# Get worksheet
sheet = workbook.Worksheets[0]
# Read comment text from cell A1
comment_text = sheet.Range["A1"].Comment.Text
# Read rich text comment from cell A2
rich_text_comment = sheet.Range["A2"].Comment.RichText.RtfText
```

---

# spire.xls python comment manipulation
## remove and modify excel comments
```python
#Get all comments of the first sheet
comments = workbook.Worksheets[0].Comments
#Change the content of the first comment
comments[0].Text = "This comment has been changed."
#Remove the second comment
comments[1].Remove()
```

---

# Spire.XLS Python Comment Fill Color
## Set the fill color of Excel cell comment
```python
#Create a workbook
workbook = Workbook()
#Get the default first worksheet
sheet = workbook.Worksheets[0]
#Create Excel font
font = workbook.CreateFont()
font.FontName = "Arial"
font.Size = 11
font.KnownColor = ExcelColors.Orange
#Add the comment
range = sheet.Range["A1"]
range.Comment.Text = "This is a comment"
range.Comment.RichText.SetFont(0, (len(range.Comment.Text) - 1), font)
#Set comment Color
range.Comment.Fill.FillType = ShapeFillType.SolidColor
range.Comment.Fill.ForeColor = Color.get_SkyBlue()
range.Comment.Visible = True
```

---

# spire.xls python comment
## set comment text rotation
```python
#Create Excel font
font = workbook.CreateFont()
font.FontName = "Arial"
font.Size = 11
font.KnownColor = ExcelColors.Orange
#Add the comment
range = sheet.Range["E1"]
range.Comment.Text = "This is a comment"
range.Comment.RichText.SetFont(0, (len(range.Comment.Text) - 1), font)
# Set its vertical and horizontal alignment 
range.Comment.VAlignment = CommentVAlignType.Center
range.Comment.HAlignment = CommentHAlignType.Right
#Set the comment text rotation
range.Comment.TextRotation = TextRotationType.LeftToRight
```

---

# spire.xls python comment positioning
## set position and alignment for Excel comments
```python
#Add comment 1 and set its position and alignment
Comment1 = sheet.Range["G5"].Comment
Comment1.IsVisible = True
Comment1.Height = 150
Comment1.Width = 300
Comment1.RichText.Text = "Spire.XLS for .Net:\nStandalone Excel component to meet your needs for conversion, data manipulation, charts in workbook etc. "
Comment1.RichText.SetFont(0, 19, font1)
Comment1.TextRotation = TextRotationType.LeftToRight

#Set the position of Comment
Comment1.Top = 20
Comment1.Left = 40

#Set the alignment of text in Comment
Comment1.VAlignment = CommentVAlignType.Center
Comment1.HAlignment = CommentHAlignType.Justified

#Add comment2 and set its size, text, position and alignment for comparison
sheet.Range["D14"].Text = "E-iceblue"
Comment2 = sheet.Range["D14"].Comment
Comment2.IsVisible = True
Comment2.Height = 150
Comment2.Width = 300
Comment2.RichText.Text = "About E-iceblue: \nWe focus on providing excellent office components for developers to operate Word, Excel, PDF, and PowerPoint documents."
Comment2.TextRotation = TextRotationType.LeftToRight
Comment2.RichText.SetFont(0, 16, font2)

#Set the position of Comment
Comment2.Top = 170
Comment2.Left = 450

#Set the alignment of text in Comment
Comment2.VAlignment = CommentVAlignType.Top
Comment2.HAlignment = CommentHAlignType.Justified
```

---

# Excel Comments Writing
## Writing regular and rich text comments to Excel cells
```python
# Regular comment
range = sheet.Range["B11"]
range.Text = "Regular comment"
range.Comment.Text = "Regular comment"
range.AutoFitColumns()
# Rich text comment
range = sheet.Range["B12"]
range.Text = "Rich text comment"
range.RichText.SetFont(0, 16, font)
range.AutoFitColumns()
range.Comment.RichText.Text = "Rich text comment"
range.Comment.RichText.SetFont(0, 4, fontGreen)
range.Comment.RichText.SetFont(5, 9, fontBlue)
```

---

# spire.xls python chart sheet conversion
## convert chart sheet to SVG format
```python
#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the chartsheet by name
cs = workbook.GetChartSheetByName("Chart1")
fs = Stream(outputFile)
cs.ToSVGStream(fs)
fs.Flush()
fs.Close()
```

---

# spire.xls python CSV to Excel conversion
## Convert CSV file to Excel format
```python
# Create a workbook
workbook = Workbook()
# Load a csv file
workbook.LoadFromFile(inputFile, ",", 1, 1)
sheet = workbook.Worksheets[0]
sheet.Range["D2:E19"].IgnoreErrorOptions = IgnoreErrorType.NumberAsText
sheet.AllocatedRange.AutoFitColumns()
# Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
```

---

# CSV to PDF Conversion
## Convert CSV file to PDF format using Spire.XLS for Python
```python
#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile, ",", 1, 1)
#Set the SheetFitToPage property as true
workbook.ConverterSetting.SheetFitToPage = True
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Autofit a column if the characters in the column exceed column width
i = 1
while i < sheet.Columns.Length:
    sheet.AutoFitColumn(i)
    i += 1
workbook.SaveToFile(outputFile, FileFormat.PDF)
workbook.Dispose()
```

---

# Excel to PDF Conversion
## Convert each worksheet to a separate PDF file
```python
# Process each worksheet in the workbook
for sheet in workbook.Worksheets:
    FileName = sheet.Name + ".pdf"
    # Save the sheet to PDF
    sheet.SaveToPdf(FileName)
```

---

# Fit Width When Converting Excel to PDF
## Configure page setup to fit content width when converting Excel to PDF
```python
for sheet in workbook.Worksheets:
    #Auto fit page height
    sheet.PageSetup.FitToPagesTall = 0
    #Fit to one page width
    sheet.PageSetup.FitToPagesWide = 1
```

---

# HTML to Excel Conversion
## Core functionality for converting HTML files to Excel format using Spire.XLS
```python
#Create a workbook
workbook = Workbook()
#Load html
workbook.LoadFromHtml(input_html_file)
#Save the document
workbook.SaveToFile(output_excel_file, ExcelVersion.Version2013)
```

---

# Spire.XLS Python Conversion
## Convert Excel to ET format
```python
#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
#convert to ET file
workbook.SaveToFile(outputFile, FileFormat.ET)
workbook.Dispose()
```

---

# spire.xls python conversion
## convert Office Open XML to Excel format
```python
# Create a workbook
workbook = Workbook()
# Load XML data from stream
fileStream = Stream(inputFile)
workbook.LoadFromXml(fileStream)
fileStream.Close()
# Save to Excel format
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
```

---

# Spire.XLS Python Range to PDF Conversion
## Convert a selected range from Excel to PDF format
```python
#Add a new sheet to workbook
workbook.Worksheets.Add("newsheet")
#Copy your area to new sheet.
workbook.Worksheets[0].Range["A9:E15"].Copy(workbook.Worksheets[1].Range["A9:E15"], False, True)
#Auto fit column width
workbook.Worksheets[1].Range["A9:E15"].AutoFitColumns()
#Save the selected range to PDF
workbook.Worksheets[1].SaveToPdf("output.pdf")
```

---

# Convert Excel Sheet to Image
## Demonstrates how to convert an Excel worksheet to an image file
```python
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn).Save(outputFile)
workbook.Dispose()
```

---

# Excel Cell Range to Image Conversion
## Convert specific cell ranges in Excel to different image formats
```python
#Get the first worksheet in Excel file
sheet = workbook.Worksheets[0]
# Specify Cell Ranges and Save to certain Image formats
sheet.ToImage(1, 1, 7, 5).Save("SpecificCellsToImage.png")
sheet.ToImage(8, 1, 15, 5).Save("SpecificCellsToImage.jpg")
sheet.ToImage(17, 1, 23, 5).Save("SpecificCellsToImage.bmp")
```

---

# Spire.XLS Python Font Directory Specification
## Specify custom font directory for Excel to PDF conversion
```python
#create a workbook
workbook = Workbook()
#Specify font directory
workbook.CustomFontFileDirectory= [("./Demos/Data/Fonts/")]
```

---

# Spire.XLS Excel to CSV Conversion
## Convert Excel worksheet to CSV format
```python
#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
#get the first sheet
sheet = workbook.Worksheets[0]
#convert to CSV file
sheet.SaveToFile(outputFile, ",", Encoding.get_UTF8())
workbook.Dispose()
```

---

# Excel to CSV Conversion
## Convert Excel worksheet to CSV with filtered values
```python
workbook = Workbook()
workbook.LoadFromFile(inputFile)
# Convert to CSV file with filtered value
workbook.Worksheets[0].SaveToFile(outputFile, ";", False)
workbook.Dispose()
```

---

# Spire.XLS Python Conversion
## Convert Excel to HTML
```python
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
options = HTMLOptions()
options.ImageEmbedded = True
sheet.SaveToHtml(outputFile)
workbook.Dispose()
```

---

# spire.xls python conversion
## convert excel sheet to html stream
```python
#Create a workbook
workbook = Workbook()
#Load the Excel document
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set the html options
options = HTMLOptions()
options.ImageEmbedded = True
#Save sheet to html stream
fileStream = Stream(outputFile)
sheet.SaveToHtml(fileStream, options)
fileStream.Close()
workbook.Dispose()
```

---

# Excel to Image Conversion Without White Space
## Convert Excel worksheet to image without white space by setting margins to zero
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
#Set the margin as 0 to remove the white space around the image
sheet.PageSetup.LeftMargin = 0
sheet.PageSetup.BottomMargin = 0
sheet.PageSetup.TopMargin = 0
sheet.PageSetup.RightMargin = 0
#convert to image
image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn)
```

---

# Excel to ODS Conversion
## Convert Excel files to ODS format using Spire.XLS for Python
```python
#create a workbook
workbook = Workbook()
#load a excel document
workbook.LoadFromFile(inputFile)
#convert to ODS file
workbook.SaveToFile(outputFile, FileFormat.ODS)
workbook.Dispose()
```

---

# spire.xls excel to ofd conversion
## convert excel file to ofd format
```python
#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
#convert to OFD file
workbook.SaveToFile(outputFile, FileFormat.OFD)
workbook.Dispose()
```

---

# Spire.XLS for Python - Save as XML
## Convert Excel workbook to Office Open XML format
```python
workbook = Workbook()
sheet = workbook.Worksheets[0]
workbook.SaveAsXml("output.xml")
workbook.Dispose()
```

---

# spire.xls python conversion
## convert Excel to PDF
```python
#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
workbook.ConverterSetting.SheetFitToPage = True
#convert to PDF file
workbook.SaveToFile(outputFile, FileFormat.PDF)
workbook.Dispose()
```

---

# Excel to PDF/A-1B Conversion
## Convert Excel files to PDF with PDF/A-1B compliance
```python
# Create a workbook
workbook = Workbook()
# Load an excel file
workbook.LoadFromFile(inputFile)
# Convert excel to PDFA/1-B
workbook.ConverterSetting.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B
workbook.SaveToFile(outputFile, FileFormat.PDF)
workbook.Dispose()
```

---

# Excel to PDF Conversion
## Simple conversion of Excel file to PDF format
```python
# Create a workbook
workbook = Workbook()
# Load an Excel document
workbook.LoadFromFile("input.xlsx")
# Convert Excel to PDF
workbook.SaveToFile("output.pdf", FileFormat.PDF)
workbook.Dispose()
```

---

# Spire.XLS for Python - Change Page Size
## Change worksheet page size to A3
```python
#Change the page size
for sheet in workbook.Worksheets:
    sheet.PageSetup.PaperSize = PaperSizeType.PaperA3
```

---

# Excel to PostScript Conversion
## Convert Excel file to PostScript format using Spire.XLS
```python
#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
workbook.SaveToFile(outputFile, FileFormat.PostScript)
workbook.Dispose()
```

---

# Excel to SVG Conversion
## Convert Excel worksheets to SVG format
```python
i = 0
for worksheet in workbook.Worksheets:
    fs = Stream("sheet-" + str(i) + ".svg")
    worksheet.ToSVGStream(fs, 0, 0, 0, 0)
    i = i + 1
```

---

# spire.xls python conversion
## convert Excel to text file
```python
# Define input and output file paths
inputFile = "./Demos/Data/ConversionSample2.xlsx"
outputFile = "ExceltoTxt.txt"

# Create a workbook
workbook = Workbook()
# Load the document from disk
workbook.LoadFromFile(inputFile)
# Get the first worksheet in excel workbook
sheet = workbook.Worksheets[0]
sheet.SaveToFile(outputFile, " ", Encoding.get_UTF8())
workbook.Dispose()
```

---

# Excel to XPS Conversion
## Convert Excel file to XPS format using Spire.XLS for Python
```python
#Create a workbook
workbook = Workbook()
# Load the document from disk
workbook.LoadFromFile(inputFile)
# Convert to XPS file
workbook.SaveToFile(outputFile, FileFormat.XPS)
workbook.Dispose()
```

---

# spire.xls python conversion
## convert workbook to HTML
```python
# Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
# Convert to HTML
workbook.SaveToHtml(outputFile)
workbook.Dispose()
```

---

# Spire.XLS Python XLS to XLSM Conversion
## Convert XLS file to XLSM format using Spire.XLS library
```python
#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Convert to xlsm
workbook.SaveToFile(outputFile, ExcelVersion.Version2007)
workbook.Dispose()
```

---

# spire.xls python autofilter
## autofilter blank cells in Excel
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Match the blank data
sheet.AutoFilters.MatchBlanks(0)
#Filter
sheet.AutoFilters.Filter()
```

---

# Apply AutoFilter for Non-Blank Cells
## This code demonstrates how to apply an auto-filter to show only non-blank cells in an Excel worksheet.
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Match the non blank data
sheet.AutoFilters.MatchNonBlanks(0)
#Filter
sheet.AutoFilters.Filter()
```

---

# spire.xls python filter
## create auto filter in Excel
```python
workbook = Workbook()
sheet = workbook.Worksheets[0]
#Create filter
sheet.AutoFilters.Range = sheet.Range["A1:J1"]
```

---

# Spire.XLS Python Data Validation
## Implement different types of data validation in Excel cells
```python
# Decimal DataValidation
sheet.Range["B11"].Text = "Input Number(3-6):"
rangeNumber = sheet.Range["B12"]
# Set the operator for the data validation
rangeNumber.DataValidation.CompareOperator = ValidationComparisonOperator.Between
# Set the value or expression associated with the data validation
rangeNumber.DataValidation.Formula1 = "3"
# The value or expression associated with the second part of the data validation
rangeNumber.DataValidation.Formula2 = "6"
# Set the data validation type
rangeNumber.DataValidation.AllowType = CellDataType.Decimal
# Set the data validation error message
rangeNumber.DataValidation.ErrorMessage = "Please input correct number!"
# Enable the error
rangeNumber.DataValidation.ShowError = True
rangeNumber.Style.KnownColor = ExcelColors.Gray25Percent

# Date DataValidation
sheet.Range["B14"].Text = "Input Date:"
rangeDate = sheet.Range["B15"]
rangeDate.DataValidation.AllowType = CellDataType.Date
rangeDate.DataValidation.CompareOperator = ValidationComparisonOperator.Between
rangeDate.DataValidation.Formula1 = "1/1/1970"
rangeDate.DataValidation.Formula2 = "12/31/1970"
rangeDate.DataValidation.ErrorMessage = "Please input correct date!"
rangeDate.DataValidation.ShowError = True
rangeDate.DataValidation.AlertStyle = AlertStyleType.Warning
rangeDate.Style.KnownColor = ExcelColors.Gray25Percent

# TextLength DataValidation
sheet.Range["B17"].Text = "Input Text:"
rangeTextLength = sheet.Range["B18"]
rangeTextLength.DataValidation.AllowType = CellDataType.TextLength
rangeTextLength.DataValidation.CompareOperator = ValidationComparisonOperator.LessOrEqual
rangeTextLength.DataValidation.Formula1 = "5"
rangeTextLength.DataValidation.ErrorMessage = "Enter a Valid String!"
rangeTextLength.DataValidation.ShowError = True
rangeTextLength.DataValidation.AlertStyle = AlertStyleType.Stop
rangeTextLength.Style.KnownColor = ExcelColors.Gray25Percent
```

---

# spire.xls python filtering
## filter cells by string in Excel
```python
#get the first worksheet
sheet=workbook.Worksheets[0]
#filter cells data which start with "South"
sheet.AutoFilters.Range = sheet.Range["D1:D24"]
filtercolumn = sheet.AutoFilters[0]
strCrt = String("South*")
sheet.AutoFilters.CustomFilter(filtercolumn, FilterOperatorType.Equal, strCrt)
sheet.AutoFilters.Filter()
```

---

# spire.xls data validation
## get data validation settings from excel cell
```python
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Cell B4 has the Decimal Validation
cell = worksheet.Range["B4"]
#Get the validation of this cell
validation = cell.DataValidation
#Get the settings
allowType = str(validation.AllowType)
data = str(validation.CompareOperator)
minimum = str(validation.Formula1)
maximum = str(validation.Formula2)
ignoreBlank = str(validation.IgnoreBlank)
```

---

# Excel List Data Validation
## Create a dropdown list for cell validation using Spire.XLS for Python
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set data validation for cell
range = sheet.Range["D10"]
range.DataValidation.ShowError = True
range.DataValidation.AlertStyle = AlertStyleType.Stop
range.DataValidation.ErrorTitle = "Error"
range.DataValidation.ErrorMessage = "Please select a city from the list"
range.DataValidation.DataRange = sheet.Range["A7:A10"]
```

---

# Remove Auto Filters in Excel
## This code demonstrates how to remove auto filters from an Excel worksheet
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Remove the auto filters.
sheet.AutoFilters.Clear()
```

---

# Spire.XLS Python Data Validation Removal
## Remove data validation from specified ranges in Excel worksheet
```python
#Create an array of rectangles, which is used to locate the ranges in worksheet.
rectangles = []
#Assign value to the first element of the array. This rectangle specifies the cells from A1 to B3.
rectangles.append(Rectangle.FromLTRB(0, 0, 1, 2))
#Remove validations in the ranges represented by rectangles.
workbook.Worksheets[0].DVTable.Remove(rectangles)
```

---

# spire.xls python data validation
## set data validation referencing a separate sheet
```python
# Get the first and second sheets
sheet1 = workbook.Worksheets[0]
sheet2 = workbook.Worksheets[1]

# Enable data validation from different sheets
sheet2.ParentWorkbook.Allow3DRangesInDataValidation = True

# Set data validation for cell B11 in sheet1 using range from sheet2
sheet1.Range["B11"].DataValidation.DataRange = sheet2.Range["A1:A7"]
```

---

# Time Data Validation in Excel
## Set time validation for a cell with specific time range
```python
#Set Time data validation for cell "D12"
range = sheet.Range["D12"]
range.DataValidation.AllowType = CellDataType.Time
range.DataValidation.CompareOperator = ValidationComparisonOperator.Between
range.DataValidation.Formula1 = "09:00"
range.DataValidation.Formula2 = "18:00"
range.DataValidation.AlertStyle = AlertStyleType.Info
range.DataValidation.ShowError = True
range.DataValidation.ErrorTitle = "Time Error"
range.DataValidation.ErrorMessage = "Please enter a valid time"
range.DataValidation.InputMessage = "Time Validation Type"
range.DataValidation.IgnoreBlank = True
range.DataValidation.ShowInput = True
```

---

# Excel Data Validation Verification
## Verify data values against Excel cell validation rules
```python
#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Cell B4 has the Decimal Validation
cell = worksheet.Range["B4"]
#Get the validation of this cell
validation = cell.DataValidation
#Get the specified data range
minimum = Double.Parse(validation.Formula1)
maximum = Double.Parse(validation.Formula2)
#Create a list to save results
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
    #Add result string to the list
    content.append(result)
```

---

# Excel Whole Number Data Validation
## Set whole number validation rules for an Excel cell
```python
#Set Whole Number data validation for cell "D12"
range = sheet.Range["D12"]
range.DataValidation.AllowType = CellDataType.Integer
range.DataValidation.CompareOperator = ValidationComparisonOperator.Between
range.DataValidation.Formula1 = "10"
range.DataValidation.Formula2 = "100"
range.DataValidation.AlertStyle = AlertStyleType.Info
range.DataValidation.ShowError = True
range.DataValidation.ErrorTitle = "Error"
range.DataValidation.ErrorMessage = "Please enter a valid number"
range.DataValidation.InputMessage = "Whole Number Validation Type"
range.DataValidation.IgnoreBlank = True
range.DataValidation.ShowInput = True
```

---

# Spire.XLS Python Chart Data Table
## Add data table to an Excel chart
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first chart
chart = sheet.Charts[0]
chart.HasDataTable = True
```

---

# spire.xls python error bars
## add error bars to excel charts
```python
#Add a line chart and then add percentage error bar to the chart
chart = sheet.Charts.Add(ExcelChartType.Line)
chart.DataRange = sheet.Range["B1:B7"]
chart.SeriesDataFromRange = False
#Set chart position
chart.TopRow = 8
chart.BottomRow = 25
chart.LeftColumn = 2
chart.RightColumn = 9
chart.ChartTitle = "Error Bar 10% Plus"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
cs1 = chart.Series[0]
cs1.CategoryLabels = sheet.Range["A2:A7"]
cs1.ErrorBar(True, ErrorBarIncludeType.Plus, ErrorBarType.Percentage, 10.0)
```

```python
# Add a column chart with standard error bars as comparison
chart2 = sheet.Charts.Add(ExcelChartType.ColumnClustered)
chart2.DataRange = sheet.Range["B1:C7"]
chart2.SeriesDataFromRange = False
#Set chart position
chart2.TopRow = 8
chart2.BottomRow = 25
chart2.LeftColumn = 10
chart2.RightColumn = 17
chart2.ChartTitle = "Standard Error Bar"
chart2.ChartTitleArea.IsBold = True
chart2.ChartTitleArea.Size = 12
cs2 = chart2.Series[0]
cs2.CategoryLabels = sheet.Range["A2:A7"]
cs2.ErrorBar(True, ErrorBarIncludeType.Minus, ErrorBarType.StandardError, 0.3)
cs3 = chart2.Series[1]
cs3.ErrorBar(True, ErrorBarIncludeType.Both, ErrorBarType.StandardError, 0.5)
```

---

# spire.xls python chart
## add picture to chart
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Add the picture in chart
chart.Shapes.AddPicture(inputFile_Img)
```

---

# spire.xls python textbox
## add textbox to chart in Excel
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first chart
chart = sheet.Charts[0]
#Add a Textbox
textbox = chart.Shapes.AddTextBox()
textbox.Width = 1200
textbox.Height = 320
textbox.Left = 1000
textbox.Top = 480
textbox.Text = "This is a textbox"
```

---

# spire.xls python trendlines
## add different types of trendlines to excel charts
```python
#select chart and set logarithmic trendline
chart = sheet.Charts[0]
chart.ChartTitle = "Logarithmic Trendline"
chart.Series[0].TrendLines.Add(TrendLineType.Logarithmic)
#select chart and set moving_average trendline
chart1 = sheet.Charts[1]
chart1.ChartTitle = "Moving Average Trendline"
chart1.Series[0].TrendLines.Add(TrendLineType.Moving_Average)
#select chart and set linear trendline
chart2 = sheet.Charts[2]
chart2.ChartTitle = "Linear Trendline"
chart2.Series[0].TrendLines.Add(TrendLineType.Linear)
#select chart and set exponential trendline
chart3 = sheet.Charts[3]
chart3.ChartTitle = "Exponential Trendline"
chart3.Series[0].TrendLines.Add(TrendLineType.Exponential)
```

---

# spire.xls python chart
## adjust bar space in chart
```python
#Get the first chart from the first worksheet
chart = workbook.Worksheets[0].Charts[0]
#Adjust the space between bars
for cs in chart.Series:
    cs.Format.Options.GapWidth = 200
    cs.Format.Options.Overlap = 0
```

---

# Apply Soft Edges Effect to Chart
## This code demonstrates how to apply a soft edges effect to a chart in an Excel worksheet
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Specify the size of the soft edge. Value can be set from 0 to 100
chart.ChartArea.Shadow.SoftEdge = 25
```

---

# Spire.XLS Chart Size and Position
## Change chart dimensions and location in Excel
```python
#Get the chart
chart = sheet.Charts[0]
#Change chart size
chart.Width = 600
chart.Height = 500
#Change chart position
chart.LeftColumn = 3
chart.TopRow = 7
```

---

# spire.xls python chart data label
## Change data label in chart
```python
#Get the chart
chart = sheet.Charts[0]
#Change data label of the first datapoint of the first series
chart.Series[0].DataPoints[0].DataLabels.Text = "changed data label"
```

---

# spire.xls python chart
## change chart data range
```python
#Get chart
chart = sheet.Charts[0]
#Change data range
chart.DataRange = sheet.Range["A1:C4"]
```

---

# Change chart major gridlines color
## Change the color of major gridlines in an Excel chart
```python
#Get the chart
chart = sheet.Charts[0]
#Change the color of major gridlines
chart.PrimaryValueAxis.MajorGridLines.LineProperties.Color = Color.get_Red()
```

---

# spire.xls python chart
## change series color
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first chart
chart = sheet.Charts[0]
#Get the second series
cs = chart.Series[1]
#Set the fill type
cs.Format.Fill.FillType = ShapeFillType.SolidColor
#Change the fill color
cs.Format.Fill.ForeColor = Color.get_Orange()
```

---

# Setting Chart Axis Titles
## This code demonstrates how to set titles for chart axes and adjust font size
```python
#Set axis title
chart.PrimaryCategoryAxis.Title = "Category Axis"
chart.PrimaryValueAxis.Title = "Value axis"
#Set font size
chart.PrimaryCategoryAxis.Font.Size = 12
chart.PrimaryValueAxis.Font.Size = 12
```

---

# spire.xls python chart to image
## Convert Excel chart to image
```python
inputFile = "./Demos/Data/ChartToImage.xlsx"
outputFile = "ChartToImage.png"
#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Save chart as image
image = workbook.SaveChartAsImage(workbook.Worksheets[0], 0)
image.Save(outputFile)
workbook.Dispose()
image.Dispose()
```

---

# spire.xls python clustered bar chart
## create clustered bar chart and 3D clustered bar chart
```python
#Add a chart
chart = sheet.Charts.Add()
#Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
chart.ChartType = ExcelChartType.BarClustered
#Chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PrimaryCategoryAxis.Title = "Country"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.TextRotationAngle = 90
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
for cs in chart.Series:
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
chart.Legend.Position = LegendPositionType.Top
```

## create 3D clustered bar chart
```python
#Add a chart
chart = sheet.Charts.Add()
#Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
chart.ChartType = ExcelChartType.Bar3DClustered
#Chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PrimaryCategoryAxis.Title = "Country"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.TextRotationAngle = 90
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
for cs in chart.Series:
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
chart.Legend.Position = LegendPositionType.Top
```

---

# spire.xls python chart
## Create clustered column charts
```python
#Add a chart to the sheet
chart = sheet.Charts.Add()
#Set data range of chart 
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
#Set position of the chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
chart.ChartType = ExcelChartType.ColumnClustered
#Chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
#Chart Axis
chart.PrimaryCategoryAxis.Title = "Country"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
for cs in chart.Series:
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
#Chart Legend
chart.Legend.Position = LegendPositionType.Top
```

## Create 3D clustered column charts
```python
#Add a chart to the sheet
chart = sheet.Charts.Add()
#Set data range of chart 
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
#Set position of the chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
chart.ChartType = ExcelChartType.Column3DClustered
#Chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
#Chart Axis
chart.PrimaryCategoryAxis.Title = "Country"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
for cs in chart.Series:
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
#Chart Legend
chart.Legend.Position = LegendPositionType.Top
```

---

# spire.xls python Box and Whisker chart
## create a Box and Whisker chart with custom series settings
```python
#add a new chart
officeChart = sheet.Charts.Add()

#set the chart title
officeChart.ChartTitle = "Yearly Vehicle Sales"

#set chart type as Box and Whisker
officeChart.ChartType = ExcelChartType.BoxAndWhisker

#set data range in the worksheet
officeChart.DataRange = sheet["A1:E17"]

#box and Whisker settings on first series
seriesA = officeChart.Series[0]
seriesA.DataFormat.ShowInnerPoints = False
seriesA.DataFormat.ShowOutlierPoints = True
seriesA.DataFormat.ShowMeanMarkers = True
seriesA.DataFormat.ShowMeanLine = False
seriesA.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian

#box and Whisker settings on second series   
seriesB = officeChart.Series[1]
seriesB.DataFormat.ShowInnerPoints = False
seriesB.DataFormat.ShowOutlierPoints = True
seriesB.DataFormat.ShowMeanMarkers = True
seriesB.DataFormat.ShowMeanLine = False
seriesB.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.InclusiveMedian

#box and Whisker settings on third series   
seriesC = officeChart.Series[2]
seriesC.DataFormat.ShowInnerPoints = False
seriesC.DataFormat.ShowOutlierPoints = True
seriesC.DataFormat.ShowMeanMarkers = True
seriesC.DataFormat.ShowMeanLine = False
seriesC.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian
```

---

# Spire.XLS Python Bubble Chart Creation
## This code demonstrates how to create a bubble chart in an Excel worksheet
```python
#Add a chart
chart = sheet.Charts.Add(ExcelChartType.Bubble)
# Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
chart.Series[0].Bubbles = sheet.Range["C2:C5"]
# Set position of chart
chart.LeftColumn = 7
chart.TopRow = 6
chart.RightColumn = 16
chart.BottomRow = 29
chart.ChartTitle = "Bubble Chart"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
```

---

# Spire.XLS Chart Based on Pivot Table
## Create a chart based on a pivot table in Excel
```python
# Get the sheet in which the pivot table is located
sheet = workbook.Worksheets[0]
pt = sheet.PivotTables[0] if isinstance(sheet.PivotTables[0], XlsPivotTable) else None
workbook.Worksheets[1].Charts.Add(ExcelChartType.BarClustered, pt)
```

---

# spire.xls python chart creation
## create chart without range data
```python
#Create a workbook
workbook = Workbook()
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Add a chart to the worksheet
chart = sheet.Charts.Add()
chart.ChartTitle = "Sample Chart"
#Add a series to the chart
series = chart.Series.Add()
#Add data 
series.EnteredDirectlyValues = [Int32(10), Int32(20), Int32(30)]
v = series.EnteredDirectlyValues
```

---

# spire.xls python custom chart
## create a chart with different chart types for different series
```python
#Add a chart based on the data from A1 to B4
chart = sheet.Charts.Add()
chart.DataRange = sheet.Range["A1:B4"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 10
chart.RightColumn = 7
chart.BottomRow = 25
#Apply different chart type to different series
cs1 = chart.Series[0]
cs1.SerieType = ExcelChartType.ColumnClustered
cs2 = chart.Series[1]
cs2.SerieType = ExcelChartType.Line
chart.ChartTitle = "Custom chart"
```

---

# Spire.XLS for Python - Create Doughnut Chart
## This code demonstrates how to create a doughnut chart in Excel using Spire.XLS for Python library
```python
# Add a new chart, set chart type as doughnut
chart = sheet.Charts.Add()
chart.ChartType = ExcelChartType.Doughnut
chart.DataRange = sheet.Range["A1:B5"]
chart.SeriesDataFromRange = False
# Set position of chart
chart.LeftColumn = 4
chart.TopRow = 2
chart.RightColumn = 12
chart.BottomRow = 22
# Chart title
chart.ChartTitle = "Market share by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
for cs in chart.Series:
    cs.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = True
chart.Legend.Position = LegendPositionType.Top
```

---

# spire.xls python chart
## create funnel chart
```python
#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as Funnel
officeChart.ChartType = ExcelChartType.Funnel

#set data range in the worksheet
officeChart.DataRange = sheet.Range["A1:B6"]

#set the chart title
officeChart.ChartTitle = "Funnel"

#format the legend and data label option
officeChart.HasLegend = False
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = True
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8
```

---

# spire.xls python histogram chart
## create and configure a histogram chart
```python
#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as Histogram
officeChart.ChartType = ExcelChartType.Histogram

#set data range in the worksheet   
officeChart.DataRange = sheet["A1:A15"]
officeChart.TopRow = 1
officeChart.BottomRow = 19
officeChart.LeftColumn = 4
officeChart.RightColumn = 12

#category axis bin settings        
officeChart.PrimaryCategoryAxis.BinWidth = 8

#gap width settings
officeChart.Series[0].DataFormat.Options.GapWidth = 6

#set the chart title and axis title
officeChart.ChartTitle = "Height Data"
officeChart.PrimaryValueAxis.Title = "Number of students"
officeChart.PrimaryCategoryAxis.Title = "Height"

#hide the legend
officeChart.HasLegend = False
```

---

# spire.xls python multi-level chart
## create multi-level category chart in Excel
```python
#Add a clustered bar chart to worksheet
chart = sheet.Charts.Add(ExcelChartType.BarClustered)
chart.ChartTitle = "Value"
chart.PlotArea.Fill.FillType = ShapeFillType.NoFill
chart.Legend.Delete()
chart.LeftColumn = 5
chart.TopRow = 1
chart.RightColumn = 14
#Set the data source of series data
chart.DataRange = sheet.Range["C2:C9"]
chart.SeriesDataFromRange = False
#Set the data source of category labels
serie = chart.Series[0]
serie.CategoryLabels = sheet.Range["A2:B9"]
#Show multi-level category labels
chart.PrimaryCategoryAxis.MultiLevelLable = True
```

---

# Spire.XLS Python Pareto Chart
## Create a Pareto chart with specific formatting and configuration
```python
#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as Pareto
officeChart.ChartType = ExcelChartType.Pareto

#set data range in the worksheet   
officeChart.DataRange = sheet["A2:B8"]

officeChart.TopRow = 1
officeChart.BottomRow = 19
officeChart.LeftColumn = 4
officeChart.RightColumn = 12
officeChart.PrimaryCategoryAxis.IsBinningByCategory = True

officeChart.PrimaryCategoryAxis.OverflowBinValue = 5
officeChart.PrimaryCategoryAxis.UnderflowBinValue = 1

#set color of Pareto line      
officeChart.Series[0].ParetoLineFormat.LineProperties.Color = Color.get_Blue()

#gap width settings
officeChart.Series[0].DataFormat.Options.GapWidth = 6

#set the chart title
officeChart.ChartTitle = "Expenses"

#hide the legend
officeChart.HasLegend = False
```

---

# Spire.XLS Python Pivot Chart
## Create a clustered column chart based on a pivot table
```python
#get the first worksheet
sheet = workbook.Worksheets[0]
#get the first pivot table in the worksheet
pivotTable = sheet.PivotTables[0]
#create a clustered column chart based on the pivot table
chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, pivotTable)
#set chart position
chart.TopRow = 10
chart.LeftColumn = 1
chart.RightColumn = 7
chart.BottomRow = 25
#set chart title
chart.ChartTitle = "Pivot Chart"
```

---

# spire.xls python radar chart creation
## Creating radar charts using spire.xls library
```python
# Add a new chart
chart = sheet.Charts.Add()
# Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
# Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
chart.ChartType = ExcelChartType.Radar
# Chart title
chart.ChartTitle = "Sale market by region"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PlotArea.Fill.Visible = False
chart.Legend.Position = LegendPositionType.Corner
```

## Creating filled radar charts using spire.xls library
```python
# Add a new chart
chart = sheet.Charts.Add()
# Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
# Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
chart.ChartType = ExcelChartType.RadarFilled
# Chart title
chart.ChartTitle = "Sale market by region"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PlotArea.Fill.Visible = False
chart.Legend.Position = LegendPositionType.Corner
```

---

# Spire.XLS Python SunBurst Chart
## Create and configure a SunBurst chart in Excel
```python
#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as SunBurst
officeChart.ChartType = ExcelChartType.SunBurst

#set data range in the worksheet   
officeChart.DataRange = sheet["A1:D16"]

officeChart.TopRow = 1
officeChart.BottomRow = 17
officeChart.LeftColumn = 6
officeChart.RightColumn = 14

#set the chart title
officeChart.ChartTitle = "Sales by quarter"

#format data labels      
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8

#hide the legend
officeChart.HasLegend = False
```

---

# spire.xls python chart
## create TreeMap chart
```python
#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as TreeMap
officeChart.ChartType = ExcelChartType.TreeMap

#set data range in the worksheet   
officeChart.DataRange = sheet["A2:C11"]
officeChart.TopRow = 1
officeChart.BottomRow = 19
officeChart.LeftColumn = 4
officeChart.RightColumn = 14

#Set the chart title
officeChart.ChartTitle = "Area by countries"

#set the Treemap label option
officeChart.Series[0].DataFormat.TreeMapLabelOption = ExcelTreeMapLabelOption.Banner

#format data labels      
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8
```

---

# Spire.XLS Python Waterfall Chart
## Creating a waterfall chart using Spire.XLS for Python
```python
#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as WaterFall
officeChart.ChartType = ExcelChartType.WaterFall

#set data range in the worksheet   
officeChart.DataRange = sheet["A2:B8"]
officeChart.TopRow = 1
officeChart.BottomRow = 19
officeChart.LeftColumn = 4
officeChart.RightColumn = 12

#set data point as total in chart
officeChart.Series[0].DataPoints[3].SetAsTotal = True
officeChart.Series[0].DataPoints[6].SetAsTotal = True

#show the connector lines between data points
officeChart.Series[0].Format.ShowConnectorLines = True

#set the chart title
officeChart.ChartTitle = "WaterFall Chart"

#format data label and legend option
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = True
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8
officeChart.Legend.Position = LegendPositionType.Right
```

---

# spire.xls python chart
## custom data markers in scatter chart
```python
#Create a Scatter-Markers chart based on the sample data
chart = sheet.Charts.Add(ExcelChartType.ScatterMarkers)
chart.DataRange = sheet.Range["A1:B7"]
chart.PlotArea.Visible = False
chart.SeriesDataFromRange = False
chart.TopRow = 5
chart.BottomRow = 22
chart.LeftColumn = 4
chart.RightColumn = 11
chart.ChartTitle = "Chart with Markers"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 10
#Format the markers in the chart by setting the background color, foreground color, type, size and transparency
cs1 = chart.Series[0]
cs1.DataFormat.MarkerBackgroundColor = Color.get_RoyalBlue()
cs1.DataFormat.MarkerForegroundColor = Color.get_WhiteSmoke()
cs1.DataFormat.MarkerSize = 7
cs1.DataFormat.MarkerStyle = ChartMarkerType.PlusSign
cs1.DataFormat.MarkerTransparencyValue = 0.8
cs2 = chart.Series[1]
cs2.DataFormat.MarkerBackgroundColor = Color.get_Pink()
cs2.DataFormat.MarkerSize = 9
cs2.DataFormat.MarkerStyle = ChartMarkerType.Triangle
cs2.DataFormat.MarkerTransparencyValue = 0.9
```

---

# spire.xls data callout configuration
## configure chart data callouts with various properties
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first chart
chart = sheet.Charts[0]
for cs in chart.Series:
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasWedgeCallout = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = True
```

---

# spire.xls python delete legend
## Delete specific legend entries from an Excel chart
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Delete the first and the second legend entries from the chart
chart.Legend.LegendEntries[0].Delete()
chart.Legend.LegendEntries[1].Delete()
```

---

# Spire.XLS for Python - Discontinuous Data Chart
## Creating a chart with discontinuous data ranges
```python
#Add a chart
chart = sheet.Charts.Add(ExcelChartType.ColumnClustered)
chart.SeriesDataFromRange = False
#Set the position of chart
chart.LeftColumn = 1
chart.TopRow = 10
chart.RightColumn = 10
chart.BottomRow = 24
#Add a series
cs1 = chart.Series.Add()
#Set the name of the cs1
cs1.Name = sheet.Range["B1"].Value
#Set discontinuous values for cs1
cs1.CategoryLabels = sheet.Range["A2:A3"].AddCombinedRange(sheet.Range["A5:A6"]).AddCombinedRange(sheet.Range["A8:A9"])
cs1.Values = sheet.Range["B2:B3"].AddCombinedRange(sheet.Range["B5:B6"]).AddCombinedRange(sheet.Range["B8:B9"])
#Set the chart type
cs1.SerieType = ExcelChartType.ColumnClustered
#Add a series
cs2 = chart.Series.Add()
cs2.Name = sheet.Range["C1"].Value
cs2.CategoryLabels = sheet.Range["A2:A3"].AddCombinedRange(sheet.Range["A5:A6"]).AddCombinedRange(sheet.Range["A8:A9"])
cs2.Values = sheet.Range["C2:C3"].AddCombinedRange(sheet.Range["C5:C6"]).AddCombinedRange(sheet.Range["C8:C9"])
cs2.SerieType = ExcelChartType.ColumnClustered
chart.ChartTitle = "Chart"
chart.ChartTitleArea.Font.Size = 20
chart.ChartTitleArea.Color = Color.get_Black()
chart.PrimaryValueAxis.HasMajorGridLines = False
```

---

# Edit Line Chart in Excel
## Add a new series to an existing line chart and set its values
```python
#Get the line chart
chart = sheet.Charts[0]
#Add a new series
cs = chart.Series.Add("Added")
#Set the values for the series
cs.Values = sheet.Range["I1:L1"]
```

---

# Spire.XLS Exploded Doughnut Chart
## Create an exploded doughnut chart in Excel using Python
```python
#Add a chart
chart = sheet.Charts.Add()
chart.ChartType = ExcelChartType.DoughnutExploded
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
#Set region of chart data
chart.DataRange = sheet.Range["A1:B5"]
chart.SeriesDataFromRange = False
#Chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
for cs in chart.Series:
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
chart.PlotArea.Fill.Visible = False
chart.Legend.Position = LegendPositionType.Top
```

---

# Extract Trendline from Excel Chart
## Extracts the equation of a trendline from an Excel chart
```python
#Get the chart from the first worksheet
chart = workbook.Worksheets[0].Charts[0]
#Get the trendline of the chart and then extract the equation of the trendline
trendLine = chart.Series[1].TrendLines[0]
formula = trendLine.Formula
```

---

# spire.xls python chart fill with picture
## Fill chart elements with image
```python
#Get the first worksheet from workbook
ws = workbook.Worksheets[0]
#Get the first chart
chart = ws.Charts[0]
# Fill chart area with image
chart.ChartArea.Fill.CustomPicture(Stream(inputImg), "None")
chart.PlotArea.Fill.Transparency = 0.9
```

---

# Excel Axis Formatting
## Format chart axis properties in Excel using Python
```python
# Add a chart
chart = sheet.Charts.Add(ExcelChartType.ColumnClustered)
chart.DataRange = sheet.Range["B1:B9"]
chart.SeriesDataFromRange = False
chart.PlotArea.Visible = False
chart.TopRow = 10
chart.BottomRow = 28
chart.LeftColumn = 2
chart.RightColumn = 10
chart.ChartTitle = "Chart with Customized Axis"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
cs1 = chart.Series[0]
cs1.CategoryLabels = sheet.Range["A2:A9"]

# Format axis
chart.PrimaryValueAxis.MajorUnit = 8
chart.PrimaryValueAxis.MinorUnit = 2
chart.PrimaryValueAxis.MaxValue = 50
chart.PrimaryValueAxis.MinValue = 0
chart.PrimaryValueAxis.IsReverseOrder = False
chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkOutside
chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkInside
chart.PrimaryValueAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionNextToAxis
chart.PrimaryValueAxis.CrossesAt = 0

# Set NumberFormat
chart.PrimaryValueAxis.NumberFormat = "$#,##0"
chart.PrimaryValueAxis.IsSourceLinked = False
```

---

# spire.xls python gauge chart
## create a gauge chart using doughnut and pie charts
```python
#Add a Doughnut chart
chart = sheet.Charts.Add(ExcelChartType.Doughnut)
chart.DataRange = sheet.Range["A1:A5"]
chart.SeriesDataFromRange = False
chart.HasLegend = True
#Set the position of chart
chart.LeftColumn = 2
chart.TopRow = 7
chart.RightColumn = 9
chart.BottomRow = 25
#Get the series 1
cs1 = chart.Series["Value"]
cs1.Format.Options.DoughnutHoleSize = 60
cs1.DataFormat.Options.FirstSliceAngle = 270
#Set the fill color
cs1.DataPoints[0].DataFormat.Fill.ForeColor = Color.get_Yellow()
cs1.DataPoints[1].DataFormat.Fill.ForeColor = Color.get_PaleVioletRed()
cs1.DataPoints[2].DataFormat.Fill.ForeColor = Color.get_DarkViolet()
cs1.DataPoints[3].DataFormat.Fill.Visible = False
#Add a series with pie chart
cs2 = chart.Series.Add("Pointer", ExcelChartType.Pie)
#Set the value
cs2.Values = sheet.Range["D2:D4"]
cs2.UsePrimaryAxis = False
cs2.DataPoints[0].DataLabels.HasValue = True
cs2.DataFormat.Options.FirstSliceAngle = 270
cs2.DataPoints[0].DataFormat.Fill.Visible = False
cs2.DataPoints[1].DataFormat.Fill.FillType = ShapeFillType.SolidColor
cs2.DataPoints[1].DataFormat.Fill.ForeColor = Color.get_Black()
cs2.DataPoints[2].DataFormat.Fill.Visible = False
```

---

# spire.xls python get chart category labels
## Extract category labels from a chart in Excel
```python
workbook = Workbook()
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Get the cell range of the category labels
cr = chart.PrimaryCategoryAxis.CategoryLabels
sb = []
for cell in cr:
    sb.append(cell.Value + "\r\n")
```

---

# Get Chart Data Point Values
## Extract values from chart data points in Excel
```python
# Get the first sheet
sheet = workbook.Worksheets[0]
# Get the chart
chart = sheet.Charts[0]
# Get the first series of the chart
cs = chart.Series[0]
for cr in cs.Values:
    # Get the range address
    range_address = cr.RangeAddress
    # Get the data point value
    value = cr.Value
```

---

# spire.xls python chart
## get worksheet containing chart
```python
#Access first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Access the first chart inside this worksheet
chart = worksheet.Charts[0]
#Get its worksheet
obj = chart.Worksheet
wSheet = Worksheet(obj)
#Set string format for displaying
result = "Sheet Name: " + worksheet.Name + "\r\nCharts' sheet Name: " + wSheet.Name
```

---

# spire.xls python chart
## Hide major gridlines of a chart
```python
#Get the chart
chart = sheet.Charts[0]
#Hide major gridlines
chart.PrimaryValueAxis.HasMajorGridLines = False
```

---

# spire.xls python line chart
## create line chart
```python
#Add a chart
chart = sheet.Charts.Add()
chart.ChartType = ExcelChartType.Line
#Set region of chart data
chart.DataRange = sheet.Range["A1:E5"]
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
#Set chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PrimaryCategoryAxis.Title = "Month"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
for cs in chart.Series:
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
chart.PlotArea.Fill.Visible = False
chart.Legend.Position = LegendPositionType.Top
```

## create line chart with circle markers
```python
#Add a chart
chart = sheet.Charts.Add()
chart.ChartType = ExcelChartType.Line
#Set region of chart data
chart.DataRange = sheet.Range["A1:E5"]
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
#Set chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PrimaryCategoryAxis.Title = "Month"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
for cs1 in chart.Series:
    cs = ChartSerie(cs1)
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
    cs.DataFormat.MarkerStyle = ChartMarkerType.Circle
chart.PlotArea.Fill.Visible = False
chart.Legend.Position = LegendPositionType.Top
```

## create 3D line chart
```python
#Add a chart
chart = sheet.Charts.Add()
chart.ChartType = ExcelChartType.Line3D
#Set region of chart data
chart.DataRange = sheet.Range["A1:E5"]
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
#Set chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PrimaryCategoryAxis.Title = "Month"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
for cs in chart.Series:
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
chart.PlotArea.Fill.Visible = False
chart.Legend.Position = LegendPositionType.Top
```

---

# spire.xls python chart
## create Pie chart
```python
#Add a chart
chart = sheet.Charts.Add(ExcelChartType.Pie)
#Set region of chart data
chart.DataRange = sheet.Range["B2:B5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 9
chart.BottomRow = 25
#Chart title
chart.ChartTitle = "Sales by year"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
cs = chart.Series[0]
cs.CategoryLabels = sheet.Range["A2:A5"]
cs.Values = sheet.Range["B2:B5"]
cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
chart.PlotArea.Fill.Visible = False
```

## create 3D Pie chart
```python
#Add a chart
chart = sheet.Charts.Add(ExcelChartType.Pie3D)
#Set region of chart data
chart.DataRange = sheet.Range["B2:B5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 9
chart.BottomRow = 25
#Chart title
chart.ChartTitle = "Sales by year"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
cs = chart.Series[0]
cs.CategoryLabels = sheet.Range["A2:A5"]
cs.Values = sheet.Range["B2:B5"]
cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
chart.PlotArea.Fill.Visible = False
```

---

# spire.xls python pyramid column chart
## create 3D clustered pyramid column chart
```python
#Add a chart
chart = sheet.Charts.Add()
#Set region of chart data
chart.DataRange = sheet.Range["B2:B5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
chart.ChartType = ExcelChartType.Pyramid3DClustered
#Chart title
chart.ChartTitle = "Sales by year"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PrimaryCategoryAxis.Title = "Year"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
cs = chart.Series[0]
cs.CategoryLabels = sheet.Range["A2:A5"]
cs.Format.Options.IsVaryColor = True
chart.Legend.Position = LegendPositionType.Top
```

---

# spire.xls python chart removal
## remove chart from worksheet
```python
#Get the first worksheet from the workbook
sheet = workbook.Worksheets[0]
#Get the first chart from the first worksheet
chart = sheet.Charts[0]
#Remove the chart
chart.Remove()
```

---

# spire.xls python chart manipulation
## resize and move chart in excel
```python
#Get the chart from the worksheet
chart = sheet.Charts[0]
#Set position of the chart
chart.LeftColumn = 5
chart.TopRow = 1
#Resize the chart
chart.Width = 500
chart.Height = 350
```

---

# Spire.XLS Python Chart Data Labels
## Apply rich text formatting to chart data labels
```python
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Get the first chart inside this worksheet
chart = worksheet.Charts[0]
#Get the first datalabel of the first series 
datalabel = chart.Series[0].DataPoints[0].DataLabels
#Set the text
datalabel.Text = "Rich Text Label"
#Show the value
chart.Series[0].DataPoints[0].DataLabels.HasValue = True
#Set styles for the text
chart.Series[0].DataPoints[0].DataLabels.Color = Color.get_Red()
chart.Series[0].DataPoints[0].DataLabels.IsBold = True
```

---

# spire.xls python 3D chart rotation
## rotate 3D chart by setting X and Y rotation values
```python
#Get the chart from the first worksheet
sheet = workbook.Worksheets[0]
chart = sheet.Charts[0]
#X rotation:
chart.Rotation = 30
#Y rotation:
chart.Elevation = 20
```

---

# Creating Scatter Chart with Spire.XLS for Python
## Core functionality for creating a scatter chart with trend line
```python
#Add a chart
chart = sheet.Charts.Add(ExcelChartType.ScatterMarkers)
#Set region of chart data
chart.DataRange = sheet.Range["B2:B10"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 11
chart.RightColumn = 10
chart.BottomRow = 28
chart.ChartTitle = "Scatter Chart"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.Series[0].CategoryLabels = sheet.Range["A2:A10"]
chart.Series[0].Values = sheet.Range["B2:B10"]
#Add a trend line for the first series
chart.Series[0].TrendLines.Add(TrendLineType.Exponential)
chart.PrimaryValueAxis.Title = "Salary"
chart.PrimaryCategoryAxis.Title = "Car Price"
```

---

# spire.xls python chart data labels
## set and format data labels in chart
```python
# Create a chart
chart = sheet.Charts.Add(ExcelChartType.LineMarkers)
chart.DataRange = sheet.Range["B1:B7"]
chart.PlotArea.Visible = False
chart.SeriesDataFromRange = False
chart.TopRow = 5
chart.BottomRow = 26
chart.LeftColumn = 2
chart.RightColumn = 11
chart.ChartTitle = "Data Labels Demo"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12

# Configure data labels
cs1 = chart.Series[0]
cs1.CategoryLabels = sheet.Range["A2:A7"]
cs1.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
cs1.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = False
cs1.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = False
cs1.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = True
cs1.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = True
cs1.DataPoints.DefaultDataPoint.DataLabels.Delimiter = ". "
cs1.DataPoints.DefaultDataPoint.DataLabels.Size = 9
cs1.DataPoints.DefaultDataPoint.DataLabels.Color = Color.get_Red()
cs1.DataPoints.DefaultDataPoint.DataLabels.FontName = "Calibri"
cs1.DataPoints.DefaultDataPoint.DataLabels.Position = DataLabelPositionType.Center
```

---

# spire.xls python chart border styling
## Set border color and style for chart series
```python
#Set CustomLineWeight property for Series line
( chart.Series[0].DataPoints[0].DataFormat.LineProperties if isinstance(chart.Series[0].DataPoints[0].DataFormat.LineProperties, XlsChartBorder) else None).CustomLineWeight = 2.5
#Set color property for Series line
( chart.Series[0].DataPoints[0].DataFormat.LineProperties if isinstance(chart.Series[0].DataPoints[0].DataFormat.LineProperties, XlsChartBorder) else None).Color = Color.get_Red()
```

---

# spire.xls python chart
## Set border width of chart markers
```python
#Get the chart from the first worksheet
chart = workbook.Worksheets[0].Charts[0]
chart.Series[0].DataFormat.MarkerBorderWidth = 1.5 #unit is pt
chart.Series[1].DataFormat.MarkerBorderWidth = 2.5 #unit is pt
```

---

# spire.xls python chart background color
## Set chart background color
```python
#Get the first worksheet from workbook and then get the first chart from the worksheet
ws = workbook.Worksheets[0]
chart = ws.Charts[0]
#Set background color
chart.ChartArea.ForeGroundColor = Color.get_LightYellow()
```

---

# spire.xls python chart
## Set color for chart area
```python
#Get the chart
chart = sheet.Charts[0]
#Set color for chart area
chart.ChartArea.Fill.ForeColor = Color.get_LightSeaGreen()
#Set color for plot area
chart.PlotArea.Fill.ForeColor = Color.get_LightGray()
```

---

# spire.xls python chart font setting
## Set font for chart data labels
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first sheet
chart = sheet.Charts[0]
#Create a font
font = workbook.CreateFont()
font.Size = 15.0
font.Color = Color.get_LightSeaGreen()
for cs in chart.Series:
    #Set font
    cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font)
```

---

# spire.xls python font formatting
## Set font for chart legend and data table
```python
#Create a font with specified size and color
font = workbook.CreateFont()
font.Size = 14.0
font.Color = Color.get_Red()
#Apply the font to chart Legend
chart.Legend.TextArea.SetFont(font)
#Apply the font to chart DataLabel
for cs in chart.Series:
    cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font)
```

---

# spire.xls python chart
## set font for chart title and axis
```python
#Set font for chart title and chart axis
worksheet = workbook.Worksheets[0]
chart = worksheet.Charts[0]
#Format the font for the chart title
chart.ChartTitleArea.Color = Color.get_Blue()
chart.ChartTitleArea.Size = 20.0
#Format the font for the chart Axis
chart.PrimaryValueAxis.Font.Color = Color.get_Gold()
chart.PrimaryValueAxis.Font.Size = 10.0
chart.PrimaryCategoryAxis.Font.Color = Color.get_Red()
chart.PrimaryCategoryAxis.Font.Size = 20.0
```

---

# spire.xls python chart legend
## set background color for chart legend
```python
ws = workbook.Worksheets[0]
chart = ws.Charts[0]
x = chart.Legend.FrameFormat if isinstance(chart.Legend.FrameFormat, XlsChartFrameFormat) else None
x.Fill.FillType = ShapeFillType.SolidColor
x.ForeGroundColor = Color.get_SkyBlue()
```

---

# Spire.XLS Python Chart
## Set number format of trendline
```python
#Get the chart from the first worksheet
chart = workbook.Worksheets[0].Charts[0]
#Get the trendline of the chart
trendLine = chart.Series[1].TrendLines[0]
#Set the number format of trendLine to "#,##0.00"
trendLine.DataLabel.NumberFormat = "#,##0.00"
```

---

# spire.xls python chart
## show leader lines in chart
```python
#Create a chart
chart = sheet.Charts.Add(ExcelChartType.BarStacked)
chart.DataRange = sheet.Range["A1:C3"]
chart.TopRow = 4
chart.LeftColumn = 2
chart.Width = 450
chart.Height = 300
#Show leader lines for data labels
for cs in chart.Series:
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
    cs.DataPoints.DefaultDataPoint.DataLabels.ShowLeaderLines = True
```

---

# spire.xls python sparkline
## create line sparklines in Excel
```python
#Add sparkline
sparklineGroup = sheet.SparklineGroups.AddGroup(SparklineType.Line)
sparklines = sparklineGroup.Add()
sparklines.Add(sheet["A2:D2"], sheet["E2"])
sparklines.Add(sheet["A3:D3"], sheet["E3"])
sparklines.Add(sheet["A4:D4"], sheet["E4"])
sparklines.Add(sheet["A5:D5"], sheet["E5"])
sparklines.Add(sheet["A6:D6"], sheet["E6"])
sparklines.Add(sheet["A7:D7"], sheet["E7"])
sparklines.Add(sheet["A8:D8"], sheet["E8"])
sparklines.Add(sheet["A9:D9"], sheet["E9"])
sparklines.Add(sheet["A10:D10"], sheet["E10"])
sparklines.Add(sheet["A11:D11"], sheet["E11"])
```

---

# spire.xls python stacked column chart
## create 2D stacked column chart
```python
#Add a chart
chart = sheet.Charts.Add()
#Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
chart.ChartType = ExcelChartType.ColumnStacked
#Chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
#Chart Axes
chart.PrimaryCategoryAxis.Title = "Country"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
for cs in chart.Series:
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
#Chart Legend
chart.Legend.Position = LegendPositionType.Top
```

## create 3D stacked column chart
```python
#Add a chart
chart = sheet.Charts.Add()
#Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
chart.ChartType = ExcelChartType.Column3DStacked
#Chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
#Chart Axes
chart.PrimaryCategoryAxis.Title = "Country"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
for cs in chart.Series:
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
#Chart Legend
chart.Legend.Position = LegendPositionType.Top
```

---

# Adding Arrow Lines to Excel
## Demonstrates how to add various types of arrow lines to an Excel worksheet
```python
#Add a Double Arrow and fill the line with solid color.
line = sheet.TypedLines.AddLine()
line.Top = 10
line.Left = 20
line.Width = 100
line.Height = 0
line.Color = Color.get_Blue()
line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow
line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
#Add an Arrow and fill the line with solid color.
line_1 = sheet.TypedLines.AddLine()
line_1.Top = 50
line_1.Left = 30
line_1.Width = 100
line_1.Height = 100
line_1.Color = Color.get_Red()
line_1.BeginArrowHeadStyle = ShapeArrowStyleType.LineNoArrow
line_1.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
#Add an Elbow Arrow Connector.
line3 = sheet.TypedLines.AddLine() if isinstance(sheet.TypedLines.AddLine(), XlsLineShape) else None
line3.LineShapeType = LineShapeType.ElbowLine
line3.Width = 30
line3.Height = 50
line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
line3.Top = 100
line3.Left = 50
#Add an Elbow Double-Arrow Connector.
line2 = sheet.TypedLines.AddLine() if isinstance(sheet.TypedLines.AddLine(), XlsLineShape) else None
line2.LineShapeType = LineShapeType.ElbowLine
line2.Width = 50
line2.Height = 50
line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow
line2.Left = 120
line2.Top = 100
#Add a Curved Arrow Connector.
line3 = sheet.TypedLines.AddLine() if isinstance(sheet.TypedLines.AddLine(), XlsLineShape) else None
line3.LineShapeType = LineShapeType.CurveLine
line3.Width = 30
line3.Height = 50
line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
line3.Top = 100
line3.Left = 200
#Add a Curved Double-Arrow Connector.
line2 = sheet.TypedLines.AddLine() if isinstance(sheet.TypedLines.AddLine(), XlsLineShape) else None
line2.LineShapeType = LineShapeType.CurveLine
line2.Width = 30
line2.Height = 50
line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
line2.Left = 250
line2.Top = 100
```

---

# spire.xls python shapes
## add line shapes to Excel worksheet
```python
#Assuming sheet is a worksheet object from an Excel workbook
#Add shape line1
line1 = sheet.Lines.AddLine(10, 2, 200, 1, LineShapeType.Line)
#Set dash style type
line1.DashStyle = ShapeDashLineStyleType.Solid
#Set color
line1.Color = Color.get_CadetBlue()
#Set weight
line1.Weight = 2
#Set end arrow style type
line1.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
#Add shape line2
line2 = sheet.Lines.AddLine(12, 2, 200, 1, LineShapeType.CurveLine)
line2.DashStyle = ShapeDashLineStyleType.Dotted
line2.Color = Color.get_OrangeRed()
line2.Weight = 2
#Add shape line3
line3 = sheet.Lines.AddLine(14, 2, 200, 1, LineShapeType.ElbowLine)
line3.DashStyle = ShapeDashLineStyleType.DashDotDot
line3.Color = Color.get_Purple()
line3.Weight = 2
#Add shape line4
line4 = sheet.Lines.AddLine(16, 2, 200, 1, LineShapeType.LineInv)
line4.DashStyle = ShapeDashLineStyleType.Dashed
line4.Color = Color.get_Green()
line4.Weight = 2
```

---

# Adding Oval Shapes to Excel Worksheet
## demonstrates how to add oval shapes with different fill styles to an Excel worksheet
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Add oval shape1
ovalShape1 = sheet.OvalShapes.AddOval(11, 2, 100, 100)
ovalShape1.Line.Weight = 0
#Fill shape with solid color
ovalShape1.Fill.FillType = ShapeFillType.SolidColor
ovalShape1.Fill.ForeColor = Color.get_DarkCyan()
#Add oval shape2
ovalShape2 = sheet.OvalShapes.AddOval(11, 5, 100, 100)
ovalShape2.Line.Weight = 1
#Fill shape with picture
ovalShape2.Line.DashStyle = ShapeDashLineStyleType.Solid
ovalShape2.Fill.CustomPicture(inputimage)
```

---

# spire.xls python shape
## Add rectangle shapes to Excel worksheet
```python
#Add rectangle shape 1------Rect
rect1 = sheet.RectangleShapes.AddRectangle(11, 2, 60, 100, RectangleShapeType.Rect)
rect1.Line.Weight = 1
#Fill shape with solid color
rect1.Fill.FillType = ShapeFillType.SolidColor
rect1.Fill.ForeColor = Color.get_DarkGreen()
#Add rectangle shape 2------RoundRect
rect2 = sheet.RectangleShapes.AddRectangle(11, 5, 60, 100, RectangleShapeType.RoundRect)
rect2.Line.Weight = 1
rect2.Fill.FillType = ShapeFillType.SolidColor
rect2.Fill.ForeColor = Color.get_DarkCyan()
```

---

# spire.xls python spinner control
## add spinner control to excel worksheet
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set text for range C11
sheet.Range["C11"].Text = "Value:"
sheet.Range["C11"].Style.Font.IsBold = True
#Set value for range B10
sheet.Range["C12"].Value2 = Int32(0)
#Add spinner control
spinner = sheet.SpinnerShapes.AddSpinner(12, 4, 20, 20)
spinner.LinkedCell = sheet.Range["C12"]
spinner.Min = 0
spinner.Max = 100
spinner.IncrementalChange = 5
spinner.Display3DShading = True
```

---

# spire.xls python arrow polyline
## adjust arrow polyline position in excel
```python
# Draw an elbow arrow
line = worksheet.TypedLines.AddLine(5, 5, 100, 100, LineShapeType.ElbowLine)
line.EndArrowHeadStyle = ShapeArrowStyleType.LineNoArrow
line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow
ad = line.ShapeAdjustValues.AddAdjustValue(GeomertyAdjustValueFormulaType.LiteralValue)
# When the parameter value is less than 0, the focus of the line is on the left side of the left point, 
# when it is equal to 0, the position is the same as the left point, it is equal to 50 in the middle of the graph, 
# and when it is equal to 100, it is the same as the right point.
ad.SetFormulaParameter([-50])
```

---

# Spire.XLS Python Copy Shapes
## Copy various shapes between worksheets in Excel
```python
workbook = Workbook()
sheet = workbook.Worksheets[0]
#Create line shape
line = sheet.TypedLines.AddLine()
line.Top = 50
line.Left = 30
line.Width = 30
line.Height = 50
line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowDiamond
line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
CopyShapes = workbook.Worksheets[1]
#Copy the line into other sheet
CopyShapes.TypedLines.AddCopy(line)
#Create a button and then copy into other sheet
button = sheet.TypedRadioButtons.Add(5, 5, 20, 20)
CopyShapes.TypedRadioButtons.AddCopy(button)
#Create a textbox and then copy into other sheet
textbox = sheet.TypedTextBoxes.AddTextBox(5, 7, 50, 100)
CopyShapes.TypedTextBoxes.AddCopy(textbox)
#Create a checkbox and then copy into other sheet
checkbox = sheet.TypedCheckBoxes.AddCheckBox(10, 1, 20, 20)
CopyShapes.TypedCheckBoxes.AddCopy(checkbox)
#Create a comboboxes and then copy into other sheet
sheet.Range["A14"].Value = "1"
sheet.Range["A15"].Value = "2"
ComboBoxes = sheet.TypedComboBoxes.AddComboBox(10, 5, 30, 30)
ComboBoxes.ListFillRange = sheet.Range["A14:A15"]
CopyShapes.TypedComboBoxes.AddCopy(ComboBoxes)
```

---

# spire.xls python shapes
## delete all shapes in worksheet
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Delete all shapes in the worksheet
for i in range(sheet.PrstGeomShapes.Count - 1, -1, -1):
    sheet.PrstGeomShapes[i].Remove()
```

---

# Delete Particular Shape in Excel
## Delete the first shape from an Excel worksheet
```python
# Get the first worksheet
sheet = workbook.Worksheets[0]
# Delete the first shape in the worksheet
sheet.PrstGeomShapes[0].Remove()
```

---

# Extract Text and Image from Excel Shapes
## This code demonstrates how to extract text and images from shapes in an Excel worksheet

```python
sheet = workbook.Worksheets[0]

# Extract text from the third shape
shape = sheet.PrstGeomShapes[2]
text = shape.Text

# Extract image from the second shape
shape = sheet.PrstGeomShapes[1]
image = shape.Fill.Picture
```

---

# Get Shape Linked Cell Range in Excel
## Core functionality to retrieve the cell range linked to shapes in an Excel worksheet
```python
#get the first worksheet
sheet=workbook.Worksheets[0]

#get PrstGeomShapes from sheet
prstGeomShapeCollection = sheet.PrstGeomShapes

#get shape
shape = prstGeomShapeCollection["Yesterday"]

#get shape linked cell range
cellAddress = shape.LinkedCell.RangeAddress

shape = prstGeomShapeCollection["NewShapes"]
cellAddress = shape.LinkedCell.RangeAddress
```

---

# spire.xls python shape visibility
## hide or unhide shapes in Excel worksheet
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Hide the second shape in the worksheet
sheet.PrstGeomShapes[1].Visible = False
#Show the second shape in the worksheet
#sheet.PrstGeomShapes[1].Visible = true
```

---

# spire.xls python shape insertion
## insert different types of shapes into an Excel worksheet with various fill styles
```python
#Add a triangle shape.
triangle = sheet.PrstGeomShapes.AddPrstGeomShape(2, 2, 100, 100, PrstGeomShapeType.Triangle)
#Fill the triangle with solid color.
triangle.Fill.ForeColor = Color.get_Yellow()
triangle.Fill.FillType = ShapeFillType.SolidColor
#Add a heart shape.
heart = sheet.PrstGeomShapes.AddPrstGeomShape(2, 5, 100, 100, PrstGeomShapeType.Heart)
#Fill the heart with gradient color.
heart.Fill.ForeColor = Color.get_Red()
heart.Fill.FillType = ShapeFillType.Gradient
#Add an arrow shape with default color.
arrow = sheet.PrstGeomShapes.AddPrstGeomShape(10, 2, 100, 100, PrstGeomShapeType.CurvedRightArrow)
#Add a cloud shape.
cloud = sheet.PrstGeomShapes.AddPrstGeomShape(10, 5, 100, 100, PrstGeomShapeType.Cloud)
#Fill the cloud with picture
cloud.Fill.FillType = ShapeFillType.Picture
```

---

# Modify Shadow Style for Shape
## This code demonstrates how to modify the shadow style properties of a shape in an Excel worksheet
```python
#Get the third shape from the worksheet.
shape = sheet.PrstGeomShapes[2]
#Set the shadow style for the shape.
shape.Shadow.Angle = 90
shape.Shadow.Transparency = 30
shape.Shadow.Distance = 10
shape.Shadow.Size = 130
shape.Shadow.Color = Color.get_Yellow()
shape.Shadow.Blur = 30
shape.Shadow.HasCustomStyle = True
```

---

# spire.xls python shape shadow style
## set shadow style for ellipse shape in Excel
```python
#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Add an ellipse shape.
ellipse = sheet.PrstGeomShapes.AddPrstGeomShape(5, 5, 150, 100, PrstGeomShapeType.Ellipse)
#Set the shadow style for the ellipse.
ellipse.Shadow.Angle = 90
ellipse.Shadow.Distance = 10
ellipse.Shadow.Size = 150
ellipse.Shadow.Color = Color.get_Gray()
ellipse.Shadow.Blur = 30
ellipse.Shadow.Transparency = 1
ellipse.Shadow.HasCustomStyle = True
```

---

# spire.xls python shape order
## change the order of shapes in excel worksheets
```python
#Bring the picture forward one level
wb.Worksheets[0].Pictures[0].ChangeLayer(ShapeLayerChangeType.BringForward)
#Bring the image in front of all other objects
wb.Worksheets[1].Pictures[0].ChangeLayer(ShapeLayerChangeType.BringToFront)
#Send the shape back one level
shape = wb.Worksheets[2].PrstGeomShapes[1] if isinstance(wb.Worksheets[2].PrstGeomShapes[1], XlsShape) else None
shape.ChangeLayer(ShapeLayerChangeType.SendBackward)
#Send the shape behind all other objects
shape = wb.Worksheets[3].PrstGeomShapes[1] if isinstance(wb.Worksheets[3].PrstGeomShapes[1], XlsShape) else None
shape.ChangeLayer(ShapeLayerChangeType.SendToBack)
```

---

# Shape to Image Conversion
## Convert Excel shapes to image format
```python
#Get the first worksheet
sheet1 = wb.Worksheets[0]
#Get the first shape from the first worksheet
shape = sheet1.PrstGeomShapes[0] if isinstance(sheet1.PrstGeomShapes[0], XlsShape) else None
#Save the shape to a image
img = shape.SaveToImage()
```

---

# Spire.XLS Python Shape Texture
## Tile picture as texture in a shape
```python
#Get a shape from worksheet
shape = sheet.PrstGeomShapes[0]
#Set fill type to texture
shape.Fill.FillType = ShapeFillType.Texture
#Apply custom picture as texture
shape.Fill.CustomTexture(image_path)
#Enable tiling of the picture as texture
shape.Fill.Tile = True
```

---

# spire.xls python style formatting
## Apply built-in styles to Excel cells
```python
#Apply title style
sheet.Range["A1:J1"].BuiltInStyle = BuiltInStyles.Title
```

---

# spire.xls python color scales
## apply color scales to data range using conditional formatting
```python
#Add color scales.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(sheet.AllocatedRange)
format = xcfs.AddCondition()
format.FormatType = ConditionalFormatType.ColorScale
```

---

# Apply Conditional Formatting in Excel
## This code demonstrates how to apply conditional formatting to Excel cells based on cell values
```python
#Create conditional formatting rule.
xcfs1 = sheet.ConditionalFormats.Add()
xcfs1.AddRange(sheet.AllocatedRange)
format1 = xcfs1.AddCondition()
format1.FormatType = ConditionalFormatType.CellValue
format1.FirstFormula = "800"
format1.Operator = ComparisonOperatorType.Greater
format1.FontColor = Color.get_Red()
format1.BackColor = Color.get_LightSalmon()

#Create conditional formatting rule.
xcfs2 = sheet.ConditionalFormats.Add()
xcfs2.AddRange(sheet.AllocatedRange)
format2 = xcfs1.AddCondition()
format2.FormatType = ConditionalFormatType.CellValue
format2.FirstFormula = "300"
format2.Operator = ComparisonOperatorType.Less
format2.FontColor = Color.get_Green()
format2.BackColor = Color.get_LightBlue()
```

---

# Spire.XLS Python Data Bars
## Apply data bars to cell range in Excel
```python
#Add data bars.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(sheet.AllocatedRange)
format = xcfs.AddCondition()
format.FormatType = ConditionalFormatType.DataBar
format.DataBar.BarColor = Color.get_CadetBlue()
```

---

# spire.xls python gradient fill
## apply gradient fill effects to excel cells
```python
#Create a workbook
workbook = Workbook()
workbook.Version = ExcelVersion.Version2010
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get "B5" cell
range = sheet.Range["B5"]
#Set alignment style
range.Style.HorizontalAlignment = HorizontalAlignType.Center
#Set gradient filling effects
range.Style.Interior.FillPattern = ExcelPatternType.Gradient
range.Style.Interior.Gradient.ForeColor = Color.FromRgb(255, 255, 255)
range.Style.Interior.Gradient.BackColor = Color.FromRgb(79, 129, 189)
range.Style.Interior.Gradient.TwoColorGradient(GradientStyleType.Horizontal, GradientVariantsType.ShadingVariants1)
```

---

# spire.xls python conditional formatting
## apply icon sets to cell range
```python
# Add icon sets
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(sheet.AllocatedRange)
format = xcfs.AddCondition()
format.FormatType = ConditionalFormatType.IconSet
format.IconSet.IconSetType = IconSetType.ThreeTrafficLights1
```

---

# Excel Color Palette Management
## Add custom colors to Excel palette and apply to cells
```python
#Create a workbook
workbook = Workbook()
#Adding Orchid color to the palette at 60th index
workbook.ChangePaletteColor(Color.get_Orchid(), 60)
#Get the first sheet
sheet = workbook.Worksheets[0]
cell = sheet.Range["B2"]
cell.Text = "Welcome to use Spire.XLS"
#Set the Orchid (custom) color to the font
cell.Style.Font.Color = Color.get_Orchid()
cell.Style.Font.Size = 20
cell.AutoFitColumns()
cell.AutoFitRows()
```

---

# spire.xls conditional formatting
## runtime conditional formatting for Excel cells
```python
def AddComparisonRule1(sheet):
    #Create conditional formatting rule
    xcfs1 = sheet.ConditionalFormats.Add()
    xcfs1.AddRange(sheet.Range["A1:D1"])
    cf1 = xcfs1.AddCondition()
    cf1.FormatType = ConditionalFormatType.CellValue
    cf1.FirstFormula = "150"
    cf1.Operator = ComparisonOperatorType.Greater
    cf1.FontColor = Color.get_Red()
    cf1.BackColor = Color.get_LightBlue()

def AddComparisonRule2(sheet):
    xcfs2 = sheet.ConditionalFormats.Add()
    xcfs2.AddRange(sheet.Range["A2:D2"])
    cf2 = xcfs2.AddCondition()
    cf2.FormatType = ConditionalFormatType.CellValue
    cf2.FirstFormula = "500"
    cf2.Operator = ComparisonOperatorType.Less
    #Set border color
    cf2.LeftBorderColor = Color.get_Pink()
    cf2.RightBorderColor = Color.get_Pink()
    cf2.TopBorderColor = Color.get_DeepSkyBlue()
    cf2.BottomBorderColor = Color.get_DeepSkyBlue()
    cf2.LeftBorderStyle = LineStyleType.Medium
    cf2.RightBorderStyle = LineStyleType.Thick
    cf2.TopBorderStyle = LineStyleType.Double
    cf2.BottomBorderStyle = LineStyleType.Double

def AddComparisonRule3(sheet):
    #Create conditional formatting rule
    xcfs1 = sheet.ConditionalFormats.Add()
    xcfs1.AddRange(sheet.Range["A3:D3"])
    cf1 = xcfs1.AddCondition()
    cf1.FormatType = ConditionalFormatType.CellValue
    cf1.FirstFormula = "300"
    cf1.SecondFormula = "500"
    cf1.Operator = ComparisonOperatorType.Between
    cf1.BackColor = Color.get_Yellow()

def AddComparisonRule4(sheet):
    #Create conditional formatting rule
    xcfs1 = sheet.ConditionalFormats.Add()
    xcfs1.AddRange(sheet.Range["A4:D4"])
    cf1 = xcfs1.AddCondition()
    cf1.FormatType = ConditionalFormatType.CellValue
    cf1.FirstFormula = "100"
    cf1.SecondFormula = "200"
    cf1.Operator = ComparisonOperatorType.NotBetween
    #Set fill pattern type
    cf1.FillPattern = ExcelPatternType.ReverseDiagonalStripe
    #Set foreground color
    cf1.Color = Color.FromRgb(255, 255, 0)
    #Set background color
    cf1.BackColor = Color.FromRgb(0, 255, 255)
```

---

# spire.xls python conditional date formatting
## format cells containing dates from the last 7 days
```python
#Highlight cells that contain a date occurring in the last 7 days.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(sheet.AllocatedRange)
conditionalFormat = xcfs.AddTimePeriodCondition(TimePeriodType.Last7Days)
conditionalFormat.BackColor = Color.get_Orange()
```

---

# spire.xls python conditional formatting
## create formula-based conditional formatting in Excel
```python
#Get the first worksheet and the first column from the workbook.
sheet = workbook.Worksheets[0]
range = sheet.Columns[0]
#Set the conditional formatting formula and apply the rule to the chosen cell range.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(range)
conditional = xcfs.AddCondition()
conditional.FormatType = ConditionalFormatType.Formula
conditional.FirstFormula = "=($A1<$B1)"
conditional.BackKnownColor = ExcelColors.Yellow
```

---

# spire.xls python font styles
## apply various font styles to excel cells
```python
#Set font style
sheet.Range["B1"].Style.Font.FontName = "Comic Sans MS"
sheet.Range["B2:D2"].Style.Font.FontName = "Corbel"
sheet.Range["B3:D7"].Style.Font.FontName = "Aleo"
#Set font size
sheet.Range["B1"].Style.Font.Size = 45
sheet.Range["B2:D3"].Style.Font.Size = 25
sheet.Range["B3:D7"].Style.Font.Size = 12
#Set excel cell data to be bold
sheet.Range["B2:D2"].Style.Font.IsBold = True
#Set excel cell data to be underline
sheet.Range["B3:B7"].Style.Font.Underline = FontUnderlineType.Single
#set excel cell data color
sheet.Range["B1"].Style.Font.Color = Color.get_CornflowerBlue()
sheet.Range["B2:D2"].Style.Font.Color = Color.get_CadetBlue()
sheet.Range["B3:D7"].Style.Font.Color = Color.get_Firebrick()
#set excel cell data to be italic
sheet.Range["B3:D7"].Style.Font.IsItalic = True
#Add strikethrough
sheet.Range["D3"].Style.Font.IsStrikethrough = True
sheet.Range["D7"].Style.Font.IsStrikethrough = True
```

---

# spire.xls python formatting
## set foreground and background colors for Excel cells
```python
#Create a new style
style = workbook.Styles.Add("newStyle1")
#Set filling pattern type
style.Interior.FillPattern = ExcelPatternType.Gradient
#Set filling Background color
style.Interior.Gradient.BackKnownColor = ExcelColors.Green
#Set filling Foreground color
style.Interior.Gradient.ForeKnownColor = ExcelColors.Yellow
#set gradient style
style.Interior.Gradient.GradientStyle = GradientStyleType.From_Center
#Apply the style to  "B2" cell
sheet.Range["B2"].CellStyleName = style.Name
sheet.Range["B2"].Text = "Test"
sheet.Range["B2"].RowHeight = 30
sheet.Range["B2"].ColumnWidth = 50

#Create a new style
style = workbook.Styles.Add("newStyle2")
#Set filling pattern type
style.Interior.FillPattern = ExcelPatternType.Gradient
#Set filling Foreground color
style.Interior.Gradient.ForeKnownColor = ExcelColors.Red
#Apply the style to  "B4" cell
sheet.Range["B4"].CellStyleName = style.Name
sheet.Range["B4"].RowHeight = 30
sheet.Range["B4"].ColumnWidth = 60
```

---

# Excel Column Formatting
## Create and apply a style to format a column in Excel
```python
#Create a workbook
workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
#Create a new style
style = workbook.Styles.Add("newStyle")
#Set the vertical alignment of the text
style.VerticalAlignment = VerticalAlignType.Center
#Set the horizontal alignment of the text
style.HorizontalAlignment = HorizontalAlignType.Center
#Set the font color of the text
style.Font.Color = Color.get_Blue()
#Shrink the text to fit in the cell
style.ShrinkToFit = True
#Set the bottom border color of the cell to OrangeRed
style.Borders[BordersLineType.EdgeBottom].Color = Color.get_OrangeRed()
#Set the bottom border type of the cell to Dotted
style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Dotted
#Apply the style to the first column
sheet.Columns[0].CellStyleName = style.Name
sheet.Columns[0].Text = "Test"
```

---

# Spire.XLS Python Row Formatting
## Format a row in Excel with custom style
```python
#Create a new style
style = workbook.Styles.Add("newStyle")
style.VerticalAlignment = VerticalAlignType.Center
style.HorizontalAlignment = HorizontalAlignType.Center
style.Font.Color = Color.get_Blue()
style.ShrinkToFit = True
style.Borders[BordersLineType.EdgeBottom].Color = Color.get_OrangeRed()
style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Dotted

#Apply the style to the second row
sheet.Rows[1].CellStyleName = style.Name
sheet.Rows[1].Text = "Test"
```

---

# spire.xls python cell formatting
## Format Excel cells with custom style including color, font, rotation, and alignment
```python
#Create a style
style = workbook.Styles.Add("newStyle")
#Set the shading color
style.Color = Color.get_DarkGray()
#Set the font color
style.Font.Color = Color.get_White()
#Set font name
style.Font.FontName = "Times New Roman"
#Set font size
style.Font.Size = 12
#Set bold for the font
style.Font.IsBold = True
#Set text rotation
style.Rotation = 45
#Set alignment
style.HorizontalAlignment = HorizontalAlignType.Center
style.VerticalAlignment = VerticalAlignType.Center
#Set the style for the specific range
workbook.Worksheets[0].Range["A1:J1"].CellStyleName = style.Name
```

---

# Get Color ARGB Data
## Extract ARGB color values from Excel cells
```python
# Get font color from cell B2
color1 = sheet.Range["B2"].Style.Font.Color
# Format ARGB data as string
argb1 = "The font color of B2: ARGB=({0},{1},{2},{3})".format(color1.A, color1.R, color1.G, color1.B)

# Get font color from cell B3
color2 = sheet.Range["B3"].Style.Font.Color
argb2 = "The font color of B3: ARGB=({0},{1},{2},{3})".format(color2.A, color2.R, color2.G, color2.B)

# Get font color from cell B4
color3 = sheet.Range["B4"].Style.Font.Color
argb3 = "The font color of B4: ARGB=({0},{1},{2},{3})".format(color3.A, color3.R, color3.G, color3.B)
```

---

# spire.xls python style
## get and set cell style
```python
#Get "B4" cell
range = sheet.Range["B4"]
#Get the style of cell
style = range.Style
style.Font.FontName = "Calibri"
style.Font.IsBold = True
style.Font.Size = 15
style.Font.Color = Color.get_CornflowerBlue()
range.Style = style
```

---

# Highlight Average Values in Excel
## This code demonstrates how to highlight cells with above and below average values in an Excel worksheet using conditional formatting.
```python
#Add conditional format.
format1 = sheet.ConditionalFormats.Add()
#Set the cell range to apply the formatting.
format1.AddRange(sheet.Range["E2:E10"])
#Add below average condition.
cf1 = format1.AddAverageCondition(AverageType.Below)
#Highlight cells below average values.
cf1.BackColor = Color.get_SkyBlue()
#Add conditional format.
format2 = sheet.ConditionalFormats.Add()
#Set the cell range to apply the formatting.
format2.AddRange(sheet.Range["E2:E10"])
#Add above average condition.
cf2 = format1.AddAverageCondition(AverageType.Above)
#Highlight cells above average values.
cf2.BackColor = Color.get_Orange()
```

---

# Spire.XLS for Python - Conditional Formatting
## Highlight duplicate and unique values in Excel cells
```python
#Use conditional formatting to highlight duplicate values in range "C2:C10" with IndianRed color.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(sheet.Range["C2:C10"])
format1 = xcfs.AddCondition()
format1.FormatType = ConditionalFormatType.DuplicateValues
format1.BackColor = Color.get_IndianRed()
#Use conditional formatting to highlight unique values in range "C2:C10" with Yellow color.
xcfs1 = sheet.ConditionalFormats.Add()
xcfs1.AddRange(sheet.Range["C2:C10"])
format2 = xcfs.AddCondition()
format2.FormatType = ConditionalFormatType.UniqueValues
format2.BackColor = Color.get_Yellow()
```

---

# spire.xls conditional formatting
## highlight top and bottom ranked values in Excel ranges
```python
#Apply conditional formatting to range "D2:D10" to highlight the top 2 values.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(sheet.Range["D2:D10"])
format1 = xcfs.AddTopBottomCondition(TopBottomType.Top, 2)
format1.FormatType = ConditionalFormatType.TopBottom
format1.BackColor = Color.get_Red()
#Apply conditional formatting to range "E2:E10" to highlight the bottom 2 values.
xcfs1 = sheet.ConditionalFormats.Add()
xcfs1.AddRange(sheet.Range["E2:E10"])
format2 = xcfs1.AddTopBottomCondition(TopBottomType.Bottom, 2)
format2.FormatType = ConditionalFormatType.TopBottom
format2.BackColor = Color.get_ForestGreen()
```

---

# Spire.XLS Python Cell Indentation
## Set indentation level for Excel cell text
```python
#Access the "B5" cell from the worksheet
cell = sheet.Range["B5"]
cell.Style.IndentLevel = 2
```

---

# Excel Cell Interior Formatting
## Apply gradient fill to Excel cell interiors
```python

#Set cell interior with gradient fill
sheet.Range["A1"].Style.Interior.FillPattern = ExcelPatternType.Gradient
sheet.Range["A1"].Style.Interior.Gradient.BackKnownColor = ExcelColors.Blue
sheet.Range["A1"].Style.Interior.Gradient.ForeKnownColor = ExcelColors.White
sheet.Range["A1"].Style.Interior.Gradient.GradientStyle = GradientStyleType.Vertical
sheet.Range["A1"].Style.Interior.Gradient.GradientVariant = GradientVariantsType.ShadingVariants1
```

---

# spire.xls python make cell active
## Set active cell and adjust visible area in Excel worksheet
```python
#Get the 2nd sheet
sheet = workbook.Worksheets[1]
#Set the 2nd sheet as an active sheet.
sheet.Activate()
#Set B2 cell as an active cell in the worksheet.
sheet.SetActiveCell(sheet.Range["B2"])
#Set the B column as the first visible column in the worksheet.
sheet.FirstVisibleColumn = 1
#Set the 2nd row as the first visible row in the worksheet.
sheet.FirstVisibleRow = 1
```

---

# spire.xls python number formatting
## apply different number formats to cells in Excel
```python
# Initialize the worksheet
sheet = workbook.Worksheets[0]

# Input a number value for the specified cell and set the number format
sheet.Range["B10"].Text = "NUMBER FORMATTING"
sheet.Range["B10"].Style.Font.IsBold = True
sheet.Range["B13"].Text = "0"
sheet.Range["C13"].NumberValue = 1234.5678
sheet.Range["C13"].NumberFormat = "0"
sheet.Range["B14"].Text = "0.00"
sheet.Range["C14"].NumberValue = 1234.5678
sheet.Range["C14"].NumberFormat = "0.00"
sheet.Range["B15"].Text = "#,##0.00"
sheet.Range["C15"].NumberValue = 1234.5678
sheet.Range["C15"].NumberFormat = "#,##0.00"
sheet.Range["B16"].Text = "$#,##0.00"
sheet.Range["C16"].NumberValue = 1234.5678
sheet.Range["C16"].NumberFormat = "$#,##0.00"
sheet.Range["B17"].Text = "0;[Red]-0"
sheet.Range["C17"].NumberValue = -1234.5678
sheet.Range["C17"].NumberFormat = "0;[Red]-0"
sheet.Range["B18"].Text = "0.00;[Red]-0.00"
sheet.Range["C18"].NumberValue = -1234.5678
sheet.Range["C18"].NumberFormat = "0.00;[Red]-0.00"
sheet.Range["B19"].Text = "#,##0;[Red]-#,##0"
sheet.Range["C19"].NumberValue = -1234.5678
sheet.Range["C19"].NumberFormat = "#,##0;[Red]-#,##0"
sheet.Range["B20"].Text = "#,##0.00;[Red]-#,##0.000"
sheet.Range["C20"].NumberValue = -1234.5678
sheet.Range["C20"].NumberFormat = "#,##0.00;[Red]-#,##0.00"
sheet.Range["B21"].Text = "0.00E+00"
sheet.Range["C21"].NumberValue = 1234.5678
sheet.Range["C21"].NumberFormat = "0.00E+00"
sheet.Range["B22"].Text = "0.00%"
sheet.Range["C22"].NumberValue = 1234.5678
sheet.Range["C22"].NumberFormat = "0.00%"
sheet.Range["B13:B22"].Style.KnownColor = ExcelColors.Gray25Percent
# AutoFit Column
sheet.AutoFitColumn(2)
sheet.AutoFitColumn(3)
```

---

# spire.xls python border formatting
## Set border styles for Excel cells
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the cell range where you want to apply border style
cr = sheet.Range[sheet.FirstRow,sheet.FirstColumn,sheet.LastRow,sheet.LastColumn]
#Apply border style 
cr.Borders.LineStyle = LineStyleType.Double
cr.Borders[BordersLineType.DiagonalDown].LineStyle = LineStyleType.none
cr.Borders[BordersLineType.DiagonalUp].LineStyle = LineStyleType.none
cr.Borders.Color = Color.get_CadetBlue()
```

---

# Spire.XLS Python DataBar Border
## Set border to existing and new data bars in Excel
```python
#get the databar format 
xcfs = sheet.ConditionalFormats[0]
cf = xcfs[0]
dataBar1 = cf.DataBar
dataBar1.BarBorder.Type = DataBarBorderType.DataBarBorderSolid
dataBar1.BarBorder.Color = Color.get_Red()

#set to new data bar
sheet["E1"].NumberValue = 200
xcfs2 = sheet.ConditionalFormats.Add()
xcfs2.AddRange(sheet.Range["E1"])
cf2 = xcfs2.AddCondition()
cf2.FormatType = ConditionalFormatType.DataBar
cf2.DataBar.BarBorder.Type = DataBarBorderType.DataBarBorderSolid
cf2.DataBar.BarBorder.Color = Color.get_Red()
cf2.DataBar.BarColor = Color.get_GreenYellow()
```

---

# spire.xls conditional formatting with formula
## Set conditional formatting based on formula conditions
```python
#Add ConditionalFormat
xcfs = sheet.ConditionalFormats.Add()
#Define the range
xcfs.AddRange(sheet.Range["B5"])
#Add condition
format = xcfs.AddCondition()
format.FormatType = ConditionalFormatType.CellValue
#If greater than 1000
format.FirstFormula = "1000"
format.Operator = ComparisonOperatorType.Greater
format.BackColor = Color.get_Orange()
```

---

# Excel Row Conditional Formatting
## Apply alternating row colors using conditional formatting in Excel
```python
# Get the first worksheet.
sheet = workbook.Worksheets[0]
# Select the range that you want to format.
dataRange = sheet.AllocatedRange
# Set conditional formatting.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(dataRange)
format1 = xcfs.AddCondition()
# Determines the cells to format.
format1.FirstFormula = "=MOD(ROW(),2)=0"
# Set conditional formatting type
format1.FormatType = ConditionalFormatType.Formula
# Set the color.
format1.BackColor = Color.get_LightSeaGreen()
# Set the backcolor of the odd rows as Yellow.
xcfs1 = sheet.ConditionalFormats.Add()
xcfs1.AddRange(dataRange)
format2 = xcfs.AddCondition()
format2.FirstFormula = "=MOD(ROW(),2)=1"
format2.FormatType = ConditionalFormatType.Formula
format2.BackColor = Color.get_Yellow()
```

---

# spire.xls python traffic lights icons
## set traffic lights icons in Excel cells using conditional formatting
```python
#Add a conditional formatting.
conditional = sheet.ConditionalFormats.Add()
conditional.AddRange(sheet.AllocatedRange)
#Add a conditional formatting of cell range and set its type to IconSet.
format = conditional.AddCondition()
format.FormatType = ConditionalFormatType.IconSet
format.IconSet.IconSetType = IconSetType.ThreeTrafficLights1
```

---

# spire.xls python conditional formatting
## apply various conditional formatting rules to excel cells
```python
def AddConditionalFormattingForExistingSheet(sheet):
    sheet.AllocatedRange.RowHeight = 15
    sheet.AllocatedRange.ColumnWidth = 16
    #Create conditional formatting rule
    xcfs1 = sheet.ConditionalFormats.Add()
    xcfs1.AddRange(sheet.Range["A1:D1"])
    cf1 = xcfs1.AddCondition()
    cf1.FormatType = ConditionalFormatType.CellValue
    cf1.FirstFormula = "150"
    cf1.Operator = ComparisonOperatorType.Greater
    cf1.FontColor = Color.get_Red()
    cf1.BackColor = Color.get_LightBlue()
    
    xcfs2 = sheet.ConditionalFormats.Add()
    xcfs2.AddRange(sheet.Range["A2:D2"])
    cf2 = xcfs2.AddCondition()
    cf2.FormatType = ConditionalFormatType.CellValue
    cf2.FirstFormula = "300"
    cf2.Operator = ComparisonOperatorType.Less
    #Set border color
    cf2.LeftBorderColor = Color.get_Pink()
    cf2.RightBorderColor = Color.get_Pink()
    cf2.TopBorderColor = Color.get_DeepSkyBlue()
    cf2.BottomBorderColor = Color.get_DeepSkyBlue()
    cf2.LeftBorderStyle = LineStyleType.Medium
    cf2.RightBorderStyle = LineStyleType.Thick
    cf2.TopBorderStyle = LineStyleType.Double
    cf2.BottomBorderStyle = LineStyleType.Double
    
    #Add data bars
    xcfs3 = sheet.ConditionalFormats.Add()
    xcfs3.AddRange(sheet.Range["A3:D3"])
    cf3 = xcfs3.AddCondition()
    cf3.FormatType = ConditionalFormatType.DataBar
    cf3.DataBar.BarColor = Color.get_CadetBlue()
    
    #Add icon sets
    xcfs4 = sheet.ConditionalFormats.Add()
    xcfs4.AddRange(sheet.Range["A4:D4"])
    cf4 = xcfs4.AddCondition()
    cf4.FormatType = ConditionalFormatType.IconSet
    cf4.IconSet.IconSetType = IconSetType.ThreeTrafficLights1
    
    #Add color scales
    xcfs5 = sheet.ConditionalFormats.Add()
    xcfs5.AddRange(sheet.Range["A5:D5"])
    cf5 = xcfs5.AddCondition()
    cf5.FormatType = ConditionalFormatType.ColorScale
    
    #Highlight duplicate values in range "A6:D6" with BurlyWood color
    xcfs6 = sheet.ConditionalFormats.Add()
    xcfs6.AddRange(sheet.Range["A6:D6"])
    cf6 = xcfs6.AddCondition()
    cf6.FormatType = ConditionalFormatType.DuplicateValues
    cf6.BackColor = Color.get_BurlyWood()
```

---

# spire.xls python text alignment
## Set vertical and horizontal alignment for Excel cells
```python
#Set the vertical alignment to Top
sheet.Range["B1:C1"].Style.VerticalAlignment = VerticalAlignType.Top
#Set the vertical alignment to Center
sheet.Range["B2:C2"].Style.VerticalAlignment = VerticalAlignType.Center
#Set the vertical alignment of to Bottom
sheet.Range["B3:C3"].Style.VerticalAlignment = VerticalAlignType.Bottom
#Set the horizontal alignment to General
sheet.Range["B4:C4"].Style.HorizontalAlignment = HorizontalAlignType.General
#Set the horizontal alignment of to Left
sheet.Range["B5:C5"].Style.HorizontalAlignment = HorizontalAlignType.Left
#Set the horizontal alignment of to Center
sheet.Range["B6:C6"].Style.HorizontalAlignment = HorizontalAlignType.Center
#Set the horizontal alignment of to Right
sheet.Range["B7:C7"].Style.HorizontalAlignment = HorizontalAlignType.Right
#Set the rotation degree
sheet.Range["B8:C8"].Style.Rotation = 45
sheet.Range["B9:C9"].Style.Rotation = 90
#Set the row height of cell
sheet.Range["B8:C9"].RowHeight = 60
```

---

# spire.xls text direction
## set text reading order in Excel cells
```python
#Access the "B5" cell from the worksheet
cell = sheet.Range["B5"]
#Add some value to the "B5" cell
cell.Text = "Hello Spire!"
#Set the reading order from right to left of the text in the "B5" cell
cell.Style.ReadingOrder = ReadingOrderType.RightToLeft
```

---

# Spire.XLS Python Style Management
## Create and apply predefined styles to Excel cells
```python
#create a workbook
workbook = Workbook()

#get the first worksheet
sheet=workbook.Worksheets[0]

#create a new style
style = workbook.Styles.Add("newStyle")
style.Font.FontName = "Calibri"
style.Font.IsBold = True
style.Font.Size = 15
style.Font.Color = Color.get_CornflowerBlue()

#get "B5" cell
range = sheet.Range["B5"]
range.Text = "Welcome to use Spire.XLS"
range.CellStyleName = style.Name
range.AutoFitColumns()
```

---

# spire.xls python style objects
## demonstrates how to create and apply style objects to Excel cells
```python
#Create a workbook and worksheet
workbook = Workbook()
sheet = workbook.Worksheets.Add("new sheet")
#Access the "B1" cell from the worksheet
cell = sheet.Range["B1"]
cell.Text = "Hello Spire!"
#Create a new style
style = workbook.Styles.Add("newStyle")
#Configure style properties
style.VerticalAlignment = VerticalAlignType.Center
style.HorizontalAlignment = HorizontalAlignType.Center
style.Font.Color = Color.get_Blue()
style.ShrinkToFit = True
style.Borders[BordersLineType.EdgeBottom].Color = Color.get_GreenYellow()
style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Medium
#Apply the style to cells
cell.Style = style
sheet.Range["B4"].Style = style
sheet.Range["B4"].Text = "Test"
sheet.Range["C3"].CellStyleName = style.Name
sheet.Range["C3"].Text = "Welcome to use Spire.XLS"
sheet.Range["D4"].Style = style
```

---

# spire.xls conditional formatting
## demonstrates various conditional formatting types in Excel
```python
# IconSet
def AddIconSet2(sheet):
    xcfs = sheet.ConditionalFormats.Add()
    xcfs.AddRange(sheet.Range["M1:O2"])
    cf = xcfs.AddCondition()
    cf.FormatType = ConditionalFormatType.IconSet
    cf.IconSet.IconSetType = IconSetType.ThreeArrows

def AddIconSet3(sheet):
    xcfs = sheet.ConditionalFormats.Add()
    xcfs.AddRange(sheet.Range["M3:O4"])
    cf = xcfs.AddCondition()
    cf.FormatType = ConditionalFormatType.IconSet
    cf.IconSet.IconSetType = IconSetType.FourArrows

def AddIconSet4(sheet):
    xcfs = sheet.ConditionalFormats.Add()
    xcfs.AddRange(sheet.Range["M5:O6"])
    cf = xcfs.AddCondition()
    cf.FormatType = ConditionalFormatType.IconSet
    cf.IconSet.IconSetType = IconSetType.FiveArrows

# ColorScale
def AddDefaultColorScale(sheet):
    xcfs = sheet.ConditionalFormats.Add()
    xcfs.AddRange(sheet.Range["A5:C6"])
    cf = xcfs.AddCondition()
    cf.FormatType = ConditionalFormatType.ColorScale

def Add3ColorScale(sheet):
    xcfs = sheet.ConditionalFormats.Add()
    xcfs.AddRange(sheet.Range["A7:C8"])
    cf = xcfs.AddCondition()
    cf.FormatType = ConditionalFormatType.ColorScale
    cf.ColorScale.MinValue.Type = ConditionValueType.Number
    cf.ColorScale.MinValue.Value = Int32(9)
    cf.ColorScale.MinColor = Color.get_Purple()

def Add2ColorScale(sheet):
    xcfs = sheet.ConditionalFormats.Add()
    xcfs.AddRange(sheet.Range["A9:C10"])
    cf = xcfs.AddCondition()
    cf.FormatType = ConditionalFormatType.ColorScale
    cf.ColorScale.MinColor = Color.get_Gold()
    cf.ColorScale.MaxColor = Color.get_SkyBlue()

# AboveAverage
def AddAboveAverage(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["A11:C12"])
    cf = conds.AddAverageCondition(AverageType.Above)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

def AddAboveAverage2(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["A13:C14"])
    cf = conds.AddAverageCondition(AverageType.BelowEqual)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_LightSkyBlue()

def AddAboveAverage3(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["A15:C16"])
    cf = conds.AddAverageCondition(AverageType.AboveStdDev3)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_LightSkyBlue()

# Top10
def AddTop10_1(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["A17:C20"])
    cf = conds.AddTopBottomCondition(TopBottomType.Top, 10)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Yellow()

def AddTop10_2(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["A21:C24"])
    cf = conds.AddTopBottomCondition(TopBottomType.Bottom, 10)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

def AddTop10_3(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["A25:C28"])
    cf = conds.AddTopBottomCondition(TopBottomType.TopPercent, 10)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Blue()

def AddTop10_4(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["A29:C32"])
    cf = conds.AddTopBottomCondition(TopBottomType.BottomPercent, 10)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Green()

# DataBar
def AddDataBar1(sheet):
    xcfs = sheet.ConditionalFormats.Add()
    xcfs.AddRange(sheet.Range["E1:G2"])
    cf = xcfs.AddCondition()
    cf.FormatType = ConditionalFormatType.DataBar
    cf.DataBar.BarColor = Color.get_Blue()
    cf.DataBar.MinPoint.Type = ConditionValueType.Percent
    cf.DataBar.ShowValue = True

def AddDataBar2(sheet):
    xcfs = sheet.ConditionalFormats.Add()
    xcfs.AddRange(sheet.Range["E3:G4"])
    cf = xcfs.AddCondition()
    cf.FormatType = ConditionalFormatType.DataBar
    cf.DataBar.BarColor = Color.get_Orange()
    cf.DataBar.MinPoint.Type = ConditionValueType.Percentile
    cf.DataBar.MinPoint.Value = Double(30.78)
    cf.DataBar.ShowValue = False

# ContainsText
def AddContainsText(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["E5:G6"])
    cf = conds.AddContainsTextCondition("abc")
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Yellow()

def AddNotContainsText(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["E7:G8"])
    cf = conds.AddNotContainsTextCondition("abc")
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

# ContainsBlank
def AddContainsBlank(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["E9:G10"])
    cf = conds.AddContainsBlanksCondition()
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Yellow()

def AddNotContainsBlank(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["E11:G12"])
    cf = conds.AddNotContainsBlanksCondition()
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

# BeginWith
def AddBeginWith(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["E15:G16"])
    cf = conds.AddBeginsWithCondition("ab")
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

# EndWith
def AddEndWith(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["E13:G14"])
    cf = conds.AddEndsWithCondition("ab")
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Yellow()

# ContainsError
def AddContainsError(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["E17:G18"])
    cf = conds.AddContainsErrorsCondition()
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Yellow()

def AddNotContainsError(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["E19:G20"])
    cf = conds.AddNotContainsErrorsCondition()
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Yellow()

# Unique
def AddUnique(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["E21:G22"])
    cf = conds.AddUniqueValuesCondition()
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Yellow()

# Duplicate
def AddDuplicate(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["E23:G24"])
    cf = conds.AddDuplicateValuesCondition()
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

# TimePeriod
def AddTimePeriod_1(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["I1:K2"])
    cf = conds.AddTimePeriodCondition(TimePeriodType.Today)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

def AddTimePeriod_2(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["I3:K4"])
    cf = conds.AddTimePeriodCondition(TimePeriodType.Last7Days)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

def AddTimePeriod_3(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["I5:K6"])
    cf = conds.AddTimePeriodCondition(TimePeriodType.LastMonth)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

def AddTimePeriod_4(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["I7:K8"])
    cf = conds.AddTimePeriodCondition(TimePeriodType.LastWeek)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

def AddTimePeriod_5(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["I9:K10"])
    cf = conds.AddTimePeriodCondition(TimePeriodType.NextMonth)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

def AddTimePeriod_6(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["I11:K12"])
    cf = conds.AddTimePeriodCondition(TimePeriodType.NextWeek)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

def AddTimePeriod_7(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["I13:K14"])
    cf = conds.AddTimePeriodCondition(TimePeriodType.ThisMonth)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

def AddTimePeriod_8(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["I15:K16"])
    cf = conds.AddTimePeriodCondition(TimePeriodType.ThisWeek)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

def AddTimePeriod_9(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["I17:K18"])
    cf = conds.AddTimePeriodCondition(TimePeriodType.Tomorrow)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()

def AddTimePeriod_10(sheet):
    conds = sheet.ConditionalFormats.Add()
    conds.AddRange(sheet.Range["I19:K20"])
    cf = conds.AddTimePeriodCondition(TimePeriodType.Yesterday)
    cf.FillPattern = ExcelPatternType.Solid
    cf.BackColor = Color.get_Pink()
```

---

# Spire.XLS Python Named Range Formula
## Insert formula with named range in Excel
```python
#Create a workbook
workbook = Workbook()
sheet = workbook.Worksheets[0]
#Set value
sheet.Range["A1"].Value = "1"
sheet.Range["A2"].Value = "1"
#Create a named range
NamedRange = workbook.NameRanges.Add("NewNamedRange")
NamedRange.NameLocal = "=SUM(A1+A2)"
#Set the formula
sheet.Range["C1"].Formula = "NewNamedRange"
```

---

# spire.xls python formula reading
## read Excel formula and its calculated value
```python
sheet = workbook.Worksheets[0]
formula = sheet.Range["C14"].Formula
value = str(sheet.Range["C14"].FormulaNumberValue)
```

---

# Register AddIn Function
## Register and use custom AddIn functions in Excel workbook
```python
inputFile = "./Demos/Data/Test.xlam"

# Create a workbook
workbook = Workbook()

# Register AddIn function
workbook.AddInFunctions.Add(inputFile, "TEST_UDF")
workbook.AddInFunctions.Add(inputFile, "TEST_UDF1")

# Get the first sheet
sheet = workbook.Worksheets[0]

# Call AddIn function
sheet.Range["A1"].Formula = "=TEST_UDF()"
sheet.Range["A2"].Formula = "=TEST_UDF1()"
```

---

# spire.xls python remove formulas
## Remove formulas from Excel cells while keeping their calculated values
```python
#Loop through worksheets.
for sheet in workbook.Worksheets:
    #Loop through cells.
    for cell in sheet.Range:
        #If the cell contain formula, get the formula value, clear cell content, and then fill the formula value into the cell.
        if cell.HasFormula:
            value = cell.FormulaValue
            cell.Clear(ExcelClearOptions.ClearContent)
            cell.Value2 = value
```

---

# Excel SubTotal Formula Implementation
## Demonstrates how to add SUBTOTAL formulas to Excel cells with different function numbers
```python
#Add SUBTOTAL formulas
sheet.Range["A5"].Formula = "=SUBTOTAL(1,A1:C3)"
sheet.Range["B5"].Formula = "=SUBTOTAL(2,A1:C3)"
sheet.Range["C5"].Formula = "=SUBTOTAL(5,A1:C3)"
#Calculate Formulas
workbook.CalculateAllValue()
```

---

# Spire.XLS Python Array Formulas
## Demonstrates how to use array formulas in Excel with Spire.XLS for Python
```python
#Write array formula
sheet.Range["A5:C6"].FormulaArray = "=LINEST(A1:A3,B1:C3,TRUE,TRUE)"
#Calculate Formulas
workbook.CalculateAllValue()
```

---

# spire.xls array R1C1 formula
## Using array R1C1 formula notation in Excel cells
```python
# Write array R1C1 formula
sheet.Range["C4"].FormulaArrayR1C1 = "=SUM(R[-3]C[-2]:R[-1]C)"
# Calculate Formulas
workbook.CalculateAllValue()
```

---

# spire.xls python formula
## using array formula in Excel
```python
#Create a workbook
workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
#Write array formula
sheet.Range["A5:C6"].FormulaArray = "=LINEST(A1:A3,B1:C3,TRUE,TRUE)"
#Calculate Formulas
workbook.CalculateAllValue()
```

---

# spire.xls python formula writing
## demonstrates how to write various types of formulas in Excel cells
```python
# String formula
sheet.Range[currentRow,1].NumberFormat = "@"
sheet.Range[currentRow,1].Text = "=\"hello\""
sheet.Range[currentRow,2].Formula = "=\"hello\""

# Integer formula
sheet.Range[currentRow,1].NumberFormat = "@"
sheet.Range[currentRow,1].Text = "=300"
sheet.Range[currentRow,2].Formula = "=300"

# Float formula
sheet.Range[currentRow,1].NumberFormat = "@"
sheet.Range[currentRow,1].Text = "=3389.639421"
sheet.Range[currentRow,2].Formula = "=3389.639421"

# Boolean formula
sheet.Range[currentRow,1].NumberFormat = "@"
sheet.Range[currentRow,1].Text = "=false"
sheet.Range[currentRow,2].Formula = "=false"

# Mathematical operations
sheet.Range[currentRow,1].NumberFormat = "@"
sheet.Range[currentRow,1].Text = "=1+2+3+4+5-6-7+8-9"
sheet.Range[currentRow,2].Formula = "=1+2+3+4+5-6-7+8-9"

# Cell reference formula
sheet.Range[currentRow,1].NumberFormat = "@"
sheet.Range[currentRow,1].Text = "=Sheet1!$B$3"
sheet.Range[currentRow,2].Formula = "=Sheet1!$B$3"

# Sheet area reference formula
sheet.Range[currentRow,1].NumberFormat = "@"
sheet.Range[currentRow,1].Text = "=AVERAGE(Sheet1!$D$3:G$3)"
sheet.Range[currentRow,2].Formula = "=AVERAGE(Sheet1!$D$3:G$3)"

# Excel functions
sheet.Range[currentRow,1].NumberFormat = "@"
sheet.Range[currentRow,1].Text = "=Count(3,5,8,10,2,34)"
sheet.Range[currentRow,2].Formula = "=Count(3,5,8,10,2,34)"

# Date and time functions
sheet.Range[currentRow,1].NumberFormat = "@"
sheet.Range[currentRow,1].Text = "=NOW()"
sheet.Range[currentRow,2].Formula = "=NOW()"
sheet.Range[currentRow,2].Style.NumberFormat = "yyyy-MM-DD"

# Logical functions
sheet.Range[currentRow,1].NumberFormat = "@"
sheet.Range[currentRow,1].Text = "=IF(4,2,2)"
sheet.Range[currentRow,2].Formula = "=IF(4,2,2)"
```

---

# spire.xls python header footer
## change font and size for header and footer
```python
#Set the new font and size for the header and footer
text = sheet.PageSetup.LeftHeader
#"Arial Unicode MS" is font name, "18" is font size
text = "&\"Arial Unicode MS\"&18 Header Footer Sample by Spire.XLS "
sheet.PageSetup.LeftHeader = text
sheet.PageSetup.RightFooter = text
```

---

# Spire.XLS Python Header Footer
## Set different header and footer for odd and even pages
```python
#Set the different header footer for Odd and Even pages
sheet.PageSetup.DifferentOddEven = 1
#Set the header with font, size, bold and color
sheet.PageSetup.OddHeaderString = "&\"Arial\"&12&B&KFFC000 Odd_Header"
sheet.PageSetup.OddFooterString = "&\"Arial\"&12&B&KFFC000 Odd_Footer"
sheet.PageSetup.EvenHeaderString = "&\"Arial\"&12&B&KFF0000 Even_Header"
sheet.PageSetup.EvenFooterString = "&\"Arial\"&12&B&KFF0000 Even_Footer"
sheet.ViewMode = ViewMode.Layout
```

---

# spire.xls python header footer
## Set different header and footer for the first page
```python
#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Set the value to show the headers/footers for first page are different from the other pages.
sheet.PageSetup.DifferentFirst = 1
#Set the header and footer for the first page.
sheet.PageSetup.FirstHeaderString = "Different First page"
sheet.PageSetup.FirstFooterString = "Different First footer"
#Set the other pages' header and footer. 
sheet.PageSetup.LeftHeader = "Demo of Spire.XLS"
sheet.PageSetup.CenterFooter = "Footer by Spire.XLS"
```

---

# spire.xls python image header footer
## add image to header and footer in excel
```python
#Set the image header
sheet.PageSetup.LeftHeaderImage = image
sheet.PageSetup.LeftHeader = "&G"
#Set the image footer
sheet.PageSetup.CenterFooterImage = image
sheet.PageSetup.CenterFooter = "&G"
#Set the view mode of the sheet
sheet.ViewMode = ViewMode.Layout
```

---

# Excel Header and Footer Setup
## Set header and footer for Excel worksheet
```python
# Get the first worksheet
Worksheet = workbook.Worksheets[0]
# Set left header with font name and size
Worksheet.PageSetup.LeftHeader = "&\"Arial Unicode MS\"&14 Spire.XLS for .Python "
# Set center footer 
Worksheet.PageSetup.CenterFooter = "Footer Text"
# Set view mode
Worksheet.ViewMode = ViewMode.Layout
```

---

# Spire.XLS for Python - Add Hyperlink to Text
## This example demonstrates how to add URL and email hyperlinks to text in Excel cells
```python
#Add url link
UrlLink = sheet.HyperLinks.Add(sheet.Range["D10"])
UrlLink.TextToDisplay = sheet.Range["D10"].Text
UrlLink.Type = HyperLinkType.Url
UrlLink.Address = "http://en.wikipedia.org/wiki/Chicago"
#Add email link
MailLink = sheet.HyperLinks.Add(sheet.Range["E10"])
MailLink.TextToDisplay = sheet.Range["E10"].Text
MailLink.Type = HyperLinkType.Url
MailLink.Address = "mailto:Amor.Aqua@gmail.com"
```

---

# Spire.XLS Python Image Hyperlink
## Add hyperlink to an image in Excel worksheet
```python
#Insert an image to a specific cell
picture = sheet.Pictures.Add(2, 1, inputFile)
#Add a hyperlink to the image
picture.SetHyperLink("https://www.e-iceblue.com/Introduce/excel-for-net-introduce.html", True)
```

---

# spire.xls python hyperlink
## get hyperlink type from excel worksheet
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Iterate all hyperlinks
for item in sheet.HyperLinks:
    #Get hyperlink address
    address = item.Address
    #Get hyperlink type
    type = item.Type
```

---

# spire.xls python hyperlink
## create hyperlink to external file in Excel
```python
#Create a workbook
workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
range = sheet.Range[1,1]
#Add hyperlink in the range
hyperlink = sheet.HyperLinks.Add(range)
#Set the link type
hyperlink.Type = HyperLinkType.File
#Set the display text
hyperlink.TextToDisplay = "Link To External File"
#Set file address
hyperlink.Address = inputFile
```

---

# spire.xls python hyperlink
## create hyperlink to another sheet cell
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
range = sheet.Range["A1"]
#Add hyperlink in the range
hyperlink = sheet.HyperLinks.Add(range)
#Set the link type
hyperlink.Type = HyperLinkType.Workbook
#Set the display text
hyperlink.TextToDisplay = "Link to Sheet2 cell C5"
#Set the address
hyperlink.Address = "Sheet2!C5"
```

---

# spire.xls python hyperlinks
## modify hyperlink properties
```python
#Get the collection of all hyperlinks in the worksheet
sheet = workbook.Worksheets[0]
#Change the values of TextToDisplay and Address property 
links = sheet.HyperLinks
links[0].TextToDisplay = "Spire.XLS for .NET"
links[0].Address = "http://www.e-iceblue.com/Introduce/excel-for-net-introduce.html"
```

---

# Spire.XLS Python Hyperlinks
## Read hyperlinks from Excel worksheet
```python
# Create a workbook and load Excel file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
# Get hyperlinks from worksheet
sheet = workbook.Worksheets[0]
address1 = sheet.HyperLinks[0].Address
address2 = sheet.HyperLinks[1].Address
workbook.Dispose()
```

---

# Remove hyperlinks from Excel worksheet
## This code demonstrates how to remove hyperlinks from an Excel worksheet while preserving the link text
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the collection of all hyperlinks in the worksheet
links = sheet.HyperLinks
#Remove all link content
sheet.Range["B1"].ClearAll()
sheet.Range["B2"].ClearAll()
sheet.Range["B3"].ClearAll()
#Remove hyperlink and keep link text
sheet.HyperLinks.RemoveAt(0)
sheet.HyperLinks.RemoveAt(0)
sheet.HyperLinks.RemoveAt(0)
```

---

# spire.xls python hyperlinks
## retrieve external file hyperlinks from Excel
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
content = []
#Retrieve external file hyperlinks
for item in sheet.HyperLinks:
    address = item.Address
    sheetName = item.Range.WorksheetName
    range = item.Range
    content.append("Cell[{0},{1}] in sheet \"" + sheetName + "\" contains File URL: {2}".format(range.Row, range.Column, address))
```

---

# Excel Hyperlinks Creation
## Create and configure hyperlinks in Excel worksheet
```python
# Set text for hyperlink labels
sheet.Range["B9"].Text = "Home page"
# Create hyperlink to website
hylink1 = sheet.HyperLinks.Add(sheet.Range["B10"])
hylink1.Type = HyperLinkType.Url
hylink1.Address = """http://www.e-iceblue.com"""
# Set text for email hyperlink
sheet.Range["B11"].Text = "Support"
# Create email hyperlink
hylink2 = sheet.HyperLinks.Add(sheet.Range["B12"])
hylink2.Type = HyperLinkType.Url
hylink2.Address = "mailto:support@e-iceblue.com"
# Set text for forum hyperlink
sheet.Range["B13"].Text = "Forum"
# Create forum hyperlink
hylink3 = sheet.HyperLinks.Add(sheet.Range["B14"])
hylink3.Type = HyperLinkType.Url
hylink3.Address = "https://www.e-iceblue.com/forum/"
```

---

# Spire.XLS Python MarkerDesigner
## Add variable array to Excel using MarkerDesigner
```python
#Create a workbook
workbook = Workbook()
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set marker designer field in cell A1
sheet.Range["A1"].Value = "&=Array"
#Fill Array
workbook.MarkerDesigner.AddArray("Array", [String("Spire.Xls"), String("Spire.Doc"), String("Spire.PDF"), String("Spire.Presentation"), String("Spire.Email")])
workbook.MarkerDesigner.Apply()
workbook.CalculateAllValue()
#AutoFit
sheet.AllocatedRange.AutoFitRows()
sheet.AllocatedRange.AutoFitColumns()
```

---

# Format Named Range Cells
## Format cells in a named range with color and bold font
```python
#Get specific named range by index
NamedRange = workbook.NameRanges[0]
#Get the cell range of the named range
range = NamedRange.RefersToRange
#Set color for the range
range.Style.Color = Color.get_Yellow()
#Set the font as bold
range.Style.Font.IsBold = True
```

---

# Spire.XLS Python Named Ranges
## Get all named ranges from an Excel workbook and access their names
```python
# Get all named ranges
ranges = workbook.NameRanges
for nameRange in ranges:
    name = nameRange.Name
```

---

# spire.xls python named range
## get address of named range
```python
# Get specific named range by index
NamedRange = workbook.NameRanges[0]
# Get the address of the named range
address = NamedRange.RefersToRange.RangeAddress
```

---

# spire.xls python named ranges
## get specific named range by index or name
```python
#Get specific named range by index
name1 = workbook.NameRanges[1].Name
#Get specific named range by name
name2 = workbook.NameRanges["NameRange3"].Name
```

---

# spire.xls python named range
## merge named range cells
```python
workbook = Workbook()
#Get specific named range by index
NamedRange = workbook.NameRanges[0]
#Get the range of the named range
range = NamedRange.RefersToRange
#Merge cells
range.Merge()
```

---

# Spire.XLS Python Named Ranges
## Creating and setting named ranges in Excel
```python
workbook = Workbook()
sheet = workbook.Worksheets[0]
# Creating a named range
NamedRange = workbook.NameRanges.Add("NewNamedRange")
# Setting the range of the named range
NamedRange.RefersToRange = sheet.Range["A8:E12"]
```

---

# Remove Named Ranges in Excel
## Demonstrates how to remove named ranges from an Excel workbook by index or by name
```python
#Remove the named range by index
workbook.NameRanges.RemoveAt(0)
#Remove the named range by name
workbook.NameRanges.Remove("NameRange2")
```

---

# spire.xls python named range
## rename named range in excel
```python
workbook = Workbook()
#Rename the named range
workbook.NameRanges[0].Name = "RenameRange"
```

---

# Excel Named Range Creation
## Create a scoped named range in Excel worksheet
```python
#Add range name
namedRange = sheet.Names.Add("Range1")
#Define the range
namedRange.RefersToRange = sheet.Range["A1:D19"]
```

---

# Excel Named Range Formula
## Create a named range and use it in a formula
```python
#Get the sheet
sheet = workbook.Worksheets[0]
#Create a named range
NamedRange = workbook.NameRanges.Add("MyNamedRange")
#Refers to range
NamedRange.RefersToRange = sheet.Range["B10:B12"]
#Set the formula of range to named range
sheet.Range["B13"].Formula = "=SUM(MyNamedRange)"
```

---

# Extract OLE Objects from Excel
## Extract OLE objects from Excel worksheets and save them as files
```python
def WriteAllBytes(fname:str, data):
    fp = open(fname, "wb")
    for d in data:
        fp.write(d)
    fp.close()

# Get the first worksheet
sheet = workbook.Worksheets[0]
# Extract ole objects
if sheet.HasOleObjects:
    for obj in sheet.OleObjects:
        type = obj.ObjectType
        # Word document
        if type is OleObjectType.WordDocument:
            WriteAllBytes(outputFile1, obj.OleData)

# Get the first worksheet
sheet = workbook.Worksheets[0]
# Extract ole objects
if sheet.HasOleObjects:
    for obj in sheet.OleObjects:
        type = obj.ObjectType
        # PDF document
        if type is OleObjectType.AdobeAcrobatDocument:
            WriteAllBytes(outputFile2, obj.OleData)

# Get the first worksheet
sheet = workbook.Worksheets[0]
# Extract ole objects
if sheet.HasOleObjects:
    for obj in sheet.OleObjects:
        type = obj.ObjectType
        # PowerPoint document
        if type is OleObjectType.PowerPointSlide:
            WriteAllBytes(outputFile3, obj.OleData)
```

---

# spire.xls python OLE Objects
## Insert OLE Objects into Excel Worksheet
```python
def GenerateImage(fileName):
    book = Workbook()
    book.LoadFromFile(fileName)
    book.Worksheets[0].PageSetup.LeftMargin = 0
    book.Worksheets[0].PageSetup.RightMargin = 0
    book.Worksheets[0].PageSetup.TopMargin = 0
    book.Worksheets[0].PageSetup.BottomMargin = 0
    return book.Worksheets[0].ToImage(1, 1, 19, 5)

workbook = Workbook()
ws = workbook.Worksheets[0]
ws.Range["A1"].Text = "Here is an OLE Object."
#insert OLE object
image = GenerateImage(inputFile)
oleObject = ws.OleObjects.Add(inputFile, image, OleLinkType.Embed)
oleObject.Location = ws.Range["B4"]
oleObject.ObjectType = OleObjectType.ExcelWorksheet
```

---

# Insert WAV file as OLE object in Excel
## This code demonstrates how to insert a WAV audio file as an OLE object in an Excel worksheet
```python
#Create a workbook
workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
#Add WAV file as OLE object with an image
with Stream(image_file_path) as fs:
    oleObject = sheet.OleObjects.Add(wav_file_path, fs, OleLinkType.Embed)
#Set the object location
oleObject.Location = sheet.Range["B4"]
#Set the object type as package
oleObject.ObjectType = OleObjectType.Package
```

---

# spire.xls python get paper dimensions
## This example demonstrates how to get the dimensions of different paper sizes in Excel using PageSetup

```python
# Create a workbook and get the first worksheet
workbook = Workbook()
sheet = workbook.Worksheets[0]

# Get dimensions for different paper sizes
sheet.PageSetup.PaperSize = PaperSizeType.A2Paper
a2_dimensions = (sheet.PageSetup.PageWidth, sheet.PageSetup.PageHeight)

sheet.PageSetup.PaperSize = PaperSizeType.PaperA3
a3_dimensions = (sheet.PageSetup.PageWidth, sheet.PageSetup.PageHeight)

sheet.PageSetup.PaperSize = PaperSizeType.PaperA4
a4_dimensions = (sheet.PageSetup.PageWidth, sheet.PageSetup.PageHeight)

sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter
letter_dimensions = (sheet.PageSetup.PageWidth, sheet.PageSetup.PageHeight)
```

---

# Spire.XLS Python Get Excel Version
## Get the version information of an Excel file
```python
# Create a workbook
workbook = Workbook()
# Get the version
version = workbook.Version
```

---

# Excel Page Setup
## Set page order type in Excel worksheets
```python
#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the reference of the PageSetup of the worksheet.
pageSetup = sheet.PageSetup
#Set the order type of the pages to over then down.
pageSetup.Order = OrderType.OverThenDown
```

---

# spire.xls page setup
## set Excel paper size to A4
```python
# Create a workbook
workbook = Workbook()
# Get the first worksheet
sheet = workbook.Worksheets[0]
# Set the paper size of the worksheet as A4 paper
sheet.PageSetup.PaperSize = PaperSizeType.PaperA4
```

---

# Excel Page Setup First Page Number
## Set the first page number for worksheet pages
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set the first page number of the worksheet pages
sheet.PageSetup.FirstPageNumber = 2
```

---

# Excel Page Setup
## Set header and footer margins in Excel worksheets
```python
#Get the PageSetup object of the first worksheet.
pageSetup = sheet.PageSetup
#Set the margins of header and footer.
pageSetup.HeaderMarginInch = 2
pageSetup.FooterMarginInch = 2
```

---

# Excel Page Setup Margins
## Set page margins for an Excel worksheet
```python
#Get the PageSetup object of the worksheet.
pageSetup = sheet.PageSetup
#Set bottom,left,right and top page margins.
pageSetup.BottomMargin = 2
pageSetup.LeftMargin = 1
pageSetup.RightMargin = 1
pageSetup.TopMargin = 3
```

---

# Spire.XLS for Python - Page Setup
## Configure various printing options for Excel worksheets
```python
#Get the reference of the PageSetup of the worksheet.
pageSetup = sheet.PageSetup
#Allow to print gridlines.
pageSetup.IsPrintGridlines = True
#Allow to print row/column headings.
pageSetup.IsPrintHeadings = True
#Allow to print worksheet in black & white mode.
pageSetup.BlackAndWhite = True
#Allow to print comments as displayed on worksheet.
pageSetup.PrintComments = PrintCommentType.InPlace
#Allow to print worksheet with draft quality.
pageSetup.Draft = True
#Allow to print cell errors as N/A.
pageSetup.PrintErrors = PrintErrorsType.NA
```

---

# Spire.XLS Page Setup
## Set page orientation to landscape
```python
# Get the first worksheet
sheet = workbook.Worksheets[0]
# Set the page orientation to Landscape
sheet.PageSetup.Orientation = PageOrientationType.Landscape
```

---

# Spire.XLS Page Setup
## Set print area for Excel worksheet
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the reference of the PageSetup of the worksheet.
pageSetup = sheet.PageSetup
#Specify the cells range of the print area.
pageSetup.PrintArea = "A1:E5"
```

---

# Excel Print Quality Setup
## Set print quality for Excel worksheet
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Set the print quality of the worksheet to 180 dpi.
sheet.PageSetup.PrintQuality = 180
```

---

# spire.xls python page setup
## set print title for Excel file
```python
#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
pageSetup = sheet.PageSetup
#Define column numbers A & B as title columns.
pageSetup.PrintTitleColumns = "$A:$B"
#Defining row numbers 1 & 2 as title rows.
pageSetup.PrintTitleRows = "$1:$2"
```

---

# spire.xls python page setup
## set sheet fit to page property
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Set the FitToPagesTall property.
sheet.PageSetup.FitToPagesTall = 1
#Set the FitToPagesWide property.
sheet.PageSetup.FitToPagesWide = 1
```

---

# spire.xls python page setup
## center worksheet on page
```python
#Get the PageSetup object of the first page.
pageSetup = sheet.PageSetup
#Set the worksheet center on page.
pageSetup.CenterHorizontally = True
pageSetup.CenterVertically = True
```

---

# spire.xls python pivot table
## Change data source of pivot table
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
Range = sheet.Range["A1:C15"]
table = workbook.Worksheets[1].PivotTables[0]
#Change data source
table.ChangeDataSource(Range)
table.Cache.IsRefreshOnLoad = False
```

---

# Spire.XLS Python Pivot Table Operations
## Clear all data fields from a pivot table
```python
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0] if isinstance(sheet.PivotTables[0], XlsPivotTable) else None
#Clear all the data fields
pt.DataFields.Clear()
pt.CalculateData()
```

---

# Pivot Table Consolidation Functions
## Apply different consolidation functions to pivot table data fields
```python
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0] if isinstance(sheet.PivotTables[0], XlsPivotTable) else None
#Apply Average consolidation function to first data field
pt.DataFields[0].Subtotal = SubtotalTypes.Average
#Apply Max consolidation function to second data field
pt.DataFields[1].Subtotal = SubtotalTypes.Max
pt.CalculateData()
```

---

# spire.xls python pivot table
## create and configure a pivot table in excel
```python
#Add a PivotTable to the worksheet
dataRange = sheet.Range["A1:C7"]
cache = workbook.PivotCaches.Add(dataRange)
pt = sheet.PivotTables.Add("Pivot Table", sheet.Range["E10"], cache)
#Drag the fields to the row area.
pf = pt.PivotFields["Product"]
pf.Axis = AxisTypes.Row
pf2 = pt.PivotFields["Month"]
pf2.Axis = AxisTypes.Row
#Drag the field to the data area.
pt.DataFields.Add(pt.PivotFields["Count"], "SUM of Count", SubtotalTypes.Sum)
#Set PivotTable style
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12
#Autofit columns generated by the pivotTable
pt.CalculateData()
sheet.AutoFitColumn(5)
sheet.AutoFitColumn(6)
```

---

# Disable Pivot Table Ribbon in Excel
## This code demonstrates how to disable the ribbon for a pivot table in Excel
```python
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0] if isinstance(sheet.PivotTables[0], XlsPivotTable) else None
#Disable ribbon for this pivot table
pt.EnableWizard = False
```

---

# spire.xls python pivot table
## expand or collapse rows in pivot table
```python
# Get pivot table
pivotTable = sheet.PivotTables[0]
# Calculate data
pivotTable.CalculateData()
# Collapse rows
(pivotTable.PivotFields["Vendor No"]).HideItemDetail("3501", True)
# Expand rows
( pivotTable.PivotFields["Vendor No"]).HideItemDetail("3502", False)
```

---

# spire.xls python pivot table
## format pivot table data field
```python
# Access the PivotTable
pt = sheet.PivotTables[0] if isinstance(sheet.PivotTables[0], XlsPivotTable) else None
# Access the data field
pivotDataField = pt.DataFields[0]
# Set data display format
pivotDataField.ShowDataAs = PivotFieldFormatType.PercentageOfColumn
```

---

# Excel Pivot Table Formatting
## Format pivot table appearance in Excel
```python
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0]
#Format appearance
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleLight10
pt.Options.ShowGridDropZone = True
pt.Options.RowLayout = PivotTableLayoutType.Tabular
```

---

# spire.xls python pivot table
## get pivot table refresh information
```python
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Get the first pivot table
pivotTable = worksheet.PivotTables[0]
#Get the refreshed information
dateTime = pivotTable.Cache.RefreshDate
refreshedBy = pivotTable.Cache.RefreshedBy
```

---

# Excel Pivot Table Refresh
## Refresh pivot table data in Excel
```python
#Get the PivotTable that was built on the data source.
pt = workbook.Worksheets[0].PivotTables[0]
#Refresh the data of PivotTable.
pt.Cache.IsRefreshOnLoad = True
```

---

# Spire.XLS Python Pivot Table
## Repeat item labels in pivot table
```python
#Create PivotTable cache and table
dataRange = sheet.Range["A1:D9"]
cache = workbook.PivotCaches.Add(dataRange)
pt = sheet2.PivotTables.Add("Pivot Table", sheet.Range["A1"], cache)
r1 = pt.PivotFields["VendorNo"]
r1.Axis = AxisTypes.Row
pt.Options.RowHeaderCaption = "VendorNo"
r1.Subtotals = SubtotalTypes.none
r1.RepeatItemLabels = True
#Repeat item labels
pt.PivotFields["OnHand"].RepeatItemLabels = True
pt.Options.RowLayout = PivotTableLayoutType.Tabular
r2 = pt.PivotFields["Desc"]
r2.Axis = AxisTypes.Row
pt.DataFields.Add(pt.PivotFields["OnHand"], "Sum of onHand", SubtotalTypes.none)
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12
```

---

# Spire.XLS Python PivotTable Format Options
## Set format options for an Excel pivot table
```python
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0]
#Set the PivotTable report is automatically formatted
pt.Options.IsAutoFormat = True
#Setting the PivotTable report shows grand totals for rows.
pt.ShowRowGrand = True
#Setting the PivotTable report shows grand totals for columns.
pt.ShowColumnGrand = True
#Setting the PivotTable report displays a custom string in cells that contain null values.
pt.DisplayNullString = True
pt.NullString = "null"
#Setting the PivotTable report's layout
pt.PageFieldOrder = PagesOrderType.DownThenOver
```

---

# spire.xls python pivot table
## set pivot field format
```python
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0]
pf = pt.PivotFields[0]
#Setting the field auto sort ascend.
pf.SortType = PivotFieldSortType.Ascending
#Setting Subtotal auto show.
pf.SubtotalTop = True
#Setting Subtotal as Count type
pf.Subtotals = SubtotalTypes.Count
#Setting the field auto show.
pf.IsAutoShow = True
```

---

# Spire.XLS for Python Pivot Table
## Show data fields in row area of pivot table
```python
#get the data in Pivot Table
pivotTable = sheet.PivotTables[0] if isinstance(sheet.PivotTables[0], XlsPivotTable) else None
#show the datafield in row
pivotTable.ShowDataFieldInRow = True
#calculate data
pivotTable.CalculateData()
```

---

# spire.xls python pivot table
## show subtotals in pivot table
```python
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["Pivot Table"]
pt = sheet.PivotTables[0]
#Show Subtotals
pt.ShowSubtotals = True
```

---

# Sort Pivot Table in Excel
## Create and sort a pivot table with specific fields
```python
#Specify the data source
dataRange = sheet.Range["A1:C9"]
cache = workbook.PivotCaches.Add(dataRange)
#Add PivotTable
pt = sheet2.PivotTables.Add("Pivot Table", sheet.Range["A1"], cache)
r1 = pt.PivotFields["No"]
r1.Axis = AxisTypes.Row
pt.Options.RowLayout = PivotTableLayoutType.Tabular
#Sort PivotField
r1.SortType = PivotFieldSortType.Descending
r2 = pt.PivotFields["Name"]
r2.Axis = AxisTypes.Row
pt.DataFields.Add(pt.PivotFields["OnHand"], "Sum of onHand", SubtotalTypes.none)
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12
```

---

# Update Pivot Table Data Source
## Update the data source of a pivot table and refresh it
```python
#Modify data of data source
data = workbook.Worksheets["Data"]
data.Range["A2"].Text = "NewValue"
data.Range["D2"].NumberValue = 28000
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0]
#Refresh and calculate
pt.Cache.IsRefreshOnLoad = True
pt.CalculateData()
```

---

# spire.xls python track changes
## accept or reject tracked changes in excel
```python
# Create workbook
workbook = Workbook()
# Accept the changes or reject the changes
# workbook.AcceptAllTrackedChanges()
workbook.RejectAllTrackedChanges()
```

---

# Detect Excel Workbook Protection
## Check if an Excel workbook is password protected

```python
inputFile = "./Demos/Data/ProtectedWorkbook.xlsx"
value = Workbook.IsPasswordProtected(inputFile)
boolvalue = "Yes" if value else "No"
```

---

# spire.xls python security
## hide formulas and protect worksheet
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Hide the formulas in the used range
sheet.AllocatedRange.IsFormulaHidden = True
#Protect the worksheet with password
sheet.Protect("e-iceblue")
```

---

# Lock specific cells in Excel worksheet
## This code demonstrates how to lock specific cells in a new Excel worksheet while leaving other cells unlocked

```python
#Create a workbook.
workbook = Workbook()
#Create an empty worksheet.
workbook.CreateEmptySheet()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Loop through all the rows in the worksheet and unlock them.
for i in range(0,255):
    sheet.Rows[i].Style.Locked = False
#Lock specific cell in the worksheet.
sheet.Range["A1"].Text = "Locked"
sheet.Range["A1"].Style.Locked = True
#Lock specific cell range in the worksheet.
sheet.Range["C1:E3"].Text = "Locked"
sheet.Range["C1:E3"].Style.Locked = True
#Set the password.
sheet.Protect("123", SheetProtectionType.All)
```

---

# Lock Specific Column in Excel
## This code demonstrates how to lock a specific column in an Excel worksheet while keeping other columns unlocked

```python
#Create a workbook.
workbook = Workbook()
#Create an empty worksheet.
workbook.CreateEmptySheet()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Loop through all the columns in the worksheet and unlock them.
for i in range(0,255):
    sheet.Rows[i].Style.Locked = False
#Lock the fourth column in the worksheet.
sheet.Columns[3].Text = "Locked"
sheet.Columns[3].Style.Locked = True
#Set the password.
sheet.Protect("123", SheetProtectionType.All)
```

---

# Excel Row Locking
## Lock specific row in Excel while keeping others unlocked
```python
#Create a workbook.
workbook = Workbook()
#Create an empty worksheet.
workbook.CreateEmptySheet()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Loop through all the rows in the worksheet and unlock them.
for i in range(0,255):
    sheet.Rows[i].Style.Locked = False
#Lock the third row in the worksheet.
sheet.Rows[2].Text = "Locked"
sheet.Rows[2].Style.Locked = True
#Set the password.
sheet.Protect("123", SheetProtectionType.All)
```

---

# Spire.XLS Python Security
## Protect specific cells in an Excel worksheet
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Protect cell
sheet.Range["B3"].Style.Locked = True
sheet.Range["C3"].Style.Locked = False
sheet.Protect("TestPassword", SheetProtectionType.All)
```

---

# spire.xls python worksheet protection
## protect worksheet with editable ranges
```python
#Define the specified ranges to allow users to edit while sheet is protected
sheet.AddAllowEditRange("EditableRanges", sheet.Range["B4:E12"])
#Protect worksheet with a password.
sheet.Protect("TestPassword", SheetProtectionType.All)
```

---

# Protect Workbook in Excel
## Protect workbook with password
```python
#Protect Workbook
workbook.Protect("e-iceblue")
```

---

# Remove Digital Signatures in Excel
## Remove all digital signatures from an Excel workbook
```python
#Remove all digital signatures.
workbook.RemoveAllDigitalSignatures()
```

---

# spire.xls python unprotect sheet
## Unprotect a password-protected Excel worksheet
```python
#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Unlock the worksheet with password
sheet.Unprotect("e-iceblue")
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
```

---

# Unlock Excel Worksheet
## Unprotect a worksheet in an Excel file
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Unlock the worksheet in a unlocked Excel file with null string.
sheet.Unprotect()
```

---

# Extract text from Excel textbox
## This code demonstrates how to extract text from a textbox in an Excel worksheet

```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the first textbox
shape = sheet.TextBoxes[0]
#Extract text from the textbox
content = []
content.append("The text extracted from the TextBox is: ")
content.append(shape.Text)
```

---

# spire.xls python get textbox by name
## demonstrates how to get a textbox by its name in an Excel worksheet
```python
#Create a workbook
workbook = Workbook()
#Get the default first worksheet
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
```

---

# spire.xls python textbox manipulation
## manipulate textbox content and alignment in excel
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the first textbox
tb = sheet.TextBoxes[0]
#Change the text of textbox
tb.Text = "Spire.XLS for Python"
#Set the alignment of textbox as center
tb.HAlignment = CommentHAlignType.Center
tb.VAlignment = CommentVAlignType.Center
```

---

# Remove Borderline of Textbox in Excel
## This code demonstrates how to create textboxes in an Excel chart and remove the borderline from a textbox
```python
chart = sheet.Charts.Add()
#Create textbox1 in the chart and input text information.
textbox1 = chart.TextBoxes.AddTextBox(50, 50, 100, 600)
textbox1.Text = "The solution with borderline"
#Create textbox2 in the chart, input text information and remove borderline.
textbox2 = chart.TextBoxes.AddTextBox(1000, 50, 100, 600)
textbox2.Text = "The solution without borderline"
textbox2.Line.Weight = 0
```

---

# Replace text in Excel TextBoxes
## This function replaces specific text in all TextBoxes within an Excel sheet
```python
def ReplaceTextInTextBox(sheet, sFind, sReplace):
    for tb in sheet.TextBoxes:
        if tb.Text != "":
            if tb.Text.__contains__(sFind):
                tb.Text = tb.Text.replace(sFind, sReplace)
```

---

# Spire.XLS Python Textbox Formatting
## Set font and background color for Excel textbox
```python
#Get the textbox which will be edited.
shape = sheet.TextBoxes[0]
#Set the font and background color for the textbox.
#Set font.
font = workbook.CreateFont()
#font.IsStrikethrough = true
font.FontName = "Century Gothic"
font.Size = 10
font.IsBold = True
font.Color = Color.get_Blue()
rto = shape.RichText
rt = RichText(rto)
rt.SetFont(0, len(shape.Text) - 1, font)
#Set background color
shape.Fill.FillType = ShapeFillType.SolidColor
shape.Fill.ForeKnownColor = ExcelColors.BlueGray
```

---

# spire.xls python textbox
## Set internal margins of textbox in Excel
```python
#Add a textbox to the sheet and set its position and size.
textbox = sheet.TextBoxes.AddTextBox(4, 2, 100, 300)
#Set the text on the textbox.
textbox.Text = "Insert TextBox in Excel and set the margin for the text"
textbox.HAlignment = CommentHAlignType.Center
textbox.VAlignment = CommentVAlignType.Center
#Set the inner margins of the contents.
textbox.InnerLeftMargin = 1
textbox.InnerRightMargin = 3
textbox.InnerTopMargin = 1
textbox.InnerBottomMargin = 1
```

---

# spire.xls python textbox
## set wrap text property for textbox
```python
#Get the text box
shape = sheet.TextBoxes[0] if isinstance(sheet.TextBoxes[0], XlsTextBoxShape) else None
#Set wrap text
shape.IsWrapText = True
```

---

# spire.xls python worksheet
## activate a specific worksheet
```python
#Create a workbook
workbook = Workbook()
#Get the second worksheet from the workbook
sheet = workbook.Worksheets[1]
#Activate the sheet
sheet.Activate()
```

---

# Spire.XLS Python Page Breaks
## Add horizontal and vertical page breaks to Excel worksheet
```python
#Add page break in Excel file.
sheet.HPageBreaks.Add(sheet.Range["E4"])
sheet.VPageBreaks.Add(sheet.Range["C4"])
```

---

# Spire.XLS Python Worksheet Management
## Add a new worksheet to an Excel workbook
```python
#Create a workbook
workbook = Workbook()
#Add a new worksheet named AddedSheet
sheet = workbook.Worksheets.Add("AddedSheet")
sheet.Range["C5"].Text = "This is a new sheet."
```

---

# spire.xls python style
## Apply style to worksheet
```python
#Create a cell style
style = workbook.Styles.Add("newStyle")
style.Color = Color.get_LightBlue()
style.Font.Color = Color.get_White()
style.Font.Size = 15
style.Font.IsBold = True
#Apply the style to the first worksheet
sheet.ApplyStyle(style)
```

---

# spire.xls python check worksheet type
## Check if a worksheet is a Dialog Sheet
```python
# Get the first worksheet
sheet = workbook.Worksheets[0]
# Check if the worksheet is a Dialog Sheet
if sheet.Type == ExcelSheetType.DialogSheet:
    # Worksheet is a Dialog Sheet
    pass
else:
    # Worksheet is not a Dialog Sheet
    pass
```

---

# Copy worksheet to another Excel file
## This code demonstrates how to copy a worksheet from one workbook to another workbook in Python using Spire.XLS
```python
#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Define a pagesetup object based on the first worksheet.
pageSetup = sheet.PageSetup
#The first five rows are repeated in each page. It can be seen in print preview.
pageSetup.PrintTitleRows = "$1:$5"
#Create another Workbook.
workbook1 = Workbook()
#Get the first worksheet in the book.
sheet1 = workbook1.Worksheets[0]
#Copy worksheet to destination worsheet in another Excel file.
sheet1.CopyFrom(sheet)
```

---

# Copy Worksheet Within Workbook
## This example demonstrates how to copy a worksheet within the same workbook.
```python
#Get the first and the second worksheets.
sheet = workbook.Worksheets[0]
sheet1 = workbook.Worksheets.Add("MySheet")
sourceRange = sheet.AllocatedRange
#Copy the first worksheet to the second one.
sheet.Copy(sourceRange, sheet1, sheet.FirstRow, sheet.FirstColumn, True)
```

---

# Copy Visible Worksheets
## Copy only visible worksheets from one workbook to another
```python
#Create workbooks
workbook = Workbook()
workbookNew = Workbook()
workbookNew.Version = ExcelVersion.Version2013
workbookNew.Worksheets.Clear()
#Loop through the worksheets
for sheet in workbook.Worksheets:
    #Judge if the worksheet is visible
    if sheet.Visibility == WorksheetVisibility.Visible:
        #Copy the sheet to new workbook
        workbookNew.Worksheets.AddCopy(sheet)
```

---

# Copy Worksheet in Excel
## Demonstrates how to copy a worksheet from one workbook to another
```python
#Get the first worksheet from source workbook
srcWorksheet = sourceWorkbook.Worksheets[0]
#Add a new worksheet to target workbook
targetWorksheet = targetWorkbook.Worksheets.Add("added")
#Copy the source worksheet to the target worksheet
targetWorksheet.CopyFrom(srcWorksheet)
```

---

# spire.xls python worksheet detection
## detect if worksheets are empty
```python
#Create a workbook
workbook = Workbook()
#Get the first worksheet
worksheet1 = workbook.Worksheets[0]
#Detect the first worksheet is empty or not
detect1 = worksheet1.IsEmpty
#Get the second worksheet
worksheet2 = workbook.Worksheets[1]
#Detect the second worksheet is empty or not
detect2 = worksheet2.IsEmpty
```

---

# spire.xls python worksheet
## fill data in worksheet
```python
#Create a workbook
workbook = Workbook()
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Fill data
worksheet.Range["A1"].Style.Font.IsBold = True
worksheet.Range["B1"].Style.Font.IsBold = True
worksheet.Range["C1"].Style.Font.IsBold = True
worksheet.Range["A1"].Text = "Month"
worksheet.Range["A2"].Text = "January"
worksheet.Range["A3"].Text = "February"
worksheet.Range["A4"].Text = "March"
worksheet.Range["A5"].Text = "April"
worksheet.Range["B1"].Text = "Payments"
worksheet.Range["B2"].NumberValue = 251
worksheet.Range["B3"].NumberValue = 515
worksheet.Range["B4"].NumberValue = 454
worksheet.Range["B5"].NumberValue = 874
worksheet.Range["C1"].Text = "Sample"
worksheet.Range["C2"].Text = "Sample1"
worksheet.Range["C3"].Text = "Sample2"
worksheet.Range["C4"].Text = "Sample3"
worksheet.Range["C5"].Text = "Sample4"
#Set width for the second column
worksheet.SetColumnWidth(2, 10)
```

---

# spire.xls python freeze panes
## freeze panes in excel worksheet
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
#Freeze Top Row
sheet.FreezePanes(2, 1)
```

---

# Get freeze pane range in Excel worksheet
## This code demonstrates how to get the freeze pane range in an Excel worksheet using Spire.XLS for Python
```python
sheet = wb.Worksheets[0]
# The row and column index of the frozen pane is passed through the out parameter.
# If it returns to 0, it means that it is not frozen
indexs = sheet.GetFreezePanes()
colIndex = indexs[1]
rowIndex = indexs[0]
```

---

# spire.xls python font list
## get list of fonts used in Excel workbook
```python
# Create a workbook
workbook = Workbook()
fonts = []
# Loop all sheets of workbook
for sheet in workbook.Worksheets:
    r = 0
    while r < sheet.Rows.Length:
        for c in sheet.Rows[r].Cells:
            # Get the font of cell and add it to list
            fonts.append(c.Style.Font)
        r += 1
strB = []
for font in fonts:
    strB.append("FontName:{0}; FontSize:{1}".format(font.FontName, font.Size))
```

---

# Spire.XLS Python Get Paper Size
## Retrieve page dimensions from Excel worksheets
```python
workbook = Workbook()
for sheet in workbook.Worksheets:
    width = sheet.PageSetup.PageWidth
    height = sheet.PageSetup.PageHeight
```

---

# Spire.XLS for Python - Get Worksheet Names
## Extracts all worksheet names from an Excel workbook
```python
# Get the names of all worksheets
worksheet_names = []
for sheet in workbook.Worksheets:
    worksheet_names.append(sheet.Name)
```

---

# spire.xls python worksheet visibility
## hide or show worksheets in Excel workbook
```python
#Hide the sheet named "Sheet1"
workbook.Worksheets["Sheet1"].Visibility = WorksheetVisibility.Hidden
#Show the second sheet
workbook.Worksheets[1].Visibility = WorksheetVisibility.Visible
```

---

# Hide Excel Worksheet Tabs
## This example demonstrates how to hide worksheet tabs in an Excel workbook
```python
# Create a workbook
workbook = Workbook()
# Hide worksheet tab
workbook.ShowTabs = False
```

---

# Spire.XLS Hide Zero Values
## Hide zero values in Excel worksheet
```python
#Get the first sheet
sheet = workbook.Worksheets[0]
#Set false to hide the zero values in sheet
sheet.IsDisplayZeros = False
```

---

# Excel Custom Document Property Linking
## Link custom document property to content in Excel workbook
```python
#Add a custom document property
workbook.CustomDocumentProperties.Add("Test", "MyNamedRange")
#Get the added document property
properties = workbook.CustomDocumentProperties
property = properties["Test"]
#Link to content 
property.LinkToContent = True
```

---

# spire.xls python worksheet
## move worksheet to a specific position
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Move worksheet
sheet.MoveWorksheet(2)
```

---

# Spire.XLS Page Break Preview
## Set zoom scale for Page Break View in Excel worksheet
```python
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Set the scale of PageBreakView mode in Excel file.
sheet.ZoomScalePageBreakView = 80
```

---

# Remove Page Breaks in Excel
## This code demonstrates how to remove vertical and horizontal page breaks in an Excel worksheet
```python
# Get the first worksheet from the workbook
sheet = workbook.Worksheets[0]
# Clear all the vertical page breaks
sheet.VPageBreaks.Clear()
# Remove the first horizontal Page Break
sheet.HPageBreaks.RemoveAt(0)
# Set the ViewMode as Preview to see how the page breaks work
sheet.ViewMode = ViewMode.Preview
```

---

# Remove Worksheet
## Remove worksheet from workbook by index
```python
# Create a workbook
workbook = Workbook()
# Remove a worksheet by sheet index
workbook.Worksheets.RemoveAt(1)
```

---

# spire.xls python page breaks
## set horizontal and vertical page breaks in Excel worksheet and change view mode
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set Excel Page Break Horizontally
sheet.HPageBreaks.Add(sheet.Range["A8"])
sheet.HPageBreaks.Add(sheet.Range["A14"])
#Set Excel Page Break Vertically
#sheet.VPageBreaks.Add(sheet.Range["B1"])
#sheet.VPageBreaks.Add(sheet.Range["C1"])
#Set view mode to Preview mode
sheet.ViewMode = ViewMode.Preview
```

---

# spire.xls python tab color
## Set tab color for Excel worksheets
```python
#Set the tab color of first sheet to be red 
worksheet = workbook.Worksheets[0]
worksheet.TabColor = Color.get_Red()
#Set the tab color of second sheet to be green 
worksheet = workbook.Worksheets[1]
worksheet.TabColor = Color.get_Green()
#Set the tab color of third sheet to be blue 
worksheet = workbook.Worksheets[2]
worksheet.TabColor = Color.get_LightBlue()
```

---

# spire.xls python worksheet view mode
## set worksheet view mode to preview
```python
#Set the view mode 
workbook.Worksheets[0].ViewMode = ViewMode.Preview
```

---

# spire.xls python gridlines
## show or hide gridlines in Excel worksheets
```python
#Get the first and second worksheet
sheet1 = workbook.Worksheets[0]
sheet2 = workbook.Worksheets[1]
#Hide grid line in the first worksheet
sheet1.GridLinesVisible = False
#Show grid line in the first worksheet
sheet2.GridLinesVisible = True
```

---

# Spire.XLS for Python - Worksheet Tabs
## Show or hide worksheet tabs in Excel
```python
# Create a workbook
workbook = Workbook()
# Show worksheet tab
workbook.ShowTabs = True
```

---

# spire.xls python worksheet
## split worksheet into panes
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Vertical and horizontal split the worksheet into four panes
sheet.FirstVisibleColumn = 2
sheet.FirstVisibleRow = 5
sheet.VerticalSplit = 4000
sheet.HorizontalSplit = 5000
#Set the active pane
sheet.ActivePane = 1
```

---

# spire.xls python unfreeze panes
## Unfreeze panes in Excel worksheet
```python
#Create a workbook
workbook = Workbook()
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Unfreeze the panes
sheet.RemovePanes()
```

---

# spire.xls python worksheet protection verification
## Verify if a worksheet is password protected
```python
#Create a workbook
workbook = Workbook()
#Get the first worksheet
worksheet = workbook.Worksheets[0]
#Verify if the first worksheet is password protected
detect = worksheet.IsPasswordProtected
```

---

# spire.xls python zoom factor
## set worksheet zoom factor
```python
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set the zoom factor of the sheet to 85
sheet.Zoom = 85
```

---

# Access Excel Document Properties
## This example demonstrates how to access custom document properties in an Excel file

```python
#Get all document properties
properties = workbook.CustomDocumentProperties
# Access document property by property name
property1 = properties["Editor"]
obj = property1.Value
# Access document property by property index
property2 = properties[0]
obj2 = property2.Value
```

---

# spire.xls python custom properties
## Add custom properties to an Excel workbook
```python
#Add a custom property to make the document as final
workbook.CustomDocumentProperties.Add("_MarkAsFinal", True)
#Add other custom properties to the workbook
workbook.CustomDocumentProperties.Add("The Editor", "E-iceblue")
workbook.CustomDocumentProperties.Add("Phone number", 81705109)
workbook.CustomDocumentProperties.Add("Revision number", 7.12)
workbook.CustomDocumentProperties.Add("Revision date", DateTime.get_Now())
```

---

# spire.xls python decrypt workbook
## Decrypt password protected Excel workbook
```python
# Detect if the Excel workbook is password protected
outValue = Workbook.IsPasswordProtected(inputFile)

if outValue:
    # Load a file with the password specified
    workbook = Workbook()
    workbook.OpenPassword = "eiceblue"
    workbook.LoadFromFile(inputFile)

    # Decrypt workbook
    workbook.UnProtect()
    workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
```

---

# Excel Version Detection
## Detect the version of an Excel workbook
```python
# Create a workbook
workbook = Workbook()
# Get the version
version = workbook.Version
```

---

# Detect VBA Macros in Excel
## Check if Excel file contains VBA macros
```python
# Create a workbook
workbook = Workbook()
# Detect if the Excel file contains VBA macros
hasMacros = workbook.HasMacros
if hasMacros:
    value = "Yes"
else:
    value = "No"
```

---

# spire.xls python workbook encryption
## encrypt workbook with password
```python
#Create a workbook
workbook = Workbook()
#Protect Workbook with the password you want
workbook.Protect("eiceblue")
workbook.Dispose()
```

---

# spire.xls python get workbook properties
## retrieve general and custom document properties from an Excel workbook
```python
# Create a workbook
workbook = Workbook()
# Load the document from disk
workbook.LoadFromFile(inputFile)
# Get the general excel properties
properties1 = workbook.DocumentProperties
sb = []
sb.append("Excel Properties:")
for i, unusedItem in enumerate(properties1):
    name = properties1[i].Name
    obj = properties1[i].Value
    t = properties1[i].PropertyType
    value = None
    if t == PropertyType.Double:
        value = Double(obj).Value
    elif t == PropertyType.DateTime:
        value = DateTime(obj).ToLongDateString()
    elif t == PropertyType.Bool:
        value = Boolean(obj).Value
    elif t == PropertyType.Int:
        value = Int32(obj).Value
    elif t == PropertyType.Int32:
        value = Int32(obj).Value
    else:
        value = String(obj).Value
    sb.append(name + ": " + str(value))
sb.append("")
# Get the custom properties
properties2 = workbook.CustomDocumentProperties
sb.append("Custom Properties:")
for i, unusedItem in enumerate(properties2):
    name = properties2[i].Name
    t = properties2[i].PropertyType
    obj = properties2[i].Value
    value = None
    if t == PropertyType.Double:
        value = Double(obj).Value
    elif t == PropertyType.DateTime:
        value = DateTime(obj).ToLongDateString()
    elif t == PropertyType.Bool:
        value = Boolean(obj).Value
    elif t == PropertyType.Int:
        value = Int32(obj).Value
    elif t == PropertyType.Int32:
        value = Int32(obj).Value
    else:
        value = String(obj).Value
    sb.append(name + ": " + str(value))
workbook.Dispose()
```

---

# spire.xls python hide window
## Hide Excel window in workbook
```python
#Create a workbook
workbook = Workbook()
#Hide window
workbook.IsHideWindow = True
```

---

# Spire.XLS Python Workbook with Macro
## Load and save Excel files with macros
```python
inputFile = "./Demos/Data/MacroSample.xls"
outputFile = "LoadAndSaveFileWithMacro.xls"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version97to2003)
workbook.Dispose()
```

---

# Merge Excel Files
## Merge multiple Excel files into a single workbook
```python
newbook = Workbook()
newbook.Version = ExcelVersion.Version2013
# Clear all worksheets
newbook.Worksheets.Clear()

# Create a temporary workbook for loading files
tempbook = Workbook()

# Process each file
for file in files:
    # Load the file
    tempbook.LoadFromFile(file)
    # Copy each worksheet to the new workbook
    for sheet in tempbook.Worksheets:
        newbook.Worksheets.AddCopy(sheet, WorksheetCopyType.CopyAll)
```

---

# spire.xls open encrypted Excel file
## Open an encrypted Excel file with password
```python
# Create a workbook
workbook = Workbook()
# Set open password
workbook.OpenPassword = password
# Load the encrypted document
workbook.LoadFromFile(encrypted_file_path)
```

---

# Spire.XLS Python Open Files
## Demonstrate how to open different Excel file formats
```python
# 1. Load file by file path
workbook1 = Workbook()
workbook1.LoadFromFile("path_to_excel_file.xlsx")

# 2. Load file by file stream
stream = Stream("path_to_excel_file.xlsx")
workbook2 = Workbook()
workbook2.LoadFromStream(stream)
stream.Dispose()

# 3. Open Microsoft Excel 97 - 2003 file
wbExcel97 = Workbook()
wbExcel97.LoadFromFile("path_to_excel97_file.xls", ExcelVersion.Version97to2003)

# 4. Open xml file
wbXML = Workbook()
wbXML.LoadFromXml("path_to_xml_file.xml")

# 5. Open csv file
wbCSV = Workbook()
wbCSV.LoadFromFile("path_to_csv_file.csv", ",", 1, 1)
```

---

# spire.xls read from stream
## Load Excel workbook from a stream
```python
workbook = Workbook()
#Open excel from a stream
fileStream = Stream(inputFile)
workbook.LoadFromStream(fileStream)
```

---

# Remove Custom Properties from Excel Workbook
## This code demonstrates how to remove custom properties from an Excel workbook
```python
# Retrieve a list of all custom document properties of the Excel file
customDocumentProperties = workbook.CustomDocumentProperties
# Remove "Editor" custom document property
customDocumentProperties.Remove("Editor")
```

---

# Spire.XLS Python File Saving
## Save Excel workbooks in different file formats
```python
# Save in Excel 97-2003 format
workbook.SaveToFile(outputFile_xls, ExcelVersion.Version97to2003)
# Save in Excel2010 xlsx format
workbook.SaveToFile(outputFile_xlsx, ExcelVersion.Version2010)
# Save in XLSB format
workbook.SaveToFile(outputFile_xlsb, ExcelVersion.Xlsb2010)
# Save in ODS format
workbook.SaveToFile(outputFile_ods, ExcelVersion.ODS)
# Save in PDF format
workbook.SaveToFile(outputFile_pdf, FileFormat.PDF)
# Save in XML format
workbook.SaveToFile(outputFile_xml, FileFormat.XML)
# Save in XPS format
workbook.SaveToFile(outputFile_xps, FileFormat.XPS)
workbook.Dispose()
```

---

# spire.xls python save to stream
## Save Excel workbook to stream
```python
outputFile = "SaveStream.xlsx"

workbook = Workbook()
#Save an excel workbook to stream
fileStream = Stream(outputFile)
workbook.SaveToStream(fileStream, FileFormat.Version2010)
fileStream.Close()
workbook.Dispose()
```

---

# spire.xls python calculation mode
## Set Excel calculation mode to manual
```python
# Set excel calculation mode as Manual
workbook.CalculationMode = ExcelCalculationMode.Manual
```

---

# spire.xls python page setup
## set worksheet margins
```python
#Set margins for top, bottom, left and right, here the unit of measure is Inch
sheet.PageSetup.TopMargin = 0.3
sheet.PageSetup.BottomMargin = 1
sheet.PageSetup.LeftMargin = 0.2
sheet.PageSetup.RightMargin = 1
#Set the header margin and footer margin
sheet.PageSetup.HeaderMarginInch = 0.1
sheet.PageSetup.FooterMarginInch = 0.5
```

---

# Excel Workbook Theme Management
## Set theme colors or copy theme from another workbook
```python
#Create workbooks
srcWorkbook = Workbook()
workbook = Workbook()

# Option 1: Copy the theme from another workbook
#workbook.CopyTheme(srcWorkbook)

# Option 2: Set a specific theme color
workbook.SetThemeColor(ThemeColorType.Dk1, Color.get_SkyBlue())
```

---

# Track Changes in Excel
## Accept or reject all tracked changes in an Excel workbook
```python
#create a workbook
workbook = Workbook()
#accept the changes or reject the changes
#workbook.AcceptAllTrackedChanges()
workbook.RejectAllTrackedChanges()
```

---

