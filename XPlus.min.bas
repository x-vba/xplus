Attribute VB_Name = "XPlus"
'The MIT License (MIT)
'Copyright © 2020 Anthony Mancini
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
Option Explicit
Public Function RGB2HEX(ByVal redColorInteger%, ByVal greenColorInteger%, ByVal blueColorInteger%) As String
RGB2HEX = WorksheetFunction.Dec2Hex(redColorInteger, 2) & WorksheetFunction.Dec2Hex(greenColorInteger, 2) & WorksheetFunction.Dec2Hex(blueColorInteger, 2)
End Function
Public Function HEX2RGB(ByVal hexColorString$, Optional ByVal singleColorNumberOrName = -1)
hexColorString = Replace(hexColorString, "#", "")
If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Red" Then
HEX2RGB = WorksheetFunction.Hex2Dec(Left(hexColorString, 2))
ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Green" Then
HEX2RGB = WorksheetFunction.Hex2Dec(Mid(hexColorString, 3, 2))
ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Blue" Then
HEX2RGB = WorksheetFunction.Hex2Dec(Right(hexColorString, 2))
Else
HEX2RGB = "(" & WorksheetFunction.Hex2Dec(Left(hexColorString, 2)) & ", " & WorksheetFunction.Hex2Dec(Mid(hexColorString, 3, 2)) & ", " & WorksheetFunction.Hex2Dec(Right(hexColorString, 2)) & ")"
End If
End Function
Public Function RGB2HSL(ByVal redColorInteger%, ByVal greenColorInteger%, ByVal blueColorInteger%, Optional ByVal singleColorNumberOrName = -1)
Dim redPrime#
Dim greenPrime#
Dim bluePrime#
redPrime = redColorInteger / 255
greenPrime = greenColorInteger / 255
bluePrime = blueColorInteger / 255
Dim colorMax#
Dim colorMin#
colorMax = WorksheetFunction.Max(redPrime, greenPrime, bluePrime)
colorMin = WorksheetFunction.Min(redPrime, greenPrime, bluePrime)
Dim deltaValue#
deltaValue = colorMax - colorMin
Dim hueValue#
Dim saturationValue#
Dim lightnessValue#
If deltaValue = 0 Then
hueValue = 0
Else
Select Case colorMax
Case redPrime
hueValue = 60 * (((greenPrime - bluePrime) / deltaValue) Mod 6)
Case greenPrime
hueValue = 60 * (((bluePrime - redPrime) / deltaValue) + 2)
Case bluePrime
hueValue = 60 * (((redPrime - greenPrime) / deltaValue) + 4)
End Select
End If
lightnessValue = (colorMax + colorMin) / 2
If deltaValue = 0 Then
saturationValue = 0
Else
saturationValue = deltaValue / (1 - Abs((2 * lightnessValue - 1)))
End If
If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Hue" Then
RGB2HSL = hueValue
ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Saturation" Then
RGB2HSL = saturationValue
ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Lightness" Then
RGB2HSL = lightnessValue
Else
RGB2HSL = "(" & Format(hueValue, "#.0") & ", " & Format(saturationValue * 100, "#.0") & "%, " & Format(lightnessValue * 100, "#.0") & "%)"
End If
End Function
Public Function HEX2HSL(ByVal hexColorString$) As String
hexColorString = Replace(hexColorString, "#", "")
Dim redValue%
Dim greenValue%
Dim blueValue%
redValue = CInt(WorksheetFunction.Hex2Dec(Left(hexColorString, 2)))
greenValue = CInt(WorksheetFunction.Hex2Dec(Mid(hexColorString, 3, 2)))
blueValue = CInt(WorksheetFunction.Hex2Dec(Right(hexColorString, 2)))
HEX2HSL = RGB2HSL(redValue, greenValue, blueValue)
End Function
Private Function ModFloat(numerator#, denominator#) As Double
Dim modValue#
modValue = numerator - Fix(numerator / denominator) * denominator
If modValue >= -2 ^ -52 Then
If modValue <= 2 ^ -52 Then
modValue = 0
End If
End If
ModFloat = modValue
End Function
Public Function HSL2RGB(ByVal hueValue#, ByVal saturationValue#, ByVal lightnessValue#, Optional ByVal singleColorNumberOrName = -1)
Dim cValue#
Dim xValue#
Dim mValue#
cValue = (1 - Abs(2 * lightnessValue - 1)) * saturationValue
xValue = cValue * (1 - Abs(ModFloat((hueValue / 60), 2) - 1))
mValue = lightnessValue - cValue / 2
Dim redValue#
Dim greenValue#
Dim blueValue#
If hueValue >= 0 And hueValue < 60 Then
redValue = cValue
greenValue = xValue
blueValue = 0
ElseIf hueValue >= 60 And hueValue < 120 Then
redValue = xValue
greenValue = cValue
blueValue = 0
ElseIf hueValue >= 120 And hueValue < 180 Then
redValue = 0
greenValue = cValue
blueValue = xValue
ElseIf hueValue >= 180 And hueValue < 240 Then
redValue = 0
greenValue = xValue
blueValue = cValue
ElseIf hueValue >= 240 And hueValue < 300 Then
redValue = xValue
greenValue = 0
blueValue = cValue
ElseIf hueValue >= 300 And hueValue < 360 Then
redValue = cValue
greenValue = 0
blueValue = xValue
End If
redValue = (redValue + mValue) * 255
greenValue = (greenValue + mValue) * 255
blueValue = (blueValue + mValue) * 255
If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Red" Then
HSL2RGB = Round(redValue, 0)
ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Green" Then
HSL2RGB = Round(greenValue, 0)
ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Blue" Then
HSL2RGB = Round(blueValue, 0)
Else
HSL2RGB = "(" & Round(redValue, 0) & ", " & Round(greenValue, 0) & ", " & Round(blueValue, 0) & ")"
End If
End Function
Public Function HSL2HEX(ByVal hueValue#, ByVal saturationValue#, ByVal lightnessValue#)
Dim redValue%
Dim greenValue%
Dim blueValue%
redValue = HSL2RGB(hueValue, saturationValue, lightnessValue, 0)
greenValue = HSL2RGB(hueValue, saturationValue, lightnessValue, 1)
blueValue = HSL2RGB(hueValue, saturationValue, lightnessValue, 2)
HSL2HEX = RGB2HEX(redValue, greenValue, blueValue)
End Function
Public Function RGB2HSV(ByVal redColorInteger%, ByVal greenColorInteger%, ByVal blueColorInteger%, Optional ByVal singleColorNumberOrName = -1)
Dim redPrime#
Dim greenPrime#
Dim bluePrime#
redPrime = redColorInteger / 255
greenPrime = greenColorInteger / 255
bluePrime = blueColorInteger / 255
Dim colorMax#
Dim colorMin#
colorMax = WorksheetFunction.Max(redPrime, greenPrime, bluePrime)
colorMin = WorksheetFunction.Min(redPrime, greenPrime, bluePrime)
Dim deltaValue#
deltaValue = colorMax - colorMin
Dim hueValue#
Dim saturationValue#
Dim valueValue#
If deltaValue = 0 Then
hueValue = 0
Else
Select Case colorMax
Case redPrime
hueValue = 60 * (((greenPrime - bluePrime) / deltaValue) Mod 6)
Case greenPrime
hueValue = 60 * (((bluePrime - redPrime) / deltaValue) + 2)
Case bluePrime
hueValue = 60 * (((redPrime - greenPrime) / deltaValue) + 4)
End Select
End If
If colorMax = 0 Then
saturationValue = 0
Else
saturationValue = deltaValue / colorMax
End If
valueValue = colorMax
If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Hue" Then
RGB2HSV = hueValue
ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Saturation" Then
RGB2HSV = saturationValue
ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Value" Then
RGB2HSV = valueValue
Else
RGB2HSV = "(" & Format(hueValue, "#.0") & ", " & Format(saturationValue * 100, "#.0") & "%, " & Format(valueValue * 100, "#.0") & "%)"
End If
End Function
Public Function WEEKDAY_NAME(Optional ByVal dayNumber As Byte) As String
If dayNumber = 0 Then
WEEKDAY_NAME = WeekdayName(Weekday(Now()))
Else
WEEKDAY_NAME = WeekdayName(dayNumber)
End If
End Function
Public Function MONTH_NAME(Optional ByVal monthNumber As Byte) As String
If monthNumber = 0 Then
MONTH_NAME = MonthName(Month(Now()))
Else
MONTH_NAME = MonthName(monthNumber)
End If
End Function
Public Function QUARTER(Optional ByVal monthNumberOrName) As Byte
If IsMissing(monthNumberOrName) Then
monthNumberOrName = MonthName(Month(Now()))
End If
If IsNumeric(monthNumberOrName) Then
monthNumberOrName = MonthName(monthNumberOrName)
End If
If monthNumberOrName = MonthName(1) Or monthNumberOrName = MonthName(2) Or monthNumberOrName = MonthName(3) Then
QUARTER = 1
End If
If monthNumberOrName = MonthName(4) Or monthNumberOrName = MonthName(5) Or monthNumberOrName = MonthName(6) Then
QUARTER = 2
End If
If monthNumberOrName = MonthName(7) Or monthNumberOrName = MonthName(8) Or monthNumberOrName = MonthName(9) Then
QUARTER = 3
End If
If monthNumberOrName = MonthName(10) Or monthNumberOrName = MonthName(11) Or monthNumberOrName = MonthName(12) Then
QUARTER = 4
End If
End Function
Public Function TIME_CONVERTER(ByVal date1 As Date, Optional ByVal secondsInteger%, Optional ByVal minutesInteger%, Optional ByVal hoursInteger%, Optional ByVal daysInteger%, Optional ByVal monthsInteger%, Optional ByVal yearsInteger%) As Date
secondsInteger = Second(date1) + secondsInteger
minutesInteger = Minute(date1) + minutesInteger
hoursInteger = Hour(date1) + hoursInteger
daysInteger = Day(date1) + daysInteger
monthsInteger = Month(date1) + monthsInteger
yearsInteger = Year(date1) + yearsInteger
TIME_CONVERTER = DateSerial(yearsInteger, monthsInteger, daysInteger) + TimeSerial(hoursInteger, minutesInteger, secondsInteger)
End Function
Public Function DAYS_OF_MONTH(Optional ByVal monthNumberOrName, Optional ByVal yearNumber%)
If IsMissing(monthNumberOrName) Then
monthNumberOrName = Month(Now())
End If
If yearNumber = 0 Then
yearNumber = Year(Now())
End If
If monthNumberOrName = 1 Or monthNumberOrName = MonthName(1) Then
DAYS_OF_MONTH = 31
ElseIf monthNumberOrName = 2 Or monthNumberOrName = MonthName(2) Then
If yearNumber Mod 4 <> 0 Then
DAYS_OF_MONTH = 28
Else
DAYS_OF_MONTH = 29
End If
ElseIf monthNumberOrName = 3 Or monthNumberOrName = MonthName(3) Then
DAYS_OF_MONTH = 31
ElseIf monthNumberOrName = 4 Or monthNumberOrName = MonthName(4) Then
DAYS_OF_MONTH = 30
ElseIf monthNumberOrName = 5 Or monthNumberOrName = MonthName(5) Then
DAYS_OF_MONTH = 31
ElseIf monthNumberOrName = 6 Or monthNumberOrName = MonthName(6) Then
DAYS_OF_MONTH = 30
ElseIf monthNumberOrName = 7 Or monthNumberOrName = MonthName(7) Then
DAYS_OF_MONTH = 31
ElseIf monthNumberOrName = 8 Or monthNumberOrName = MonthName(8) Then
DAYS_OF_MONTH = 31
ElseIf monthNumberOrName = 9 Or monthNumberOrName = MonthName(9) Then
DAYS_OF_MONTH = 30
ElseIf monthNumberOrName = 10 Or monthNumberOrName = MonthName(10) Then
DAYS_OF_MONTH = 31
ElseIf monthNumberOrName = 11 Or monthNumberOrName = MonthName(11) Then
DAYS_OF_MONTH = 30
ElseIf monthNumberOrName = 12 Or monthNumberOrName = MonthName(12) Then
DAYS_OF_MONTH = 31
Else
DAYS_OF_MONTH = "#NotAValidMonthNumberOrName"
End If
End Function
Public Function ENVIRONMENT(ByVal environmentVariableNameString$) As String
Dim WshShell As Object
Set WshShell = CreateObject("Wscript.Shell")
ENVIRONMENT = WshShell.ExpandEnvironmentStrings("%" & environmentVariableNameString & "%")
End Function
Private Function GetActiveWorkbookPath()
Dim filePath$
filePath = ThisWorkbook.Path & "\" & ThisWorkbook.Name
GetActiveWorkbookPath = filePath
End Function
Public Function FILE_CREATION_TIME(Optional ByVal filePath$) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FILE_CREATION_TIME = FSO.GetFile(GetActiveWorkbookPath()).DateCreated
Else
If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
FILE_CREATION_TIME = FSO.GetFile(ThisWorkbook.Path & "\" & filePath).DateCreated
Else
FILE_CREATION_TIME = FSO.GetFile(filePath).DateCreated
End If
End If
End Function
Public Function FILE_LAST_MODIFIED_TIME(Optional ByVal filePath$) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FILE_LAST_MODIFIED_TIME = FSO.GetFile(GetActiveWorkbookPath()).DateLastModified
Else
If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
FILE_LAST_MODIFIED_TIME = FSO.GetFile(ThisWorkbook.Path & "\" & filePath).DateLastModified
Else
FILE_LAST_MODIFIED_TIME = FSO.GetFile(filePath).DateLastModified
End If
End If
End Function
Public Function FILE_DRIVE(Optional ByVal filePath$) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FILE_DRIVE = FSO.GetFile(GetActiveWorkbookPath()).Drive
Else
If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
FILE_DRIVE = FSO.GetFile(ThisWorkbook.Path & "\" & filePath).Drive
Else
FILE_DRIVE = FSO.GetFile(filePath).Drive
End If
End If
End Function
Public Function FILE_NAME(Optional ByVal filePath$) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FILE_NAME = FSO.GetFile(GetActiveWorkbookPath()).Name
Else
If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
FILE_NAME = FSO.GetFile(ThisWorkbook.Path & "\" & filePath).Name
Else
FILE_NAME = FSO.GetFile(filePath).Name
End If
End If
End Function
Public Function FILE_FOLDER(Optional ByVal filePath$) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FILE_FOLDER = FSO.GetFile(GetActiveWorkbookPath()).ParentFolder
Else
If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
FILE_FOLDER = FSO.GetFile(ThisWorkbook.Path & "\" & filePath).ParentFolder
Else
FILE_FOLDER = FSO.GetFile(filePath).ParentFolder
End If
End If
End Function
Public Function FILE_PATH(Optional ByVal filePath$) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FILE_PATH = FSO.GetFile(GetActiveWorkbookPath()).Path
Else
If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
FILE_PATH = FSO.GetFile(ThisWorkbook.Path & "\" & filePath).Path
Else
FILE_PATH = FSO.GetFile(filePath).Path
End If
End If
End Function
Public Function FILE_SIZE(Optional ByVal filePath$, Optional ByVal byteSize$) As Double
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim totalBytes#
If filePath = "" Then
totalBytes = FSO.GetFile(GetActiveWorkbookPath()).Size
Else
If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
totalBytes = FSO.GetFile(ThisWorkbook.Path & "\" & filePath).Size
Else
totalBytes = FSO.GetFile(filePath).Size
End If
End If
Select Case LCase(byteSize)
Case "kb"
totalBytes = totalBytes / (2 ^ 10)
Case "mb"
totalBytes = totalBytes / (2 ^ 20)
Case "gb"
totalBytes = totalBytes / (2 ^ 30)
End Select
FILE_SIZE = totalBytes
End Function
Public Function FILE_TYPE(Optional ByVal filePath$) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FILE_TYPE = FSO.GetFile(GetActiveWorkbookPath()).Type
Else
If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
FILE_TYPE = FSO.GetFile(ThisWorkbook.Path & "\" & filePath).Type
Else
FILE_TYPE = FSO.GetFile(filePath).Type
End If
End If
End Function
Public Function FILE_EXTENSION(Optional ByVal filePath$) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim fileName$
If filePath = "" Then
fileName = FSO.GetFile(GetActiveWorkbookPath()).Name
Else
If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
fileName = FSO.GetFile(ThisWorkbook.Path & "\" & filePath).Name
Else
fileName = FSO.GetFile(filePath).Name
End If
End If
FILE_EXTENSION = Right(fileName, Len(fileName) - InStrRev(fileName, "."))
End Function
Public Function READ_FILE(ByVal filePath$, Optional ByVal lineNumber%) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim fileName$
Dim fileStream As Object
If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
filePath = ThisWorkbook.Path & "\" & filePath
ElseIf FSO.FileExists(filePath) Then
filePath = filePath
Else
READ_FILE = "#FileDoesntExist!"
End If
Set fileStream = FSO.GetFile(filePath)
Set fileStream = fileStream.OpenAsTextStream(1, -2)
If lineNumber > 0 Then
Dim fileLinesArray() As String
fileLinesArray = SPLIT(fileStream.ReadAll(), vbCrLf)
READ_FILE = fileLinesArray(lineNumber)
Else
READ_FILE = fileStream.ReadAll()
End If
End Function
Public Function WRITE_FILE(ByVal filePath$, ByVal fileText$, Optional ByVal appendModeFlag As Boolean) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim fileName$
Dim fileStream As Object
If InStr(filePath, "\") = 0 Then
If InStr(filePath, "/") = 0 Then
filePath = ThisWorkbook.Path & "\" & filePath
End If
ElseIf Right(filePath, 1) = "\" Or Right(filePath, 1) = "/" Then
If Not FSO.FolderExists(Left(filePath, InStrRev(filePath, "\"))) Then
WRITE_FILE = "#FolderDoesNotExist!"
Exit Function
End If
ElseIf Not FSO.FolderExists(filePath) Then
WRITE_FILE = "#FolderDoesNotExist!"
Exit Function
End If
Set fileStream = FSO.CreateTextFile(filePath, Not appendModeFlag)
fileStream.Write fileText
WRITE_FILE = "Successfully wrote to: " & filePath
End Function
Public Function PATH_JOIN(ParamArray pathArray()) As String
Dim individualPath
Dim combinedPath$
Dim individualRange As Range
For Each individualPath In pathArray
If TypeName(individualPath) = "Range" Then
For Each individualRange In individualPath
combinedPath = combinedPath & individualRange.Value & "\"
Next
Else
combinedPath = combinedPath & CStr(individualPath) & "\"
End If
Next
combinedPath = Left(combinedPath, Len(combinedPath) - 1)
PATH_JOIN = combinedPath
End Function
Public Function INTERPOLATE_NUMBER(ByVal startingNumber#, ByVal endingNumber#, ByVal interpolationPercentage#) As Double
INTERPOLATE_NUMBER = startingNumber + ((endingNumber - startingNumber) * interpolationPercentage)
End Function
Public Function INTERPOLATE_PERCENT(ByVal startingNumber#, ByVal endingNumber#, ByVal interpolationNumber#) As Double
INTERPOLATE_PERCENT = (interpolationNumber - startingNumber) / (endingNumber - startingNumber)
End Function
Public Function USER_NAME() As String
Dim WshNetwork As Object
Set WshNetwork = CreateObject("WScript.Network")
USER_NAME = WshNetwork.UserName
End Function
Public Function USER_DOMAIN() As String
Dim WshNetwork As Object
Set WshNetwork = CreateObject("WScript.Network")
USER_DOMAIN = WshNetwork.UserDomain
End Function
Public Function COMPUTER_NAME() As String
Dim WshNetwork As Object
Set WshNetwork = CreateObject("WScript.Network")
COMPUTER_NAME = WshNetwork.ComputerName
End Function
Public Function RANGE_COMMENT(ByVal range1 As Range, Optional ByVal excludeUsername As Boolean) As String
Application.Volatile
If excludeUsername Then
RANGE_COMMENT = Mid(range1.Comment.Text, InStr(range1.Comment.Text, ":") + 1)
Else
RANGE_COMMENT = range1.Comment.Text
End If
End Function
Public Function RANGE_HYPERLINK(ByVal range1 As Range) As String
Application.Volatile
RANGE_HYPERLINK = range1.Hyperlinks(1).Name
End Function
Public Function RANGE_NUMBER_FORMAT(ByVal range1 As Range) As String
Application.Volatile
RANGE_NUMBER_FORMAT = range1.NumberFormat
End Function
Public Function RANGE_FONT(ByVal range1 As Range) As String
Application.Volatile
RANGE_FONT = range1.Font.Name
End Function
Public Function RANGE_NAME(ByVal range1 As Range) As String
Application.Volatile
RANGE_NAME = range1.Name.Name
End Function
Public Function RANGE_WIDTH(ByVal range1 As Range) As Double
Application.Volatile
RANGE_WIDTH = range1.Width
End Function
Public Function RANGE_HEIGHT(ByVal range1 As Range) As Double
Application.Volatile
RANGE_HEIGHT = range1.Height
End Function
Public Function SHEET_NAME(Optional ByVal sheetNameOrNumber) As String
Application.Volatile
If IsMissing(sheetNameOrNumber) Then
SHEET_NAME = Application.Caller.Parent.Name
Else
SHEET_NAME = Sheets(sheetNameOrNumber).Name
End If
End Function
Public Function SHEET_CODE_NAME(Optional ByVal sheetNameOrNumber) As String
Application.Volatile
If IsMissing(sheetNameOrNumber) Then
SHEET_CODE_NAME = Application.Caller.Parent.CodeName
Else
SHEET_CODE_NAME = Sheets(sheetNameOrNumber).CodeName
End If
End Function
Public Function SHEET_TYPE(Optional ByVal sheetNameOrNumber) As String
Application.Volatile
Dim sheetTypeInteger%
If IsMissing(sheetNameOrNumber) Then
sheetTypeInteger = Application.Caller.Parent.Type
Else
sheetTypeInteger = Sheets(sheetNameOrNumber).Type
End If
Select Case sheetTypeInteger
Case xlChart
SHEET_TYPE = "Chart"
Case xlDialogSheet
SHEET_TYPE = "Dialog Sheet"
Case xlExcel4IntlMacroSheet
SHEET_TYPE = "Excel Version 4 International Macro Sheet"
Case xlExcel4MacroSheet
SHEET_TYPE = "Excel Version 4 Macro Sheet"
Case xlWorksheet
SHEET_TYPE = "Worksheet"
End Select
End Function
Public Function WORKBOOK_TITLE(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_TITLE = ThisWorkbook.BuiltinDocumentProperties("Title")
Else
WORKBOOK_TITLE = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Title")
End If
End Function
Public Function WORKBOOK_SUBJECT(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_SUBJECT = ThisWorkbook.BuiltinDocumentProperties("Subject")
Else
WORKBOOK_SUBJECT = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Subject")
End If
End Function
Public Function WORKBOOK_AUTHOR(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_AUTHOR = ThisWorkbook.BuiltinDocumentProperties("Author")
Else
WORKBOOK_AUTHOR = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Author")
End If
End Function
Public Function WORKBOOK_MANAGER(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_MANAGER = ThisWorkbook.BuiltinDocumentProperties("Manager")
Else
WORKBOOK_MANAGER = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Manager")
End If
End Function
Public Function WORKBOOK_COMPANY(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_COMPANY = ThisWorkbook.BuiltinDocumentProperties("Company")
Else
WORKBOOK_COMPANY = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Company")
End If
End Function
Public Function WORKBOOK_CATEGORY(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_CATEGORY = ThisWorkbook.BuiltinDocumentProperties("Category")
Else
WORKBOOK_CATEGORY = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Category")
End If
End Function
Public Function WORKBOOK_KEYWORDS(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_KEYWORDS = ThisWorkbook.BuiltinDocumentProperties("Keywords")
Else
WORKBOOK_KEYWORDS = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Keywords")
End If
End Function
Public Function WORKBOOK_COMMENTS(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_COMMENTS = ThisWorkbook.BuiltinDocumentProperties("Comments")
Else
WORKBOOK_COMMENTS = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Comments")
End If
End Function
Public Function WORKBOOK_HYPERLINK_BASE(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_HYPERLINK_BASE = ThisWorkbook.BuiltinDocumentProperties("Hyperlink Base")
Else
WORKBOOK_HYPERLINK_BASE = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Hyperlink Base")
End If
End Function
Public Function WORKBOOK_REVISION_NUMBER(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_REVISION_NUMBER = ThisWorkbook.BuiltinDocumentProperties("Revision Number")
Else
WORKBOOK_REVISION_NUMBER = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Revision Number")
End If
End Function
Public Function WORKBOOK_CREATION_DATE(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_CREATION_DATE = ThisWorkbook.BuiltinDocumentProperties("Creation Date")
Else
WORKBOOK_CREATION_DATE = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Creation Date")
End If
End Function
Public Function WORKBOOK_LAST_SAVE_TIME(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_LAST_SAVE_TIME = ThisWorkbook.BuiltinDocumentProperties("Last Save Time")
Else
WORKBOOK_LAST_SAVE_TIME = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Last Save Time")
End If
End Function
Public Function WORKBOOK_LAST_AUTHOR(Optional ByVal workbookNameOrNumber) As String
Application.Volatile
If IsMissing(workbookNameOrNumber) Then
WORKBOOK_LAST_AUTHOR = ThisWorkbook.BuiltinDocumentProperties("Last Author")
Else
WORKBOOK_LAST_AUTHOR = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Last Author")
End If
End Function
Public Function RANDOM_SAMPLE(ByVal rangeArray As Range)
Dim randomNumber%
randomNumber = WorksheetFunction.RandBetween(1, rangeArray.Count)
RANDOM_SAMPLE = rangeArray(randomNumber).Value
End Function
Public Function RANDOM_RANGE(ByVal startNumber%, ByVal stopNumber%, ByVal stepNumber%) As Integer
Dim randomNumber%
randomNumber = WorksheetFunction.RandBetween(startNumber / stepNumber, stopNumber / stepNumber) * stepNumber
RANDOM_RANGE = randomNumber
End Function
Public Function RANDOM_SAMPLE_PERCENT(ByVal valueRange As Range, ByVal percentRange As Range)
Application.Volatile
Dim i%
Dim cumulativeSum#
Dim cumulativePercentage#
Dim dataGrid()
ReDim dataGrid(1 To valueRange.Count, 1 To 3)
For i = 1 To valueRange.Count
cumulativeSum = cumulativeSum + percentRange(i).Value
Next
For i = 1 To valueRange.Count
cumulativePercentage = cumulativePercentage + percentRange(i).Value / cumulativeSum
dataGrid(i, 1) = valueRange(i).Value
dataGrid(i, 2) = percentRange(i).Value
dataGrid(i, 3) = cumulativePercentage
Next
Dim randomNumber#
Dim randomValue
randomNumber = Rnd
For i = 2 To valueRange.Count
If randomNumber > dataGrid(i - 1, 3) And randomNumber <= dataGrid(i, 3) Then
randomValue = dataGrid(i, 1)
End If
Next
If IsEmpty(randomValue) Then
randomValue = dataGrid(1, 1)
End If
RANDOM_SAMPLE_PERCENT = randomValue
End Function
Public Function FIRST_UNIQUE(ByVal range1 As Range, ByVal rangeArray As Range) As Boolean
Dim individualRange As Range
For Each individualRange In rangeArray
If range1.Value = individualRange.Value Then
If range1.Address = individualRange.Address Then
FIRST_UNIQUE = True
Exit For
Else
FIRST_UNIQUE = False
Exit For
End If
End If
Next
End Function
Private Function BubbleSort(ByRef sortableArray, Optional ByVal descendingFlag As Boolean)
Dim i%
Dim swapOccuredBool As Boolean
Dim arrayLength%
arrayLength = UBound(sortableArray) - LBound(sortableArray) + 1
Dim sortedArray()
ReDim sortedArray(arrayLength)
For i = 0 To arrayLength - 1
sortedArray(i) = sortableArray(i)
Next
Dim temporaryValue
Do
swapOccuredBool = False
For i = 0 To arrayLength - 1
If (sortedArray(i)) < sortedArray(i + 1) Then
temporaryValue = sortedArray(i)
sortedArray(i) = sortedArray(i + 1)
sortedArray(i + 1) = temporaryValue
swapOccuredBool = True
End If
Next
Loop While swapOccuredBool
If descendingFlag = True Then
BubbleSort = sortedArray
Else
Dim ascendingArray()
ReDim ascendingArray(arrayLength)
For i = 0 To arrayLength - 1
ascendingArray(i) = sortedArray(arrayLength - i - 1)
Next
BubbleSort = ascendingArray
End If
End Function
Function SORT_RANGE(ByVal range1 As Range, ByVal rangeArray As Range, Optional ByVal descendingFlag As Boolean)
Dim variantArray()
ReDim variantArray(rangeArray.Count)
Dim returnArray()
ReDim returnArray(rangeArray.Count)
Dim returnBoolean As Boolean
Dim i%
For i = 1 To rangeArray.Count
variantArray(i) = rangeArray(i)
Next
returnArray = BubbleSort(variantArray, descendingFlag)
Dim k%
k = 1
Do Until range1.Address = rangeArray(k).Address
k = k + 1
Loop
If descendingFlag Then
SORT_RANGE = returnArray(k - 1)
Else
SORT_RANGE = returnArray(k)
End If
End Function
Public Function REVERSE_RANGE(ByVal range1 As Range, ByVal rangeArray As Range)
Dim i%
For i = 1 To rangeArray.Count
If range1.Address = rangeArray(i).Address Then
REVERSE_RANGE = rangeArray(rangeArray.Count - i + 1).Value
Exit Function
End If
Next
End Function
Public Function COLUMNIFY(ByVal columnRangeArray As Range, ByVal rowRangeArray As Range)
Dim i%
Dim individualRange As Range
i = 0
For Each individualRange In columnRangeArray
i = i + 1
If Application.Caller.Address = individualRange.Address Then
Exit For
End If
Next
COLUMNIFY = rowRangeArray(ColumnIndex:=i)
End Function
Public Function ROWIFY(ByVal rowRangeArray As Range, ByVal columnRangeArray As Range)
Dim i%
Dim individualRange As Range
i = 0
For Each individualRange In rowRangeArray
i = i + 1
If Application.Caller.Address = individualRange.Address Then
Exit For
End If
Next
ROWIFY = columnRangeArray(RowIndex:=i)
End Function
Public Function SUMN(ByVal rangeArray As Range, ByVal nthNumber%, Optional ByVal startAtBeginningFlag As Boolean)
Dim i%
Dim sumValue#
For i = 1 To rangeArray.Count
If startAtBeginningFlag Then
If i Mod nthNumber = 1 Then
sumValue = sumValue + rangeArray(i).Value
End If
Else
If i Mod nthNumber = 0 Then
sumValue = sumValue + rangeArray(i).Value
End If
End If
Next
SUMN = sumValue
End Function
Public Function AVERAGEN(ByVal rangeArray As Range, ByVal nthNumber%, Optional ByVal startAtBeginningFlag As Boolean)
Dim i%
Dim sumValue#
Dim countValue%
For i = 1 To rangeArray.Count
If startAtBeginningFlag Then
If i Mod nthNumber = 1 Then
countValue = countValue + 1
sumValue = sumValue + rangeArray(i).Value
End If
Else
If i Mod nthNumber = 0 Then
countValue = countValue + 1
sumValue = sumValue + rangeArray(i).Value
End If
End If
Next
AVERAGEN = sumValue / countValue
End Function
Public Function MAXN(ByVal rangeArray As Range, ByVal nthNumber%, Optional ByVal startAtBeginningFlag As Boolean)
Dim i%
Dim sumValue#
Dim maxValue
maxValue = Null
For i = 1 To rangeArray.Count
If startAtBeginningFlag Then
If i Mod nthNumber = 1 Then
If IsNull(maxValue) Then
maxValue = rangeArray(i).Value
ElseIf maxValue < rangeArray(i).Value Then
maxValue = rangeArray(i).Value
End If
End If
Else
If i Mod nthNumber = 0 Then
If IsNull(maxValue) Then
maxValue = rangeArray(i).Value
ElseIf maxValue < rangeArray(i).Value Then
maxValue = rangeArray(i).Value
End If
End If
End If
Next
MAXN = maxValue
End Function
Public Function MINN(ByVal rangeArray As Range, ByVal nthNumber%, Optional ByVal startAtBeginningFlag As Boolean)
Dim i%
Dim sumValue#
Dim minValue
minValue = Null
For i = 1 To rangeArray.Count
If startAtBeginningFlag Then
If i Mod nthNumber = 1 Then
If IsNull(minValue) Then
minValue = rangeArray(i).Value
ElseIf minValue > rangeArray(i).Value Then
minValue = rangeArray(i).Value
End If
End If
Else
If i Mod nthNumber = 0 Then
If IsNull(minValue) Then
minValue = rangeArray(i).Value
ElseIf minValue > rangeArray(i).Value Then
minValue = rangeArray(i).Value
End If
End If
End If
Next
MINN = minValue
End Function
Public Function SUMHIGH(ByVal rangeArray As Range, ByVal numberSummed%)
Dim i%
Dim sumValue#
For i = 1 To numberSummed
sumValue = sumValue + WorksheetFunction.Large(rangeArray, i)
Next
SUMHIGH = sumValue
End Function
Public Function SUMLOW(ByVal rangeArray As Range, ByVal numberSummed%)
Dim i%
Dim sumValue#
For i = 1 To numberSummed
sumValue = sumValue + WorksheetFunction.Small(rangeArray, i)
Next
SUMLOW = sumValue
End Function
Public Function AVERAGEHIGH(ByVal rangeArray As Range, ByVal numberAveraged%)
Dim i%
Dim sumValue#
For i = 1 To numberAveraged
sumValue = sumValue + WorksheetFunction.Large(rangeArray, i)
Next
AVERAGEHIGH = sumValue / numberAveraged
End Function
Public Function AVERAGELOW(ByVal rangeArray As Range, ByVal numberAveraged%)
Dim i%
Dim sumValue#
For i = 1 To numberAveraged
sumValue = sumValue + WorksheetFunction.Small(rangeArray, i)
Next
AVERAGELOW = sumValue / numberAveraged
End Function
Public Function INRANGE(ByVal valueOrRange, ByVal searchRange As Range) As Boolean
Dim individualValueRange As Range
Dim individualSearchRange As Range
If TypeName(valueOrRange) = "Range" Then
For Each individualValueRange In valueOrRange
For Each individualSearchRange In searchRange
If individualValueRange.Value = individualSearchRange.Value Then
INRANGE = True
Exit Function
End If
Next
Next
Else
For Each individualSearchRange In searchRange
If valueOrRange = individualSearchRange Then
INRANGE = True
Exit Function
End If
Next
End If
INRANGE = False
End Function
Public Function REGEX_SEARCH(ByVal string1$, ByVal stringPattern$, Optional ByVal globalFlag As Boolean, Optional ByVal ignoreCaseFlag As Boolean, Optional ByVal multilineFlag As Boolean) As String
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
Dim searchResults As Object
Dim regexResult As Object
With Regex
.Global = globalFlag
.IgnoreCase = ignoreCaseFlag
.MultiLine = multilineFlag
.Pattern = stringPattern
End With
Set searchResults = Regex.Execute(string1)
REGEX_SEARCH = searchResults(0).Value
End Function
Public Function REGEX_TEST(ByVal string1$, ByVal stringPattern$, Optional ByVal globalFlag As Boolean, Optional ByVal ignoreCaseFlag As Boolean, Optional ByVal multilineFlag As Boolean) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
Dim searchResults As Object
Dim regexResult As Object
With Regex
.Global = globalFlag
.IgnoreCase = ignoreCaseFlag
.MultiLine = multilineFlag
.Pattern = stringPattern
End With
REGEX_TEST = Regex.Test(string1)
End Function
Public Function REGEX_REPLACE(ByVal string1$, ByVal stringPattern$, ByVal replacementString$, Optional ByVal globalFlag As Boolean, Optional ByVal ignoreCaseFlag As Boolean, Optional ByVal multilineFlag As Boolean) As String
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
Dim searchResults As Object
Dim regexResult As Object
With Regex
.Global = globalFlag
.IgnoreCase = ignoreCaseFlag
.MultiLine = multilineFlag
.Pattern = stringPattern
End With
REGEX_REPLACE = Regex.Replace(string1, replacementString)
End Function
Public Function CAPITALIZE(ByVal string1$) As String
CAPITALIZE = UCase(Left(string1, 1)) & LCase(Mid(string1, 2))
End Function
Public Function LEFT_FIND(ByVal string1$, ByVal searchString$) As String
LEFT_FIND = Left(string1, InStr(1, string1, searchString) - 1)
End Function
Public Function RIGHT_FIND(ByVal string1$, ByVal searchString$) As String
RIGHT_FIND = Right(string1, Len(string1) - InStrRev(string1, searchString))
End Function
Public Function LEFT_SEARCH(ByVal string1$, ByVal searchString$) As String
LEFT_SEARCH = Left(string1, InStr(1, string1, searchString, vbTextCompare) - 1)
End Function
Public Function RIGHT_SEARCH(ByVal string1$, ByVal searchString$) As String
RIGHT_SEARCH = Right(string1, Len(string1) - InStrRev(string1, searchString, Compare:=vbTextCompare))
End Function
Public Function SUBSTR(ByVal string1$, ByVal startCharacterNumber%, ByVal endCharacterNumber%) As String
SUBSTR = Mid(string1, startCharacterNumber, endCharacterNumber - startCharacterNumber)
End Function
Public Function SUBSTR_SEARCH(ByVal string1$, ByVal leftSearchString$, ByVal rightSearchString$, Optional ByVal noninclusiveFlag As Boolean) As String
Dim leftCharacterNumber%
Dim rightCharacterNumber%
leftCharacterNumber = InStr(1, string1, leftSearchString)
rightCharacterNumber = InStrRev(string1, rightSearchString)
If noninclusiveFlag = True Then
leftCharacterNumber = leftCharacterNumber + 1
rightCharacterNumber = rightCharacterNumber - 1
End If
SUBSTR_SEARCH = Mid(string1, leftCharacterNumber, rightCharacterNumber - leftCharacterNumber + 1)
End Function
Public Function REPEAT(ByVal string1$, ByVal numberOfRepeats%) As String
Dim i%
Dim combinedString$
For i = 1 To numberOfRepeats
combinedString = combinedString & string1
Next
REPEAT = combinedString
End Function
Public Function FORMATTER(ByVal formatString$, ParamArray textArray()) As String
Dim i As Byte
Dim individualTextItem
Dim individualRange As Range
i = 0
For Each individualTextItem In textArray
If TypeName(individualTextItem) = "Range" Then
For Each individualRange In individualTextItem
i = i + 1
formatString = Replace(formatString, "{" & i & "}", individualRange.Value)
Next
Else
i = i + 1
formatString = Replace(formatString, "{" & i & "}", individualTextItem)
End If
Next
FORMATTER = formatString
End Function
Public Function ZFILL(ByVal string1$, ByVal fillLength As Byte, Optional ByVal fillCharacter$, Optional ByVal rightToLeftFlag As Boolean) As String
If fillCharacter = "" Then
fillCharacter = "0"
End If
While Len(string1) < fillLength
If rightToLeftFlag = False Then
string1 = fillCharacter + string1
Else
string1 = string1 + fillCharacter
End If
Wend
ZFILL = string1
End Function
Public Function TEXT_JOIN(ByVal rangeArray As Range, Optional ByVal delimiterCharacter$, Optional ByVal ignoreEmptyCellsFlag As Boolean) As String
Dim individualRange As Range
Dim combinedString$
For Each individualRange In rangeArray
If ignoreEmptyCellsFlag Then
If Not IsEmpty(individualRange.Value) Then
combinedString = combinedString & individualRange.Value & delimiterCharacter
End If
Else
combinedString = combinedString & individualRange.Value & delimiterCharacter
End If
Next
If delimiterCharacter <> "" Then
combinedString = Left(combinedString, InStrRev(combinedString, delimiterCharacter) - 1)
End If
TEXT_JOIN = combinedString
End Function
Public Function SPLIT_TEXT(ByVal string1$, ByVal substringNumber%, Optional ByVal delimiterString$) As String
If delimiterString = "" Then
SPLIT_TEXT = SPLIT(string1, " ")(substringNumber - 1)
Else
SPLIT_TEXT = SPLIT(string1, delimiterString)(substringNumber - 1)
End If
End Function
Public Function COUNT_WORDS(ByVal string1$, Optional ByVal delimiterString$) As Integer
Dim stringArray() As String
If delimiterString = "" Then
stringArray = SPLIT(string1, " ")
Else
stringArray = SPLIT(string1, delimiterString)
End If
COUNT_WORDS = UBound(stringArray) - LBound(stringArray) + 1
End Function
Public Function CAMEL_CASE(ByVal string1$) As String
Dim i%
Dim stringArray() As String
stringArray = SPLIT(string1, " ")
stringArray(0) = LCase(stringArray(0))
For i = 1 To (UBound(stringArray) - LBound(stringArray))
stringArray(i) = UCase(Left(stringArray(i), 1)) & LCase(Mid(stringArray(i), 2))
Next
CAMEL_CASE = Join(stringArray, "")
End Function
Public Function KEBAB_CASE(ByVal string1$) As String
KEBAB_CASE = LCase(Join(SPLIT(string1, " "), "-"))
End Function
Public Function REMOVE_CHARACTERS(ByVal string1$, ParamArray removedCharacters()) As String
Dim i%
Dim individualCharacter
For Each individualCharacter In removedCharacters
If Len(individualCharacter) > 1 Then
For i = 1 To Len(individualCharacter)
string1 = Replace(string1, Mid(individualCharacter, i, 1), "")
Next
Else
string1 = Replace(string1, individualCharacter, "")
End If
Next
REMOVE_CHARACTERS = string1
End Function
Private Function NumberOfUppercaseLetters(ByVal string1$) As Integer
Dim i%
Dim numberOfUppercase%
For i = 1 To Len(string1)
If Asc(Mid(string1, i, 1)) >= 65 Then
If Asc(Mid(string1, i, 1)) <= 90 Then
numberOfUppercase = numberOfUppercase + 1
End If
End If
Next
NumberOfUppercaseLetters = numberOfUppercase
End Function
Public Function COMPANY_CASE(ByVal string1$) As String
Dim i%
Dim k%
Dim origionalString$
Dim stringArray() As String
Dim splitCharacters$
origionalString = string1
string1 = LCase(string1)
splitCharacters = " ./()-_,*&1234567890"
For k = 1 To Len(splitCharacters)
stringArray = SPLIT(string1, Mid(splitCharacters, k, 1))
For i = 0 To UBound(stringArray) - LBound(stringArray)
If NumberOfUppercaseLetters(SPLIT(origionalString, Mid(splitCharacters, k, 1))(i)) <= 1 Then
stringArray(i) = UCase(Left(stringArray(i), 1)) & Mid(stringArray(i), 2)
Else
If UCase(Join(stringArray, Mid(splitCharacters, k, 1))) = origionalString Then
stringArray(i) = UCase(Left(stringArray(i), 1)) & Mid(stringArray(i), 2)
Else
stringArray(i) = SPLIT(origionalString, Mid(splitCharacters, k, 1))(i)
End If
End If
Next
string1 = Join(stringArray, Mid(splitCharacters, k, 1))
Next
Dim companyAbbreviationArray() As String
companyAbbreviationArray = SPLIT("AB|AG|GmbH|LLC|LLP|NV|PLC|SA|A. en P.|ACE|AD|AE|AL|AmbA|ANS|ApS|AS|ASA|AVV|BVBA|CA|CVA|d.d.|d.n.o.|d.o.o.|DA|e.V.|EE|EEG|EIRL|ELP|EOOD|EPE|EURL|GbR|GCV|GesmbH|GIE|HB|hf|IBC|j.t.d.|k.d.|k.d.d.|k.s.|KA/S|KB|KD|KDA|KG|KGaA|KK|Kol. SrK|Kom. SrK|LDC|Ltï¿½e.|NT|OE|OHG|Oy|OYJ|Oï¿½|PC Ltd|PMA|PMDN|PrC|PT|RAS|S. de R.L.|S. en N.C.|SA de CV|SAFI|SAS|SC|SCA|SCP|SCS|SENC|SGPS|SK|SNC|SOPARFI|sp|Sp. z.o.o.|SpA|spol s.r.o.|SPRL|TD|TLS|v.o.s.|VEB|VOF|BYSHR", "|")
Dim stringArrayLength%
stringArray = SPLIT(string1, " ")
stringArrayLength = UBound(stringArray) - LBound(stringArray)
Dim companyAbbreviationString
For Each companyAbbreviationString In companyAbbreviationArray
If InStrRev(LCase(string1), " " & LCase(companyAbbreviationString)) = (Len(string1) - Len(companyAbbreviationString)) Then
If InStrRev(LCase(string1), " " & LCase(companyAbbreviationString)) <> 0 Then
COMPANY_CASE = Left(string1, InStrRev(LCase(string1), LCase(companyAbbreviationString)) - 1) & companyAbbreviationString
Exit Function
End If
End If
Next
COMPANY_CASE = string1
End Function
Public Function HAMMING(string1$, string2$) As Integer
If Len(string1) <> Len(string2) Then
HAMMING = CVErr(xlErrValue)
End If
Dim totalDistance%
totalDistance = 0
Dim i%
For i = 1 To Len(string1)
If Mid(string1, i, 1) <> Mid(string2, i, 1) Then
totalDistance = totalDistance + 1
End If
Next
HAMMING = totalDistance
End Function
Public Function LEVENSHTEIN(string1$, string2$) As Long
If string1 = string2 Then
LEVENSHTEIN = 0
Exit Function
ElseIf string1 = Empty Then
LEVENSHTEIN = Len(string2)
Exit Function
ElseIf string2 = Empty Then
LEVENSHTEIN = Len(string1)
Exit Function
End If
Dim numberOfRows%
Dim numberOfColumns%
numberOfRows = Len(string1)
numberOfColumns = Len(string2)
Dim distanceArray() As Integer
ReDim distanceArray(numberOfRows, numberOfColumns)
Dim r%
Dim c%
For r = 0 To numberOfRows
For c = 0 To numberOfColumns
distanceArray(r, c) = 0
Next
Next
For r = 1 To numberOfRows
distanceArray(r, 0) = r
Next
For c = 1 To numberOfColumns
distanceArray(0, c) = c
Next
Dim operationCost%
For c = 1 To numberOfColumns
For r = 1 To numberOfRows
If Mid(string1, r, 1) = Mid(string2, c, 1) Then
operationCost = 0
Else
operationCost = 1
End If
distanceArray(r, c) = Min3(distanceArray(r - 1, c) + 1, distanceArray(r, c - 1) + 1, distanceArray(r - 1, c - 1) + operationCost)
Next
Next
LEVENSHTEIN = distanceArray(numberOfRows, numberOfColumns)
End Function
Private Function Min3(integer1%, integer2%, integer3%) As Integer
If integer1 <= integer2 And integer1 <= integer3 Then
Min3 = integer1
ElseIf integer2 <= integer1 And integer2 <= integer3 Then
Min3 = integer2
ElseIf integer3 <= integer1 And integer3 <= integer2 Then
Min3 = integer3
End If
End Function
Public Function LEV_STR(range1 As Range, rangeArray As Range) As String
Dim lngBestDistance&
Dim lngCurrentDistance&
Dim strRange1Value$
Dim strRange1Address$
Dim strBestMatch$
Dim rngCell As Range
lngBestDistance = -1
strRange1Value = range1.Value
strRange1Address = range1.Address
For Each rngCell In rangeArray.Cells
If rngCell.Address <> strRange1Address Then
lngCurrentDistance = LEVENSHTEIN(strRange1Value, rngCell.Value)
If lngCurrentDistance = 0 Then
strBestMatch = rngCell.Value
GoTo Match
ElseIf lngBestDistance = -1 Then
lngBestDistance = lngCurrentDistance
strBestMatch = rngCell.Value
ElseIf lngCurrentDistance < lngBestDistance Then
lngBestDistance = lngCurrentDistance
strBestMatch = rngCell.Value
End If
End If
Next
Match:
LEV_STR = strBestMatch
End Function
Public Function LEV_STR_OPT(range1 As Range, rangeArray As Range, numberOfLeftCharactersBound%, plusOrMinusLengthBound%) As String
Dim lngBestDistance&
Dim lngCurrentDistance&
Dim strRange1Value$
Dim strRange1Address$
Dim strBestMatch$
Dim rngCell As Range
lngBestDistance = -1
strRange1Value = range1.Value
strRange1Address = range1.Address
For Each rngCell In rangeArray.Cells
If Left(rngCell.Value, numberOfLeftCharactersBound) = Left(strRange1Value, numberOfLeftCharactersBound) Then
If Len(strRange1Value) < Len(rngCell.Value) + plusOrMinusLengthBound Then
If Len(strRange1Value) > Len(rngCell.Value) - plusOrMinusLengthBound Then
If rngCell.Address <> strRange1Address Then
lngCurrentDistance = LEVENSHTEIN(strRange1Value, rngCell.Value)
If lngCurrentDistance = 0 Then
strBestMatch = rngCell.Value
GoTo Match
ElseIf lngBestDistance = -1 Then
lngBestDistance = lngCurrentDistance
strBestMatch = rngCell.Value
ElseIf lngCurrentDistance < lngBestDistance Then
lngBestDistance = lngCurrentDistance
strBestMatch = rngCell.Value
End If
End If
End If
End If
End If
Next
Match:
LEV_STR_OPT = strBestMatch
End Function
Public Function DAMERAU(string1$, string2$) As Integer
If string1 = string2 Then
DAMERAU = 0
ElseIf string1 = Empty Then
DAMERAU = Len(string2)
ElseIf string2 = Empty Then
DAMERAU = Len(string1)
End If
Dim inf&
Dim da As Object
inf = Len(string1) + Len(string2)
Set da = CreateObject("Scripting.Dictionary")
Dim i%
For i = 1 To Len(string1)
If da.exists(Mid(string1, i, 1)) = False Then
da.Add Mid(string1, i, 1), "0"
End If
Next
For i = 1 To Len(string2)
If da.exists(Mid(string2, i, 1)) = False Then
da.Add Mid(string2, i, 1), "0"
End If
Next
Dim H() As Long
ReDim H(Len(string1) + 1, Len(string2) + 1)
Dim k%
For i = 0 To (Len(string1) + 1)
For k = 0 To (Len(string2) + 1)
H(i, k) = 0
Next
Next
For i = 0 To Len(string1)
H(i + 1, 0) = inf
H(i + 1, 1) = i
Next
For k = 0 To Len(string2)
H(0, k + 1) = inf
H(1, k + 1) = k
Next
Dim db&
Dim i1&
Dim k1&
Dim cost&
For i = 1 To Len(string1)
db = 0
For k = 1 To Len(string2)
i1 = CInt(da(Mid(string2, k, 1)))
k1 = db
cost = 1
If Mid(string1, i, 1) = Mid(string2, k, 1) Then
cost = 0
db = k
End If
H(i + 1, k + 1) = Min4(H(i, k) + cost, H(i + 1, k) + 1, H(i, k + 1) + 1, H(i1, k1) + (i - i1 - 1) + 1 + (k - k1 - 1))
Next
If da.exists(Mid(string1, i, 1)) Then
da.Remove Mid(string1, i, 1)
da.Add Mid(string1, i, 1), CStr(i)
Else
da.Add Mid(string1, i, 1), CStr(i)
End If
Next
DAMERAU = H(Len(string1) + 1, Len(string2) + 1)
End Function
Private Function Min4(integer1%, integer2%, integer3%, integer4%) As Long
If integer1 <= integer2 And integer1 <= integer3 And integer1 <= integer4 Then
Min4 = integer1
ElseIf integer2 <= integer1 And integer2 <= integer3 And integer2 <= integer4 Then
Min4 = integer2
ElseIf integer3 <= integer1 And integer3 <= integer2 And integer3 <= integer4 Then
Min4 = integer3
ElseIf integer4 <= integer1 And integer4 <= integer2 And integer4 <= integer3 Then
Min4 = integer4
End If
End Function
Public Function DAM_STR(range1 As Range, rangeArray As Range) As String
Dim lngBestDistance&
Dim lngCurrentDistance&
Dim strRange1Value$
Dim strRange1Address$
Dim strBestMatch$
Dim rngCell As Range
lngBestDistance = -1
strRange1Value = range1.Value
strRange1Address = range1.Address
For Each rngCell In rangeArray.Cells
If rngCell.Address <> strRange1Address Then
lngCurrentDistance = DAMERAU(strRange1Value, rngCell.Value)
If lngCurrentDistance = 0 Then
strBestMatch = rngCell.Value
GoTo Match
ElseIf lngBestDistance = -1 Then
lngBestDistance = lngCurrentDistance
strBestMatch = rngCell.Value
ElseIf lngCurrentDistance < lngBestDistance Then
lngBestDistance = lngCurrentDistance
strBestMatch = rngCell.Value
End If
End If
Next
Match:
DAM_STR = strBestMatch
End Function
Public Function DAM_STR_OPT(range1 As Range, rangeArray As Range, numberOfLeftCharactersBound&, plusOrMinusLengthBound) As String
Dim lngBestDistance&
Dim lngCurrentDistance&
Dim strRange1Value$
Dim strRange1Address$
Dim strBestMatch$
Dim rngCell As Range
lngBestDistance = -1
strRange1Value = range1.Value
strRange1Address = range1.Address
For Each rngCell In rangeArray.Cells
If Left(rngCell.Value, numberOfLeftCharactersBound) = Left(strRange1Value, numberOfLeftCharactersBound) Then
If Len(strRange1Value) < Len(rngCell.Value) + plusOrMinusLengthBound Then
If Len(strRange1Value) > Len(rngCell.Value) - plusOrMinusLengthBound Then
If rngCell.Address <> strRange1Address Then
lngCurrentDistance = DAMERAU(strRange1Value, rngCell.Value)
If lngCurrentDistance = 0 Then
strBestMatch = rngCell.Value
GoTo Match
ElseIf lngBestDistance = -1 Then
lngBestDistance = lngCurrentDistance
strBestMatch = rngCell.Value
ElseIf lngCurrentDistance < lngBestDistance Then
lngBestDistance = lngCurrentDistance
strBestMatch = rngCell.Value
End If
End If
End If
End If
End If
Next
Match:
DAM_STR_OPT = strBestMatch
End Function
Public Function PARTIAL_LOOKUP(range1 As Range, rangeArray As Range) As String
PARTIAL_LOOKUP = DAM_STR(range1, rangeArray)
End Function
Public Function DISPLAY_TEXT(ParamArray textArray()) As String
Dim combinedString$
Dim individualTextItem
Dim individualRange As Range
For Each individualTextItem In textArray
If TypeName(individualTextItem) = "Range" Then
For Each individualRange In individualTextItem
combinedString = combinedString & individualRange.Value & vbCrLf & vbCrLf
Next
Else
combinedString = combinedString & individualTextItem & vbCrLf & vbCrLf
End If
Next
If Application.Caller.Parent.Parent.Name = ActiveWorkbook.Name Then
If Application.Caller.Worksheet.Name = ActiveCell.Worksheet.Name Then
If Application.Caller.Address = ActiveCell.Address Then
MsgBox combinedString, , "Cell " & Replace(Application.Caller.Address, "$", "") & " Contents"
End If
End If
End If
DISPLAY_TEXT = combinedString
End Function
Public Function JSONIFY(ByVal indentLevel As Byte, ParamArray stringArray()) As String
Dim i As Byte
Dim jsonString$
Dim individualTextItem
Dim individualRange As Range
Dim indentString$
jsonString = "["
For i = 1 To indentLevel
indentString = indentString & " "
Next
If indentLevel > 0 Then
jsonString = jsonString & Chr(10)
End If
For Each individualTextItem In stringArray
If TypeName(individualTextItem) = "Range" Then
For Each individualRange In individualTextItem
jsonString = jsonString & indentString
If IsNumeric(individualRange.Value) Then
jsonString = jsonString & individualRange.Value & ","
Else
jsonString = jsonString & Chr(34) & individualRange.Value & Chr(34) & ","
End If
If indentLevel > 0 Then
jsonString = jsonString & Chr(10)
End If
Next
Else
jsonString = jsonString & indentString
If IsNumeric(individualTextItem) Then
jsonString = jsonString & individualTextItem & ","
Else
jsonString = jsonString & Chr(34) & individualTextItem & Chr(34) & ","
End If
If indentLevel > 0 Then
jsonString = jsonString & Chr(10)
End If
End If
Next
jsonString = Left(jsonString, InStrRev(jsonString, ",") - 1)
If indentLevel > 0 Then
jsonString = jsonString & Chr(10)
End If
jsonString = jsonString & "]"
JSONIFY = jsonString
End Function
Public Function UUID_FOUR() As String
Dim firstGroup$
Dim secondGroup$
Dim thirdGroup$
Dim fourthGroup$
Dim fifthGroup$
Dim sixthGroup$
firstGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(0, 4294967295#), 8) & "-"
secondGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(0, 65535), 4) & "-"
thirdGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(16384, 20479), 4) & "-"
fourthGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(32768, 49151), 4) & "-"
fifthGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(0, 65535), 4)
sixthGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(0, 4294967295#), 8)
UUID_FOUR = firstGroup & secondGroup & thirdGroup & fourthGroup & fifthGroup & sixthGroup
End Function
Public Function HIDDEN(ByVal string1$, ByVal hiddenFlag As Boolean, Optional ByVal hideString$) As String
If hiddenFlag Then
If hideString = "" Then
HIDDEN = "********"
Else
HIDDEN = hideString
End If
Else
HIDDEN = string1
End If
End Function
Public Function ISERRORALL(ByVal range1 As Range) As Boolean
Dim rangeValue
rangeValue = range1.Value
If IsError(rangeValue) Then
ISERRORALL = True
ElseIf Left(rangeValue, 1) = "#" Then
If Right(rangeValue, 1) = "!" Or Right(rangeValue, 1) = "?" Then
ISERRORALL = True
Else
ISERRORALL = False
End If
Else
ISERRORALL = False
End If
End Function
Public Function JAVASCRIPT(ByVal jsFuncCode$, ByVal jsFuncName$, Optional ByVal argument1, Optional ByVal argument2, Optional ByVal argument3, Optional ByVal argument4, Optional ByVal argument5, Optional ByVal argument6, Optional ByVal argument7, Optional ByVal argument8, Optional ByVal argument9, Optional ByVal argument10, Optional ByVal argument11, Optional ByVal argument12, Optional ByVal argument13, Optional ByVal argument14, Optional ByVal argument15, Optional ByVal argument16)
Dim ScriptContoller As Object
Set ScriptContoller = CreateObject("ScriptControl")
ScriptContoller.Language = "JScript"
ScriptContoller.addCode jsFuncCode
JAVASCRIPT = ScriptContoller.Run(jsFuncName, argument1, argument2, argument3, argument4, argument5, argument6, argument7, argument8, argument9, argument10, argument11, argument12, argument13, argument14, argument15, argument16)
End Function
Public Function SUMSHEET(ByVal partialSheetName$, Optional ByVal range1 As Range)
Dim sumValue
Dim individualSheet As Worksheet
For Each individualSheet In Worksheets
If InStr(individualSheet.Name, partialSheetName) > 0 Then
If range1 Is Nothing Then
sumValue = sumValue + individualSheet.Range(Application.Caller.Address).Value
Else
sumValue = sumValue + individualSheet.Range(range1.Address).Value
End If
End If
Next
SUMSHEET = sumValue
End Function
Public Function AVERAGESHEET(ByVal partialSheetName$, Optional ByVal range1 As Range)
Dim sumValue
Dim countValue%
Dim individualSheet As Worksheet
For Each individualSheet In Worksheets
If InStr(individualSheet.Name, partialSheetName) > 0 Then
If range1 Is Nothing Then
sumValue = sumValue + individualSheet.Range(Application.Caller.Address).Value
countValue = countValue + 1
Else
sumValue = sumValue + individualSheet.Range(range1.Address).Value
countValue = countValue + 1
End If
End If
Next
AVERAGESHEET = (sumValue / countValue)
End Function
Public Function MAXSHEET(ByVal partialSheetName$, Optional ByVal range1 As Range)
Dim maxValue
Dim currentValue
Dim individualSheet As Worksheet
For Each individualSheet In Worksheets
If InStr(individualSheet.Name, partialSheetName) > 0 Then
If range1 Is Nothing Then
currentValue = individualSheet.Range(Application.Caller.Address).Value
Else
currentValue = individualSheet.Range(range1.Address).Value
End If
If IsEmpty(maxValue) Then
maxValue = currentValue
End If
If currentValue > maxValue Then
maxValue = currentValue
End If
End If
Next
MAXSHEET = maxValue
End Function
Public Function MINSHEET(ByVal partialSheetName$, Optional ByVal range1 As Range)
Dim minValue
Dim currentValue
Dim individualSheet As Worksheet
For Each individualSheet In Worksheets
If InStr(individualSheet.Name, partialSheetName) > 0 Then
If range1 Is Nothing Then
currentValue = individualSheet.Range(Application.Caller.Address).Value
Else
currentValue = individualSheet.Range(range1.Address).Value
End If
If IsEmpty(minValue) Then
minValue = currentValue
End If
If currentValue < minValue Then
minValue = currentValue
End If
End If
Next
MINSHEET = minValue
End Function
Public Function HTML_TABLEIFY(ByVal rangeTable As Range) As String
Dim i%
Dim htmlTableString$
Dim individualRange As Range
htmlTableString = htmlTableString & "<table>" & vbCrLf
htmlTableString = htmlTableString & "  <thead>" & vbCrLf
htmlTableString = htmlTableString & "    <tr>" & vbCrLf
For Each individualRange In rangeTable.Rows(1).Cells
htmlTableString = htmlTableString & "      <th>" & individualRange.Value & "</th>" & vbCrLf
Next
htmlTableString = htmlTableString & "    </tr>" & vbCrLf
htmlTableString = htmlTableString & "  </thead>" & vbCrLf
htmlTableString = htmlTableString & "  <tbody>" & vbCrLf
For i = 1 To rangeTable.Rows.Count - 1
htmlTableString = htmlTableString & "    <tr>" & vbCrLf
For Each individualRange In rangeTable.Rows(i + 1).Cells
htmlTableString = htmlTableString & "      <td>" & individualRange.Value & "</td>" & vbCrLf
Next
htmlTableString = htmlTableString & "    </tr>" & vbCrLf
Next
htmlTableString = htmlTableString & "  <tbody>" & vbCrLf
htmlTableString = htmlTableString & "</table>" & vbCrLf
HTML_TABLEIFY = htmlTableString
End Function
Public Function HTML_ESCAPE(ByVal string1$) As String
string1 = Replace(string1, "&", "&amp;")
string1 = Replace(string1, Chr(34), "&quot;")
string1 = Replace(string1, "'", "&apos;")
string1 = Replace(string1, "<", "&lt;")
string1 = Replace(string1, ">", "&gt;")
HTML_ESCAPE = string1
End Function
Public Function HTML_UNESCAPE(ByVal string1$) As String
string1 = Replace(string1, "&amp;", "&")
string1 = Replace(string1, "&quot;", Chr(34))
string1 = Replace(string1, "&apos;", "'")
string1 = Replace(string1, "&lt;", "<")
string1 = Replace(string1, "&gt;", ">")
HTML_UNESCAPE = string1
End Function
Public Function SPEAK_TEXT(ParamArray textArray()) As String
Dim combinedString$
Dim individualTextItem
Dim individualRange As Range
For Each individualTextItem In textArray
If TypeName(individualTextItem) = "Range" Then
For Each individualRange In individualTextItem
combinedString = combinedString & individualRange.Value & " "
Next
Else
combinedString = combinedString & individualTextItem & " "
End If
Next
If Application.Caller.Parent.Parent.Name = ActiveWorkbook.Name Then
If Application.Caller.Worksheet.Name = ActiveCell.Worksheet.Name Then
If Application.Caller.Address = ActiveCell.Address Then
Application.Speech.SPEAK combinedString, True
End If
End If
End If
SPEAK_TEXT = combinedString
End Function
Public Function IS_EMAIL(ByVal string1$) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = "^[a-zA-Z0-9_.]*?[@][a-zA-Z0-9.]*?[.][a-zA-Z]{2,15}$"
End With
IS_EMAIL = Regex.Test(string1)
End Function
Public Function IS_PHONE(ByVal string1$) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = "^\s*[+]{0,1}[0-9]{0,1}[\s-]{0,1}\({0,1}([0-9]{3})\){0,1}[\s-]{0,1}([0-9]{3})[\s-]{0,1}([0-9]{4})$"
End With
IS_PHONE = Regex.Test(string1)
End Function
Public Function IS_CREDIT_CARD(ByVal string1$) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
Dim regexPattern$
regexPattern = regexPattern & "(3[47][0-9]{13})|"
regexPattern = regexPattern & "(3(0[0-5]|[68][0-9])?[0-9]{11})|"
regexPattern = regexPattern & "(6(011|5[0-9]{2})[0-9]{12})|"
regexPattern = regexPattern & "((2131|1800|35[0-9]{3})[0-9]{11})"
regexPattern = regexPattern & "(5[1-5][0-9]{14})|"
regexPattern = regexPattern & "(4[0-9]{12}([0-9]{3})?)|"
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = regexPattern
End With
IS_CREDIT_CARD = Regex.Test(string1)
End Function
Public Function IS_URL(ByVal string1$) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = "http(s){0,1}://www.[a-zA-Z0-9_.]*?[.][a-zA-Z]{2,15}"
End With
IS_URL = Regex.Test(string1)
End Function
Public Function IS_IP_FOUR(ByVal string1$) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = "^((2[0-4]\d|25[0-5]|1\d\d|\d{1,2})[.]){3}(2[0-4]\d|25[0-5]|1\d\d|\d{1,2})$"
End With
IS_IP_FOUR = Regex.Test(string1)
End Function
Public Function CREDIT_CARD_NAME(ByVal string1$) As String
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
Regex.Global = True
Regex.IgnoreCase = True
Regex.MultiLine = True
Regex.Pattern = "(3[47][0-9]{13})"
If Regex.Test(string1) Then
CREDIT_CARD_NAME = "Amex"
Exit Function
End If
Regex.Pattern = "(3(0[0-5]|[68][0-9])?[0-9]{11})"
If Regex.Test(string1) Then
CREDIT_CARD_NAME = "Diners"
Exit Function
End If
Regex.Pattern = "(6(011|5[0-9]{2})[0-9]{12})"
If Regex.Test(string1) Then
CREDIT_CARD_NAME = "Discover"
Exit Function
End If
Regex.Pattern = "((2131|1800|35[0-9]{3})[0-9]{11})"
If Regex.Test(string1) Then
CREDIT_CARD_NAME = "JCB"
Exit Function
End If
Regex.Pattern = "(5[1-5][0-9]{14})"
If Regex.Test(string1) Then
CREDIT_CARD_NAME = "MasterCard"
Exit Function
End If
Regex.Pattern = "(4[0-9]{12}([0-9]{3})?)"
If Regex.Test(string1) Then
CREDIT_CARD_NAME = "Visa"
Exit Function
End If
CREDIT_CARD_NAME = "#NotAValidCreditCardNumber!"
End Function
Public Function FORMAT_FRACTION(ByVal decimal1#) As String
FORMAT_FRACTION = Trim(WorksheetFunction.Text(decimal1, "# ?/?"))
End Function
Public Function FORMAT_PHONE(ByVal string1$) As String
If IS_PHONE(string1) Then
string1 = Trim(string1)
string1 = Replace(string1, "+", "")
string1 = Replace(string1, "-", "")
string1 = Replace(string1, "(", "")
string1 = Replace(string1, ")", "")
string1 = Replace(string1, " ", "")
FORMAT_PHONE = WorksheetFunction.Text(CLng(string1), "[<=9999999]###-####;(###) ###-####")
Else
FORMAT_PHONE = "#NotAValidPhoneNumber!"
End If
End Function
Public Function FORMAT_CREDIT_CARD(ByVal string1$) As String
If IS_CREDIT_CARD(string1) Then
FORMAT_CREDIT_CARD = Left(string1, 4) & "-" & Mid(string1, 5, 4) & "-" & Mid(string1, 9, 4) & "-" & Mid(string1, 13)
Else
FORMAT_CREDIT_CARD = "#NotAValidCreditCardNumber!"
End If
End Function
