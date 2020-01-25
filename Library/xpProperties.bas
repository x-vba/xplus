Attribute VB_Name = "xpProperties"
'@Module: This module contains a set of functions for getting properties from Ranges, Worksheets, and Workbooks.

Option Explicit


Public Function RANGE_COMMENT( _
    ByVal range1 As Range, _
    Optional ByVal excludeUsername As Boolean) _
As String

    '@Description: This function gets the comment of the selected cell. It also includes an optional parameter that if set to TRUE will remove the Username from the comment.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the cell we want to get the comment from
    '@Param: excludeUsername if set to TRUE, will remove the Username from the comment
    '@Returns: Returns a string of the comment from the cell
    '@Example: =RANGE_COMMENT(A1) -> "Anthony: This is my comment"; Where the cell contains a comment
    '@Example: =RANGE_COMMENT(A1, TRUE) -> "This is my comment"

    Application.Volatile
    
    If excludeUsername Then
        RANGE_COMMENT = Mid(range1.Comment.Text, InStr(range1.Comment.Text, ":") + 1)
    Else
        RANGE_COMMENT = range1.Comment.Text
    End If

End Function


Public Function RANGE_HYPERLINK( _
    ByVal range1 As Range) _
As String

    '@Description: This function gets the hyperlink of the selected cell.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the cell we want to get the hyperlink from
    '@Returns: Returns a string of the hyperlink from the cell
    '@Example: =RANGE_HYPERLINK(A1) -> "https://www.microsoft.com"; Where the cell has a link to https://www.microsoft.com

    Application.Volatile

    RANGE_HYPERLINK = range1.Hyperlinks(1).Name

End Function


Public Function RANGE_NUMBER_FORMAT( _
    ByVal range1 As Range) _
As String

    '@Description: This function gets the number format of the selected cell.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the cell we want to get the number format from
    '@Returns: Returns a string of the number format from the cell
    '@Example: =RANGE_NUMBER_FORMAT(A1) -> "General"; Where the cell has the default number format
    '@Example: =RANGE_NUMBER_FORMAT(A2) -> "Accounting"; Where the cell uses the Accounting number format

    Application.Volatile

    RANGE_NUMBER_FORMAT = range1.NumberFormat

End Function


Public Function RANGE_FONT( _
    ByVal range1 As Range) _
As String

    '@Description: This function gets the name of the font of the selected cell.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the cell we want to get the font name from
    '@Returns: Returns a string of the number format from the cell
    '@Example: =RANGE_FONT(A1) -> "Calibri"; Where the cell has the font style Calibri
    '@Example: =RANGE_FONT(A2) -> "Arial"; Where the cell has the font style Arial

    Application.Volatile

    RANGE_FONT = range1.Font.Name

End Function


Public Function RANGE_NAME( _
    ByVal range1 As Range) _
As String

    '@Description: This function gets the name of the selected cell when using the Name Manager to create Named Ranges.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the cell we want to get the name from
    '@Returns: Returns a string of the name of the cell
    '@Example: =RANGE_NAME(A1) -> "Hello_World"; Where the name of the cell has been named to "Hello_World"

    Application.Volatile

    RANGE_NAME = range1.Name.Name

End Function


Public Function RANGE_WIDTH( _
    ByVal range1 As Range) _
As Double

    '@Description: This function gets the width of the selected cell
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the cell we want to get the width from
    '@Returns: Returns the width of the cell as a Double
    '@Example: =RANGE_WIDTH(A1) -> 20
    
    Application.Volatile

    RANGE_WIDTH = range1.Width

End Function


Public Function RANGE_HEIGHT( _
    ByVal range1 As Range) _
As Double

    '@Description: This function gets the height of the selected cell
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the cell we want to get the height from
    '@Returns: Returns the height of the cell as a Double
    '@Example: =GET_RANGE_WIDTH(A1) -> 14

    Application.Volatile

    RANGE_HEIGHT = range1.Height

End Function


Public Function SHEET_NAME( _
    Optional ByVal sheetNameOrNumber As Variant) _
As String

    '@Description: This function returns the name of the sheet the function resides in, or if a number/name is provided, returns the name of the sheet that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: sheetNameOrNumber is the name or number of the sheet
    '@Returns: Returns the name of the sheet
    '@Example: =SHEET_NAME() -> "Sheet1"; Where this function resides in Sheet1
    '@Example: =SHEET_NAME("Sheet2") -> "Sheet2"
    '@Example: =SHEET_NAME(2) -> "Sheet2"

    Application.Volatile

    If IsMissing(sheetNameOrNumber) Then
        SHEET_NAME = Application.Caller.Parent.Name
    Else
        SHEET_NAME = Sheets(sheetNameOrNumber).Name
    End If

End Function


Public Function SHEET_CODE_NAME( _
    Optional ByVal sheetNameOrNumber As Variant) _
As String

    '@Description: This function returns the code name of the sheet the function resides in, or if a number/name is provided, returns the code name of the sheet that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: sheetNameOrNumber is the name or number of the sheet
    '@Returns: Returns the code name of the sheet
    '@Example: =SHEET_CODE_NAME() -> "Sheet1"; Where this function resides in Sheet1
    '@Example: =SHEET_CODE_NAME("MySheet") -> "Sheet1"
    '@Example: =SHEET_CODE_NAME(1) -> "Sheet1"

    Application.Volatile

    If IsMissing(sheetNameOrNumber) Then
        SHEET_CODE_NAME = Application.Caller.Parent.CodeName
    Else
        SHEET_CODE_NAME = Sheets(sheetNameOrNumber).CodeName
    End If

End Function


Public Function SHEET_TYPE( _
    Optional ByVal sheetNameOrNumber As Variant) _
As String

    '@Description: This function returns the type of the sheet the function resides in, or if a number/name is provided, returns the type of the sheet that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: sheetNameOrNumber is the name or number of the sheet
    '@Returns: Returns the code name of the sheet
    '@Example: =SHEET_CODE_NAME() -> "Worksheet"
    '@Example: =SHEET_CODE_NAME("MyChart") -> "Chart"
    '@Example: =SHEET_CODE_NAME(2) -> "Chart"

    Application.Volatile

    Dim sheetTypeInteger As Integer

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


Public Function WORKBOOK_TITLE( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the title of the workbook that the function resides in, or if a number/name is provided, returns the title of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the title of the workbook
    '@Note: Workbook title can be set in File->Info->Properties
    '@Example: =WORKBOOK_TITLE() -> "MyWorkbook"
    '@Example: =WORKBOOK_TITLE("Otherbook.xlsx") -> "MyOtherWorksheet"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_TITLE = ThisWorkbook.BuiltinDocumentProperties("Title")
    Else
        WORKBOOK_TITLE = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Title")
    End If

End Function


Public Function WORKBOOK_SUBJECT( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the subject of the workbook that the function resides in, or if a number/name is provided, returns the subject of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the subject of the workbook
    '@Note: Workbook subject can be set in File->Info->Properties
    '@Example: =WORKBOOK_SUBJECT() -> "MySubject"
    '@Example: =WORKBOOK_SUBJECT("Otherbook.xlsx") -> "MyOtherSubject"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_SUBJECT = ThisWorkbook.BuiltinDocumentProperties("Subject")
    Else
        WORKBOOK_SUBJECT = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Subject")
    End If

End Function


Public Function WORKBOOK_AUTHOR( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the author of the workbook that the function resides in, or if a number/name is provided, returns the author of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the author of the workbook
    '@Note: Workbook author can be set in File->Info->Properties
    '@Example: =WORKBOOK_AUTHOR() -> "John Doe"
    '@Example: =WORKBOOK_AUTHOR("Otherbook.xlsx") -> "Jane Doe"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_AUTHOR = ThisWorkbook.BuiltinDocumentProperties("Author")
    Else
        WORKBOOK_AUTHOR = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Author")
    End If

End Function


Public Function WORKBOOK_MANAGER( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the manager of the workbook that the function resides in, or if a number/name is provided, returns the manager of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the manager of the workbook
    '@Note: Workbook manager can be set in File->Info->Properties
    '@Example: =WORKBOOK_MANAGER() -> "Manager John"
    '@Example: =WORKBOOK_MANAGER("Otherbook.xlsx") -> "Manager Jane"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_MANAGER = ThisWorkbook.BuiltinDocumentProperties("Manager")
    Else
        WORKBOOK_MANAGER = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Manager")
    End If

End Function


Public Function WORKBOOK_COMPANY( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the company of the workbook that the function resides in, or if a number/name is provided, returns the company of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the company of the workbook
    '@Note: Workbook company can be set in File->Info->Properties
    '@Example: =WORKBOOK_COMPANY() -> "Hello Company"
    '@Example: =WORKBOOK_COMPANY("Otherbook.xlsx") -> "World Company"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_COMPANY = ThisWorkbook.BuiltinDocumentProperties("Company")
    Else
        WORKBOOK_COMPANY = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Company")
    End If

End Function


Public Function WORKBOOK_CATEGORY( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the category of the workbook that the function resides in, or if a number/name is provided, returns the category of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the category of the workbook
    '@Note: Workbook category can be set in File->Info->Properties
    '@Example: =WORKBOOK_CATEGORY() -> "Category1"
    '@Example: =WORKBOOK_CATEGORY("Otherbook.xlsx") -> "Category2"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_CATEGORY = ThisWorkbook.BuiltinDocumentProperties("Category")
    Else
        WORKBOOK_CATEGORY = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Category")
    End If

End Function


Public Function WORKBOOK_KEYWORDS( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the keywords of the workbook that the function resides in, or if a number/name is provided, returns the keywords of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the keywords of the workbook
    '@Note: Workbook keywords can be set in File->Info->Properties
    '@Example: =WORKBOOK_KEYWORDS() -> "accounting, jan, hello"
    '@Example: =WORKBOOK_KEYWORDS("Otherbook.xlsx") -> "finance, feb, world"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_KEYWORDS = ThisWorkbook.BuiltinDocumentProperties("Keywords")
    Else
        WORKBOOK_KEYWORDS = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Keywords")
    End If

End Function


Public Function WORKBOOK_COMMENTS( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the comments of the workbook that the function resides in, or if a number/name is provided, returns the comments of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the comments of the workbook
    '@Note: Workbook comments can be set in File->Info->Properties
    '@Example: =WORKBOOK_COMMENTS() -> "This is my workbook"
    '@Example: =WORKBOOK_COMMENTS("Otherbook.xlsx") -> "This is my other workbook"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_COMMENTS = ThisWorkbook.BuiltinDocumentProperties("Comments")
    Else
        WORKBOOK_COMMENTS = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Comments")
    End If

End Function


Public Function WORKBOOK_HYPERLINK_BASE( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the hyperlink base of the workbook that the function resides in, or if a number/name is provided, returns the hyperlink base of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the hyperlink base of the workbook
    '@Note: Workbook hyperlink base can be set in File->Info->Properties
    '@Example: =WORKBOOK_HYPERLINK_BASE() -> "http://myhyperlinkbase-example.com"
    '@Example: =WORKBOOK_HYPERLINK_BASE("Otherbook.xlsx") -> "http://myotherhyperlinkbase-example.com"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_HYPERLINK_BASE = ThisWorkbook.BuiltinDocumentProperties("Hyperlink Base")
    Else
        WORKBOOK_HYPERLINK_BASE = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Hyperlink Base")
    End If

End Function


Public Function WORKBOOK_REVISION_NUMBER( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the revision number of the workbook that the function resides in, or if a number/name is provided, returns the revision number of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the revision number of the workbook
    '@Example: =WORKBOOK_REVISION_NUMBER() -> 1
    '@Example: =WORKBOOK_REVISION_NUMBER("Otherbook.xlsx") -> 2

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_REVISION_NUMBER = ThisWorkbook.BuiltinDocumentProperties("Revision Number")
    Else
        WORKBOOK_REVISION_NUMBER = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Revision Number")
    End If

End Function


Public Function WORKBOOK_CREATION_DATE( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the creation date of the workbook that the function resides in, or if a number/name is provided, returns the creation date of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the creation date of the workbook
    '@Example: =WORKBOOK_CREATION_DATE() -> "1/1/2020 10:00:00 PM"
    '@Example: =WORKBOOK_CREATION_DATE("Otherbook.xlsx") -> "1/5/2020 8:00:00 PM"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_CREATION_DATE = ThisWorkbook.BuiltinDocumentProperties("Creation Date")
    Else
        WORKBOOK_CREATION_DATE = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Creation Date")
    End If

End Function


Public Function WORKBOOK_LAST_SAVE_TIME( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the last save time of the workbook that the function resides in, or if a number/name is provided, returns the last save time of the workbook that resides at that number/name
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the last save time of the workbook
    '@Example: =WORKBOOK_LAST_SAVE_TIME() -> "1/3/2020 10:00:00 PM"
    '@Example: =WORKBOOK_LAST_SAVE_TIME("Otherbook.xlsx") -> "1/10/2020 8:00:00 PM"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_LAST_SAVE_TIME = ThisWorkbook.BuiltinDocumentProperties("Last Save Time")
    Else
        WORKBOOK_LAST_SAVE_TIME = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Last Save Time")
    End If

End Function


Public Function WORKBOOK_LAST_AUTHOR( _
    Optional ByVal workbookNameOrNumber As Variant) _
As String

    '@Description: This function returns the last author of the workbook that the function resides in, or if a number/name is provided, returns the comments of the workbook that resides at that number/name. Last author is the person who last saved the workbook.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: workbookNameOrNumber is the name or number of the workbook
    '@Returns: Returns the last author of the workbook
    '@Example: =WORKBOOK_LAST_AUTHOR() -> "John Doe"
    '@Example: =WORKBOOK_LAST_AUTHOR("Otherbook.xlsx") -> "Jane Doe"

    Application.Volatile

    If IsMissing(workbookNameOrNumber) Then
        WORKBOOK_LAST_AUTHOR = ThisWorkbook.BuiltinDocumentProperties("Last Author")
    Else
        WORKBOOK_LAST_AUTHOR = Workbooks(workbookNameOrNumber).BuiltinDocumentProperties("Last Author")
    End If

End Function
