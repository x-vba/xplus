Attribute VB_Name = "xpUtilities"
'@Module: This module contains a set of basic miscellaneous utility functions

Option Explicit

Public Function DISPLAY_TEXT( _
    ParamArray textArray() As Variant) _
As String

    '@Description: This function takes the range of the cell that this function resides, and then an array of text, and when this function is recalculated manually by the user (for example when pressing the F2 key while on the cell) a textbox will display all the text in the cell, making it easier to read and manage large strings of text in a single cell.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: textArray() is an array of ranges, strings, or number that will be displayed
    '@Returns: Returns all the strings in the text array combined as well as displays all the text in the text array
    '@Example: =DISPLAY_TEXT("hello", "world") -> "hello world" and displays the text in a textbox
    '@Example: =DISPLAY_TEXT(A1:A2) -> "hello world" and displays the text in a textbox, where A1="hello" and A2="world"
    '@Example: =DISPLAY_TEXT(B1:B2, "Three") -> "One Two Three" and displays the text in a textbox, where B1="One" and B2="Two"

    Dim combinedString As String
    Dim individualTextItem As Variant
    Dim individualRange As Range
    
    
    For Each individualTextItem In textArray
    
        ' If range use .Value call
        If TypeName(individualTextItem) = "Range" Then
            For Each individualRange In individualTextItem
                combinedString = combinedString & individualRange.Value & vbCrLf & vbCrLf
            Next
            
        ' Else just get the value directly
        Else
            combinedString = combinedString & individualTextItem & vbCrLf & vbCrLf
        End If
    Next

    
    ' If the function is called within the active cell of the same workbook and same sheet
    If Application.Caller.Parent.Parent.Name = ActiveWorkbook.Name Then
        If Application.Caller.Worksheet.Name = ActiveCell.Worksheet.Name Then
            If Application.Caller.Address = ActiveCell.Address Then
                MsgBox combinedString, , "Cell " & Replace(Application.Caller.Address, "$", "") & " Contents"
            End If
        End If
    End If

    DISPLAY_TEXT = combinedString

End Function


Public Function JSONIFY( _
    ByVal indentLevel As Byte, _
    ParamArray stringArray() As Variant) _
As String

    '@Description: This function takes an array of strings and numbers and returns the array as a JSON string. This function takes into account formatting for numbers, and supports specifying the indentation level.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: indentLevel is an optional number that specifying the indentation level. Leaving this argument out will result in no indentation
    '@Param: stringArray() is an array of strings and number in the following format: {"Hello", "World"}
    '@Returns: Returns a JSON valid string of all elements in the array
    '@Example: =JSONIFY(0, "Hello", "World", "1", "2", 3, 4.5) -> "["Hello","World",1,2,3,4.5]"
    '@Example: =JSONIFY(0, A1:A6) -> "["Hello","World",1,2,3,4.5]"

    Dim i As Byte
    Dim jsonString As String
    Dim individualTextItem As Variant
    Dim individualRange As Range
    Dim indentString As String
    
    ' Setting up some base JSON features and the indenting
    jsonString = "["
    
    For i = 1 To indentLevel
        indentString = indentString & " "
    Next
    
    If indentLevel > 0 Then
        jsonString = jsonString & Chr(10)
    End If
    
    
    ' Creating the contents of the JSON string
    For Each individualTextItem In stringArray
    
        ' In cases of ranges
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
            
        ' In cases of text
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

    '@Description: This function generates a unique ID based on the UUID V4 specification. This function is useful for generating unique IDs of a fixed character length.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string unique ID based on UUID V4. The format of the string will always be in the form of "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx" where each x is a hex digit, and y is either 8, 9, A, or B.
    '@Example: =UUID_FOUR() -> "3B4BDC26-E76E-4D6C-9E05-7AE7D2D68304"
    '@Example: =UUID_FOUR() -> "D5761256-8385-4FDA-AD56-6AEF0AD6B9A5"
    '@Example: =UUID_FOUR() -> "CDCAE2F5-B52F-4C90-A38A-42BD58BCED4B"

    Dim firstGroup As String
    Dim secondGroup As String
    Dim thirdGroup As String
    Dim fourthGroup As String
    Dim fifthGroup As String
    Dim sixthGroup As String

    firstGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(0, 4294967295#), 8) & "-"
    secondGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(0, 65535), 4) & "-"
    thirdGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(16384, 20479), 4) & "-"
    fourthGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(32768, 49151), 4) & "-"
    fifthGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(0, 65535), 4)
    sixthGroup = WorksheetFunction.Dec2Hex(WorksheetFunction.RandBetween(0, 4294967295#), 8)

    UUID_FOUR = firstGroup & secondGroup & thirdGroup & fourthGroup & fifthGroup & sixthGroup

End Function


Public Function HIDDEN( _
    ByVal string1 As String, _
    ByVal hiddenFlag As Boolean, _
    Optional ByVal hideString As String) _
As String

    '@Description: This function takes the value in a cell and visibly hides it if the hidden flag set to TRUE. If TRUE, the value will appear as "********", with the option to set the hidden characters to a different set of text.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be hidden
    '@Param: hiddenFlag if set to TRUE will hide string1
    '@Param: hideString is an optional string that if set will be used instead of "********"
    '@Returns: Returns a string to hide string1 if hideFlag is TRUE
    '@Example: =HIDDEN("Hello World", FALSE) -> "Hello World"
    '@Example: =HIDDEN("Hello World", TRUE) -> "********"
    '@Example: =HIDDEN("Hello World", TRUE, "[Hidden Text]") -> "[Hidden Text]"
    '@Example: =HIDDEN("Hello World", USER_NAME()="Anthony") -> "********"

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


Public Function ISERRORALL( _
    ByVal range1 As Range) _
As Boolean

    '@Description: This function is an extension of Excel's =ISERROR(). It returns TRUE for all of Excel's built in errors, similar to =ISERROR() but also returns TRUE for User-Defined Error Strings. User-Defined Error Strings are strings that start with character "#" and end with either the character "!" or "?". This is similar to the format of errors in Excel, such as "#DIV/0!", "#VALUE!", "#NAME?", "#REF!", etc. User-Defined Error Strings are used all throughout XPlus, so this is a useful function for checking errors in XPlus functions. Additionally, users can create their own User-Defined Error Strings in Excel and use this function to check for those errors.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the range that will be checked for an error
    '@Returns: Returns TRUE if the range contains an Excel error or a User-Defined Error String
    '@Example: =ISERRORALL("Not an Error") -> FALSE
    '@Example: =ISERRORALL(1/0) -> TRUE
    '@Example: =ISERRORALL("#UserDefinedErrorString!") -> TRUE
    '@Example: =ISERRORALL("#UserDefinedErrorString?") -> TRUE
    '@Example: =ISERRORALL("UserDefinedErrorString") -> FALSE; The format for the User-Defined Error String is incorrect since it is missing the character "#" at the beginning, and either "!" or "?" at the end

    Dim rangeValue As Variant
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


Public Function COUNTERRORALL( _
    ParamArray rangeArray() As Variant) _
As Integer

    '@Description: This function takes a range or multiple ranges, and returns a count of all Errors and User-Defined Error Strings within those ranges. User-Defined Error Strings are strings that start with character "#" and end with either the character "!" or "?". This is similar to the format of errors in Excel, such as "#DIV/0!", "#VALUE!", "#NAME?", "#REF!", etc. User-Defined Error Strings are used all throughout XPlus, so this is a useful function for checking errors in XPlus functions. Additionally, users can create their own User-Defined Error Strings in Excel and use this function to check for those errors.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Potentially update this function so that it accepts strings as well as ranges
    '@Param: rangeArray is the range or multiple ranges whose errors will be counted
    '@Returns: Returns the number of errors counted
    '@Example: =COUNTERRORALL(A1:A6) -> 4; Where A1="Hello World", A2="#DIV/0!", A3="#ErrorMessage!", A4="#ErrorMessage?", A5="#NAME?", A6="12345678"

    Dim errorCount As Integer
    Dim individualRange As Variant
    Dim individualCell As Range
    
    For Each individualRange In rangeArray
        For Each individualCell In individualRange
            If WorksheetFunction.IsError(individualCell.Value) Then
                errorCount = errorCount + 1
            ElseIf Left(individualCell.Value, 1) = "#" Then
                If Right(individualCell.Value, 1) = "!" Or Right(individualCell.Value, 1) = "?" Then
                    errorCount = errorCount + 1
                End If
            End If
        Next
    Next
    
    COUNTERRORALL = errorCount

End Function


Public Function JAVASCRIPT( _
    ByVal jsFuncCode As String, _
    ByVal jsFuncName As String, _
    Optional ByVal argument1 As Variant, _
    Optional ByVal argument2 As Variant, _
    Optional ByVal argument3 As Variant, _
    Optional ByVal argument4 As Variant, _
    Optional ByVal argument5 As Variant, _
    Optional ByVal argument6 As Variant, _
    Optional ByVal argument7 As Variant, _
    Optional ByVal argument8 As Variant, _
    Optional ByVal argument9 As Variant, _
    Optional ByVal argument10 As Variant, _
    Optional ByVal argument11 As Variant, _
    Optional ByVal argument12 As Variant, _
    Optional ByVal argument13 As Variant, _
    Optional ByVal argument14 As Variant, _
    Optional ByVal argument15 As Variant, _
    Optional ByVal argument16 As Variant) _
As Variant

    '@Description: This function executes JavaScript code using Microsoft's JScript scripting language. It takes 3 arguments, the JavaScript code that will be executed, the name of the JavaScript function that will be executed, and up to 16 optional arguments to be used in the JavaScript function that is called. One thing to note is that ES5 syntax should be used in the JavaScript code, as ES6 features are unlikely to be supported.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: jsFuncCode is a string of the JavaScript source code that will be executed
    '@Param: jsFuncName is the name of the JavaScript function that will be called
    '@Param: argument1 - argument16 are optional arguments used in the JScript function call
    '@Returns: Returns the result of the JavaScript function that is called
    '@Example: =JAVASCRIPT("function helloFunc(){return 'Hello World!'}", "helloFunc") -> "Hello World!"
    '@Example: =JAVASCRIPT("function addTwo(a, b){return a + b}","addTwo",12,24) -> 36

    Dim ScriptContoller As Object
    Set ScriptContoller = CreateObject("ScriptControl")
    
    ScriptContoller.Language = "JScript"
    ScriptContoller.addCode jsFuncCode

    JAVASCRIPT = ScriptContoller.Run(jsFuncName, _
        argument1, argument2, argument3, argument4, _
        argument5, argument6, argument7, argument8, _
        argument9, argument10, argument11, argument12, _
        argument13, argument14, argument15, argument16)

End Function


Public Function SUMSHEET( _
    ByVal partialSheetName As String, _
    Optional ByVal range1 As Range) _
As Variant

    '@Description: This function sums up the value of the same cell in multiple sheets based on a partial sheet name.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: partialSheetName is a string with the partial name of a sheet. For example, if you set this argument to the string "Dat" all sheets with the string "Dat" in their name will be the sheets that are summed up
    '@Param: range1 is an optional paramter to set the cell that will be summed. By default, the cell this function resides will be the one that is summed up in the other sheets, but if range1 is set, that is the range that will be summed up.
    '@Returns: Returns the sum of all cells that pass the partial sheet name criteria
    '@Example: =SUMSHEET("- Data") -> 20; Where this function resides in cell C2 and the workbook contains the sheets "Jan - Data", "Feb - Data", "Mar - Data", "HelloWorld", "SumSheet", cell C2 in sheets "Jan - Data" (which contains value 5), "Feb - Data" (which contains value 7), "Mar - Data" (which contains value 8) will be summed up
    '@Example: =SUMSHEET("- Data", A1) -> 6; Same as the above example except cell A1 will be used instead of C2 and where A1 contains 1, 2, and 3 for values in the other sheets

    Dim sumValue As Variant
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


Public Function AVERAGESHEET( _
    ByVal partialSheetName As String, _
    Optional ByVal range1 As Range) _
As Variant

    '@Description: This function averages the value of the same cell in multiple sheets based on a partial sheet name.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: partialSheetName is a string with the partial name of a sheet. For example, if you set this argument to the string "Dat" all sheets with the string "Dat" in their name will be the sheets that are averaged
    '@Param: range1 is an optional paramter to set the cell that will be averaged. By default, the cell this function resides will be the one that is averaged in the other sheets, but if range1 is set, that is the range that will be averaged.
    '@Returns: Returns the average of all cells that pass the partial sheet name criteria
    '@Example: =AVERAGESHEET("- Data") -> 6.67; Where this function resides in cell C2 and the workbook contains the sheets "Jan - Data", "Feb - Data", "Mar - Data", "HelloWorld", "SumSheet", cell C2 in sheets "Jan - Data" (which contains value 5), "Feb - Data" (which contains value 7), "Mar - Data" (which contains value 8) will be averaged
    '@Example: =AVERAGESHEET("- Data", A1) -> 2; Same as the above example except cell A1 will be used instead of C2 and where A1 contains 1, 2, and 3 for values in the other sheets

    Dim sumValue As Variant
    Dim countValue As Integer
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


Public Function MAXSHEET( _
    ByVal partialSheetName As String, _
    Optional ByVal range1 As Range) _
As Variant

    '@Description: This function gets the max value of the same cell in multiple sheets based on a partial sheet name.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: partialSheetName is a string with the partial name of a sheet. For example, if you set this argument to the string "Dat" all sheets with the string "Dat" in their name will be the sheets that the max value is picked from
    '@Param: range1 is an optional paramter to set the cell that will be maxed. By default, the cell this function resides will be the one that is maxed in the other sheets, but if range1 is set, that is the range that will be maxed.
    '@Returns: Returns the max of all cells that pass the partial sheet name criteria
    '@Example: =MAXSHEET("- Data") -> 8; Where this function resides in cell C2 and the workbook contains the sheets "Jan - Data", "Feb - Data", "Mar - Data", "HelloWorld", "SumSheet", cell C2 in sheets "Jan - Data" (which contains value 5), "Feb - Data" (which contains value 7), "Mar - Data" (which contains value 8) will be maxed
    '@Example: =MAXSHEET("- Data", A1) -> 3; Same as the above example except cell A1 will be used instead of C2 and where A1 contains 1, 2, and 3 for values in the other sheets

    Dim maxValue As Variant
    Dim currentValue As Variant
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


Public Function MINSHEET( _
    ByVal partialSheetName As String, _
    Optional ByVal range1 As Range) _
As Variant

    '@Description: This function gets the min value of the same cell in multiple sheets based on a partial sheet name.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: partialSheetName is a string with the partial name of a sheet. For example, if you set this argument to the string "Dat" all sheets with the string "Dat" in their name will be the sheets that the min value is picked from
    '@Param: range1 is an optional paramter to set the cell that will be mined. By default, the cell this function resides will be the one that is mined in the other sheets, but if range1 is set, that is the range that will be mined.
    '@Returns: Returns the min of all cells that pass the partial sheet name criteria
    '@Example: =MINSHEET("- Data") -> 5; Where this function resides in cell C2 and the workbook contains the sheets "Jan - Data", "Feb - Data", "Mar - Data", "HelloWorld", "SumSheet", cell C2 in sheets "Jan - Data" (which contains value 5), "Feb - Data" (which contains value 7), "Mar - Data" (which contains value 8) will be mined
    '@Example: =MINSHEET("- Data", A1) -> 1; Same as the above example except cell A1 will be used instead of C2 and where A1 contains 1, 2, and 3 for values in the other sheets

    Dim minValue As Variant
    Dim currentValue As Variant
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


Public Function HTML_TABLEIFY( _
    ByVal rangeTable As Range) _
As String

    '@Description: This function takes a range in a table format and generates an HTML table from it. It uses the first row in the range chosen as the headers, and all other data as row data.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Add a Boolean parameter that adds hooks and some styling to the table
    '@Param: rangeTable is a range that will be formatted as an HTML table string.
    '@Returns: Returns an HTML table string with data from the range populated in it
    '@Example: =HTML_TABLEIFY(A1:C5) -> <table>...</table>

    Dim i As Integer
    Dim htmlTableString As String
    Dim individualRange As Range
    
    htmlTableString = htmlTableString & "<table>" & vbCrLf
    
    
    ' Generating the Table Head
    htmlTableString = htmlTableString & "  <thead>" & vbCrLf
    htmlTableString = htmlTableString & "    <tr>" & vbCrLf
    
    For Each individualRange In rangeTable.Rows(1).Cells
        htmlTableString = htmlTableString & "      <th>" & individualRange.Value & "</th>" & vbCrLf
    Next
    
    htmlTableString = htmlTableString & "    </tr>" & vbCrLf
    htmlTableString = htmlTableString & "  </thead>" & vbCrLf
    
    
    ' Generating the Table Body
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


Public Function HTML_ESCAPE( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and escapes the HTML characters in it. For example, the character ">" will be escaped into "%gt;"
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will have its characters HTML escaped
    '@Returns: Returns an HTML escaped string
    '@Example: =HTML_ESCAPE("<p>Hello World</p>") -> "&lt;p&gt;Hello World&lt;/p&gt;"

    string1 = Replace(string1, "&", "&amp;")
    string1 = Replace(string1, Chr(34), "&quot;")
    string1 = Replace(string1, "'", "&apos;")
    string1 = Replace(string1, "<", "&lt;")
    string1 = Replace(string1, ">", "&gt;")
    
    HTML_ESCAPE = string1

End Function


Public Function HTML_UNESCAPE( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and unescapes the HTML characters in it. For example, the character "%gt;" will be escaped into ">"
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will have its characters HTML unescaped
    '@Returns: Returns an HTML unescaped string
    '@Example: =HTML_UNESCAPE("&lt;p&gt;Hello World&lt;/p&gt;") -> "<p>Hello World</p>"

    string1 = Replace(string1, "&amp;", "&")
    string1 = Replace(string1, "&quot;", Chr(34))
    string1 = Replace(string1, "&apos;", "'")
    string1 = Replace(string1, "&lt;", "<")
    string1 = Replace(string1, "&gt;", ">")

    HTML_UNESCAPE = string1

End Function


Public Function SPEAK_TEXT( _
    ParamArray textArray() As Variant) _
As String

    '@Description: This function takes the range of the cell that this function resides, and then an array of text, and when this function is recalculated manually by the user (for example when pressing the F2 key while on the cell) this function will use Microsoft's text-to-speech to speak out the text through the speakers or microphone.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: textArray() is an array of ranges, strings, or number that will be displayed
    '@Returns: Returns all the strings in the text array combined as well as displays all the text in the text array
    '@Example: =SPEAK_TEXT("Hello", "World") -> "Wello World" and the text will be spoken through the speaker
    '@Example: =SPEAK_TEXT(A1:A2) -> "Hello World" and the text will be spoken through the speaker, where A1="Hello" and A2="World"
    '@Example: =SPEAK_TEXT(B1:B2, "Three") -> "One Two Three" and the text will be spoken through the speaker, where B1="One" and B2="Two"

    Dim combinedString As String
    Dim individualTextItem As Variant
    Dim individualRange As Range
    
    For Each individualTextItem In textArray
    
        ' If range use .Value call
        If TypeName(individualTextItem) = "Range" Then
            For Each individualRange In individualTextItem
                combinedString = combinedString & individualRange.Value & " "
            Next
            
        ' Else just get the value directly
        Else
            combinedString = combinedString & individualTextItem & " "
        End If
    Next
    
    ' If the function is called within the active cell of the same workbook and same sheet
    If Application.Caller.Parent.Parent.Name = ActiveWorkbook.Name Then
        If Application.Caller.Worksheet.Name = ActiveCell.Worksheet.Name Then
            If Application.Caller.Address = ActiveCell.Address Then
                Application.Speech.SPEAK combinedString, True
            End If
        End If
    End If

    SPEAK_TEXT = combinedString

End Function


Public Function EVALUATE_FORMULA( _
    ByVal formulaText As String, _
    ParamArray rangeArray() As Variant) _
As Variant

    '@Description: This function takes a formula as a string with placeholders in it, and executes and returns the value of that formula with the values from the placeholder used as inputs. Placeholders are in the form of {1}, {2}, {3}, etc., with the first placeholder starting at 1
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: formulaText is the formula with placeholders as text
    '@Param: rangeArray() is any number of ranges to use as inputs and that will be replaced with the placeholders
    '@Returns: Returns the executed value from the formula
    '@Note: This function only support ranges in the rangeArray, as the values in the placeholders are replaced with the addresses of the ranges used as inputs
    '@Warning: Any evaluation function in any programming language can result in security vulnerabilites if misused. Particularly, when the inputs for this function are user inputs from other sources, it is possible for malicious inputs and functions to be executed. As a result, use this function with care, and please research more examples of eval function security vulnerabilites and best practices.
    '@Example: =EVALUATE_FORMULA("=SUM({1})+AVERAGE({2})", A1:A3, A4:A6) -> 80; Where A1:A3=[10, 20, 30] and A4:A6=[15, 20, 25]
    '@Example: =EVALUATE_FORMULA("=SUM({1})/COUNT({1})", A1:A3, A4:A6) -> 20; Where A1:A3=[10, 20, 30]; Notice I've used the {1} placeholder twice here

    Dim i As Integer
    Dim individualRange As Variant
    
    For i = 1 To (UBound(rangeArray) - LBound(rangeArray) + 1)
        For Each individualRange In rangeArray
            formulaText = Replace(formulaText, "{" & i & "}", individualRange.Address)
        Next
    Next
    
    EVALUATE_FORMULA = Application.Evaluate(formulaText)

End Function


Public Function ISBADERROR( _
    ParamArray rangeArray() As Variant) _
As Boolean

    '@Description: This function is similar to Excel's Built-in ISERROR() except that it only returns TRUE for #NULL!, #NAME?, #REF!, #DIV/0!, and #NUM! Errors
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: rangeArray() is a range or multiple ranges that may contain errors
    '@Returns: Returns TRUE if there is one of the listed bad errors in the range, or else FALSE
    '@Warning: Excel generates a lot of errors when using common formulas, and some of these are errors the user intends to create, where as some errors are likely to be unintended errors. For example, users typically do not intend to generate a #DIV/0 error or a #REF! error on purpose. This function attempts to only consider the latter errors (errors that may likely be unintentional). As a result, this formula can be interpretted as attempting to signal for errors that are likely unintentional and maybe should be explicitly handled by the users. However, a FALSE value from this formula DOES NOT mean that there are no errors in the spreadsheet or the ranges that this formula is operating on are free from error.
    '@Example: =ISBADERROR(#NAME?) -> TRUE; #NAME? is unlikely to be generated by the users, as it occurs when the user attempts to use a function that doesn't exist
    '@Example: =ISBADERROR(#NUM!) -> TRUE; #NUM! is often generated in Math functions where invalid inputs are used
    '@Example: =ISBADERROR(#DIV/0!) -> TRUE; #DIV/0! is not typically generated by users intentionally
    '@Example: =ISBADERROR(#REF!) -> TRUE; #REF! is often generated when deleting rows that a function points to for an input, and is typically unlikely to be generated by users (except sometimes in the case of using the INDIRECT() function)
    '@Example: =ISBADERROR(#NULL!) -> TRUE; #NULL! is often generated when using incorrect range references in formulas
    '@Example: =ISBADERROR(#N/A) -> FALSE; #N/A may sometimes be intentionally generated by the users
    '@Example: =ISBADERROR(#VALUE!) -> FALSE; #VALUE! may sometimes be intentionally generated by the users
    '@Example: =ISBADERROR(A1:A3) -> TRUE; Where A1=#NAME?, A2=#N/A, A3=#VALUE!
    '@Example: =ISBADERROR(A1, A2, A3) -> TRUE; Where A1=#NAME?, A2=#N/A, A3=#VALUE!

    Dim individualRange As Variant
    Dim individualCell As Range
    
    For Each individualRange In rangeArray
        For Each individualCell In individualRange
            If Not IsEmpty(individualCell) Then
                ' #NULL! Error
                If individualCell.Value = CVErr(2000) Then
                    ISBADERROR = True
                    Exit Function
                End If
            
                ' #NAME? Error
                If individualCell.Value = CVErr(2029) Then
                    ISBADERROR = True
                    Exit Function
                End If
                
                ' #REF! Error
                If individualCell.Value = CVErr(2023) Then
                    ISBADERROR = True
                    Exit Function
                End If
                
                ' #DIV/0! Error
                If individualCell.Value = CVErr(2007) Then
                    ISBADERROR = True
                    Exit Function
                End If
                
                ' #NUM! Error
                If individualCell.Value = CVErr(2036) Then
                    ISBADERROR = True
                    Exit Function
                End If
            End If
        Next
    Next
    
    ISBADERROR = False

End Function


Public Function REFERENCE_EXISTS( _
    ByVal referenceName As String, _
    Optional ByVal partialNameFlag As Boolean) _
As Boolean

    '@Description: This takes a string name of a VBA reference library, and checks if the current workbook currently includes the VBA library as a reference.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: referenceName is the name of a reference we want to check if it exists
    '@Param: partialNameFlag if set to TRUE, will perform checks based on partial names instead of requiring the exact name of the reference
    '@Returns: Returns TRUE if the reference exists, and FALSE if it doesn't
    '@Note: This function is case-insensitive
    '@Example: =REFERENCE_EXISTS("Excel") -> TRUE; The "Excel" Library is typically included as a reference
    '@Example: =REFERENCE_EXISTS("VBA") -> TRUE; The "VBA" Library is typically included as a reference
    '@Example: =REFERENCE_EXISTS("vba") -> TRUE; The "VBA" Library is typically included as a reference, and this function works on case-insensitive checks
    '@Example: =REFERENCE_EXISTS("MSHTML") -> FALSE; The "MSHTML" Library is typically not included as a reference, but can be included by the user
    '@Example: =REFERENCE_EXISTS("VB") -> FALSE; There is typically no library named "VB"
    '@Example: =REFERENCE_EXISTS("VB", TRUE) -> TRUE; Since the partialNameFlag is set to TRUE and since a reference to the "VBA" library exists, this will match the "VBA" library

    Application.Volatile

    Dim individualReference As Variant
    
    For Each individualReference In ThisWorkbook.VBProject.References
        If Not partialNameFlag Then
            If UCase(individualReference.Name) = UCase(referenceName) Then
                REFERENCE_EXISTS = True
                Exit Function
            End If
        Else
            If InStr(1, individualReference.Name, referenceName, vbTextCompare) >= 0 Then
                REFERENCE_EXISTS = True
                Exit Function
            End If
        End If
    Next
    
    REFERENCE_EXISTS = False

End Function


Public Function ADDIN_EXISTS( _
    ByVal addinName As String, _
    Optional ByVal partialNameFlag As Boolean) _
As Boolean

    '@Description: This takes a string name of an Excel Addin, and checks if Excel currently includes the Excel Addin.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: addinName is the name of a Addin we want to check if it exists
    '@Param: partialNameFlag if set to TRUE, will perform checks based on partial names instead of requiring the exact name of the Addin
    '@Returns: Returns TRUE if the Addin exists, and FALSE if it doesn't
    '@Note: This function is case-insensitive. Also, this function checks if an Addin exists, not if the Addin is currently installed. For example, many versions of Excel include the Solver Addin, but by default this Addin is not active in many cases. ADDIN_EXISTS() will return TRUE for the Solver Addin even if it isn't currently installed. For a function that check if an Addin is currently installed, used the ADDIN_INSTALLED() function.
    '@Example: =ADDIN_EXISTS("SOLVER.XLAM") -> TRUE; Most versions of Excel will have the Solver Addin
    '@Example: =ADDIN_EXISTS("solver.xlam") -> TRUE; This function is case-insensitive
    '@Example: =ADDIN_EXISTS("NonExistantAddin.xlam") -> FALSE; As this Addin doesn't currently exist
    '@Example: =ADDIN_EXISTS("SOLVER") -> FALSE; To use partial matches, use the partialNameFlag
    '@Example: =ADDIN_EXISTS("SOLVER", TRUE) -> TRUE; As the partialNameFlag is set and so "SOLVER" will match "SOLVER.XLAM"

    Application.Volatile

    Dim individualAddin As AddIn
    
    For Each individualAddin In Application.AddIns
        If Not partialNameFlag Then
            If UCase(individualAddin.Name) = UCase(addinName) Then
                ADDIN_EXISTS = True
                Exit Function
            End If
        Else
            If InStr(1, individualAddin.Name, addinName, vbTextCompare) >= 0 Then
                ADDIN_EXISTS = True
                Exit Function
            End If
        End If
    Next
    
    ADDIN_EXISTS = False

End Function


Public Function ADDIN_INSTALLED( _
    ByVal addinName As String, _
    Optional ByVal partialNameFlag As Boolean) _
As Boolean

    '@Description: This takes a string name of an Excel Addin, and checks if the Addin is currently installed and active.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: addinName is the name of a Addin we want to check if it is installed
    '@Param: partialNameFlag if set to TRUE, will perform checks based on partial names instead of requiring the exact name of the Addin
    '@Returns: Returns TRUE if the Addin is installed, and FALSE if it doesn't
    '@Note: This function is case-insensitive
    '@Example: =ADDIN_INSTALLED("SOLVER.XLAM") -> TRUE; Most versions of Excel will have the Solver Addin, and I currently have it installed
    '@Example: =ADDIN_INSTALLED("solver.xlam") -> TRUE; This function is case-insensitive
    '@Example: =ADDIN_INSTALLED("EUROTOOL.XLAM") -> FALSE; Many versions of Excel will have the Eurotools Addin, and it currently exists, but I currently don't have it installed, so this function returned FALSE
    '@Example: =ADDIN_INSTALLED("SOLVER", TRUE) -> TRUE; As the partialNameFlag is set and so "SOLVER" will match "SOLVER.XLAM"

    Application.Volatile

    Dim individualAddin As AddIn
    
    For Each individualAddin In Application.AddIns
        If Not partialNameFlag Then
            If UCase(individualAddin.Name) = UCase(addinName) Then
                If individualAddin.Installed Then
                    ADDIN_INSTALLED = True
                    Exit Function
                End If
            End If
        Else
            If InStr(1, individualAddin.Name, addinName, vbTextCompare) >= 0 Then
                If individualAddin.Installed Then
                    ADDIN_INSTALLED = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    ADDIN_INSTALLED = False

End Function
