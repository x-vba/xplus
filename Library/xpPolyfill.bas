Attribute VB_Name = "xpPolyfill"
'@Module: This module contains a set of functions that act as polyfills for functions in later versions of Excel. For example, MAXIF() is available in some later versions of Excel, but a user may not have access to this function if they are using an older version of Excel. In this case, this module adds a polyfill called MAX_IF() which works very similar to the MAXIF() function

Option Explicit


Public Function CONCAT_TEXT( _
    ParamArray rangeOrStringArray() As Variant) _
As String

    '@Description: This function takes multiple ranges and strings and concatenates all of them together. It is a polyfill for the CONCAT() function.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: rangeOrStringArray is any number of strings and ranges that will be concatenated together
    '@Returns: Returns a concatenated string
    '@Example: =CONCAT_TEXT(A1:A2, B1, "Two", B2) -> "HelloWorldOneTwoThree"; Where A1:A2=["Hello", "World"] and B1="One", B2="Three"

    Dim individualElement As Variant
    Dim individualRange As Range

    For Each individualElement In rangeOrStringArray
        If TypeName(individualElement) = "Range" Then
            For Each individualRange In individualElement
                CONCAT_TEXT = CONCAT_TEXT + individualRange.Value
            Next
        Else
            CONCAT_TEXT = CONCAT_TEXT + individualElement
        End If
    Next

End Function


Public Function MAX_IF( _
    ByVal maxRange As Range, _
    ByVal criteriaRange As Range, _
    ByVal criteriaValue As Variant) _
As Variant

    '@Description: This function takes a max range, a criteria range, and then a criteria value, and finds the maximum value given the criteria
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: maxRange is the range that the max value will be chosen from given the criteria
    '@Param: criteriaRange is the range that will be checked against the criteria
    '@Param: criteriaValue is the value of the criteria. The criteria allowed is similar to the SUMIF() function
    '@Returns: Returns the max value that passes the criteria
    '@Example: =MAX_IF(A1:A3, B1:B3, 10) -> 2; Where A1:A3=[1, 2, 3] and B1:B3=[20, 10, 5]
    '@Example: =MAX_IF(A1:A3, B1:B3, "<10") -> 3; Where A1:A3=[1, 2, 3] and B1:B3=[20, 10, 5]

    MAX_IF = MAX_IFS(maxRange, criteriaRange, criteriaValue)

End Function


Public Function MAX_IFS( _
    ByVal maxRange As Range, _
    ParamArray criteraRangeAndCriteria() As Variant) _
As Variant

    '@Description: This function takes a max range, and then any number or criteria ranges and criteria values, and returns the max value in the max range conditional on the values passing the criteria. It uses very similar criteria and syntax to the Excel Built-in SUMIFS().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: maxRange is the range that the max value will be chosen from given the criteria
    '@Param: criteraRangeAndCriteria is either a range that will be checked against a criteria, or the criteria. These values should alternate between criteria range and criteria value
    '@Returns: Returns the max value that passes all the criteria
    '@Example: =MAX_IFS(A1:A3, B1:B3, ">=10", C1:C3, "A") -> 2; Where A1:A3=[1, 2, 3], B1:B3=[20, 10, 5], and C1:C3=["A", "A", "C"]

    Dim i As Integer
    Dim k As Integer
    Dim maxValue As Variant
    Dim temporaryValueHolder As Variant
    Dim individualRange As Range
    Dim criteraRangeLength As Integer
    Dim currentCriteria As Variant
    Dim maxArray() As Variant
    
    criteraRangeLength = UBound(criteraRangeAndCriteria) - LBound(criteraRangeAndCriteria)
    
    ReDim maxArray(maxRange.Count, 1)
    For i = 1 To maxRange.Count
        maxArray(i, 0) = maxRange(i).Value
        maxArray(i, 1) = True
    Next
    
    For i = 0 To criteraRangeLength Step 2
        
        ' Checking if the criteria is a single cell, and if so set its value as
        ' the current criteria
        currentCriteria = criteraRangeAndCriteria(i + 1)
        If TypeName(currentCriteria) = "Range" Then
            If currentCriteria.Count = 1 Then
                currentCriteria = currentCriteria.Value
            End If
        End If
        
        ' Check if string, and then check for >, <, or <> symbols at the beginning
        If TypeName(currentCriteria) = "String" Then
        
            ' The not equal to case
            If Left(currentCriteria, 2) = "<>" Then
                temporaryValueHolder = CDbl(Mid(currentCriteria, 3))
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value <> temporaryValueHolder Then
                        If maxArray(k, 1) <> False Then
                            maxArray(k, 1) = True
                        End If
                    Else
                        maxArray(k, 1) = False
                    End If
                Next
                
            ' The greater than or equal to case
            ElseIf Left(currentCriteria, 2) = ">=" Then
                temporaryValueHolder = CDbl(Mid(currentCriteria, 3))
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value >= temporaryValueHolder Then
                        If maxArray(k, 1) <> False Then
                            maxArray(k, 1) = True
                        End If
                    Else
                        maxArray(k, 1) = False
                    End If
                Next
            
            ' The less than or equal to case
            ElseIf Left(currentCriteria, 2) = "<=" Then
                temporaryValueHolder = CDbl(Mid(currentCriteria, 3))
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value <= temporaryValueHolder Then
                        If maxArray(k, 1) <> False Then
                            maxArray(k, 1) = True
                        End If
                    Else
                        maxArray(k, 1) = False
                    End If
                Next
                
            ' The greater than case
            ElseIf Left(currentCriteria, 1) = ">" Then
                temporaryValueHolder = CDbl(Mid(currentCriteria, 2))
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value > temporaryValueHolder Then
                        If maxArray(k, 1) <> False Then
                            maxArray(k, 1) = True
                        End If
                    Else
                        maxArray(k, 1) = False
                    End If
                Next
            
            
            ' The less than case
            ElseIf Left(currentCriteria, 1) = "<" Then
                temporaryValueHolder = CDbl(Mid(currentCriteria, 2))
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value < temporaryValueHolder Then
                        If maxArray(k, 1) <> False Then
                            maxArray(k, 1) = True
                        End If
                    Else
                        maxArray(k, 1) = False
                    End If
                Next
            
            ' The pure string equality case when the string doesn't specify some greater than,
            ' less than, or not equal to criteria
            Else
                temporaryValueHolder = currentCriteria
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If CStr(criteraRangeAndCriteria(i)(k).Value) = temporaryValueHolder Then
                        If maxArray(k, 1) <> False Then
                            maxArray(k, 1) = True
                        End If
                    Else
                        maxArray(k, 1) = False
                    End If
                Next
            
            End If
        
        ' If numeric, then has to be purely equality comparison
        ElseIf IsNumeric(currentCriteria) Then
            temporaryValueHolder = currentCriteria
            For k = 1 To criteraRangeAndCriteria(i).Count
                If criteraRangeAndCriteria(i)(k).Value = temporaryValueHolder Then
                    If maxArray(k, 1) <> False Then
                        maxArray(k, 1) = True
                    End If
                Else
                    maxArray(k, 1) = False
                End If
            Next
        
        ' If range, then perform comparison to everything within the criteria
        ElseIf TypeName(currentCriteria) = "Range" Then
            For Each individualRange In currentCriteria
                temporaryValueHolder = individualRange.Value
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value = temporaryValueHolder Then
                        If maxArray(k, 1) <> False Then
                            maxArray(k, 1) = True
                        End If
                    Else
                        maxArray(k, 1) = False
                    End If
                Next
            Next
        
        End If
    Next

    ' Finding the max given all TRUE values
    For i = 1 To maxRange.Count
        If maxArray(i, 1) Then
            If IsEmpty(maxValue) Then
                maxValue = maxArray(i, 0)
            ElseIf maxValue < maxArray(i, 0) Then
                maxValue = maxArray(i, 0)
            End If
        End If
    Next

    ' If no criteria ia met, return an error string
    If IsEmpty(maxValue) Then
        MAX_IFS = "#NoCriteriaSatisfied!"
    Else
        MAX_IFS = maxValue
    End If

End Function


Public Function MIN_IF( _
    ByVal minRange As Range, _
    ByVal criteriaRange As Range, _
    ByVal criteriaValue As Variant) _
As Variant

    '@Description: This function takes a min range, a criteria range, and then a criteria value, and finds the minimum value given the criteria
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: minRange is the range that the min value will be chosen from given the criteria
    '@Param: criteriaRange is the range that will be checked against the criteria
    '@Param: criteriaValue is the value of the criteria. The criteria allowed is similar to the SUMIF() function
    '@Returns: Returns the min value that passes the criteria
    '@Example: =MIN_IF(A1:A3, B1:B3, 5) -> 3; Where A1:A3=[1, 2, 3] and B1:B3=[20, 10, 5]
    '@Example: =MIN_IF(A1:A3, B1:B3, "<=10") -> 2; Where A1:A3=[1, 2, 3] and B1:B3=[20, 10, 5]

    MIN_IF = MIN_IFS(minRange, criteriaRange, criteriaValue)

End Function


Public Function MIN_IFS( _
    ByVal minRange As Range, _
    ParamArray criteraRangeAndCriteria() As Variant) _
As Variant

    '@Description: This function takes a min range, and then any number or criteria ranges and criteria values, and returns the min value in the min range conditional on the values passing the criteria. It uses very similar criteria and syntax to the Excel Built-in SUMIFS().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: minRange is the range that the min value will be chosen from given the criteria
    '@Param: criteraRangeAndCriteria is either a range that will be checked against a criteria, or the criteria. These values should alternate between criteria range and criteria value
    '@Returns: Returns the min value that passes all the criteria
    '@Example: =MIN_IFS(A1:A3, B1:B3, ">=10", C1:C3, "A") -> 1; Where A1:A3=[1, 2, 3], B1:B3=[20, 10, 5], and C1:C3=["A", "A", "C"]

    Dim i As Integer
    Dim k As Integer
    Dim minValue As Variant
    Dim temporaryValueHolder As Variant
    Dim individualRange As Range
    Dim criteraRangeLength As Integer
    Dim currentCriteria As Variant
    Dim minArray() As Variant
    
    criteraRangeLength = UBound(criteraRangeAndCriteria) - LBound(criteraRangeAndCriteria)
    
    ReDim minArray(minRange.Count, 1)
    For i = 1 To minRange.Count
        minArray(i, 0) = minRange(i).Value
        minArray(i, 1) = True
    Next
    
    For i = 0 To criteraRangeLength Step 2
        
        ' Checking if the criteria is a single cell, and if so set its value as
        ' the current criteria
        currentCriteria = criteraRangeAndCriteria(i + 1)
        If TypeName(currentCriteria) = "Range" Then
            If currentCriteria.Count = 1 Then
                currentCriteria = currentCriteria.Value
            End If
        End If
        
        ' Check if string, and then check for >, <, or <> symbols at the beginning
        If TypeName(currentCriteria) = "String" Then
        
            ' The not equal to case
            If Left(currentCriteria, 2) = "<>" Then
                temporaryValueHolder = CDbl(Mid(currentCriteria, 3))
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value <> temporaryValueHolder Then
                        If minArray(k, 1) <> False Then
                            minArray(k, 1) = True
                        End If
                    Else
                        minArray(k, 1) = False
                    End If
                Next
            
            ' The greater than or equal to case
            ElseIf Left(currentCriteria, 2) = ">=" Then
                temporaryValueHolder = CDbl(Mid(currentCriteria, 3))
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value >= temporaryValueHolder Then
                        If minArray(k, 1) <> False Then
                            minArray(k, 1) = True
                        End If
                    Else
                        minArray(k, 1) = False
                    End If
                Next
            
            ' The less than or equal to case
            ElseIf Left(currentCriteria, 2) = "<=" Then
                temporaryValueHolder = CDbl(Mid(currentCriteria, 3))
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value <= temporaryValueHolder Then
                        If minArray(k, 1) <> False Then
                            minArray(k, 1) = True
                        End If
                    Else
                        minArray(k, 1) = False
                    End If
                Next
                
            ' The greater than case
            ElseIf Left(currentCriteria, 1) = ">" Then
                temporaryValueHolder = CDbl(Mid(currentCriteria, 2))
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value > temporaryValueHolder Then
                        If minArray(k, 1) <> False Then
                            minArray(k, 1) = True
                        End If
                    Else
                        minArray(k, 1) = False
                    End If
                Next
            
            
            ' The less than case
            ElseIf Left(currentCriteria, 1) = "<" Then
                temporaryValueHolder = CDbl(Mid(currentCriteria, 2))
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value < temporaryValueHolder Then
                        If minArray(k, 1) <> False Then
                            minArray(k, 1) = True
                        End If
                    Else
                        minArray(k, 1) = False
                    End If
                Next
            
            ' The pure string equality case when the string doesn't specify some greater than,
            ' less than, or not equal to criteria
            Else
                temporaryValueHolder = currentCriteria
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If CStr(criteraRangeAndCriteria(i)(k).Value) = temporaryValueHolder Then
                        If minArray(k, 1) <> False Then
                            minArray(k, 1) = True
                        End If
                    Else
                        minArray(k, 1) = False
                    End If
                Next
            
            End If
        
        ' If numeric, then has to be purely equality comparison
        ElseIf IsNumeric(currentCriteria) Then
            temporaryValueHolder = currentCriteria
            For k = 1 To criteraRangeAndCriteria(i).Count
                If criteraRangeAndCriteria(i)(k).Value = temporaryValueHolder Then
                    If minArray(k, 1) <> False Then
                        minArray(k, 1) = True
                    End If
                Else
                    minArray(k, 1) = False
                End If
            Next
        
        ' If range, then perform comparison to everything within the criteria
        ElseIf TypeName(currentCriteria) = "Range" Then
            For Each individualRange In currentCriteria
                temporaryValueHolder = individualRange.Value
                For k = 1 To criteraRangeAndCriteria(i).Count
                    If criteraRangeAndCriteria(i)(k).Value = temporaryValueHolder Then
                        If minArray(k, 1) <> False Then
                            minArray(k, 1) = True
                        End If
                    Else
                        minArray(k, 1) = False
                    End If
                Next
            Next
        
        End If
    Next

    ' Finding the max given all TRUE values
    For i = 1 To minRange.Count
        If minArray(i, 1) Then
            If IsEmpty(minValue) Then
                minValue = minArray(i, 0)
            ElseIf minValue > minArray(i, 0) Then
                minValue = minArray(i, 0)
            End If
        End If
    Next

    ' If no criteria ia met, return an error string
    If IsEmpty(minValue) Then
        MIN_IFS = "#NoCriteriaSatisfied!"
    Else
        MIN_IFS = minValue
    End If

End Function


Public Function TEXT_JOIN( _
    ByVal rangeArray As Range, _
    Optional ByVal delimiterCharacter As String, _
    Optional ByVal ignoreEmptyCellsFlag As Boolean) _
As String

    '@Description: This function takes a range of cells and combines all the text together, optionally allowing a character delimiter between all the combined strings, and optionally allowing blank cells to be ignored when combining the text. Finally note that this function is very similar to the TEXTJOIN function available in Excel 2019, and thus is a polyfill for that function for earlier versions of Excel.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: rangeArray is the range with all the strings we want to combine
    '@Param: delimiterCharacter is an optional character that will be used as the delimiter between the combined text. By default, no delimiter character will be used.
    '@Param: ignoreEmptyCellsFlag if set to TRUE will skip combining empty cells into the combined string, and is useful when specifying a delimiter so that the delimiter does not repeat for empty cells.
    '@Returns: Returns a new combined string containing the strings in the range delimited by the delimiter character.
    '@Example: =TEXT_JOIN(A1:A3) -> "123"; Where A1:A3 contains ["1", "2", "3"]
    '@Example: =TEXT_JOIN(A1:A3, "--") -> "1--2--3"; Where A1:A3 contains ["1", "2", "3"]
    '@Example: =TEXT_JOIN(A1:A3, "--") -> "1----3"; Where A1:A3 contains ["1", "", "3"]
    '@Example: =TEXT_JOIN(A1:A3, "-") -> "1--3"; Where A1:A3 contains ["1", "", "3"]
    '@Example: =TEXT_JOIN(A1:A3, "-", TRUE) -> "1-3"; Where A1:A3 contains ["1", "", "3"]

    Dim individualRange As Range
    Dim combinedString As String
    
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

