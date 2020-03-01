Attribute VB_Name = "xpRange"
'@Module: This module contains a set of functions for manipulating and working with ranges of cells.

Option Explicit


Public Function FIRST_UNIQUE( _
    ByVal range1 As Range, _
    ByVal rangeArray As Range) _
As Boolean

    '@Description: This function takes a single cell and an large range of cells and returns TRUE if the cell selected is the first unique value in the larger array of cells, and returns FALSE if it is not the first unique value.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the range we want to check if the value is the first unique value in the rangeArray
    '@Param: rangeArray is the group of cells we are checking to see if range1 is the first unique occurrence in the rangeArray
    '@Returns: Returns TRUE if the cell selected is the first unique value in the range array, and FALSE if it isn't
    '@Example: =FIRST_UNIQUE(A1, $A$1:$A$10) -> TRUE, where A1 is the first unique occurrence of the word "Hello" in the range array
    '@Example: =FIRST_UNIQUE(A5, $A$1:$A$10) -> FALSE, where A5 is the second unique occurrence of the word "Hello" in the range array

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


Public Function COUNT_UNIQUE( _
    ParamArray rangeArray() As Variant) _
As Integer
    
    '@Description: This function counts the number of unique occurances of values within a range or multiple ranges
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: rangeArray is the group of cells we are counting the unique values of
    '@Returns: Returns the number of unique values
    '@Example: =COUNT_UNIQUE(A1:A5) -> 3; Where A1-A5 contains ["A", "A", "B", "A", "C"]
    
    Dim individualValue As Variant
    Dim individualRange As Range
    Dim uniqueDictionary As Object
    Dim uniqueCount As Integer
    
    Set uniqueDictionary = CreateObject("Scripting.Dictionary")
    
    For Each individualValue In rangeArray
        If TypeName(individualValue) = "Range" Then
            For Each individualRange In individualValue
                If Not uniqueDictionary.Exists(individualRange.Value) Then
                    uniqueDictionary.Add individualRange.Value, 0
                    uniqueCount = uniqueCount + 1
                End If
            Next
        Else
            If Not uniqueDictionary.Exists(individualValue) Then
                uniqueDictionary.Add individualValue, 0
                uniqueCount = uniqueCount + 1
            End If
        End If
    Next
    
    COUNT_UNIQUE = uniqueCount
    
End Function


Private Function BubbleSort( _
    ByRef sortableArray As Variant, _
    Optional ByVal descendingFlag As Boolean) _
As Variant

    '@Description: This function is an implementation of Bubble Sort
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: descendingFlag changes the sort to descending
    '@Returns: Returns the a sorted array
    '@Example: =BubbleSort({1,3,2}) -> {1,2,3}
    '@Example: =BubbleSort({1,3,2}, True) -> {3,2,1}

    Dim i As Integer
    Dim swapOccuredBool As Boolean
    Dim arrayLength As Integer
    arrayLength = UBound(sortableArray) - LBound(sortableArray) + 1
    
    Dim sortedArray() As Variant
    ReDim sortedArray(arrayLength)
    
    For i = 0 To arrayLength - 1
        sortedArray(i) = sortableArray(i)
    Next
    
    Dim temporaryValue As Variant
    
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
        Dim ascendingArray() As Variant
        ReDim ascendingArray(arrayLength)
        
        For i = 0 To arrayLength - 1
            ascendingArray(i) = sortedArray(arrayLength - i - 1)
        Next
        
        BubbleSort = ascendingArray
    End If
    
End Function


Function SORT_RANGE( _
    ByVal range1 As Range, _
    ByVal rangeArray As Range, _
    Optional ByVal descendingFlag As Boolean) _
As Variant

    '@Description: This function takes a single cell and a large range of cells and sorts the cells in ascending or descending order.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the range containing a single cell that we want to sort
    '@Param: rangeArray is the group of cells we are sorting
    '@Param: descendingFlag is a Boolean value that if set to TRUE will set the sort to Descending
    '@Returns: Returns the value of the cell sorted
    '@Example: =SORT_RANGE(A1, $A$1:$A$4) -> 1, where A1="3", A2="1", A3="4", A4="2"
    '@Example: =SORT_RANGE(A1, $A$1:$A$4, TRUE) -> 4, where A1="3", A2="1", A3="4", A4="2"

    ' Sorting the values from rangeArray
    Dim variantArray() As Variant
    ReDim variantArray(rangeArray.Count)
    Dim returnArray() As Variant
    ReDim returnArray(rangeArray.Count)
    Dim returnBoolean As Boolean
    Dim i As Integer
    
    For i = 1 To rangeArray.Count
        variantArray(i) = rangeArray(i)
    Next
    
    returnArray = BubbleSort(variantArray, descendingFlag)
    
    
    ' Returning the value in the rangeArray based on the address of range1
    Dim k As Integer
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


Public Function REVERSE_RANGE( _
    ByVal range1 As Range, _
    ByVal rangeArray As Range) _
As Variant

    '@Description: This function takes a single cell and a large range of cells and reverses all values in the range.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the range we want to be reversed in order of the rangeArray
    '@Param: rangeArray is the group of cells we are reversing the order of
    '@Returns: Returns the value of the cell in the reversed position
    '@Example: =REVERSE_RANGE(A1, $A$1:$A$3) -> "C", where A1="A", A2="B", A3="C"
    '@Example: =REVERSE_RANGE(A2, $A$1:$A$3) -> "B", where A1="A", A2="B", A3="C"
    '@Example: =REVERSE_RANGE(A3, $A$1:$A$3) -> "A", where A1="A", A2="B", A3="C"

    Dim i As Integer
    
    For i = 1 To rangeArray.Count
        If range1.Address = rangeArray(i).Address Then
            REVERSE_RANGE = rangeArray(rangeArray.Count - i + 1).Value
            Exit Function
        End If
    Next

End Function


Public Function COLUMNIFY( _
    ByVal columnRangeArray As Range, _
    ByVal rowRangeArray As Range) _
As Variant

    '@Description: This function takes 2 ranges, a column range which will be filled in with data in the row range, allowing for easily converting a row of data into a column of data
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: columnRangeArray is a range that will be populated with data from a rowRangeArray
    '@Param: rowRangeArray is a range that will be used to populate the columnRangeArray
    '@Returns: Returns the value at the same location in the rowRangeArray
    '@Example: =COLUMNIFY(A1:A2, B1:C1) -> "B"; Where this function resides in cell A1 and where B1="B" and C1="C"
    '@Example: =COLUMNIFY(A1:A2, B1:C1) -> "C"; Where this function resides in cell A2 and where B1="B" and C1="C"

    Dim i As Integer
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


Public Function ROWIFY( _
    ByVal rowRangeArray As Range, _
    ByVal columnRangeArray As Range) _
As Variant

    '@Description: This function takes 2 ranges, a row range which will be filled in with data in the column range, allowing for easily converting a column of data into a row of data
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: rowRangeArray is a range that will be populated with data from a columnRangeArray
    '@Param: columnRangeArray is a range that will be used to populate the rowRangeArray
    '@Returns: Returns the value at the same location in the columnRangeArray
    '@Example: =ROWIFY(B1:C1, A1:A2) -> "A1"; Where this function resides in cell B1 and where A1="A1" and A2="A2"
    '@Example: =ROWIFY(B1:C1, A1:A2) -> "A2"; Where this function resides in cell C1 and where A1="A1" and A2="A2"

    Dim i As Integer
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


Public Function SUMN( _
    ByVal rangeArray As Range, _
    ByVal nthNumber As Integer, _
    Optional ByVal startAtBeginningFlag As Boolean) _
As Variant

    '@Description: This function sums up every Nth value of a range. For example, if you have a range that is 4 cells long, and set the nthNumber to 2, then only the 2nd and 4th cell value will be summed up. Optionally, a third parameter can be set to TRUE, and if so the summing will start at the first cell. For example, for 4 cells in a range and for the nthNumber set to 2, the 1st and 3rd cell will be summed.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Note: If the range chosen is more than 1 cell in width, the summing will occur in left-to-right and then top-to-bottom order
    '@Param: rangeArray is the range that will be summed up
    '@Param: nthNumber is the number which will determine which cells are summed
    '@Param: startAtBeginningFlag is an optional value that if set to TRUE will make the sum start at the first cell instead of at the nth cell
    '@Returns: Returns the sum of the nth cells
    '@Example: =SUMN(A1:A4, 2) -> 6; Where A1=1, A2=2, A3=3, A4=4
    '@Example: =SUMN(A1:A4, 2, TRUE) -> 4; Where A1=1, A2=2, A3=3, A4=4

    Dim i As Integer
    Dim sumValue As Double
    
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


Public Function AVERAGEN( _
    ByVal rangeArray As Range, _
    ByVal nthNumber As Integer, _
    Optional ByVal startAtBeginningFlag As Boolean) _
As Variant

    '@Description: This function averages up every Nth value of a range. For example, if you have a range that is 4 cells long, and set the nthNumber to 2, then only the 2nd and 4th cell value will be averaged up. Optionally, a third parameter can be set to TRUE, and if so the averaging will start at the first cell. For example, for 4 cells in a range and for the nthNumber set to 2, the 1st and 3rd cell will be averaged.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Note: If the range chosen is more than 1 cell in width, the averaging will occur in left-to-right and then top-to-bottom order
    '@Param: rangeArray is the range that will be averaged up
    '@Param: nthNumber is the number which will determine which cells are averaged
    '@Param: startAtBeginningFlag is an optional value that if set to TRUE will make the average start at the first cell instead of at the nth cell
    '@Returns: Returns the average of the nth cells
    '@Example: =AVERAGEN(A1:A4, 2) -> 3; Where A1=1, A2=2, A3=3, A4=4
    '@Example: =AVERAGEN(A1:A4, 2, TRUE) -> 2; Where A1=1, A2=2, A3=3, A4=4

    Dim i As Integer
    Dim sumValue As Double
    Dim countValue As Integer
    
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


Public Function MAXN( _
    ByVal rangeArray As Range, _
    ByVal nthNumber As Integer, _
    Optional ByVal startAtBeginningFlag As Boolean) _
As Variant

    '@Description: This function maxes up every Nth value of a range. For example, if you have a range that is 4 cells long, and set the nthNumber to 2, then only the 2nd and 4th cell value will be maxed up. Optionally, a third parameter can be set to TRUE, and if so the maxing will start at the first cell. For example, for 4 cells in a range and for the nthNumber set to 2, the 1st and 3rd cell will be maxed.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Note: If the range chosen is more than 1 cell in width, the maxing will occur in left-to-right and then top-to-bottom order
    '@Param: rangeArray is the range that will be maxed up
    '@Param: nthNumber is the number which will determine which cells are maxed
    '@Param: startAtBeginningFlag is an optional value that if set to TRUE will make the max start at the first cell instead of at the nth cell
    '@Returns: Returns the max of the nth cells
    '@Example: =MAXN(A1:A4, 2) -> 4; Where A1=1, A2=2, A3=3, A4=4
    '@Example: =MAXN(A1:A4, 2, TRUE) -> 3; Where A1=1, A2=2, A3=3, A4=4

    Dim i As Integer
    Dim sumValue As Double
    Dim maxValue As Variant
    
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


Public Function MINN( _
    ByVal rangeArray As Range, _
    ByVal nthNumber As Integer, _
    Optional ByVal startAtBeginningFlag As Boolean) _
As Variant

    '@Description: This function mins up every Nth value of a range. For example, if you have a range that is 4 cells long, and set the nthNumber to 2, then only the 2nd and 4th cell value will be minned up. Optionally, a third parameter can be set to TRUE, and if so the minning will start at the first cell. For example, for 4 cells in a range and for the nthNumber set to 2, the 1st and 3rd cell will be minned.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Note: If the range chosen is more than 1 cell in width, the minning will occur in left-to-right and then top-to-bottom order
    '@Param: rangeArray is the range that will be minned up
    '@Param: nthNumber is the number which will determine which cells are minned
    '@Param: startAtBeginningFlag is an optional value that if set to TRUE will make the min start at the first cell instead of at the nth cell
    '@Returns: Returns the min of the nth cells
    '@Example: =MINN(A1:A4, 2) -> 2; Where A1=1, A2=2, A3=3, A4=4
    '@Example: =MINN(A1:A4, 2, TRUE) -> 1; Where A1=1, A2=2, A3=3, A4=4

    Dim i As Integer
    Dim sumValue As Double
    Dim minValue As Variant
    
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


Public Function SUMHIGH( _
    ByVal rangeArray As Range, _
    ByVal numberSummed As Integer) _
As Variant

    '@Description: This function returns the sum of the top values of the number specified in the second argument. For example, if the second argument is 3, only the top 3 values will be summed
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Check if the Large worksheet function is available in older versions of Excel
    '@Param: rangeArray is the range that will be summed
    '@Param: numberSummed is the number of the top values that will be summed
    '@Returns: Returns the sum of the top numbers specified
    '@Example: =SUMHIGH(A1:A4, 2) -> 7; Where A1=1, A2=2, A3=3, A4=4; 4 and 3 are summed to 7
    '@Example: =SUMHIGH(A1:A4, 3) -> 9; Where A1=1, A2=2, A3=3, A4=4; 4, 3, and 2 are summed to 9

    Dim i As Integer
    Dim sumValue As Double
    
    For i = 1 To numberSummed
        sumValue = sumValue + WorksheetFunction.Large(rangeArray, i)
    Next
    
    SUMHIGH = sumValue

End Function


Public Function SUMLOW( _
    ByVal rangeArray As Range, _
    ByVal numberSummed As Integer) _
As Variant

    '@Description: This function returns the sum of the bottom values of the number specified in the second argument. For example, if the second argument is 3, only the bottom 3 values will be summed
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Check if the Small worksheet function is available in older versions of Excel
    '@Param: rangeArray is the range that will be summed
    '@Param: numberSummed is the number of the bottom values that will be summed
    '@Returns: Returns the sum of the bottom numbers specified
    '@Example: =SUMLOW(A1:A4, 2) -> 3; Where A1=1, A2=2, A3=3, A4=4; 1 and 2 are summed to 3
    '@Example: =SUMLOW(A1:A4, 3) -> 6; Where A1=1, A2=2, A3=3, A4=4; 1, 2, and 3 are summed to 6

    Dim i As Integer
    Dim sumValue As Double
    
    For i = 1 To numberSummed
        sumValue = sumValue + WorksheetFunction.Small(rangeArray, i)
    Next
    
    SUMLOW = sumValue

End Function


Public Function AVERAGEHIGH( _
    ByVal rangeArray As Range, _
    ByVal numberAveraged As Integer) _
As Variant

    '@Description: This function returns the average of the top values of the number specified in the second argument. For example, if the second argument is 3, only the top 3 values will be averaged
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Check if the Large worksheet function is available in older versions of Excel
    '@Param: rangeArray is the range that will be averaged
    '@Param: numberAveraged is the number of the top values that will be averaged
    '@Returns: Returns the average of the top numbers specified
    '@Example: =AVERAGEHIGH(A1:A4, 2) -> 3.5; Where A1=1, A2=2, A3=3, A4=4; 4 and 3 are averaged to 3.5
    '@Example: =AVERAGEHIGH(A1:A4, 3) -> 3; Where A1=1, A2=2, A3=3, A4=4; 4, 3, and 2 are averaged to 3

    Dim i As Integer
    Dim sumValue As Double
    
    For i = 1 To numberAveraged
        sumValue = sumValue + WorksheetFunction.Large(rangeArray, i)
    Next
    
    AVERAGEHIGH = sumValue / numberAveraged

End Function


Public Function AVERAGELOW( _
    ByVal rangeArray As Range, _
    ByVal numberAveraged As Integer) _
As Variant

    '@Description: This function returns the average of the bottom values of the number specified in the second argument. For example, if the second argument is 3, only the bottom 3 values will be averaged
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Check if the Small worksheet function is available in older versions of Excel
    '@Param: rangeArray is the range that will be averaged
    '@Param: numberAveraged is the number of the bottom values that will be averaged
    '@Returns: Returns the average of the bottom numbers specified
    '@Example: =AVERAGELOW(A1:A4, 2) -> 1.5; Where A1=1, A2=2, A3=3, A4=4; 1 and 2 are averaged as 1.5
    '@Example: =AVERAGELOW(A1:A4, 3) -> 2; Where A1=1, A2=2, A3=3, A4=4; 1, 2, and 3 are averaged to 2

    Dim i As Integer
    Dim sumValue As Double
    
    For i = 1 To numberAveraged
        sumValue = sumValue + WorksheetFunction.Small(rangeArray, i)
    Next
    
    AVERAGELOW = sumValue / numberAveraged

End Function


Public Function INRANGE( _
    ByVal valueOrRange As Variant, _
    ByVal searchRange As Range) _
As Boolean

    '@Description: This function takes a range or a value, and a second range, and returns TRUE if the first range or value is within the second range. Otherwise it returns FALSE
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: valueOrRange is the range or value that will be checked if it exists in the search range
    '@Param: searchRange is the range that contains values that will be checked in the second range
    '@Returns: Returns TRUE if the first value or range contains a value in the second range
    '@Example: =INRANGE(A1:A3, B1:B3) -> TRUE; Where A1="One", A2="Two", A3="Three", B1="Five", B2="Six", and B3="One"; TRUE since "One" occurs in both ranges
    '@Example: =INRANGE(A1:A3, B1:B3) -> TRUE; Where A1="One", A2="Two", A3="Three", B1="Five", B2="Six", and B3="Seven"; FALSE since the ranges have no values in common
    '@Example: =INRANGE("Five", B1:B3) -> TRUE; B1="Five", B2="Six", and B3="Seven"; TRUE since "Five" is in the search range

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


Public Function COUNT_UNIQUE_COLORS( _
    ParamArray rangeArray() As Variant) _
As Integer

    '@Description: This function counts the number of unique background colors of the cells in a range or multiple ranges
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: rangeArray is a range or multiple ranges whose colors will be counted
    '@Returns: Returns the number of unique background colors of all the cells
    '@Example: =COUNT_UNIQUE_COLORS(A1:A3) -> 3; Where all the cells have a unique background color
    '@Example: =COUNT_UNIQUE_COLORS(A1:A3) -> 2; Where A1 and A2 have the same background color

    Dim colorCount As Integer
    Dim colorDictionary As Object
    Dim individualRange As Variant
    Dim individualCell As Range
    
    Set colorDictionary = CreateObject("Scripting.Dictionary")
    
    For Each individualRange In rangeArray
        For Each individualCell In individualRange
            If Not colorDictionary.Exists(individualCell.Interior.Color) Then
                colorDictionary.Add individualCell.Interior.Color, "0"
                colorCount = colorCount + 1
            End If
        Next
    Next
    
    COUNT_UNIQUE_COLORS = colorCount

End Function


Public Function ALTERNATE_COLUMNS( _
    ByVal rangeGrid As Range, _
    ByVal outputRange As Range) _
As Variant

    '@Description: This function takes a grid of cells, and converts the grid into a single columns where the values of the grid alternate between the columns in the grid.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: rangeGrid is a grid of cells
    '@Param: outputRange is the column that will be populated with the data from the grid
    '@Returns: Returns one of the values from the grid in alternating column order
    '@Example: =ALTERNATE_COLUMNS($A$1:$B$2, $C$1:$C$4) -> "A1 Value"; Where this function is the 1st cell in the column
    '@Example: =ALTERNATE_COLUMNS($A$1:$B$2, $C$1:$C$4) -> "B1 Value"; Where this function is the 2nd cell in the column
    '@Example: =ALTERNATE_COLUMNS($A$1:$B$2, $C$1:$C$4) -> "A2 Value"; Where this function is the 3rd cell in the column
    '@Example: =ALTERNATE_COLUMNS($A$1:$B$2, $C$1:$C$4) -> "B2 Value"; Where this function is the 4th cell in the column

    Dim cellPosition As Integer
    Dim individualRange As Range
    Dim addressFoundFlag As Boolean
    
    For Each individualRange In outputRange
        cellPosition = cellPosition + 1
        If individualRange.Address = Application.Caller.Address Then
            addressFoundFlag = True
            Exit For
        End If
    Next
    
    If addressFoundFlag Then
        ALTERNATE_COLUMNS = rangeGrid(cellPosition).Value
    Else
        ALTERNATE_COLUMNS = ""
    End If

End Function


Public Function ALTERNATE_ROWS( _
    ByVal rangeGrid As Range, _
    ByVal outputRange As Range) _
As Variant

    '@Description: This function takes a grid of cells, and converts the grid into a single columns where the values of the grid alternate between the rows in the grid.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: rangeGrid is a grid of cells
    '@Param: outputRange is the column that will be populated with the data from the grid
    '@Returns: Returns one of the values from the grid in alternating row order
    '@Example: =ALTERNATE_ROWS($A$1:$B$2, $C$1:$C$4) -> "A1 Value"; Where this function is the 1st cell in the column
    '@Example: =ALTERNATE_ROWS($A$1:$B$2, $C$1:$C$4) -> "A2 Value"; Where this function is the 2nd cell in the column
    '@Example: =ALTERNATE_ROWS($A$1:$B$2, $C$1:$C$4) -> "B1 Value"; Where this function is the 3rd cell in the column
    '@Example: =ALTERNATE_ROWS($A$1:$B$2, $C$1:$C$4) -> "B2 Value"; Where this function is the 4th cell in the column

    Dim cellPosition As Integer
    Dim individualRange As Range
    Dim addressFoundFlag As Boolean
    
    For Each individualRange In outputRange
        If individualRange.Address = Application.Caller.Address Then
            addressFoundFlag = True
            Exit For
        End If
        cellPosition = cellPosition + 1
    Next
    
    Dim rowNumber As Integer
    Dim cellNumber As Integer
    
    rowNumber = cellPosition Mod rangeGrid.Rows().Count
    cellNumber = WorksheetFunction.Floor_Math(cellPosition / rangeGrid.Rows().Count) + 1
    
    If addressFoundFlag Then
        ALTERNATE_ROWS = rangeGrid(rowNumber * rangeGrid.Rows().Count + cellNumber).Value
    Else
        ALTERNATE_ROWS = ""
    End If

End Function

