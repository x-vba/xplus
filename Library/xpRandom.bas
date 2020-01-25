Attribute VB_Name = "xpRandom"
'@Module: This module contains a set of functions for generating and sampling random data

Option Explicit


Public Function RANDOM_SAMPLE( _
    ByVal rangeArray As Range) _
As Variant

    '@Description: This function takes an array of cells and returns a random value from the cells chosen
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Check if there is native randbetween function besides the worksheet function
    '@Param: rangeArray a single cell or multiple cells
    '@Returns: Returns a random cell value from the array of cells chosen
    '@Example: =RANDOM_SAMPLE(A1:A5) -> "Hello World"; where "Hello World" is the value in cell A3

    Dim randomNumber As Integer
    
    randomNumber = WorksheetFunction.RandBetween(1, rangeArray.Count)
    
    RANDOM_SAMPLE = rangeArray(randomNumber).Value

End Function


Public Function RANDOM_RANGE( _
    ByVal startNumber As Integer, _
    ByVal stopNumber As Integer, _
    ByVal stepNumber As Integer) _
As Integer

    '@Description: This function takes 3 numbers, a start number, a stop number, and a step number, and returns a random number between the start number and stop number that is an interval of the step number.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Check if there is native randbetween function besides the worksheet function
    '@Param: startNumber is the beginning value of the range
    '@Param: stopNumber is the end value of the range
    '@Param: stepNumber is the step of the range
    '@Returns: Returns a random number between the start and stop that is a multiple of the step
    '@Example: =RANDOM_RANGE(50,100,10) -> 60

    Dim randomNumber As Integer
    
    randomNumber = WorksheetFunction.RandBetween(startNumber / stepNumber, stopNumber / stepNumber) * stepNumber
    
    RANDOM_RANGE = randomNumber

End Function


Public Function RANDOM_SAMPLE_PERCENT( _
    ByVal valueRange As Range, _
    ByVal percentRange As Range) _
As Variant

    '@Description: This function takes 2 ranges, one with values that will be sampled, and the other with the percentage chance that the value will be sampled, and returns a value from the sample.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: valueRange is the range containing values of which one will be sampled
    '@Param: percentRange is the range that contains the percent chances of the values in the valueRange
    '@Returns: Returns a random value from the valueRange
    '@Example: =RANDOM_SAMPLE_PERCENT(A1:A2, B1:B2) -> "Hello" ;Assuming the valueRange contains ["Hello", "World"], and percentRange contains [.9, .1]

    Application.Volatile

    ' Creating a datagrid to perform the search on
    Dim i As Integer
    Dim cumulativeSum As Double
    Dim cumulativePercentage As Double
    Dim dataGrid() As Variant
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
    
    
    ' Getting the random value
    Dim randomNumber As Double
    Dim randomValue As Variant
    
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

