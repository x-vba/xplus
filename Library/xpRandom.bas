Attribute VB_Name = "xpRandom"
'@Module: This module contains a set of functions for generating and sampling random data.

Option Explicit


Public Function RANDOM_SAMPLE( _
    ByVal rangeArray As Range) _
As Variant

    '@Description: This function takes an array of cells and returns a random value from the cells chosen
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Check if there is native randbetween function besides the worksheet function
    '@Param: rangeArray a single cell or multiple cells where the sample will be pulled from
    '@Returns: Returns a random cell value from the array of cells chosen
    '@Example: =RANDOM_SAMPLE(A1:A5) -> "Hello"; where "Hello" is the value in cell A3, and where A3 was the chosen random cell
    '@Example: =RANDOM_SAMPLE(A1:A5) -> "World"; where "World" is the value in cell A2, and where A2 was the chosen random cell

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
    '@Example: =RANDOM_RANGE(50, 100, 10) -> 60
    '@Example: =RANDOM_RANGE(50, 100, 10) -> 50
    '@Example: =RANDOM_RANGE(50, 100, 10) -> 90
    '@Example: =RANDOM_RANGE(0, 10, 2) -> 8
    '@Example: =RANDOM_RANGE(0, 10, 2) -> 0
    '@Example: =RANDOM_RANGE(0, 10, 2) -> 4
    '@Example: =RANDOM_RANGE(0, 10, 2) -> 10

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
    '@Warning: Internally this function sums up all the number in the percentRange to calculate percentage chances of a sample. For example, if the percentRange contains the values 10 and 90, the first value will have a 10/(10+90) = 10% chance of being chosen. Similarly, if the values in the percentRange contains the values 5 and 45, the first value will have a 5/(5+45) = 10% chance of being chosen. This means you have to be careful when choosing percentages in the percentRange, as if you choose 0.1 and 0.8 for the percentRange, the percentage chance the first value is chosen is NOT 10%, but rather 0.1/(0.1+0.8) = 11.1%. Thus, you should be careful to only interpret a 0.1 in the percentRange as a 10% chance only if the values in the percentRange actually sum up to 1.0.
    '@Example: =RANDOM_SAMPLE_PERCENT(A1:A2, B1:B2) -> "Hello"; Assuming the valueRange contains ["Hello", "World"], and percentRange contains [.9, .1]
    '@Example: =RANDOM_SAMPLE_PERCENT(A1:A2, B1:B2) -> "Hello"; Assuming the valueRange contains ["Hello", "World"], and percentRange contains [.9, .1]
    '@Example: =RANDOM_SAMPLE_PERCENT(A1:A2, B1:B2) -> "Hello"; Assuming the valueRange contains ["Hello", "World"], and percentRange contains [.9, .1]
    '@Example: =RANDOM_SAMPLE_PERCENT(A1:A2, B1:B2) -> "Hello"; Assuming the valueRange contains ["Hello", "World"], and percentRange contains [.9, .1]
    '@Example: =RANDOM_SAMPLE_PERCENT(A1:A2, B1:B2) -> "Hello"; Assuming the valueRange contains ["Hello", "World"], and percentRange contains [.9, .1]
    '@Example: =RANDOM_SAMPLE_PERCENT(A1:A2, B1:B2) -> "Hello"; Assuming the valueRange contains ["Hello", "World"], and percentRange contains [.9, .1]
    '@Example: =RANDOM_SAMPLE_PERCENT(A1:A2, B1:B2) -> "Hello"; Assuming the valueRange contains ["Hello", "World"], and percentRange contains [.9, .1]
    '@Example: =RANDOM_SAMPLE_PERCENT(A1:A2, B1:B2) -> "World"; Assuming the valueRange contains ["Hello", "World"], and percentRange contains [.9, .1]; Notice how "World" shows up less since there is only a 10% chance it is chosen.

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


Public Function RANDBOOL() As Boolean

    '@Description: This function generates a random Boolean (TRUE or FALSE) value
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns either TRUE or FALSE based on the random value choosen
    '@Example: =RANDBOOL() -> TRUE
    '@Example: =RANDBOOL() -> FALSE
    '@Example: =RANDBOOL() -> TRUE
    '@Example: =RANDBOOL() -> TRUE
    '@Example: =RANDBOOL() -> FALSE
    '@Example: =RANDBOOL() -> FALSE

    RANDBOOL = CBool(WorksheetFunction.RandBetween(0, 1))

End Function


Public Function RANDBETWEENS( _
    ParamArray startOrEndNumberArray() As Variant) _
As Variant

    '@Description: This function is similar to RANDBETWEEN, except that it allows multiple ranges from which to pick a random number. One of the ranges from which to generate a random number between is chosen at an equal probably.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns either TRUE or FALSE based on the random value choosen
    '@Note: This function always requires an even number of inputs. Essentially, when using multiple numbers, the 1st and 2nd will make up a range from which to pull a random number between, the 3rd and 4th will make a different range, and so on. If an even number is used, this function will return a User-Defined Error. See the ISERRORALL() function for how to handle these numbers.
    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 6
    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 5002
    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 8
    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 3
    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 5010
    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 2
    '@Example: =RANDBETWEENS(5, 10, 15, 20, 25, 30, 35, 40) -> 32

    Dim pickNumber As Byte

    ' Checking for ParamArray length, as it needs to be even or it won't be
    ' possible to generate and min and max number.
    If (UBound(startOrEndNumberArray) - LBound(startOrEndNumberArray) + 1) Mod 2 = 1 Then
        RANDBETWEENS = "#NotAnEvenNumberOfParameters!"
    End If

    pickNumber = WorksheetFunction.Ceiling_Math(WorksheetFunction.RandBetween(1, (UBound(startOrEndNumberArray) - LBound(startOrEndNumberArray) + 1)) / 2) * 2
    
    RANDBETWEENS = WorksheetFunction.RandBetween(startOrEndNumberArray(pickNumber - 2), startOrEndNumberArray(pickNumber - 1))

End Function

