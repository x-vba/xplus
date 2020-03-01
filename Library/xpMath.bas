Attribute VB_Name = "xpMath"
'@Module: This module contains a set of basic mathematical functions where those functions don't already exist as base Excel functions.

Option Explicit


Public Function INTERPOLATE_NUMBER( _
    ByVal startingNumber As Double, _
    ByVal endingNumber As Double, _
    ByVal interpolationPercentage As Double) _
As Double

    '@Description: This function takes three numbers, a starting number, an ending number, and an interpolation percent, and linearly interpolates the number at the given percentage between the starting and ending number.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: startingNumber is the beginning number of the interpolation
    '@Param: endingNumber is the ending number of the interpolation
    '@Param: interpolationPercentage is the percentage that will be interpolated linearly between the startingNumber and the endingNumber
    '@Returns: Returns the linearly interpolated number between the two points
    '@Example: =INTERPOLATE_NUMBER(10, 20, 0.5) -> 15; Where 0.5 would be 50% between 10 and 20
    '@Example: =INTERPOLATE_NUMBER(16, 124, 0.64) -> 85.12; Where 0.64 would be 64% between 16 and 124

    INTERPOLATE_NUMBER = startingNumber + ((endingNumber - startingNumber) * interpolationPercentage)

End Function


Public Function INTERPOLATE_PERCENT( _
    ByVal startingNumber As Double, _
    ByVal endingNumber As Double, _
    ByVal interpolationNumber As Double) _
As Double

    '@Description: This function takes three numbers, a starting number, an ending number, and an interpolation number, and linearly interpolates the percentage location of the interpolated number between the starting and ending number.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: startingNumber is the beginning number of the interpolation
    '@Param: endingNumber is the ending number of the interpolation
    '@Param: interpolationNumber is the number that will be interpolated linearly between the startingNumber and the endingNumber to calculate a percentage
    '@Returns: Returns the linearly interpolated percent between the two points given the interpolation number
    '@Example: =INTERPOLATE_PERCENT(10, 18, 12) -> 0.25; As 12 is 25% of the way from 10 to 18
    '@Example: =INTERPOLATE_PERCENT(10, 20, 15) -> 0.5; As 15 is 50% of the way from 10 to 20

    INTERPOLATE_PERCENT = (interpolationNumber - startingNumber) / (endingNumber - startingNumber)

End Function

