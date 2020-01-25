Attribute VB_Name = "xpColor"
'@Module: This module contains a set of functions for working with colors

Option Explicit


Public Function RGB2HEX( _
    ByVal redColorInteger As Integer, _
    ByVal greenColorInteger As Integer, _
    ByVal blueColorInteger As Integer) _
As String

    '@Description: This function converts an RGB color value into a HEX color value
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: redColorInteger is the red value
    '@Param: greenColorInteger is the green value
    '@Param: blueColorInteger is the blue value
    '@Returns: Returns a string with the HEX value of the color
    '@Example: =RGB2HEX(255,255,255) -> "FFFFFF"

    RGB2HEX = WorksheetFunction.Dec2Hex(redColorInteger, 2) & WorksheetFunction.Dec2Hex(greenColorInteger, 2) & WorksheetFunction.Dec2Hex(blueColorInteger, 2)
    
End Function

Public Function HEX2RGB( _
    ByVal hexColorString As String, _
    Optional ByVal singleColorNumberOrName As Variant = -1) _
As Variant

    '@Description: This function converts a HEX color value into an RGB color value
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: hexColorString is the color in HEX format
    '@Param: singleColorNumberOrName is a number or string that specifies which of the single color values to return. If this is set to 0 or "Red", the red value will be returned. If this is set to 1 or "Green", the green value will be returned. If this is set to 2 or "Blue", the blue value will be returned.
    '@Returns: Returns a string with the RGB value of the color or the number of the individual color chosen
    '@Example: =HEX2RGB("FFFFFF") -> "(255, 255, 255)"
    '@Example: =HEX2RGB("FF0109", 0) -> 255; The red color
    '@Example: =HEX2RGB("FF0109", "Red") -> 255; The red color
    '@Example: =HEX2RGB("FF0109", 1) -> 1; The green color
    '@Example: =HEX2RGB("FF0109", "Green") -> 1; The green color
    '@Example: =HEX2RGB("FF0109", 2) -> 9; The blue color
    '@Example: =HEX2RGB("FF0109", "Blue") -> 9; The blue color

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


Public Function RGB2HSL( _
    ByVal redColorInteger As Integer, _
    ByVal greenColorInteger As Integer, _
    ByVal blueColorInteger As Integer, _
    Optional ByVal singleColorNumberOrName As Variant = -1) _
As Variant

    '@Description: This function converts an RGB color value into an HSL color value and returns a string of the HSL value, or optionally a single value from the HSL value.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: redColorInteger is the red value
    '@Param: greenColorInteger is the green value
    '@Param: blueColorInteger is the blue value
    '@Param: singleColorNumberOrName is a number or string that specifies which of the single color values to return. If this is set to 0 or "Hue", the hue value will be returned. If this is set to 1 or "Saturation", the saturation value will be returned. If this is set to 2 or "Lightness", the lightness value will be returned.
    '@Returns: Returns a string with the HSL value of the color
    '@Example: =RGB2HSL(8,64,128) -> "(212.0°, 88.2%, 26.7%)"
    '@Example: =RGB2HSL(8,64,128,0) -> 212
    '@Example: =RGB2HSL(8,64,128,"Saturation") -> .882
    '@Example: =RGB2HSL(8,64,128,2) -> .267

    ' Calculating values needed to calculate HSL
    Dim redPrime As Double
    Dim greenPrime As Double
    Dim bluePrime As Double
    
    redPrime = redColorInteger / 255
    greenPrime = greenColorInteger / 255
    bluePrime = blueColorInteger / 255
    
    Dim colorMax As Double
    Dim colorMin As Double
    
    colorMax = WorksheetFunction.Max(redPrime, greenPrime, bluePrime)
    colorMin = WorksheetFunction.Min(redPrime, greenPrime, bluePrime)
    
    Dim deltaValue As Double
    
    deltaValue = colorMax - colorMin
    
    Dim hueValue As Double
    Dim saturationValue As Double
    Dim lightnessValue As Double
    
    
    ' Calculating Hue
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
    
    
    ' Calculating Lightness
    lightnessValue = (colorMax + colorMin) / 2
    
    
    ' Calculating Saturation
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


Public Function HEX2HSL( _
    ByVal hexColorString As String) _
As String

    '@Description: This function converts a HEX color value into an HSL color value
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: hexColorString is the hex color
    '@Returns: Returns a string with the HSL value of the color
    '@Example: =HEX2HSL("084080") -> "(212.0, 88.2%, 26.7%)"
    '@Example: =HEX2HSL("#084080") -> "(212.0, 88.2%, 26.7%)"

    hexColorString = Replace(hexColorString, "#", "")

    Dim redValue As Integer
    Dim greenValue As Integer
    Dim blueValue As Integer
    
    redValue = CInt(WorksheetFunction.Hex2Dec(Left(hexColorString, 2)))
    greenValue = CInt(WorksheetFunction.Hex2Dec(Mid(hexColorString, 3, 2)))
    blueValue = CInt(WorksheetFunction.Hex2Dec(Right(hexColorString, 2)))

    HEX2HSL = RGB2HSL(redValue, greenValue, blueValue)

End Function


Private Function ModFloat( _
    numerator As Double, _
    denominator As Double) _
As Double

    '@Description: This function performs modulus operations with floats as the Mod operator in VBA does not support floats.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Find out if numerator and denominator are the correct names for Modulo operation
    '@Param: numerator is the left value of the Mod
    '@Param: denominator is the right value of the Mod
    '@Returns: Returns a double with ModFloat operator performed on it
    '@Example: =ModFloat(3.55, 2) -> 1.55

    Dim modValue As Double

    modValue = numerator - Fix(numerator / denominator) * denominator

    If modValue >= -2 ^ -52 Then
        If modValue <= 2 ^ -52 Then
            modValue = 0
        End If
    End If
    
    ModFloat = modValue
    
End Function


Public Function HSL2RGB( _
    ByVal hueValue As Double, _
    ByVal saturationValue As Double, _
    ByVal lightnessValue As Double, _
    Optional ByVal singleColorNumberOrName As Variant = -1) _
As Variant

    '@Description: This function converts an HSL color value into an RGB color value.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: hueValue is the hue value
    '@Param: saturationValue is the saturation value
    '@Param: lightnessValue is the lightness value
    '@Param: singleColorNumberOrName is a number or string that specifies which of the single color values to return. If this is set to 0 or "Red", the red value will be returned. If this is set to 1 or "Green", the green value will be returned. If this is set to 2 or "Blue", the blue value will be returned.
    '@Returns: Returns a string with the RGB value of the color or an individual color value
    '@Example: =HSL2RGB(212, .882, .267) -> "(8, 64, 128)"
    '@Example: =HSL2RGB(212, .882, .267, 0) -> 8
    '@Example: =HSL2RGB(212, .882, .267, "Red") -> 8
    '@Example: =HSL2RGB(212, .882, .267, 1) -> 64
    '@Example: =HSL2RGB(212, .882, .267, "Green") -> 64
    '@Example: =HSL2RGB(212, .882, .267, 2) -> 128
    '@Example: =HSL2RGB(212, .882, .267, "Blue") -> 128

    Dim cValue As Double
    Dim xValue As Double
    Dim mValue As Double
    
    cValue = (1 - Abs(2 * lightnessValue - 1)) * saturationValue
    xValue = cValue * (1 - Abs(ModFloat((hueValue / 60), 2) - 1))
    mValue = lightnessValue - cValue / 2
    
    Dim redValue As Double
    Dim greenValue As Double
    Dim blueValue As Double
    
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


Public Function HSL2HEX( _
    ByVal hueValue As Double, _
    ByVal saturationValue As Double, _
    ByVal lightnessValue As Double) _
As Variant

    '@Description: This function converts an HSL color value into a HEX color value.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Code the formula directly instead of using an additional conversion to speed up the function
    '@Param: hueValue is the hue value
    '@Param: saturationValue is the saturation value
    '@Param: lightnessValue is the lightness value
    '@Returns: Returns a string with the HEX value of the color
    '@Example: =HSL2RGB(212, .882, .267) -> "(8, 64, 128)"

    Dim redValue As Integer
    Dim greenValue As Integer
    Dim blueValue As Integer

    redValue = HSL2RGB(hueValue, saturationValue, lightnessValue, 0)
    greenValue = HSL2RGB(hueValue, saturationValue, lightnessValue, 1)
    blueValue = HSL2RGB(hueValue, saturationValue, lightnessValue, 2)

    HSL2HEX = RGB2HEX(redValue, greenValue, blueValue)

End Function


Public Function RGB2HSV( _
    ByVal redColorInteger As Integer, _
    ByVal greenColorInteger As Integer, _
    ByVal blueColorInteger As Integer, _
    Optional ByVal singleColorNumberOrName As Variant = -1) _
As Variant

    '@Description: This function converts an RGB color value into an HSV color value.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: redColorInteger is the red value
    '@Param: greenColorInteger is the green value
    '@Param: blueColorInteger is the blue value
    '@Param: singleColorNumberOrName is a number or string that specifies which of the single color values to return. If this is set to 0 or "Hue", the hue value will be returned. If this is set to 1 or "Saturation", the saturation value will be returned. If this is set to 2 or "Value", the value value will be returned.
    '@Returns: Returns a string with the RGB value of the color or an individual color value
    '@Example: =RGB2HSV(8, 64, 128) -> "(212.0, 93.8%, 50.2%)"
    '@Example: =RGB2HSV(8, 64, 128, 0) -> 212
    '@Example: =RGB2HSV(8, 64, 128, "Red") -> 212
    '@Example: =RGB2HSV(8, 64, 128, 1) -> .938
    '@Example: =RGB2HSV(8, 64, 128, "Green") -> .938
    '@Example: =RGB2HSV(8, 64, 128, 2) -> .502
    '@Example: =RGB2HSV(8, 64, 128, "Blue") -> .502

    ' Calculating values needed to calculate HSV
    Dim redPrime As Double
    Dim greenPrime As Double
    Dim bluePrime As Double
    
    redPrime = redColorInteger / 255
    greenPrime = greenColorInteger / 255
    bluePrime = blueColorInteger / 255
    
    Dim colorMax As Double
    Dim colorMin As Double
    
    colorMax = WorksheetFunction.Max(redPrime, greenPrime, bluePrime)
    colorMin = WorksheetFunction.Min(redPrime, greenPrime, bluePrime)
    
    Dim deltaValue As Double
    
    deltaValue = colorMax - colorMin
    
    Dim hueValue As Double
    Dim saturationValue As Double
    Dim valueValue As Double

    ' Calculating Hue
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
    
    
    ' Calculating Saturation
    If colorMax = 0 Then
        saturationValue = 0
    Else
        saturationValue = deltaValue / colorMax
    End If
    
    
    ' Calculating Value
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
