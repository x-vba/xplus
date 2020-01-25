Attribute VB_Name = "XPlus"
'The MIT License (MIT)
'Copyright © 2020 Anthony Mancini
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Option Explicit
'@Module: This module contains a set of functions for working with colors



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
    '@Example: =RGB2HSL(8,64,128) -> "(212.0ï¿½, 88.2%, 26.7%)"
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

'@Module: This module contains a set of functions for working with dates and times.



Public Function WEEKDAY_NAME( _
    Optional ByVal dayNumber As Byte _
) As String

    '@Description: This function takes a weekday number and returns the name of the day of the week
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: dayNumber is a number that should be between 1 and 7, with 1 being Sunday and 7 being Saturday. If no dayNumber is given, the value will default to the current day of the week.
    '@Returns: Returns the day of the week as a string
    '@Example: =WEEKDAY_NAME(4) -> Wednesday
    '@Example: To get today's weekday name: =WEEKDAY_NAME()

    If dayNumber = 0 Then
        WEEKDAY_NAME = WeekdayName(Weekday(Now()))
    Else
        WEEKDAY_NAME = WeekdayName(dayNumber)
    End If

End Function


Public Function MONTH_NAME( _
    Optional ByVal monthNumber As Byte _
) As String

    '@Description: This function takes a month number and returns the name of the month
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: monthNumber is a number that should be between 1 and 12, with 1 being January and 12 being December. If no monthNumber is given, the value will default to the current month.
    '@Returns: Returns the month name as a string
    '@Example: =MONTH_NAME(1) -> January
    '@Example: To get today's month name: =MONTH_NAME()

    If monthNumber = 0 Then
        MONTH_NAME = MonthName(Month(Now()))
    Else
        MONTH_NAME = MonthName(monthNumber)
    End If

End Function


Public Function QUARTER( _
    Optional ByVal monthNumberOrName As Variant _
) As Byte
    
    '@Description: This function takes a month as a number and returns the quarter of the year the month resides
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Look further into DatePart function and see if its a better choice for generating the quarter of the year. Also look into adding the month name as well as an option for this function
    '@Param: monthNumberOrName is a number that should be between 1 and 12, with 1 being January and 12 being December
    '@Returns: Returns the quarter of the month as a number
    '@Example: =QUARTER(4) -> 2
    '@Example: =QUARTER(12) -> 2
    '@Example: To get today's quarter: =QUARTER()
    
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


Public Function TIME_CONVERTER( _
    ByVal date1 As Date, _
    Optional ByVal secondsInteger As Integer, _
    Optional ByVal minutesInteger As Integer, _
    Optional ByVal hoursInteger As Integer, _
    Optional ByVal daysInteger As Integer, _
    Optional ByVal monthsInteger As Integer, _
    Optional ByVal yearsInteger As Integer) _
As Date
    
    '@Description: This function takes a date, and then a series of optional arguments for a number of seconds, minutes, hours, days, and years, and then converts the date given to a new date adding in the other date argument values
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: date1 is the origional date that will be converted into a new date
    '@Param: secondsInteger is the number of seconds that will be added
    '@Param: minutesInteger is the number of minutes that will be added
    '@Param: hoursInteger is the number of hours that will be added
    '@Param: daysInteger is the number of days that will be added
    '@Param: monthsInteger is the number of months that will be added
    '@Param: yearsInteger is the number of years that will be added
    '@Returns: Returns a new date with all the date arguments added to it
    '@Note: You can skip earlier date arguments in the function by putting a 0 in place. For example, if we only wanted to changed the month, which is the 5th argument, we can do =TIME_CONVERTER(A1,0,0,0,2) which will add 2 months to the date chosen
    '@Example: Where A1 contains the date 1/1/2000 1:00; =TIME_CONVERTER(A1,60) -> 1/1/2000 1:01
    '@Example: Where A1 contains the date 1/1/2000 1:00; =TIME_CONVERTER(A1,0,5) -> 1/1/2000 1:05
    '@Example: Where A1 contains the date 1/1/2000 1:00; =TIME_CONVERTER(A1,0,0,2) -> 1/1/2000 3:00
    '@Example: Where A1 contains the date 1/1/2000 1:00; =TIME_CONVERTER(A1,0,0,0,4) -> 1/5/2000 1:00
    '@Example: Where A1 contains the date 1/1/2000 1:00; =TIME_CONVERTER(A1,0,0,0,0,1) -> 2/1/2000 1:00
    '@Example: Where A1 contains the date 1/1/2000 1:00; =TIME_CONVERTER(A1,0,0,0,0,0,5) -> 1/1/2005 1:00
    '@Example: Where A1 contains the date 1/1/2000 1:00; =TIME_CONVERTER(A1,60,5,3,10,5,15) -> 6/11/2015 4:06
    
    secondsInteger = Second(date1) + secondsInteger
    minutesInteger = Minute(date1) + minutesInteger
    hoursInteger = Hour(date1) + hoursInteger
    daysInteger = Day(date1) + daysInteger
    monthsInteger = Month(date1) + monthsInteger
    yearsInteger = Year(date1) + yearsInteger
    
    TIME_CONVERTER = DateSerial(yearsInteger, monthsInteger, daysInteger) + TimeSerial(hoursInteger, minutesInteger, secondsInteger)

End Function


Public Function DAYS_OF_MONTH( _
    Optional ByVal monthNumberOrName As Variant, _
    Optional ByVal yearNumber As Integer) _
As Variant

    '@Description: This function takes a month number or month name and returns the number of days in the month. Optionally, a year number can be specified. If no year number is provided, the current year will be used. Finally, note that the month name or number argument is optional and if omitted will use the current month.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: monthNumberOrName is the number or name of the month we want to get the number of days from. If omitted the current month will be used
    '@Param: yearNumber is the year that will be used. If omitted, the current year will be used.
    '@Returns: Returns the number of days in the month and year specified
    '@Example: =DAYS_OF_MONTH() -> 31; Where the current month is January
    '@Example: =DAYS_OF_MONTH(1) -> 31
    '@Example: =DAYS_OF_MONTH("January") -> 31
    '@Example: =DAYS_OF_MONTH(2, 2019) -> 28
    '@Example: =DAYS_OF_MONTH(2, 2020) -> 29

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



'@Module: This module contains a set of functions for gathering information on the environment that Excel is being run on, such as getting the location of the SystemDrive or the location of the temporary folder on the system



Public Function ENVIRONMENT( _
    ByVal environmentVariableNameString As String) _
As String

    '@Description: This function takes a string of the name of an environment variable and returns the value of that EV as a string. A list of EV key/value pairs can be found by using the SET command on the command prompt.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Check out if using Environ() or Environ$() has better performance
    '@Param: environmentVariableNameString is the string of the environment variable name
    '@Returns: Returns a string of the environment variable value associated with that name
    '@Note: A list of Environment Variable Key/Value pairs can be found by using the Set command on the Command Prompt
    '@Example: =ENVIRONMENT("HOMEDRIVE") -> "C:"
    '@Example: =ENVIRONMENT("OS") -> "Windows_NT"
    
    Dim WshShell As Object
    Set WshShell = CreateObject("Wscript.Shell")

    ENVIRONMENT = WshShell.ExpandEnvironmentStrings("%" & environmentVariableNameString & "%")

End Function



'@Module: This module contains a set of functions for gathering info about the excel workbook that is currently open



Private Function GetActiveWorkbookPath() As Variant

    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    
    GetActiveWorkbookPath = filePath

End Function


Public Function FILE_CREATION_TIME( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the file creation time of the file specified in the file path argument. If no file path is specified, the current Excel workbook is used. Also, if a full path isn't used, a path relative to the folder the workbook resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system
    '@Returns: Returns the file creation time of the file as a string
    '@Example: =FILE_CREATION_TIME() -> "1/1/2020 1:23:45 PM"
    '@Example: =FILE_CREATION_TIME("C:\hello\world.txt") -> "1/1/2020 5:55:55 PM"
    '@Example: =FILE_CREATION_TIME("vba.txt") -> "12/25/2000 1:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook this function resides in

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


Public Function FILE_LAST_MODIFIED_TIME( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the file last modified time of the file specified in the file path argument. If no file path is specified, the current Excel workbook is used. Also, if a full path isn't used, a path relative to the folder the workbook resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system
    '@Returns: Returns the file last modified time of the file as a string
    '@Example: =FILE_CREATION_TIME() -> "1/1/2020 2:23:45 PM"
    '@Example: =FILE_CREATION_TIME("C:\hello\world.txt") -> "1/1/2020 7:55:55 PM"
    '@Example: =FILE_CREATION_TIME("vba.txt") -> "12/25/2000 3:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook this function resides in

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


Public Function FILE_DRIVE( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the drive of the file specified in the file path argument. If no file path is specified, the current Excel workbook is used. Also, if a full path isn't used, a path relative to the folder the workbook resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system
    '@Returns: Returns the file drive of the file as a string
    '@Example: =FILE_CREATION_TIME() -> "A:"; Where the current workbook resides on the A: drive
    '@Example: =FILE_CREATION_TIME("C:\hello\world.txt") -> "C:"
    '@Example: =FILE_CREATION_TIME("vba.txt") -> "B:"; Where "vba.txt" resides in the same folder as the workbook this function resides in, and where the workbook resides in the B: drive

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


Public Function FILE_NAME( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the name of the file specified in the file path argument. If no file path is specified, the current Excel workbook is used. Also, if a full path isn't used, a path relative to the folder the workbook resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system
    '@Returns: Returns the name of the file as a string
    '@Example: =FILE_NAME() -> "MyWorkbook.xlsm"
    '@Example: =FILE_NAME("C:\hello\world.txt") -> "world.txt"
    '@Example: =FILE_NAME("vba.txt") -> "vba.txt"; Where "vba.txt" resides in the same folder as the workbook this function resides in

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


Public Function FILE_FOLDER( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the path of the folder of the file specified in the file path argument. If no file path is specified, the current Excel workbook is used. Also, if a full path isn't used, a path relative to the folder the workbook resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system
    '@Returns: Returns the path of the folder of the file as a string
    '@Example: =FILE_FOLDER() -> "C:\my_excel_files"
    '@Example: =FILE_FOLDER("C:\hello\world.txt") -> "C:\hello"
    '@Example: =FILE_FOLDER("vba.txt") -> "C:\my_excel_files"; Where "vba.txt" resides in the same folder as the workbook this function resides in

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


Public Function FILE_PATH( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the path of the file specified in the file path argument. If no file path is specified, the current Excel workbook is used. Also, if a full path isn't used, a path relative to the folder the workbook resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system
    '@Returns: Returns the path of the file as a string
    '@Example: =FILE_PATH() -> "C:\my_excel_files\MyWorkbook.xlsx"
    '@Example: =FILE_PATH("C:\hello\world.txt") -> "C:\hello\world.txt"
    '@Example: =FILE_PATH("vba.txt") -> "C:\hello\world.txt"; Where "vba.txt" resides in the same folder as the workbook this function resides in

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


Public Function FILE_SIZE( _
    Optional ByVal filePath As String, _
    Optional ByVal byteSize As String) _
As Double

    '@Description: This function returns the file size of the file specified in the file path argument, with the option to set if the file size is returned in Bytes, Kilobytes, Megabytes, or Gigabytes. If no file path is specified, the current Excel workbook is used. Also, if a full path isn't used, a path relative to the folder the workbook resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system
    '@Param: byteSize is a string of value "KB", "MB", or "GB"
    '@Returns: Returns the size of the file as a Double
    '@Example: =FILE_SIZE() -> 1024
    '@Example: =FILE_SIZE(,"KB") -> 1
    '@Example: =FILE_SIZE("vba.txt", "KB") -> 0.25; Where "vba.txt" resides in the same folder as the workbook this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim totalBytes As Double
    
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


Public Function FILE_TYPE( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the file type of the file specified in the file path argument. If no file path is specified, the current Excel workbook is used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system
    '@Returns: Returns the file type of the file as a string
    '@Example: FILE_TYPE() -> "Microsoft Excel Macro-Enabled Worksheet"
    '@Example: FILE_TYPE("C:\hello\world.txt") -> "Text Document"
    '@Example: FILE_TYPE("vba.txt") -> "Text Document"; Where "vba.txt" resides in the same folder as the workbook this function resides in

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


Public Function FILE_EXTENSION( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the extension of the file specified in the file path argument. If no file path is specified, the current Excel workbook is used. Also, if a full path isn't used, a path relative to the folder the workbook resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system
    '@Returns: Returns the extension of the file as a string
    '@Example: =FILE_EXTENSION() = "xlsx"
    '@Example: =FILE_EXTENSION("C:\hello\world.txt") -> "txt"
    '@Example: =FILE_EXTENSION("vba.txt") -> "txt"; Where "vba.txt" resides in the same folder as the workbook this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim fileName As String
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


Public Function READ_FILE( _
    ByVal filePath As String, _
    Optional ByVal lineNumber As Integer) _
As String

    '@Description: This function reads the file specified in the file path argument and returns it's contents. Optionally, a line number can be specified so that only a single line is read. If no file path is specified, the current Excel workbook is used. Also, if a full path isn't used, a path relative to the folder the workbook resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system
    '@Param: lineNumber is the number of the line that will be read, and if left blank all the file contents will be read. Note that the first line starts at line number 1.
    '@Returns: Returns the contents of the file as a string
    '@Warning: This function may run very slowly when running it on large files. Also, for files that are not in text format (such as compressed zip files) this file contents returned will not be in a usable format.
    '@Example: =READ_FILE("C:\hello\world.txt") -> "Hello" World
    '@Example: =READ_FILE("vba.txt") -> "This is my VBA text file"; Where "vba.txt" resides in the same folder as the workbook this function resides in
    '@Example: =READ_FILE("multline.txt", 1) -> "This is line 1";
    '@Example: =READ_FILE("multline.txt", 2) -> "This is line 2";

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim fileName As String
    Dim fileStream As Object
    
    ' Checking if the file exists in the current directory, and then if it
    ' exists in the path specified, and if it doesn't exist in either, returns
    ' a "#FileDoesntExist!"
    If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
        filePath = ThisWorkbook.Path & "\" & filePath
    ElseIf FSO.FileExists(filePath) Then
        filePath = filePath
    Else
        READ_FILE = "#FileDoesntExist!"
    End If
    
    Set fileStream = FSO.GetFile(filePath)
    Set fileStream = fileStream.OpenAsTextStream(1, -2)
    
    
    ' If lineNumber is positive, read a line, else read the whole contents
    If lineNumber > 0 Then
        Dim fileLinesArray() As String
        
        fileLinesArray = SPLIT(fileStream.ReadAll(), vbCrLf)
        READ_FILE = fileLinesArray(lineNumber)
    Else
        READ_FILE = fileStream.ReadAll()
    End If

End Function


Public Function WRITE_FILE( _
    ByVal filePath As String, _
    ByVal fileText As String, _
    Optional ByVal appendModeFlag As Boolean) _
As String

    '@Description: This function creates and writes to the file specified in the file path argument. If no file path is specified, the current Excel workbook is used. Also, if a full path isn't used, a path relative to the folder the workbook resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system
    '@Param: fileText is the text that will be written to the file
    '@Param: appendModeFlag is a Boolean value that if set to TRUE will append to the existing file instead of creating a new file and writing over the contents.
    '@Returns: Returns a message stating the file written
    '@Warning: Be careful when writing files, as misuse of this function can results in files being overwritten accidently as well as creating large numbers of files if misused.
    '@Example: =WRITE_FILE("hello.txt","Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim fileName As String
    Dim fileStream As Object
    
    
    ' Checking if the folder exists if the path is an absolute path
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
    
    
    ' Writing to the file
    Set fileStream = FSO.CreateTextFile(filePath, Not appendModeFlag)
    fileStream.Write fileText
    
    WRITE_FILE = "Successfully wrote to: " & filePath

End Function


Public Function PATH_JOIN( _
    ParamArray pathArray() As Variant) _
As String

    '@Description: This function combines multiple strings or a range of values into a file path by placing seperators "\" between the arguments
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: pathArray is an array of ranges and strings that will be combined
    '@Returns: Returns a string with the combined file path
    '@Example: =PATH_JOIN(A1:A3) -> "C:\hello\world.txt"
    '@Example: =PATH_JOIN("C:", "hello", "world.txt") -> "C:\hello\world.txt"

    Dim individualPath As Variant
    Dim combinedPath As String
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

'@Module: This module contains a set of basic mathematical functions where those functions don't already exist in the base Excel functions



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
    '@Returns: Returns a the linearly interpolated number between the two points
    '@Example: =INTERPOLATE_NUMBER(10,20,0.5) -> 15
    '@Example: =INTERPOLATE_NUMBER(16,124,0.64) -> 85.12

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
    '@Returns: Returns a the linearly interpolated number between the two points
    '@Example: =INTERPOLATE(10,20,0.5) -> 15
    '@Example: =INTERPOLATE(16,124,0.64) -> 85.12

    INTERPOLATE_PERCENT = (interpolationNumber - startingNumber) / (endingNumber - startingNumber)

End Function


'@Module: This module contains a set of functions for getting the UserName and ComputerName and other values from the Network Module.



Public Function USER_NAME() As String

    '@Description: This function takes no arguments and returns a string of the USERNAME of the computer
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the username
    '@Example: =USER_NAME() -> "Anthony"

    Dim WshNetwork As Object
    Set WshNetwork = CreateObject("WScript.Network")

    USER_NAME = WshNetwork.UserName

End Function

Public Function USER_DOMAIN() As String

    '@Description: This function takes no arguments and returns a string of the USERDOMAIN of the computer
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the user domain of the computer
    '@Example: =USER_DOMAIN() -> "DESKTOP-XYZ1234"

    Dim WshNetwork As Object
    Set WshNetwork = CreateObject("WScript.Network")

    USER_DOMAIN = WshNetwork.UserDomain

End Function

Public Function COMPUTER_NAME() As String

    '@Description: This function takes no arguments and returns a string of the COMPUTERNAME of the computer
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the computer name of the computer
    '@Example: =COMPUTER_NAME() -> "DESKTOP-XYZ1234"

    Dim WshNetwork As Object
    Set WshNetwork = CreateObject("WScript.Network")

    COMPUTER_NAME = WshNetwork.ComputerName

End Function

'@Module: This module contains a set of functions for getting properties from Ranges, Worksheets, and Workbooks.



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

'@Module: This module contains a set of functions for generating and sampling random data



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


'@Module: This module contains a set of functions for manipulating and working with ranges of cells



Public Function FIRST_UNIQUE( _
    ByVal range1 As Range, _
    ByVal rangeArray As Range) _
As Boolean

    '@Description: This function takes a single cell and an large range of cells and returns TRUE if the cell selected is the first unique value in the larger array of cells, and returns FALSE if it is not the first unique value.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the range we want to check if the value is the first unique value in the larger range
    '@Param: rangeArray is the group of cells we are checking to see if range1 is the first unique occurance in the range
    '@Returns: Returns TRUE if the cell selected is the first unique value in the range array, and FALSE if it isn't
    '@Example: =FIRST_UNIQUE(A1,$A$1:$A$10) -> TRUE, where A1 is the first unique occurance of the word "Hello" in the range array
    '@Example: =FIRST_UNIQUE(A5,$A$1:$A$10) -> FALSE, where A5 is the second unique occurance of the word "Hello" in the range array

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

    '@Description: This function takes a single cell and an large range of cells and sorts the cells in ascending or descending order.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the range containing a single cell that we want to sort
    '@Param: rangeArray is the group of cells we are sorting
    '@Param: descendingFlag is a Boolean value that if set to TRUE will set the sort to Descending
    '@Returns: Returns the value of the cell sorted
    '@Example: =SORT_RANGE(A1,$A$1:$A$3) -> 1, where A1="3", A2="1", A3="4", A4="2"
    '@Example: =SORT_RANGE(A1,$A$1:$A$3, TRUE) -> 4, where A1="3", A2="1", A3="4", A4="2"

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

    '@Description: This function takes a single cell and an large range of cells and reverses all values in the range.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: range1 is the range we want to be reversed in order of the larger range
    '@Param: rangeArray is the group of cells we are reversing the order of
    '@Returns: Returns the value of the cell in the reversed position
    '@Example: =REVERSE_RANGE(A1,$A$1:$A$3) -> "C", where A1="A", A2="B", A3="C"
    '@Example: =REVERSE_RANGE(A2,$A$1:$A$3) -> "B", where A1="A", A2="B", A3="C"
    '@Example: =REVERSE_RANGE(A3,$A$1:$A$3) -> "A", where A1="A", A2="B", A3="C"

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

    '@Description: This function takes 2 ranges, a column range which will be filled in with data in the row range, allowing for easily convering a row of data into a column of data
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: columnRangeArray is a range that will be populated with data from a rowRangeArray
    '@Param: rowRangeArray is a range that will be used to populate the columnRangeArray
    '@Returns: Returns the value at the same location in the rowRangeArray
    '@Example: =COLUMNIFY(A1:A2,B1:C1) -> "B"; Where this function resides in cell A1 and where B1="B" and C1="C"
    '@Example: =COLUMNIFY(A1:A2,B1:C1) -> "C"; Where this function resides in cell A2 and where B1="B" and C1="C"

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

    '@Description: This function takes 2 ranges, a row range which will be filled in with data in the column range, allowing for easily convering a column of data into a row of data
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: rowRangeArray is a range that will be populated with data from a columnRangeArray
    '@Param: columnRangeArray is a range that will be used to populate the rowRangeArray
    '@Returns: Returns the value at the same location in the columnRangeArray
    '@Example: =ROWIFY(B1:C1,A1:A2) -> "A1"; Where this function resides in cell B1 and where A1="A1" and A2="A2"
    '@Example: =ROWIFY(B1:C1,A1:A2) -> "A2"; Where this function resides in cell C1 and where A1="A1" and A2="A2"

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
    '@Example: =SUMHIGH(A1:A4, 2) -> 7; Where A1=1, A2=2, A3=3, A4=4
    '@Example: =SUMHIGH(A1:A4, 3) -> 9; Where A1=1, A2=2, A3=3, A4=4

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
    '@Example: =SUMLOW(A1:A4, 2) -> 3; Where A1=1, A2=2, A3=3, A4=4
    '@Example: =SUMLOW(A1:A4, 3) -> 6; Where A1=1, A2=2, A3=3, A4=4

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
    '@Example: =AVERAGEHIGH(A1:A4, 2) -> 3.5; Where A1=1, A2=2, A3=3, A4=4
    '@Example: =AVERAGEHIGH(A1:A4, 3) -> 3; Where A1=1, A2=2, A3=3, A4=4

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
    '@Example: =AVERAGEHIGH(A1:A4, 2) -> 3.5; Where A1=1, A2=2, A3=3, A4=4
    '@Example: =AVERAGEHIGH(A1:A4, 3) -> 3; Where A1=1, A2=2, A3=3, A4=4

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

'@Module: This module contains a set of functions for performing Regular Expressions, which are a type of string pattern matching. For more info on Regular Expression Pattern matching, please check "https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference"



Public Function REGEX_SEARCH( _
    ByVal string1 As String, _
    ByVal stringPattern As String, _
    Optional ByVal globalFlag As Boolean, _
    Optional ByVal ignoreCaseFlag As Boolean, _
    Optional ByVal multilineFlag As Boolean) _
As String

    '@Description: This function takes a string that we will perform the Regular Expression on and a Regular Expression string pattern, and returns the first value of the matched string. This function also contains optional arguments for various Regular Expression flags.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the regex will be performed on
    '@Param: stringPattern is a string of the regex pattern
    '@Param: globalFlag is a boolean value that if set TRUE will perform a global search
    '@Param: ignoreCaseFlag is a boolean value that if set TRUE will perform a case insensitive search
    '@Param: multilineFlag is a boolean value that if set TRUE will perform a mulitline search
    '@Returns: Returns a string of the regex value that is found
    '@Example: =REGEX_SEARCH("Hello World","[a-z]{2}\s[W]") -> "lo W";

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

Public Function REGEX_TEST( _
    ByVal string1 As String, _
    ByVal stringPattern As String, _
    Optional ByVal globalFlag As Boolean, _
    Optional ByVal ignoreCaseFlag As Boolean, _
    Optional ByVal multilineFlag As Boolean) _
As Boolean

    '@Description: This function takes a string that we will perform the Regular Expression on and a Regular Expression string pattern, and returns TRUE if the pattern is found in the string. This function also contains optional arguments for various Regular Expression flags.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the regex will be performed on
    '@Param: stringPattern is a string of the regex pattern
    '@Param: globalFlag is a boolean value that if set TRUE will perform a global search
    '@Param: ignoreCaseFlag is a boolean value that if set TRUE will perform a case insensitive search
    '@Param: multilineFlag is a boolean value that if set TRUE will perform a mulitline search
    '@Returns: Returns TRUE if the regex value that is found, or FALSE if it isn't
    '@Example: =REGEX_TEST("Hello World","[a-z]{2}\s[W]") -> TRUE;

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

Public Function REGEX_REPLACE( _
    ByVal string1 As String, _
    ByVal stringPattern As String, _
    ByVal replacementString As String, _
    Optional ByVal globalFlag As Boolean, _
    Optional ByVal ignoreCaseFlag As Boolean, _
    Optional ByVal multilineFlag As Boolean) _
As String

    '@Description: This function takes a string that we will perform the Regular Expression on, a Regular Expression string pattern, and a string that we will replace if the pattern is found, and returns a new string with the replacement string in place of the pattern. This function also contains optional arguments for various Regular Expression flags.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the regex will be performed on
    '@Param: stringPattern is a string of the regex pattern
    '@Param: replacementString is a string that will be replaced if the pattern is found
    '@Param: globalFlag is a boolean value that if set TRUE will perform a global search
    '@Param: ignoreCaseFlag is a boolean value that if set TRUE will perform a case insensitive search
    '@Param: multilineFlag is a boolean value that if set TRUE will perform a mulitline search
    '@Returns: Returns a new string with the replaced string value
    '@Example: =REGEX_REPLACE("Hello World","[W][a-z]{4}", "VBA") -> "Hello VBA"

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


'@Module: This module contains a set of basic functions for manipulating strings



Public Function CAPITALIZE( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the same string with the first character capitalized and all other characters lowercased
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the capitalization will be performed on
    '@Returns: Returns a new string with the first character capitalized and all others lowercased
    '@Example: =CAPITALIZE("hello World") -> "Hello world"

    CAPITALIZE = UCase(Left(string1, 1)) & LCase(Mid(string1, 2))
    
End Function


Public Function LEFT_FIND( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the left of the first search string found within string1. Similar to Excel's built-in =SEARCH() function, this function is case-sensitive. For a case-insensitive version of this function, see =LEFT_SEARCH().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the left of the first search string within string1
    '@Example: =LEFT_FIND("Hello World","r") -> "Hello Wo"
    '@Example: =LEFT_FIND("Hello World","R") -> "#VALUE!"; Since string1 does not contain "R" in it.

    LEFT_FIND = Left(string1, InStr(1, string1, searchString) - 1)

End Function


Public Function RIGHT_FIND( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the right of the last search string found within string 1. Similar to Excel's built-in =SEARCH() function, this function is case-sensitive. For a case-insensitive version of this function, see =RIGHT_SEARCH().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the right of the last search string within string1
    '@Example: =RIGHT_FIND("Hello World","o") -> "rld"
    '@Example: =RIGHT_FIND("Hello World","O") -> "#VALUE!"; Since string1 does not contain "O" in it.

    RIGHT_FIND = Right(string1, Len(string1) - InStrRev(string1, searchString))

End Function


Public Function LEFT_SEARCH( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the left of the first search string found within string1. Similar to Excel's built-in =FIND() function, this function is NOT case-sensitive (it's case-insensitive). For a case-sensitive version of this function, see =LEFT_FIND().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the left of the first search string within string1
    '@Example: =LEFT_SEARCH("Hello World","r") -> "Hello Wo"
    '@Example: =LEFT_SEARCH("Hello World","R") -> "Hello Wo"

    LEFT_SEARCH = Left(string1, InStr(1, string1, searchString, vbTextCompare) - 1)

End Function


Public Function RIGHT_SEARCH( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the right of the last search string found within string 1. Similar to Excel's built-in =FIND() function, this function is NOT case-sensitive (it's case-insensitive). For a case-sensitive version of this function, see =RIGHT_FIND().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the right of the last search string within string1
    '@Example: =RIGHT_SEARCH("Hello World","o") -> "rld"
    '@Example: =RIGHT_SEARCH("Hello World","O") -> "rld"

    RIGHT_SEARCH = Right(string1, Len(string1) - InStrRev(string1, searchString, Compare:=vbTextCompare))

End Function


Public Function SUBSTR( _
    ByVal string1 As String, _
    ByVal startCharacterNumber As Integer, _
    ByVal endCharacterNumber As Integer) _
As String

    '@Description: This function takes a string and a starting character number and ending character number, and returns the substring between these two numbers. The total number of characters returned will be endCharacterNumber - startCharacterNumber.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that we will get a substring from
    '@Param: startCharacterNumber is the character number of the start of the substring, with 1 being the first character in the string
    '@Param: endCharacterNumber is the character number of the end of the substring
    '@Returns: Returns a substring between the two numbers.
    '@Example: =SUBSTR("Hello World",2,6) -> "ello"

    SUBSTR = Mid(string1, startCharacterNumber, endCharacterNumber - startCharacterNumber)

End Function


Public Function SUBSTR_SEARCH( _
    ByVal string1 As String, _
    ByVal leftSearchString As String, _
    ByVal rightSearchString As String, _
    Optional ByVal noninclusiveFlag As Boolean) _
As String

    '@Description: This function takes a string and a left character and right character, and returns a substring between those two characters. The left character will find the first matching character starting from the left, and the right character will find the first matching character starting from the right. Finally, and optional final parameter can be set to TRUE to make the substring noninclusive of the two searched characters.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Currently this function only supports single characters, but in the future include the ability to use substrings of longer length than 1
    '@Param: string1 is the string that we will get a substring from
    '@Param: leftSearchString is the string that will be searched from the left
    '@Param: rightSearchString is the string that will be searched from the right
    '@Param: noninclusiveFlag is an optional paramater that if set to TRUE will result in the substring not including the left and right searched chracters
    '@Returns: Returns a substring between the two chracters.
    '@Example: =SUBSTR_SEARCH("Hello World", "e", "o") -> "ello Wo"
    '@Example: =SUBSTR_SEARCH("Hello World", "e", "o", TRUE) -> "llo W"
    '@Example: =SUBSTR_SEARCH("Phone Number: 123 456 789 - Name: John Doe", ":", "-", TRUE) -> " 123 456 789 "

    Dim leftCharacterNumber As Integer
    Dim rightCharacterNumber As Integer
    
    leftCharacterNumber = InStr(1, string1, leftSearchString)
    rightCharacterNumber = InStrRev(string1, rightSearchString)
    
    If noninclusiveFlag = True Then
        leftCharacterNumber = leftCharacterNumber + 1
        rightCharacterNumber = rightCharacterNumber - 1
    End If
    
    SUBSTR_SEARCH = Mid(string1, leftCharacterNumber, rightCharacterNumber - leftCharacterNumber + 1)

End Function
    
    
Public Function REPEAT( _
    ByVal string1 As String, _
    ByVal numberOfRepeats As Integer) _
As String

    '@Description: This function takes a string and a left character and right character, and returns a substring between those two characters. The left character will find the first matching character starting from the left, and the right character will find the first matching character starting from the right. Finally, and optional final parameter can be set to TRUE to make the substring noninclusive of the two searched characters.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Currently this function only supports single characters, but in the future include the ability to use substrings of longer length than 1
    '@Param: string1 is the string that will be repeated
    '@Param: numberOfRepeats is the number of times string1 will be repeated
    '@Returns: Returns a string repeated multiple times based on the numberOfRepeats
    '@Example: =REPEAT("=",10) -> "=========="

    Dim i As Integer
    Dim combinedString As String

    For i = 1 To numberOfRepeats
        combinedString = combinedString & string1
    Next

    REPEAT = combinedString

End Function


Public Function FORMATTER( _
    ByVal formatString As String, _
    ParamArray textArray() As Variant) _
As String

    '@Description: This function takes a formatter string and then an array of ranges or strings, and replaces the format placeholders with the values in the range or strings. The format needed is "{1} - {2}" where the "{1}" and "{2}" will be replaced with the values given in the text array.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: formatString is the string that will be used as the format and which will be replaced with the individual strings
    '@Param: textArray are the ranges or strings that will be placed within the slots of the format string
    '@Returns: Returns a new string with the individual strings in the placeholder slots of the format string
    '@Example: =FORMATTER("Hello {1}", "World") -> "Hello World"
    '@Example: =FORMATTER("{1} {2}", "Hello", "World") -> "Hello World"
    '@Example: =FORMATTER("{1}.{2}@{3}", "FirstName", "LastName", "email.com") -> "FirstName.LastName@email.com"
    '@Example: =FORMATTER("{1}.{2}@{3}", A1:A3) -> "FirstName.LastName@email.com"; where A1="FirstName", A2="LastName", and A3="email.com"
    '@Example: =FORMATTER("{1}.{2}@{3}", A1, A2, A3) -> "FirstName.LastName@email.com"; where A1="FirstName", A2="LastName", and A3="email.com"

    Dim i As Byte
    Dim individualTextItem As Variant
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


Public Function ZFILL( _
    ByVal string1 As String, _
    ByVal fillLength As Byte, _
    Optional ByVal fillCharacter As String, _
    Optional ByVal rightToLeftFlag As Boolean) _
As String

    '@Description: This function pads zeros to the left of a string until the string is at least the length of the fill length. Optional parameters can be used to pad with a different character than 0, and to pad from right to left instead of from the default left to right.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be filled
    '@Param: fillLength is the length that string1 will be padded to. In cases where string1 is of greater length than this argument, no padding will occur.
    '@Param: fillCharacter is an optional string that will change the character that will be padded with
    '@Param: rightToLeftFlag is a Boolean parameter that if set to TRUE will result in padding from right to leftt instead of left to right
    '@Returns: Returns a new padded string of the length of specified by fillLength at minimum
    '@Example: =ZFILL(123, 5) -> "00123"
    '@Example: =ZFILL(5678, 5) -> "05678"
    '@Example: =ZFILL(12345678, 5) -> "12345678"
    '@Example: =ZFILL(123, 5, "X") -> "XX123"
    '@Example: =ZFILL(123, 5, "X", TRUE) -> "123XX"

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


Public Function TEXT_JOIN( _
    ByVal rangeArray As Range, _
    Optional ByVal delimiterCharacter As String, _
    Optional ByVal ignoreEmptyCellsFlag As Boolean) _
As String

    '@Description: This function takes a range of cells and combines all the text together, optionally allowing a character delimiter between all the combined strings, and optionally allowing blank cells to be ignored when combining the text. Finally note that this function is very similar to the TEXTJOIN function available in Excel 2019
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: rangeArray is the range with all the strings we want to combine
    '@Param: delimiterCharacter is an optional character that will be used as the delimiter between the combined text. By default no delimiter character will be used.
    '@Param: ignoreEmptyCellsFlag if set to TRUE will skip combining empty cells into the combined string, and is useful when specifying a delimiter so that the delimiter does not repeat for empty cells.
    '@Returns: Returns a new combined string containing the strings in the range delimited by the delimiter character.
    '@Example: Where A1:A3 contains ["1", "2", "3"] =TEXT_JOIN(A1:A3) -> "123"
    '@Example: Where A1:A3 contains ["1", "2", "3"] =TEXT_JOIN(A1:A3, "--") -> "1--2--3"
    '@Example: Where A1:A3 contains ["1", "", "3"] =TEXT_JOIN(A1:A3, "--") -> "1----3"
    '@Example: Where A1:A3 contains ["1", "", "3"] =TEXT_JOIN(A1:A3, "-", TRUE) -> "1--3"

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


Public Function SPLIT_TEXT( _
    ByVal string1 As String, _
    ByVal substringNumber As Integer, _
    Optional ByVal delimiterString As String) _
As String
    
    '@Description: This function takes a string and a number, splits the string by the space characters, and returns the substring in the position of the number specified in the second argument.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be split and a substring returned
    '@Param: substringNumber is the number of the substring that will be chosen
    '@Param: delimiterString is an optional parameter that can be used to specify a different delimiter
    '@Returns: Returns a substring of the split text in the location specified
    '@Example: =SPLIT_TEXT("Hello World", 1) -> "Hello"
    '@Example: =SPLIT_TEXT("Hello World", 2) -> "World"
    '@Example: =SPLIT_TEXT("One Two Three", 2) -> "Two"
    '@Example: =SPLIT_TEXT("One-Two-Three", 2, "-") -> "Two"
    
    If delimiterString = "" Then
        SPLIT_TEXT = SPLIT(string1, " ")(substringNumber - 1)
    Else
        SPLIT_TEXT = SPLIT(string1, delimiterString)(substringNumber - 1)
    End If

End Function


Public Function COUNT_WORDS( _
    ByVal string1 As String, _
    Optional ByVal delimiterString As String) _
As Integer

    '@Description: This function takes a string and returns the number of words in the string
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Note: If the number given is higher than the number of words, its possible that the string contains excess whitespace. Try using the =TRIM() function first to remove the excess whitespace
    '@Param: string1 is the string whose number of words will be counted
    '@Param: delimiterString is an optional parameter that can be used to specify a different delimiter
    '@Returns: Returns a substring of the split text in the location specified
    '@Example: =COUNT_WORDS("Hello World") -> 2
    '@Example: =COUNT_WORDS("One Two Three") -> 3
    '@Example: =COUNT_WORDS("One-Two-Three", "-") -> 3

    Dim stringArray() As String

    If delimiterString = "" Then
        stringArray = SPLIT(string1, " ")
    Else
        stringArray = SPLIT(string1, delimiterString)
    End If
    
    COUNT_WORDS = UBound(stringArray) - LBound(stringArray) + 1

End Function


Public Function CAMEL_CASE( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the same string in camel case, removing all the spaces.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be camel cased
    '@Returns: Returns a new string in camel case, where the first character of the first word is lowercase, and uppercased for all other words
    '@Example: =CAMEL_CASE("Hello World") -> "helloWorld"
    '@Example: =CAMEL_CASE("One Two Three") -> "oneTwoThree"

    Dim i As Integer
    Dim stringArray() As String
    
    stringArray = SPLIT(string1, " ")
    stringArray(0) = LCase(stringArray(0))
    
    For i = 1 To (UBound(stringArray) - LBound(stringArray))
        stringArray(i) = UCase(Left(stringArray(i), 1)) & LCase(Mid(stringArray(i), 2))
    Next
    
    CAMEL_CASE = Join(stringArray, "")

End Function


Public Function KEBAB_CASE( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the same string in kebab case.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be kebab cased
    '@Returns: Returns a new string in kebab case, where all letters are lowercase and seperated by a "-" character
    '@Example: =KEBAB_CASE("Hello World") -> "hello-world"
    '@Example: =KEBAB_CASE("One Two Three") -> "one-two-three"

    KEBAB_CASE = LCase(Join(SPLIT(string1, " "), "-"))

End Function


Public Function REMOVE_CHARACTERS( _
    ByVal string1 As String, _
    ParamArray removedCharacters() As Variant) _
As String

    '@Description: This function takes a string and either another string or mutliple strings and removes all characters from the first string that are in the second string.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Consider adding a Boolean flag that will make non-case sensitive replacements
    '@Note: This function is case sensitive. If you want to remove the "H" from "Hello World" you would need to use "H" as a removed character, not "h".
    '@Param: string1 is the string that will have characters removed
    '@Param: removedCharacters is an array of strings that will be removed from string1
    '@Returns: Returns the origional string with characters removed
    '@Example: =REMOVE_CHARACTERS("Hello World", "l") -> "Heo Word"
    '@Example: =REMOVE_CHARACTERS("Hello World", "lo") -> "He Wrd"
    '@Example: =REMOVE_CHARACTERS("Hello World", "l", "o") -> "He Wrd"
    '@Example: =REMOVE_CHARACTERS("Hello World", "lod") -> "He Wr"
    '@Example: =REMOVE_CHARACTERS("One Two Three", "o", "t") -> "One Two Three"; Nothing is replaced since this function is case sensitive
    '@Example: =REMOVE_CHARACTERS("One Two Three", "O", "T") -> "ne wo hree"

    Dim i As Integer
    Dim individualCharacter As Variant
    
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


Private Function NumberOfUppercaseLetters( _
    ByVal string1 As String) _
As Integer

    '@Description: This function returns the number of uppercase letter found within a string based on the ASCII character code range for uppercase letters
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose uppercase letters will be counted
    '@Returns: Returns the number of uppercase letters

    Dim i As Integer
    Dim numberOfUppercase As Integer
    
    For i = 1 To Len(string1)
        If Asc(Mid(string1, i, 1)) >= 65 Then
            If Asc(Mid(string1, i, 1)) <= 90 Then
                numberOfUppercase = numberOfUppercase + 1
            End If
        End If
    Next
    
    NumberOfUppercaseLetters = numberOfUppercase

End Function


Public Function COMPANY_CASE( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and uses an algorithm to return the string in Company Case. The standard =PROPER() function in Excel will not capitalize company names properly, as it only capitalizes based on space characters, so a name like "j.p. morgan" will be incorrectly formatted as "J.p. Morgan" instead of the correct "J.P. Morgan". Additionally =PROPER() may incorrectly lowercase company abbreviations, such as the last "H" in "GmbH", as =PROPER() returns "Gmbh" instead of the correct "GmbH". This function attempts to adjust for these issues when a string is a company name.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Warning: There is no perfect algorithm for correctly formatting company names, and while this function can give better performance for correct formatting when compared to =PROPER(), if the performance of this function isn't as accurate as one needs, another solution would be to try Partial Lookup functions in the String Metrics Module and compare that to a known list of well formatted company strings.
    '@Param: string1 is the string that will be formatted
    '@Returns: Returns the origional string in a Company Case format
    '@Example: =COMPANY_CASE("hello world") -> "Hello World"
    '@Example: =COMPANY_CASE("x.y.z company & co.") -> "X.Y.Z Company & Co."
    '@Example: =COMPANY_CASE("x.y.z plc") -> "X.Y.Z PLC"
    '@Example: =COMPANY_CASE("one company gmbh") -> "One Company GmbH"
    '@Example: =COMPANY_CASE("three company s. en n.c.") -> "Three Company S. en N.C."
    '@Example: =COMPANY_CASE("FOUR COMPANY SPOL S.R.O.") -> "Four Company spol s.r.o."
    '@Example: =COMPANY_CASE("five company bvba") -> "Five Company BVBA"

    Dim i As Integer
    Dim k As Integer
    Dim origionalString As String
    Dim stringArray() As String
    Dim splitCharacters As String
    
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
    
    
    ' Checking the final words in the string to see if they are one of the
    ' company abbreviation strings, and if it is, replace the ending with
    ' the correct cases of the company abbreviation
    Dim companyAbbreviationArray() As String
    companyAbbreviationArray = SPLIT("AB|AG|GmbH|LLC|LLP|NV|PLC|SA|A. en P.|ACE|AD|AE|AL|AmbA|ANS|ApS|AS|ASA|AVV|BVBA|CA|CVA|d.d.|d.n.o.|d.o.o.|DA|e.V.|EE|EEG|EIRL|ELP|EOOD|EPE|EURL|GbR|GCV|GesmbH|GIE|HB|hf|IBC|j.t.d.|k.d.|k.d.d.|k.s.|KA/S|KB|KD|KDA|KG|KGaA|KK|Kol. SrK|Kom. SrK|LDC|Ltï¿½e.|NT|OE|OHG|Oy|OYJ|Oï¿½|PC Ltd|PMA|PMDN|PrC|PT|RAS|S. de R.L.|S. en N.C.|SA de CV|SAFI|SAS|SC|SCA|SCP|SCS|SENC|SGPS|SK|SNC|SOPARFI|sp|Sp. z.o.o.|SpA|spol s.r.o.|SPRL|TD|TLS|v.o.s.|VEB|VOF|BYSHR", "|")

    Dim stringArrayLength As Integer

    stringArray = SPLIT(string1, " ")
    stringArrayLength = UBound(stringArray) - LBound(stringArray)

    Dim companyAbbreviationString As Variant
    
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


'@Module: This module contains a set of functions for performing fuzzy string matches. It can be useful when you have 2 columns containing text that is close but not 100% the same. However, since the functions in this module only perform fuzzy matches, there is no guarentee that there will be 100% accuracy in the matches. However, for small groups of string where each string is very different than the other (such as a small group of fairly disimilar names), these functions can be highly accurate. Finally, some of the functions in this Module will take a long time to calculate for large numbers of cells, as the number of calculations for some functions will grow exponentially, but for small sets of data (such as 100 strings to compare), these functions perform fairly quickly.



'========================================
'  Hamming Distance
'========================================

Public Function HAMMING( _
    string1 As String, _
    string2 As String) _
As Integer

    '@Description: This function takes two strings of the same length and calculates the Hamming Distance between them. Hamming Distance measures how close two strings are by checking how many Subsitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the first string
    '@Param: string2 is the second string that will be compared to the first string
    '@Returns: Returns an integer of the Hamming Distance between two string
    '@Example: =HAMMING("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =HAMMING("Cat", "Bag") -> 2; 2 changes are needed, changing the "B" to "C" and the "g" to "t"
    '@Example: =HAMMING("Cat", "Dog") -> 3; Every single character needs to be substituted in this case

    If Len(string1) <> Len(string2) Then
        HAMMING = CVErr(xlErrValue)
    End If
    
    Dim totalDistance As Integer
    totalDistance = 0
    
    Dim i As Integer
    
    For i = 1 To Len(string1)
        If Mid(string1, i, 1) <> Mid(string2, i, 1) Then
            totalDistance = totalDistance + 1
        End If
    Next
    
    HAMMING = totalDistance
    
End Function



'========================================
'  Levenshtein Distance
'========================================

Public Function LEVENSHTEIN( _
    string1 As String, _
    string2 As String) _
As Long

    '@Description: This function takes two strings of any length and calculates the Levenshtein Distance between them. Levenshtein Distance measures how close two strings are by checking how many Insertions, Deletions, or Subsitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers. Unlike Hamming Distance, Levenshtein Distance works for strings of any length and includes 2 more operations. However, calculation time will be slower than Hamming Distance for same length strings, so if you know the two strings are the same length, its preferred to use Hamming Distance.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the first string
    '@Param: string2 is the second string that will be compared to the first string
    '@Returns: Returns an integer of the Levenshtein Distance between two string
    '@Example: =LEVENSHTEIN("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =LEVENSHTEIN("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
    '@Example: =LEVENSHTEIN("Cat", "Cta") -> 2; Since the "t" in "Cta" needs to be substituted into an "a", and the final character "a" needs to be substituted into a "t"

    ' **Error Checking**
    ' Quick returns for common errors
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
    

    ' **Algorithm Code**
    ' Creating the distance metrix and filling it with values
    Dim numberOfRows As Integer
    Dim numberOfColumns As Integer
    
    numberOfRows = Len(string1)
    numberOfColumns = Len(string2)
    
    Dim distanceArray() As Integer
    ReDim distanceArray(numberOfRows, numberOfColumns)
    
    Dim r As Integer
    Dim c As Integer
    
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
    
    ' Non-recursive Levenshtein Distance matrix walk
    Dim operationCost As Integer
    
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


Private Function Min3( _
    integer1 As Integer, _
    integer2 As Integer, _
    integer3 As Integer) _
As Integer

    '@Description: This function takes 3 integers and returns the minimum of them.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Check if WorksheetFunction.Min() is quicker than this function at calculating the minimum value, or check if there are alternative ways to calculate the min.
    '@Param: integer1 - integer3 are the integers to be compared
    '@Returns: Returns a the minimum integer
    '@Example: =Min3(4,10,6) -> 4

    If integer1 <= integer2 And integer1 <= integer3 Then
        Min3 = integer1
    ElseIf integer2 <= integer1 And integer2 <= integer3 Then
        Min3 = integer2
    ElseIf integer3 <= integer1 And integer3 <= integer2 Then
        Min3 = integer3
    End If

End Function


Public Function LEV_STR( _
    range1 As Range, _
    rangeArray As Range) _
As String

    '@Description: This function takes two ranges and calculates the string that is the result of the lowest Levenshtein Distance. The first range is a single cell which will be compared to all other cells in the second range and whichever value produces the lowest Levenshtein Distance will be returned.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: See if I can replace the first argument as a range with a string instead.
    '@Warning: This function will require exponential numbers of calculations for large amounts of strings. In cases where the number of strings are very large (more than 1000 strings), a better solution would be to use an external program other than Excel.
    '@Param: range1 contains the string we want to find the closest string in the rangeArray to
    '@Param: rangeArray is a range of all strings that will be compared to the string in range1
    '@Returns: Returns the string that is closest from the rangeArray
    '@Example: Where A1:A3 contains ["Bat", "Hello", "Dog"] =LEV_STR("Cat", A1:A3) -> "Bat"; Since "Bat" will have the lowest Levenshtein Distance of all 3 strings when compared to the string "Cat"

    Dim lngBestDistance As Long
    Dim lngCurrentDistance As Long
    Dim strRange1Value As String
    Dim strRange1Address As String
    Dim strBestMatch As String
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


Public Function LEV_STR_OPT( _
    range1 As Range, _
    rangeArray As Range, _
    numberOfLeftCharactersBound As Integer, _
    plusOrMinusLengthBound As Integer) _
As String

    '@Description: This function is the same as LEV_STR except that it adds two more arguments which can be used to optimize the speed of searches when the number of strings to search is very large. Since the number of calculations will increase exponentially to find the best fit string, this function can exclude a lot of strings that are unlikely to have the lowest Levenshtein Distance. The additional two parameters are a paramter that first checks a certain number of characters at the left of the strings and if the strings don't have the same characters on the left, then that string is excluded. The second of the two parameters sets the maximum length difference between the two strings, and if the length of string2 is not within the bounds of string1 length +/- the length bound, then this string is excluded. Setting high values for these parameters will essentially conver this function into a slightly slower version of LEV_STR.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: See if I can replace the first argument as a range with a string instead.
    '@Warning: This function will work for a much larger number of strings than LEV_STR, however care must be taken before using it. This function will perform well in cases where the group of strings are likely to have a large number of differnces between each individual string and where it is likely that the leftmost charaters of the string will be the same. An example might comparing two sets of company names for the companies in a stock index, as they are likely to be fairly different but likely will have the same leftmost characters between the two lists.
    '@Param: range1 contains the string we want to find the closest string in the rangeArray to
    '@Param: rangeArray is a range of all strings that will be compared to the string in range1
    '@Param: numberOfLeftCharactersBound is the number of left characters that will be checked first on both strings before calculating their Levenshtein Distance
    '@Param: plusOrMinusLengthBound is the number plus or minus the length of the first string that will be checked compared to the second string before calculating their Levenshtein Distance
    '@Returns: Returns the string that is closest from the rangeArray
    '@Example: Where A1:A3 contains ["Car", "C Programming Langauge", "Dog"] =LEV_STR_OPT("Cat", A1:A3, 1, 2) -> "Car"; The calculation won't be performed on "Dog" since "Dog" doesn't start with the character "C", and "C Programming Langauge" won't be calculated either since its length is greating than LEN("Cat") +/- 2 (its length is not between 0-5 characters long).

    Dim lngBestDistance As Long
    Dim lngCurrentDistance As Long
    Dim strRange1Value As String
    Dim strRange1Address As String
    Dim strBestMatch As String
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



'========================================
'  Damerau-Levenshtein Distance
'========================================

Public Function DAMERAU( _
    string1 As String, _
    string2 As String) _
As Integer

    '@Description: This function takes two strings of any length and calculates the Damerau-Levenshtein Distance between them. Damerau-Levenshtein Distance differs from Levenshtein Distance in that it includes an additional operation, called Transpositions, which occurs when two adjacent characters are swapped. Thus, Damerau-Levenshtein Distance calculates the number of Insertions, Deletions, Substitutions, and Transpositons needed to convert string1 into string2. As a result, this function is good when it is likely that spelling errors have occured between two string where the error is simply a transposition of 2 adjacent characters.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the first string
    '@Param: string2 is the second string that will be compared to the first string
    '@Returns: Returns an integer of the Damerau-Levenshtein Distance between two string
    '@Example: =DAMERAU("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =DAMERAU("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
    '@Example: =DAMERAU("Cat", "Cta") -> 1; Since the "t" and "a" can be transposed as they are adjacent to each other. Notice how LEVENSHTEIN("Cat","Cta")=2 but DAMERAU("Cat","Cta")=1

    ' **Error Checking**
    ' Quick returns for common errors
    If string1 = string2 Then
        DAMERAU = 0
    ElseIf string1 = Empty Then
        DAMERAU = Len(string2)
    ElseIf string2 = Empty Then
        DAMERAU = Len(string1)
    End If
    
    Dim inf As Long
    Dim da As Object
    inf = Len(string1) + Len(string2)
    Set da = CreateObject("Scripting.Dictionary")
    
    ' 35 - 38 = filling the dictionary
    Dim i As Integer
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
    
    ' 39 = creating h matrix
    Dim H() As Long
    ReDim H(Len(string1) + 1, Len(string2) + 1)
    
    Dim k As Integer
    For i = 0 To (Len(string1) + 1)
        For k = 0 To (Len(string2) + 1)
            H(i, k) = 0
        Next
    Next
    
    ' 40 - 45 = updating the matrix
    For i = 0 To Len(string1)
        H(i + 1, 0) = inf
        H(i + 1, 1) = i
    Next
    For k = 0 To Len(string2)
        H(0, k + 1) = inf
        H(1, k + 1) = k
    Next
    

    ' 46 - 60 = running the array
    Dim db As Long
    Dim i1 As Long
    Dim k1 As Long
    Dim cost As Long
    
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
            
            H(i + 1, k + 1) = Min4(H(i, k) + cost, _
                                   H(i + 1, k) + 1, _
                                   H(i, k + 1) + 1, _
                                   H(i1, k1) + (i - i1 - 1) + 1 + (k - k1 - 1))
                           
            
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


Private Function Min4( _
    integer1 As Integer, _
    integer2 As Integer, _
    integer3 As Integer, _
    integer4 As Integer) _
As Long

    '@Description: This function takes 4 integers and returns the minimum of them.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Check if WorksheetFunction.Min() is quicker than this function at calculating the minimum value, or check if there are alternative ways to calculate the min.
    '@Param: integer1 - integer4 are the integers to be compared
    '@Returns: Returns a the minimum integer
    '@Example: =Min4(4,10,6,3) -> 3

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


Public Function DAM_STR( _
    range1 As Range, _
    rangeArray As Range) _
As String

    '@Description: This function takes two ranges and calculates the string that is the result of the lowest Damerau-Levenshtein Distance. The first range is a single cell which will be compared to all other cells in the second range and whichever value produces the lowest Damerau-Levenshtein Distance will be returned.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: See if I can replace the first argument as a range with a string instead.
    '@Warning: This function will require exponential numbers of calculations for large amounts of strings. In cases where the number of strings are very large (more than 1000 strings), a better solution would be to use an external program other than Excel. Also this function will perform well in the case of comparing two lists with the same content but with spelling errors, but in cases where transpositions are unlikely, this LEV_STR should be used as this function will be slower.
    '@Param: range1 contains the string we want to find the closest string in the rangeArray to
    '@Param: rangeArray is a range of all strings that will be compared to the string in range1
    '@Returns: Returns the string that is closest from the rangeArray
    '@Example: Where A1:A3 contains ["Bath", "Hello", "Cta"] =DAM_STR("Cat", A1:A3) -> "Cta"; LEV_STR will actually return "Bath" in this case since it comes first in the range and since "Bath" and "Cta" will actually both have a LEV=2, but while "Bath" with have DAM=2, for "Cta" only one operation is required (a single Transposition instead of a Substitution and a Deletion) and thus for "Cta" DAM=1

    Dim lngBestDistance As Long
    Dim lngCurrentDistance As Long
    Dim strRange1Value As String
    Dim strRange1Address As String
    Dim strBestMatch As String
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


Public Function DAM_STR_OPT( _
    range1 As Range, _
    rangeArray As Range, _
    numberOfLeftCharactersBound As Long, _
    plusOrMinusLengthBound) _
As String

    '@Description: This function is the same as DAM_STR except that it adds two more arguments which can be used to optimize the speed of searches when the number of strings to search is very large. Since the number of calculations will increase exponentially to find the best fit string, this function can exclude a lot of strings that are unlikely to have the lowest Damerauï¿½Levenshtein Distance. The additional two parameters are a paramter that first checks a certain number of characters at the left of the strings and if the strings don't have the same characters on the left, then that string is excluded. The second of the two parameters sets the maximum length difference between the two strings, and if the length of string2 is not within the bounds of string1 length +/- the length bound, then this string is excluded. Setting high values for these parameters will essentially conver this function into a slightly slower version of DAM_STR.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: See if I can replace the first argument as a range with a string instead.
    '@Warning: This function will work for a much larger number of strings than LEV_STR, however care must be taken before using it. This function will perform well in cases where the group of strings are likely to have a large number of differnces between each individual string and where it is likely that the leftmost charaters of the string will be the same. An example might comparing two sets of company names for the companies in a stock index, as they are likely to be fairly different but likely will have the same leftmost characters between the two lists.
    '@Param: range1 contains the string we want to find the closest string in the rangeArray to
    '@Param: rangeArray is a range of all strings that will be compared to the string in range1
    '@Param: numberOfLeftCharactersBound is the number of left characters that will be checked first on both strings before calculating their Levenshtein Distance
    '@Param: plusOrMinusLengthBound is the number plus or minus the length of the first string that will be checked compared to the second string before calculating their Levenshtein Distance
    '@Returns: Returns the string that is closest from the rangeArray
    '@Example: Where A1:A3 contains ["Car", "C Programming Langauge", "Dog"] =DAM_STR_OPT("Cat", A1:A3, 1, 2) -> "Car"; The calculation won't be performed on "Dog" since "Dog" doesn't start with the character "C", and "C Programming Langauge" won't be calculated either since its length is greating than LEN("Cat") +/- 2 (its length is not between 0-5 characters long).
    
    Dim lngBestDistance As Long
    Dim lngCurrentDistance As Long
    Dim strRange1Value As String
    Dim strRange1Address As String
    Dim strBestMatch As String
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


Public Function PARTIAL_LOOKUP( _
    range1 As Range, _
    rangeArray As Range) _
As String

    '@Description: This function takes two ranges and calculates the string that is the closest match. It works similar to a VLOOKUP, expect that it works with partial matches as well as exact matches.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: See if I can replace the first argument as a range with a string instead.
    '@Note: This function is an alias for DAM_STR, and works in the exact same way.
    '@Warning: This function will require exponential numbers of calculations for large amounts of strings. In cases where the number of strings are very large (more than 1000 strings), a better solution would be to use an external program other than Excel. Also this function will perform well in the case of comparing two lists with the same content but with spelling errors, but in cases where transpositions are unlikely, this LEV_STR should be used as this function will be slower.
    '@Param: range1 contains the string we want to find the closest string in the rangeArray to
    '@Param: rangeArray is a range of all strings that will be compared to the string in range1
    '@Returns: Returns the string that is closest from the rangeArray
    '@Example: Where A1:A3 contains ["Bath", "Hello", "Cta"] =DAM_STR("Cat", A1:A3) -> "Cta"; LEV_STR will actually return "Bath" in this case since it comes first in the range and since "Bath" and "Cta" will actually both have a LEV=2, but while "Bath" with have DAM=2, for "Cta" only one operation is required (a single Transposition instead of a Substitution and a Deletion) and thus for "Cta" DAM=1

    PARTIAL_LOOKUP = DAM_STR(range1, rangeArray)

End Function


'@Module: This module contains a set of basic miscellanous utility functions


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
    '@Param: stringArray() is an array of strings and number in the following format: {"Hello", "World"}
    '@Param: indentLevel is an optional number that specifying the indentation level. Leaving this argument out will result in no indentation
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

    '@Description: This function is an extension of Excel's =ISERROR(). It returns TRUE for all of Excel's built in errors, similar to =ISERROR() but also returns TRUE for User-Defined Error Strings. User-Defined Error Strings are strings that start with character "#" and end with either the character "!" or "?". This is similar to the format of errors in Excel, such as "#DIV/0!", "#VALUE!", "#NAME?", "#REF!", etc. User-Defined Error Strings are used all throughout X+, so this is a useful function for checking errors in X+ functions. Additionally, users can create their own User-Defined Error Strings in Excel and use this function to check for those errors.
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

    '@Description: This function executes JavaScript code using Microsoft's JScript scripting language. It takes 3 arguments, the JavaScript code that will be executed, the name of the JavaScript function that will be executed, and up to 16 optional arguments to be used in the JavaScript function that is called. One thing to note is that ES5 syntax should be used in the JavaScript code, as ES6 features are unlikley to be supported.
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
    '@Example: =HTML_TABLEIFY(A1:C5) -> <table>...</table>;

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
    '@Example: =HTML_ESCAPE("&lt;p&gt;Hello World&lt;/p&gt;") -> "<p>Hello World</p>"

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


'@Module: This module contains a set of functions for validating some commonly used string, such as validators for email addresses and phone numbers.



Public Function IS_EMAIL( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid email address.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Param: string1 is the string we are checking if its a valid email
    '@Returns: Returns TRUE if the string is a valid email, and FALSE if its invalid
    '@Example: =IS_EMAIL("JohnDoe@testmail.com") -> TRUE
    '@Example: =IS_EMAIL("JohnDoe@test/mail.com") -> FALSE
    '@Example: =IS_EMAIL("not_an_email_address") -> FALSE

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


Public Function IS_PHONE( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a phone number is valid.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Todo: Add a second argument that lets the user add a country code and uses a different regex for phone number formats for that country. Also make the regx more robust so it can include more common formats.
    '@Param: string1 is the string we are checking if its a valid phone number
    '@Returns: Returns TRUE if the string is a valid phone number, and FALSE if its invalid
    '@Example: =IS_PHONE("123 456 7890") -> TRUE
    '@Example: =IS_PHONE("1234567890") -> TRUE
    '@Example: =IS_PHONE("1-234-567-890") -> FALSE; Not enough digits
    '@Example: =IS_PHONE("1-234-567-8905") -> TRUE
    '@Example: =IS_PHONE("+1-234-567-890") -> FALSE; Not enough digits
    '@Example: =IS_PHONE("+1-234-567-8905") -> TRUE
    '@Example: =IS_PHONE("+1-(234)-567-8905") -> TRUE
    '@Example: =IS_PHONE("+1 (234) 567 8905") -> TRUE
    '@Example: =IS_PHONE("1(234)5678905") -> TRUE
    '@Example: =IS_PHONE("123-456-789") -> FALSE; Not enough digits
    '@Example: =IS_PHONE("Hello World") -> FALSE; Not a phone number

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


Public Function IS_CREDIT_CARD( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid credit card from one of the major card issuing companies.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string we are checking if its a valid credit card number
    '@Returns: Returns TRUE if the string is a valid credit card number, and FALSE if its invalid. Currently supports these cards: Visa, MasterCard, Discover, Amex, Diners, JCB
    '@Example: =IS_CREDIT_CARD("5111567856785678") -> TRUE; This is a valid Mastercard number
    '@Example: =IS_CREDIT_CARD("511156785678567") -> FALSE; Not enough digits
    '@Example: =IS_CREDIT_CARD("9999999999999999") -> FALSE; Enough digits, but not a valid card number
    '@Example: =IS_CREDIT_CARD("Hello World") -> FALSE

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
        
    Dim regexPattern As String
    
    ' Regex for Amex
    regexPattern = regexPattern & "(3[47][0-9]{13})|"
    
    ' Regex for Diners
    regexPattern = regexPattern & "(3(0[0-5]|[68][0-9])?[0-9]{11})|"
    
    ' Regex for Discover
    regexPattern = regexPattern & "(6(011|5[0-9]{2})[0-9]{12})|"
    
    ' Regex for JCB
    regexPattern = regexPattern & "((2131|1800|35[0-9]{3})[0-9]{11})"
    
    ' Regex for MasterCard
    regexPattern = regexPattern & "(5[1-5][0-9]{14})|"
    
    ' Regex for Visa
    regexPattern = regexPattern & "(4[0-9]{12}([0-9]{3})?)|"
    
    With Regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = regexPattern
    End With

    IS_CREDIT_CARD = Regex.Test(string1)

End Function


Public Function IS_URL( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid URL address.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Param: string1 is the string we are checking if its a valid URL
    '@Returns: Returns TRUE if the string is a valid URL, and FALSE if its invalid
    '@Example: =IS_URL("https://www.wikipedia.org/") -> TRUE
    '@Example: =IS_URL("http://www.wikipedia.org/") -> TRUE
    '@Example: =IS_URL("hello_world") -> FALSE

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


Public Function IS_IP_FOUR( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid IPv4 address.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Param: string1 is the string we are checking if its a valid IPv4 address
    '@Returns: Returns TRUE if the string is a valid IPv4, and FALSE if its invalid
    '@Example: =IS_IP_FOUR("0.0.0.0") -> TRUE
    '@Example: =IS_IP_FOUR("100.100.100.100") -> TRUE
    '@Example: =IS_IP_FOUR("255.255.255.255") -> TRUE
    '@Example: =IS_IP_FOUR("255.255.255.256") -> FALSE; as the final 256 makes the address outside of the bounds of IPv4
    '@Example: =IS_IP_FOUR("0.0.0") -> FALSE; as the fourth octet is missing

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


Public Function CREDIT_CARD_NAME( _
    ByVal string1 As String) _
As String

    '@Description: This function checks if a string is a valid credit card from one of the major card issuing companies, and then returns the name of the credit card name. This function assumes no spaces or hyphens (if you have card numbers with spaces or hyphens you can remove these using =SUBSTITUTE("-", "") function.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the credit card string
    '@Returns: Returns the name of the credit card. Currently supports these cards: Visa, MasterCard, Discover, Amex, Diners, JCB
    '@Example: =CREDIT_CARD_NAME("5111567856785678") -> "MasterCard"; This is a valid Mastercard number
    '@Example: =CREDIT_CARD_NAME("not_a_card_number") -> #VALUE!

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
    
    Regex.Global = True
    Regex.IgnoreCase = True
    Regex.MultiLine = True

    ' Regex for Amex
    Regex.Pattern = "(3[47][0-9]{13})"
    If Regex.Test(string1) Then
        CREDIT_CARD_NAME = "Amex"
        Exit Function
    End If
    
    ' Regex for Diners
    Regex.Pattern = "(3(0[0-5]|[68][0-9])?[0-9]{11})"
    If Regex.Test(string1) Then
        CREDIT_CARD_NAME = "Diners"
        Exit Function
    End If
    
    ' Regex for Discover
    Regex.Pattern = "(6(011|5[0-9]{2})[0-9]{12})"
    If Regex.Test(string1) Then
        CREDIT_CARD_NAME = "Discover"
        Exit Function
    End If
    
    ' Regex for JCB
    Regex.Pattern = "((2131|1800|35[0-9]{3})[0-9]{11})"
    If Regex.Test(string1) Then
        CREDIT_CARD_NAME = "JCB"
        Exit Function
    End If
    
    ' Regex for MasterCard
    Regex.Pattern = "(5[1-5][0-9]{14})"
    If Regex.Test(string1) Then
        CREDIT_CARD_NAME = "MasterCard"
        Exit Function
    End If
    
    ' Regex for Visa
    Regex.Pattern = "(4[0-9]{12}([0-9]{3})?)"
    If Regex.Test(string1) Then
        CREDIT_CARD_NAME = "Visa"
        Exit Function
    End If
    
    CREDIT_CARD_NAME = "#NotAValidCreditCardNumber!"

End Function


Public Function FORMAT_FRACTION( _
    ByVal decimal1 As Double) _
As String

    '@Description: This function takes a decimal number and formats it as a close rounded fraction.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: decimal1 is decimal number that will be formatted
    '@Returns: Returns a string of a decimal formatted as a fraction
    '@Example: =FORMAT_FRACTION(".33") -> "1/3"
    '@Example: =FORMAT_FRACTION(".35") -> "1/3"
    '@Example: =FORMAT_FRACTION(".37") -> "3/8"
    '@Example: =FORMAT_FRACTION(".7") -> "2/3"
    '@Example: =FORMAT_FRACTION("2.5") -> "2 1/5"

    FORMAT_FRACTION = Trim(WorksheetFunction.Text(decimal1, "# ?/?"))

End Function


Public Function FORMAT_PHONE( _
    ByVal string1 As String) _
As String

    '@Description: This function checks if a string is a phone number and if it is, formats the phone number as a more readable string.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Add a second argument that lets the user add a country code and uses a different format.
    '@Param: string1 is a phone number
    '@Returns: Returns a string formatted as a more readable phone number
    '@Example: =FORMAT_PHONE("123 456 7890") -> "(123) 456-7890"

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

Public Function FORMAT_CREDIT_CARD( _
    ByVal string1 As String) _
As String

    '@Description: This function checks if a string is a valid credit card, and if it is formats it in a more readable way. The format used is XXXX-XXXX-XXXX-XXXX.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is credit card number
    '@Returns: Returns a string formatted as a more readable credit card number
    '@Example: =FORMAT_CREDIT_CARD("5111567856785678") -> "5111-5678-5678-5678"

    If IS_CREDIT_CARD(string1) Then
        FORMAT_CREDIT_CARD = Left(string1, 4) & "-" & Mid(string1, 5, 4) & "-" & Mid(string1, 9, 4) & "-" & Mid(string1, 13)
    Else
        FORMAT_CREDIT_CARD = "#NotAValidCreditCardNumber!"
    End If

End Function

