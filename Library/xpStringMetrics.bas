Attribute VB_Name = "xpStringMetrics"
'@Module: This module contains a set of functions for performing fuzzy string matches. It can be useful when you have 2 columns containing text that is close but not 100% the same. However, since the functions in this module only perform fuzzy matches, there is no guarentee that there will be 100% accuracy in the matches. However, for small groups of string where each string is very different than the other (such as a small group of fairly disimilar names), these functions can be highly accurate. Finally, some of the functions in this Module will take a long time to calculate for large numbers of cells, as the number of calculations for some functions will grow exponentially, but for small sets of data (such as 100 strings to compare), these functions perform fairly quickly.

Option Explicit


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

    '@Description: This function is the same as DAM_STR except that it adds two more arguments which can be used to optimize the speed of searches when the number of strings to search is very large. Since the number of calculations will increase exponentially to find the best fit string, this function can exclude a lot of strings that are unlikely to have the lowest Damerau–Levenshtein Distance. The additional two parameters are a paramter that first checks a certain number of characters at the left of the strings and if the strings don't have the same characters on the left, then that string is excluded. The second of the two parameters sets the maximum length difference between the two strings, and if the length of string2 is not within the bounds of string1 length +/- the length bound, then this string is excluded. Setting high values for these parameters will essentially conver this function into a slightly slower version of DAM_STR.
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

