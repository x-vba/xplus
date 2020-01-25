Attribute VB_Name = "xpStringManipulation"
'@Module: This module contains a set of basic functions for manipulating strings

Option Explicit


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
    companyAbbreviationArray = SPLIT("AB|AG|GmbH|LLC|LLP|NV|PLC|SA|A. en P.|ACE|AD|AE|AL|AmbA|ANS|ApS|AS|ASA|AVV|BVBA|CA|CVA|d.d.|d.n.o.|d.o.o.|DA|e.V.|EE|EEG|EIRL|ELP|EOOD|EPE|EURL|GbR|GCV|GesmbH|GIE|HB|hf|IBC|j.t.d.|k.d.|k.d.d.|k.s.|KA/S|KB|KD|KDA|KG|KGaA|KK|Kol. SrK|Kom. SrK|LDC|Ltée.|NT|OE|OHG|Oy|OYJ|OÜ|PC Ltd|PMA|PMDN|PrC|PT|RAS|S. de R.L.|S. en N.C.|SA de CV|SAFI|SAS|SC|SCA|SCP|SCS|SENC|SGPS|SK|SNC|SOPARFI|sp|Sp. z.o.o.|SpA|spol s.r.o.|SPRL|TD|TLS|v.o.s.|VEB|VOF|BYSHR", "|")

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

