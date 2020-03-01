Attribute VB_Name = "xpStringManipulation"
'@Module: This module contains a set of basic functions for manipulating strings.

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
    '@Example: =LEFT_FIND("Hello World", "r") -> "Hello Wo"
    '@Example: =LEFT_FIND("Hello World", "R") -> "#VALUE!"; Since string1 does not contain "R" in it.

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
    '@Example: =RIGHT_FIND("Hello World", "o") -> "rld"
    '@Example: =RIGHT_FIND("Hello World", "O") -> "#VALUE!"; Since string1 does not contain "O" in it.

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
    '@Example: =LEFT_SEARCH("Hello World", "r") -> "Hello Wo"
    '@Example: =LEFT_SEARCH("Hello World", "R") -> "Hello Wo"

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
    '@Example: =RIGHT_SEARCH("Hello World", "o") -> "rld"
    '@Example: =RIGHT_SEARCH("Hello World", "O") -> "rld"

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
    '@Example: =SUBSTR("Hello World", 2, 6) -> "ello"

    SUBSTR = Mid(string1, startCharacterNumber, endCharacterNumber - startCharacterNumber)

End Function


Public Function SUBSTR_FIND( _
    ByVal string1 As String, _
    ByVal leftSearchString As String, _
    ByVal rightSearchString As String, _
    Optional ByVal noninclusiveFlag As Boolean) _
As String

    '@Description: This function takes a string and a left string and right string, and returns a substring between those two strings. The left string will find the first matching string starting from the left, and the right string will find the first matching string starting from the right. Finally, and optional final parameter can be set to TRUE to make the substring noninclusive of the two searched strings. SUBSTR_FIND is case-sensitive. For case-insensitive version, see SUBSTR_SEARCH
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that we will get a substring from
    '@Param: leftSearchString is the string that will be searched from the left
    '@Param: rightSearchString is the string that will be searched from the right
    '@Param: noninclusiveFlag is an optional parameter that if set to TRUE will result in the substring not including the left and right searched characters
    '@Returns: Returns a substring between the two strings.
    '@Example: =SUBSTR_FIND("Hello World", "e", "o") -> "ello Wo"
    '@Example: =SUBSTR_FIND("Hello World", "e", "o", TRUE) -> "llo W"
    '@Example: =SUBSTR_FIND("One Two Three", "ne ", " Thr") -> "ne Two Thr"
    '@Example: =SUBSTR_FIND("One Two Three", "NE ", " THR") -> "#VALUE!"; Since SUBSTR_FIND() is case-sensitive
    '@Example: =SUBSTR_FIND("One Two Three", "ne ", " Thr", TRUE) -> "Two"
    '@Example: =SUBSTR_FIND("Country Code: +51; Area Code: 315; Phone Number: 762-5929;", "Area Code: ", "; Phone", TRUE) -> 315
    '@Example: =SUBSTR_FIND("Country Code: +313; Area Code: 423; Phone Number: 284-2468;", "Area Code: ", "; Phone", TRUE) -> 423
    '@Example: =SUBSTR_FIND("Country Code: +171; Area Code: 629; Phone Number: 731-5456;", "Area Code: ", "; Phone", TRUE) -> 629

    Dim leftCharacterNumber As Integer
    Dim rightCharacterNumber As Integer
    
    leftCharacterNumber = InStr(1, string1, leftSearchString)
    rightCharacterNumber = InStrRev(string1, rightSearchString)
    
    If noninclusiveFlag = True Then
        leftCharacterNumber = leftCharacterNumber + Len(leftSearchString)
        rightCharacterNumber = rightCharacterNumber - Len(rightSearchString)
    End If
    
    SUBSTR_FIND = Mid(string1, leftCharacterNumber, rightCharacterNumber - leftCharacterNumber + Len(rightSearchString))

End Function


Public Function SUBSTR_SEARCH( _
    ByVal string1 As String, _
    ByVal leftSearchString As String, _
    ByVal rightSearchString As String, _
    Optional ByVal noninclusiveFlag As Boolean) _
As String

    '@Description: This function takes a string and a left string and right string, and returns a substring between those two strings. The left string will find the first matching string starting from the left, and the right string will find the first matching string starting from the right. Finally, and optional final parameter can be set to TRUE to make the substring noninclusive of the two searched strings. SUBSTR_SEARCH is case-insensitive. For case-sensitive version, see SUBSTR_FIND
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Param: string1 is the string that we will get a substring from
    '@Param: leftSearchString is the string that will be searched from the left
    '@Param: rightSearchString is the string that will be searched from the right
    '@Param: noninclusiveFlag is an optional parameter that if set to TRUE will result in the substring not including the left and right searched characters
    '@Returns: Returns a substring between the two strings.
    '@Example: =SUBSTR_SEARCH("Hello World", "e", "o") -> "ello Wo"
    '@Example: =SUBSTR_SEARCH("Hello World", "e", "o", TRUE) -> "llo W"
    '@Example: =SUBSTR_SEARCH("One Two Three", "ne ", " Thr") -> "ne Two Thr"
    '@Example: =SUBSTR_SEARCH("One Two Three", "NE ", " THR") -> "ne Two Thr"; No error, since SUBSTR_SEARCH is case-insensitive
    '@Example: =SUBSTR_SEARCH("One Two Three", "ne ", " Thr", TRUE) -> "Two"
    '@Example: =SUBSTR_SEARCH("Country Code: +51; Area Code: 315; Phone Number: 762-5929;", "Area Code: ", "; Phone", TRUE) -> 315
    '@Example: =SUBSTR_SEARCH("Country Code: +313; Area Code: 423; Phone Number: 284-2468;", "Area Code: ", "; Phone", TRUE) -> 423
    '@Example: =SUBSTR_SEARCH("Country Code: +171; Area Code: 629; Phone Number: 731-5456;", "Area Code: ", "; Phone", TRUE) -> 629

    Dim leftCharacterNumber As Integer
    Dim rightCharacterNumber As Integer
    
    leftCharacterNumber = InStr(1, string1, leftSearchString, vbTextCompare)
    rightCharacterNumber = InStrRev(string1, rightSearchString, Compare:=vbTextCompare)
    
    If noninclusiveFlag = True Then
        leftCharacterNumber = leftCharacterNumber + Len(leftSearchString)
        rightCharacterNumber = rightCharacterNumber - Len(rightSearchString)
    End If
    
    SUBSTR_SEARCH = Mid(string1, leftCharacterNumber, rightCharacterNumber - leftCharacterNumber + Len(rightSearchString))

End Function

    
Public Function REPEAT( _
    ByVal string1 As String, _
    ByVal numberOfRepeats As Integer) _
As String

    '@Description: This function repeats string1 based on the number of repeats specified in the second argument
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be repeated
    '@Param: numberOfRepeats is the number of times string1 will be repeated
    '@Returns: Returns a string repeated multiple times based on the numberOfRepeats
    '@Example: =REPEAT("Hello", 2) -> HelloHello"
    '@Example: =REPEAT("=", 10) -> "=========="

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

    '@Description: This function takes a formatter string and then an array of ranges or strings, and replaces the format placeholders with the values in the range or strings. The format syntax is "{1} - {2}" where the "{1}" and "{2}" will be replaced with the values given in the text array.
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
    Optional ByVal fillCharacter As String = "0", _
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
    
    While Len(string1) < fillLength
        If rightToLeftFlag = False Then
            string1 = fillCharacter + string1
        Else
            string1 = string1 + fillCharacter
        End If
    Wend
    
    ZFILL = string1

End Function


Public Function SPLIT_TEXT( _
    ByVal string1 As String, _
    ByVal substringNumber As Integer, _
    Optional ByVal delimiterString As String = " ") _
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
    
    SPLIT_TEXT = SPLIT(string1, delimiterString)(substringNumber - 1)

End Function


Public Function COUNT_WORDS( _
    ByVal string1 As String, _
    Optional ByVal delimiterString As String = " ") _
As Integer

    '@Description: This function takes a string and returns the number of words in the string
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Note: If the number given is higher than the number of words, its possible that the string contains excess whitespace. Try using the =TRIM() function first to remove the excess whitespace
    '@Param: string1 is the string whose number of words will be counted
    '@Param: delimiterString is an optional parameter that can be used to specify a different delimiter
    '@Returns: Returns the number of words in the string
    '@Example: =COUNT_WORDS("Hello World") -> 2
    '@Example: =COUNT_WORDS("One Two Three") -> 3
    '@Example: =COUNT_WORDS("One-Two-Three", "-") -> 3

    Dim stringArray() As String

    stringArray = SPLIT(string1, delimiterString)
    
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

    '@Description: This function takes a string and either another string or multiple strings and removes all characters from the first string that are in the second string.
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


Public Function REVERSE_TEXT( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and reverses all the characters in it so that the returned string is backwards
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be reversed
    '@Returns: Returns the origional string in reverse
    '@Example: =REVERSE_TEXT("Hello World") -> "dlroW olleH"

    Dim i As Integer
    Dim reversedString As String
    
    For i = 1 To Len(string1)
        reversedString = reversedString & Mid(string1, Len(string1) - i + 1, 1)
    Next
    
    REVERSE_TEXT = reversedString

End Function


Public Function REVERSE_WORDS( _
    ByVal string1 As String, _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string and reverses all the words in it so that the returned string's words are backwards. By default, this function uses the space character as a delimiter, but you can optionally specify a different delimiter.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose words will be reversed
    '@Param: delimiterCharacter is the delimiter that will be used, with the default being " "
    '@Returns: Returns the origional string with it's words reversed
    '@Example: =REVERSE_WORDS("Hello World") -> "World Hello"
    '@Example: =REVERSE_WORDS("One Two Three") -> "Three Two One"
    '@Example: =REVERSE_WORDS("One-Two-Three", "-") -> "Three-Two-One"

    Dim i As Integer
    Dim stringArray() As String
    Dim stringArrayLength As Integer
    Dim reversedStringArray() As String
    
    stringArray = SPLIT(string1, delimiterCharacter)
    stringArrayLength = (UBound(stringArray) - LBound(stringArray))
    
    ReDim reversedStringArray(stringArrayLength)
    
    For i = 0 To stringArrayLength
        reversedStringArray(i) = stringArray(stringArrayLength - i)
    Next
    
    REVERSE_WORDS = Join(reversedStringArray, delimiterCharacter)

End Function


Public Function INDENT( _
    ByVal string1 As String, _
    Optional ByVal indentAmount As Byte = 4) _
As String

    '@Description: This function takes a string and indents all of its lines by a specified number of space characters (or 4 space characters if left blank)
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be indented
    '@Param: indentAmount is the amount of " " characters that will be indented to the left of string1
    '@Returns: Returns the origional string indented by a specified number of space characters
    '@Example: =INDENT("Hello") -> "    Hello"
    '@Example: =INDENT("Hello", 4) -> "    Hello"
    '@Example: =INDENT("Hello", 3) -> "   Hello"
    '@Example: =INDENT("Hello", 2) -> "  Hello"
    '@Example: =INDENT("Hello", 1) -> " Hello"

    Dim i As Integer
    Dim stringArray() As String

    stringArray = SPLIT(string1, Chr(10))
    
    string1 = ""
    For i = 1 To indentAmount
        string1 = string1 & " "
    Next
    
    For i = 0 To (UBound(stringArray) - LBound(stringArray))
        stringArray(i) = string1 & stringArray(i)
    Next

    INDENT = Join(stringArray, Chr(10))

End Function


Public Function DEDENT( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and dedents all of its lines so that there are no space characters to the left or right of each line
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be dedented
    '@Returns: Returns the origional string dedented on each line
    '@Note: Unlike the Excel built-in TRIM() function, this function will dedent every single line, so for strings that span multiple lines in a cell, this will dedent all lines.
    '@Example: =DEDENT("    Hello") -> "Hello"

    Dim i As Integer
    Dim stringArray() As String

    stringArray = SPLIT(string1, Chr(10))
    
    For i = 0 To (UBound(stringArray) - LBound(stringArray))
        stringArray(i) = Trim(stringArray(i))
    Next

    DEDENT = Join(stringArray, Chr(10))

End Function


Public Function SHORTEN( _
    ByVal string1 As String, _
    Optional ByVal shortenWidth As Integer = 80, _
    Optional ByVal placeholderText As String = "[...]", _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string and shortens it with placeholder text so that it is no longer in length than the specified width.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be shortened
    '@Param: shortenWidth is the max width of the string. By default this is set to 80
    '@Param: placeholderText is the text that will be placed at the end of the string if it is longer than the shortenWidth. By default this placeholder string is "[...]
    '@Param: delimiterCharacter is the character that will be used as the word delimiter. By default this is the space character " "
    '@Returns: Returns a shortened string with placeholder text if it is longer than the shorten width
    '@Example: =SHORTEN("Hello World One Two Three", 20) -> "Hello World [...]"; Only the first two words and the placeholder will result in a string that is less than or equal to 20 in length
    '@Example: =SHORTEN("Hello World One Two Three", 15) -> "Hello [...]"; Only the first word and the placeholder will result in a string that is less than or equal to 15 in length
    '@Example: =SHORTEN("Hello World One Two Three") -> "Hello World One Two Three"; Since this string is shorter than the default 80 shorten width value, no placeholder will be used and the string wont be shortened
    '@Example: =SHORTEN("Hello World One Two Three", 15, "-->") -> "Hello World -->"; A new placeholder is used
    '@Example: =SHORTEN("Hello_World_One_Two_Three", 15, "-->", "_") -> "Hello_World_-->"; A new placeholder andd delimiter is used

    Dim shortenedString As String
    Dim individualString As Variant
    Dim stringArray() As String
    
    ' In cases where the origional string is less than the threshold needed to
    ' shorten the string, simply return the origional string
    If Len(string1) <= (shortenWidth - Len(placeholderText) - Len(delimiterCharacter)) Then
        SHORTEN = string1
        Exit Function
    End If
    
    stringArray = SPLIT(string1, delimiterCharacter)

    For Each individualString In stringArray
        If Len(shortenedString & individualString) > (shortenWidth - Len(placeholderText) - Len(delimiterCharacter)) Then
            shortenedString = shortenedString & placeholderText
            Exit For
        Else
            shortenedString = shortenedString & individualString & delimiterCharacter
        End If
    Next

    SHORTEN = shortenedString

End Function


Public Function INSPLIT( _
    ByVal string1 As String, _
    ByVal splitString As String, _
    Optional ByVal delimiterCharacter As String = " ") _
As Boolean

    '@Description: This function takes a search string and checks if it exists within a larger string that is split by a delimiter character.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be checked if it exists within the splitString after the split
    '@Param: splitString is the string that will be split and of which string1 will be searched in
    '@Param: delimiterCharacter is the character that will be used as the delimiter for the split. By default this is the space character " "
    '@Returns: Returns TRUE if string1 is found in splitString after the split occurs
    '@Example: =INSPLIT("Hello", "Hello World One Two Three") -> TRUE; Since "Hello" is found within the searchString after being split
    '@Example: =INSPLIT("NotInString", "Hello World One Two Three") -> FALSE; Since "NotInString" is not found within the searchString after being split
    '@Example: =INSPLIT("Hello", "Hello-World-One-Two-Three", "-") -> TRUE; Since "Hello" is found and since the delimiter is set to "-"

    Dim individualString As Variant
    
    For Each individualString In SPLIT(splitString, delimiterCharacter)
        If string1 = individualString Then
            INSPLIT = True
            Exit Function
        End If
    Next
    
    INSPLIT = False

End Function


Public Function ELITE_CASE( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the string with characters replaced by similar in appearance numbers
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will have characters replaced
    '@Returns: Returns the string with characters replaced with similar in appearance numbers
    '@Example: =ELITE_CASE("Hello World") -> "H3110 W0r1d"

    string1 = Replace(string1, "o", "0", Compare:=vbTextCompare)
    string1 = Replace(string1, "l", "1", Compare:=vbTextCompare)
    string1 = Replace(string1, "z", "2", Compare:=vbTextCompare)
    string1 = Replace(string1, "e", "3", Compare:=vbTextCompare)
    string1 = Replace(string1, "a", "4", Compare:=vbTextCompare)
    string1 = Replace(string1, "s", "5", Compare:=vbTextCompare)
    string1 = Replace(string1, "t", "7", Compare:=vbTextCompare)

    ELITE_CASE = string1

End Function


Public Function SCRAMBLE_CASE( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string scrambles the case on each character in the string
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose character's cases will be scrambled
    '@Returns: Returns the origional string with cases scrambled
    '@Example: =SCRAMBLE_CASE("Hello World") -> "helLo WORlD"
    '@Example: =SCRAMBLE_CASE("Hello World") -> "HElLo WorLD"
    '@Example: =SCRAMBLE_CASE("Hello World") -> "hELlo WOrLd"

    Dim i As Integer

    For i = 1 To Len(string1)
        If WorksheetFunction.RandBetween(0, 1) = 1 Then
            Mid(string1, i, 1) = UCase(Mid(string1, i, 1))
        Else
            Mid(string1, i, 1) = LCase(Mid(string1, i, 1))
        End If
    Next
    
    SCRAMBLE_CASE = string1

End Function


Public Function LEFT_SPLIT( _
    ByVal string1 As String, _
    ByVal numberOfSplit As Integer, _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string, splits it based on a delimiter, and returns all characters to the left of the specified position of the split.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be split to get a substring
    '@Param: numberOfSplit is the number of the location within the split that we will get all characters to the left of
    '@Param: delimiterCharacter is the delimiter that will be used for the split. By default, the delimiter will be the space character " "
    '@Returns: Returns all characters to the left of the number of the split
    '@Example: =LEFT_SPLIT("Hello World One Two Three", 1) -> "Hello"
    '@Example: =LEFT_SPLIT("Hello World One Two Three", 2) -> "Hello World"
    '@Example: =LEFT_SPLIT("Hello World One Two Three", 3) -> "Hello World One"
    '@Example: =LEFT_SPLIT("Hello World One Two Three", 10) -> "Hello World One Two Three"
    '@Example: =LEFT_SPLIT("Hello-World-One-Two-Three", 2, "-") -> "Hello-World"

    Dim i As Integer
    Dim newString As String
    Dim stringArray() As String
    Dim stringArrayLength As Integer
    
    numberOfSplit = numberOfSplit - 1
    stringArray = SPLIT(string1, delimiterCharacter)
    stringArrayLength = (UBound(stringArray) - LBound(stringArray) + 1)
    
    ' Checking if the number of split is greater than the length of the split
    ' array, and if so returns the origional string
    If numberOfSplit >= stringArrayLength Then
        LEFT_SPLIT = string1
        Exit Function
    End If
    
    For i = 0 To numberOfSplit
        If i = numberOfSplit Then
            newString = newString & stringArray(i)
        Else
            newString = newString & stringArray(i) & delimiterCharacter
        End If
    Next
    
    LEFT_SPLIT = newString

End Function


Public Function RIGHT_SPLIT( _
    ByVal string1 As String, _
    ByVal numberOfSplit As Integer, _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string, splits it based on a delimiter, and returns all characters to the right of the specified position of the split.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be split to get a substring
    '@Param: numberOfSplit is the number of the location within the split that we will get all characters to the right of
    '@Param: delimiterCharacter is the delimiter that will be used for the split. By default, the delimiter will be the space character " "
    '@Returns: Returns all characters to the right of the number of the split
    '@Example: =RIGHT_SPLIT("Hello World One Two Three", 1) -> "Three"
    '@Example: =RIGHT_SPLIT("Hello World One Two Three", 2) -> "Two Three"
    '@Example: =RIGHT_SPLIT("Hello World One Two Three", 3) -> "One Two Three"
    '@Example: =RIGHT_SPLIT("Hello World One Two Three", 10) -> "Hello World One Two Three"
    '@Example: =RIGHT_SPLIT("Hello-World-One-Two-Three", 2, "-") -> "Two-Three"

    Dim i As Integer
    Dim newString As String
    Dim stringArray() As String
    Dim stringArrayLength As Integer
    
    numberOfSplit = numberOfSplit - 1
    stringArray = SPLIT(string1, delimiterCharacter)
    stringArrayLength = (UBound(stringArray) - LBound(stringArray) + 1)
    
    ' Checking if the number of split is greater than the length of the split
    ' array, and if so returns the origional string
    If numberOfSplit >= stringArrayLength Then
        RIGHT_SPLIT = string1
        Exit Function
    End If
    
    For i = 0 To numberOfSplit
        If i = numberOfSplit Then
            newString = newString & stringArray(stringArrayLength - (numberOfSplit - i) - 1)
        Else
            newString = newString & stringArray(stringArrayLength - (numberOfSplit - i) - 1) & delimiterCharacter
        End If
    Next
    
    RIGHT_SPLIT = newString

End Function


Public Function SUBSTITUTE_ALL( _
    ByVal string1 As String, _
    ByVal oldTextRange As Range, _
    ByVal newTextRange As Range) _
As String

    '@Description: This function takes a string, and old text range, and a new text range, and replaces all the strings in the old text range with the ones in the new text range.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will have substrings replaced
    '@Param: oldTextRange is a range containing the text that will be replaced by the text in the newTextRange
    '@Param: newTextRange is the replacement range
    '@Returns: Returns the origional string with all the replacements from the two ranges
    '@Example: =SUBSTITUTE_ALL("Hello World", A1:A2, B1:B2) -> "One Two"; Where A1:A2=["Hello", "World"] and B1:B2=["One", "Two"]

    ' Throwing an error if the ranges are not the same length and shape
    If oldTextRange.Count <> newTextRange.Count Then
        SUBSTITUTE_ALL = "#OldAndNewRangeNotSameLength!"
        Exit Function
    End If
    
    If oldTextRange.Columns.Count <> newTextRange.Columns.Count Then
        SUBSTITUTE_ALL = "#OldAndNewRangeNotSameLength!"
        Exit Function
    End If
    
    If oldTextRange.Rows.Count <> newTextRange.Rows.Count Then
        SUBSTITUTE_ALL = "#OldAndNewRangeNotSameLength!"
        Exit Function
    End If
    
    ' Performing the substitutions
    Dim i As Integer
    
    For i = 1 To oldTextRange.Count
        string1 = Replace(string1, oldTextRange(i), newTextRange(i))
    Next
    
    SUBSTITUTE_ALL = string1

End Function


Public Function INSTRING( _
    ByVal string1 As String, _
    ParamArray stringArray() As Variant) _
As Boolean

    '@Description: This function takes a string, and then any number of substrings, and will check if any of the substrings can be found within the string
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be checked against the substrings
    '@Param: stringArray is any number of substrings that will be used to check if they exist within the string
    '@Returns: Returns Boolean TRUE if any of the substrings can be found within the string, and FALSE if none are found
    '@Example: =INSTRING("112 - 312 - 221 - 132", "111", "222", "333") -> FALSE; None of the substrings are found
    '@Example: =INSTRING("123 - 222 - 122 - 311", "111", "222", "333") -> TRUE; "222" is found
    '@Example: =INSTRING("222 - 322 - 233 - 232", "111", "222", "333") -> TRUE; "222" is found
    '@Example: =INSTRING("312 - 131 - 123 - 333", "111", "222", "333") -> TRUE; "333" is found
    '@Example: =INSTRING("212 - 232 - 213 - 323", "111", "222", "333") -> FALSE
    '@Example: =INSTRING("111 - 212 - 222 - 333", "111", "222", "333") -> TRUE; "111", "222", and "333" are found

    Dim individualString As Variant
    
    For Each individualString In stringArray
        If InStr(1, string1, individualString) > 0 Then
            INSTRING = True
            Exit Function
        End If
    Next
    
    INSTRING = False

End Function


Public Function TRIM_CHAR( _
    ByVal string1 As String, _
    Optional ByVal trimCharacter As String = " ") _
As String

    '@Description: This function takes a string trims characters to the left and right of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Allow more than 1 character to be used for trimming
    '@Param: string1 is the string that will be trimmed
    '@Param: trimCharacter is an optional character that will be trimmed from the string. By default, this character will be the space character " "
    '@Returns: Returns the origional string with characters trimmed from the left and right
    '@Note: This function currently supports only single characters for trimming
    '@Example: =TRIM_CHAR("   Hello World   ") -> "Hello World"
    '@Example: =TRIM_CHAR("---Hello World---", "-") -> "Hello World"

    While Left(string1, 1) = trimCharacter
        Mid(string1, 1) = Chr(1)
        string1 = Replace(string1, Chr(1), "")
    Wend
    
    While Right(string1, 1) = trimCharacter
        Mid(string1, Len(string1)) = Chr(1)
        string1 = Replace(string1, Chr(1), "")
    Wend
    
    TRIM_CHAR = string1

End Function


Public Function TRIM_LEFT( _
    ByVal string1 As String, _
    Optional ByVal trimCharacter As String = " ") _
As String

    '@Description: This function takes a string trims characters to the left of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Allow more than 1 character to be used for trimming
    '@Param: string1 is the string that will be trimmed
    '@Param: trimCharacter is an optional character that will be trimmed from the string. By default, this character will be the space character " "
    '@Returns: Returns the origional string with characters trimmed from the left only
    '@Note: This function currently supports only single characters for trimming
    '@Example: =TRIM_LEFT("   Hello World   ") -> "Hello World   "
    '@Example: =TRIM_LEFT("---Hello World---", "-") -> "Hello World---"

    While Left(string1, 1) = trimCharacter
        Mid(string1, 1) = Chr(1)
        string1 = Replace(string1, Chr(1), "")
    Wend
    
    TRIM_LEFT = string1

End Function


Public Function TRIM_RIGHT( _
    ByVal string1 As String, _
    Optional ByVal trimCharacter As String = " ") _
As String

    '@Description: This function takes a string trims characters to the right of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Allow more than 1 character to be used for trimming
    '@Param: string1 is the string that will be trimmed
    '@Param: trimCharacter is an optional character that will be trimmed from the string. By default, this character will be the space character " "
    '@Returns: Returns the origional string with characters trimmed from the right only
    '@Note: This function currently supports only single characters for trimming
    '@Example: =TRIM_RIGHT("   Hello World   ") -> "   Hello World"
    '@Example: =TRIM_RIGHT("---Hello World---", "-") -> "---Hello World"
    
    While Right(string1, 1) = trimCharacter
        Mid(string1, Len(string1)) = Chr(1)
        string1 = Replace(string1, Chr(1), "")
    Wend
    
    TRIM_RIGHT = string1

End Function


Public Function COUNT_UPPERCASE_CHARACTERS( _
    ByVal string1 As String) _
As Integer

    '@Description: This function takes a string and counts the number of uppercase characters in it
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose characters will be counted
    '@Returns: Returns the number of uppercase characters in the string
    '@Example: =COUNT_UPPERCASE_CHARACTERS("Hello World") -> 2; As the "H" and the "E" are the only 2 uppercase characters in the string

    Dim i As Integer
    Dim characterAsciiCode As Byte
    Dim uppercaseCounter As Integer
    
    For i = 1 To Len(string1)
        characterAsciiCode = Asc(Mid(string1, i, 1))
        If characterAsciiCode >= 65 And characterAsciiCode <= 90 Then
            uppercaseCounter = uppercaseCounter + 1
        End If
    Next
    
    COUNT_UPPERCASE_CHARACTERS = uppercaseCounter

End Function


Public Function COUNT_LOWERCASE_CHARACTERS( _
    ByVal string1 As String) _
As Integer

    '@Description: This function takes a string and counts the number of lowercase characters in it
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose characters will be counted
    '@Returns: Returns the number of lowercase characters in the string
    '@Example: =COUNT_LOWERCASE_CHARACTERS("Hello World") -> 8; As the "ello" and the "orld" are lowercase

    Dim i As Integer
    Dim characterAsciiCode As Byte
    Dim lowercaseCounter As Integer
    
    For i = 1 To Len(string1)
        characterAsciiCode = Asc(Mid(string1, i, 1))
        If characterAsciiCode >= 97 And characterAsciiCode <= 122 Then
            lowercaseCounter = lowercaseCounter + 1
        End If
    Next
    
    COUNT_LOWERCASE_CHARACTERS = lowercaseCounter

End Function
