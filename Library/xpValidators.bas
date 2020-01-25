Attribute VB_Name = "xpValidators"
'@Module: This module contains a set of functions for validating some commonly used string, such as validators for email addresses and phone numbers.

Option Explicit


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

