Attribute VB_Name = "xpMeta"
'@Module: This module contains a set of functions that return information on the XPlus library, such as the version number, credits, and a link to the documentation.

Option Explicit


Public Function VERSION() As String

    '@Description: This function returns the version number of XPlus
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the XPlus version number
    '@Example: =VERSION() -> "1.2.0"; Where the version of XPlus you are using is 1.2.0

    VERSION = "1.2.0"

End Function


Public Function CREDITS() As String

    '@Description: This function returns credits for the XPlus library
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the XPlus credits
    '@Example: =CREDITS() -> "Copyright (c) 2020 Anthony Mancini. XPlus is Licensed under an MIT License."

    CREDITS = "Copyright (c) 2020 Anthony Mancini. XPlus is Licensed under an MIT License."

End Function


Public Function DOCUMENTATION() As String

    '@Description: This function returns a link to the Documentation for XPlus
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the XPlus Documentation link
    '@Example: =DOCUMENTATION() -> "https://x-vba.com"

    DOCUMENTATION = "https://x-vba.com"

End Function

