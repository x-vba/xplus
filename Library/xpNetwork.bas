Attribute VB_Name = "xpNetwork"
'@Module: This module contains a set of functions for getting the UserName and ComputerName and other values from the Network Module.

Option Explicit


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
