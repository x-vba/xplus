Attribute VB_Name = "xpEnvironment"
'@Module: This module contains a set of functions for gathering information on the environment that Excel is being run on, such as the UserName of the computer, the OS Excel is being run on, and other Environment Variable values.

Option Explicit


Public Function ENVIRONMENT( _
    ByVal environmentVariableNameString As String) _
As String

    '@Description: This function takes a string of the name of an environment variable and returns the value of that EV as a string.
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Param: environmentVariableNameString is the string of the environment variable name.
    '@Returns: Returns a string of the environment variable value associated with that name.
    '@Note: A list of Environment Variable Key/Value pairs can be found by using the Set command on the Command Prompt.
    '@Example: =ENVIRONMENT("HOMEDRIVE") -> "C:"
    '@Example: =ENVIRONMENT("PUBLIC") -> "C:\Users\Public"

    ENVIRONMENT = Environ(environmentVariableNameString)

End Function


Public Function OS() As String

    '@Description: This function returns the Operating System name. Currently it will return either "Windows" or "Mac" depending on the OS used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the name of the Operating System
    '@Example: =OS() -> "Windows"; When running this function on Windows
    '@Example: =OS() -> "Mac"; When running this function on MacOS

    #If Mac Then
        OS = "Mac"
    #Else
        OS = "Windows"
    #End If

End Function


Public Function USER_NAME() As String

    '@Description: This function takes no arguments and returns a string of the USERNAME of the computer
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Returns: Returns a string of the username
    '@Example: =USER_NAME() -> "Anthony"

    #If Mac Then
        USER_NAME = Environ("USER")
    #Else
        USER_NAME = Environ("USERNAME")
    #End If

End Function


Public Function USER_DOMAIN() As String

    '@Description: This function takes no arguments and returns a string of the USERDOMAIN of the computer
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Returns: Returns a string of the user domain of the computer
    '@Example: =USER_DOMAIN() -> "DESKTOP-XYZ1234"
    
    #If Mac Then
        USER_DOMAIN = Environ("HOST")
    #Else
        USER_DOMAIN = Environ("USERDOMAIN")
    #End If

End Function


Public Function COMPUTER_NAME() As String

    '@Description: This function takes no arguments and returns a string of the COMPUTERNAME of the computer
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Returns: Returns a string of the computer name of the computer
    '@Example: =COMPUTER_NAME() -> "DESKTOP-XYZ1234"

    COMPUTER_NAME = Environ("COMPUTERNAME")

End Function
