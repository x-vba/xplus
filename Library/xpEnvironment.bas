Attribute VB_Name = "xpEnvironment"
'@Module: This module contains a set of functions for gathering information on the environment that Excel is being run on, such as getting the location of the SystemDrive or the location of the temporary folder on the system

Option Explicit


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


