Attribute VB_Name = "xpFile"
'@Module: This module contains a set of functions for gathering info on files. It includes functions for gathering file info on the current workbook, as well as functions for reading and writing to files, and functions for manipulating file path strings.

Option Explicit


Private Function GetActiveWorkbookPath() As Variant

    '@Description: This function returns the path of the current workbook
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current workbook path

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
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file creation time of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
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
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file last modified time of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FILE_LAST_MODIFIED_TIME() -> "1/1/2020 2:23:45 PM"
    '@Example: =FILE_LAST_MODIFIED_TIME("C:\hello\world.txt") -> "1/1/2020 7:55:55 PM"
    '@Example: =FILE_LAST_MODIFIED_TIME("vba.txt") -> "12/25/2000 3:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook this function resides in

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
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file drive of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FILE_DRIVE() -> "A:"; Where the current workbook resides on the A: drive
    '@Example: =FILE_DRIVE("C:\hello\world.txt") -> "C:"
    '@Example: =FILE_DRIVE("vba.txt") -> "B:"; Where "vba.txt" resides in the same folder as the workbook this function resides in, and where the workbook resides in the B: drive

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
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the name of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
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
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the path of the folder where the file resides in as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
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
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the path of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
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
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Param: byteSize is a string of value "KB", "MB", or "GB"
    '@Returns: Returns the size of the file as a Double
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
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
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file type of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
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
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the extension of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
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

    '@Description: This function reads the file specified in the file path argument and returns it's contents. Optionally, a line number can be specified so that only a single line is read. If a full path isn't used, a path relative to the folder the workbook resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Param: lineNumber is the number of the line that will be read, and if left blank all the file contents will be read. Note that the first line starts at line number 1.
    '@Returns: Returns the contents of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
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
        
        fileLinesArray = Split(fileStream.ReadAll(), vbCrLf)
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
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Param: fileText is the text that will be written to the file
    '@Param: appendModeFlag is a Boolean value that if set to TRUE will append to the existing file instead of creating a new file and writing over the contents.
    '@Returns: Returns a message stating the file written to successfully
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: Be careful when writing files, as misuse of this function can results in files being overwritten accidently as well as creating large numbers of files accidently.
    '@Example: =WRITE_FILE("C:\MyWorkbookFolder\hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"
    '@Example: =WRITE_FILE("hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"; Where the Workbook resides in "C:\MyWorkbookFolder\"

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

    '@Description: This function combines multiple strings or a range of values into a file path by placing the separator "\" between the arguments
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


Public Function COUNT_FILES( _
    Optional ByVal filePath As String) _
As Integer

    '@Description: This function returns the number of files at the specified folder path. If no path is given, the current workbook path will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the number of files in the folder
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes the count for hidden files as well. For example, when a workbook is open, a hidden file for the workbook is created, so if you run this function in the same folder as the workbook and notice the file count is one higher than expected, it is likely due to the hidden file.
    '@Example: =COUNT_FILES() -> 6
    '@Example: =COUNT_FILES("C:\hello") -> 10

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        COUNT_FILES = FSO.GetFolder(FSO.GetParentFolderName(GetActiveWorkbookPath())).Files.Count
    Else
        If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
            COUNT_FILES = FSO.GetFolder(ThisWorkbook.Path & "\" & filePath).Files.Count
        Else
            COUNT_FILES = FSO.GetFolder(filePath).Files.Count
        End If
    End If

End Function


Public Function COUNT_FOLDERS( _
    Optional ByVal filePath As String) _
As Integer

    '@Description: This function returns the number of folders at the specified folder path. If no path is given, the current workbook path will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the number of folders in the folder
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes the count for hidden folders as well. Hidden folders are often prefixed with a . character at the beginning
    '@Example: =COUNT_FOLDERS() -> 2
    '@Example: =COUNT_FOLDERS("C:\hello") -> 20

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        COUNT_FOLDERS = FSO.GetFolder(FSO.GetParentFolderName(GetActiveWorkbookPath())).SubFolders.Count
    Else
        If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
            COUNT_FOLDERS = FSO.GetFolder(ThisWorkbook.Path & "\" & filePath).SubFolders.Count
        Else
            COUNT_FOLDERS = FSO.GetFolder(filePath).SubFolders.Count
        End If
    End If

End Function


Public Function COUNT_FILES_AND_FOLDERS( _
    Optional ByVal filePath As String) _
As Integer

    '@Description: This function returns the number of files and folders at the specified folder path. If no path is given, the current workbook path will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the number of files and folders in the folder
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes the count for hidden files and folders as well
    '@Example: =COUNT_FILES_AND_FOLDERS() -> 8
    '@Example: =COUNT_FILES_AND_FOLDERS("C:\hello") -> 30

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        COUNT_FILES_AND_FOLDERS = FSO.GetFolder(FSO.GetParentFolderName(GetActiveWorkbookPath())).Files.Count + FSO.GetFolder(FSO.GetParentFolderName(GetActiveWorkbookPath())).SubFolders.Count
    Else
        If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
            COUNT_FILES_AND_FOLDERS = FSO.GetFolder(ThisWorkbook.Path & "\" & filePath).Files.Count + FSO.GetFolder(ThisWorkbook.Path & "\" & filePath).SubFolders.Count
        Else
            COUNT_FILES_AND_FOLDERS = FSO.GetFolder(filePath).Files.Count + FSO.GetFolder(filePath).SubFolders.Count
        End If
    End If

End Function


Public Function GET_FILE_NAME( _
    Optional ByVal filePath As String, _
    Optional ByVal fileNumber As Integer = -1) _
As String

    '@Description: This function returns the name of a file in a folder given the number of the file in the list of all files
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Param: fileNumber is the number of the file in the folder. For example, if there are 3 files in a folder, this should be a number between 1 and 3
    '@Returns: Returns the name of the specified file
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes hidden files as well. For example, when a workbook is open, a hidden file for the workbook is created, so if you run this function in the same folder as the workbook and notice the file count is one higher than expected, it is likely due to the hidden file.
    '@Example: =GET_FILE_NAME(,1) -> "hello.txt"
    '@Example: =GET_FILE_NAME(,1) -> "world.txt"
    '@Example: =GET_FILE_NAME("C:\hello", 1) -> "one.txt"
    '@Example: =GET_FILE_NAME("C:\hello", 1) -> "two.txt"
    '@Example: =GET_FILE_NAME("C:\hello", 1) -> "three.txt"

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim fileCounter As Integer
    Dim individualFile As Object
    Dim fileCollection As Object
    
    If filePath = "" Then
        Set fileCollection = FSO.GetFolder(FSO.GetParentFolderName(GetActiveWorkbookPath())).Files
    Else
        If FSO.FileExists(ThisWorkbook.Path & "\" & filePath) Then
            Set fileCollection = FSO.GetFolder(ThisWorkbook.Path & "\" & filePath).Files
        Else
            Set fileCollection = FSO.GetFolder(filePath).Files
        End If
    End If
    
    For Each individualFile In fileCollection
        fileCounter = fileCounter + 1
        If fileNumber = -1 Then
            GET_FILE_NAME = individualFile.Name
            Exit Function
        ElseIf fileCounter = fileNumber Then
            GET_FILE_NAME = individualFile.Name
            Exit Function
        End If
    Next

End Function
