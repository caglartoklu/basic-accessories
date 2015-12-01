' File: mdlFiles.bas
' Includes the functions and subroutines about files and folders.


Option Compare Database
Option Explicit


' Function: GetPathSeparator
' Returns the path separator for the operating system.
' Since Microsoft Access is a Windows only software,
' the path separator is always: \
'
' Returns:
' the path separator of the operating system, \ for Windows.
Public Function GetPathSeparator() As String
    GetPathSeparator = CHR(92) ' \
End Function


' Function: PathCombine
' Combines 2 paths ensuring that there is a path separator between them.
'
' Parameters:
' path1 - A path, head in this case
' path2 - A path, tail in this case
'
' Returns:
' Combined path
Public Function PathCombine(ByVal path1 As String, ByVal path2 As String) As String
    Dim result As String
    result = ""
    If EndsWith(path1, GetPathSeparator()) Then
        result = path1 & path2
    Else
        result = path1 & GetPathSeparator() & path2
    End If
    PathCombine = result
End Function


' Function: DirName
' Returns the containing path.
'
' Parameters:
' path - A path, it can be directory or a path to a file.
'
' Returns:
' The containing path.
Public Function DirName(path As String) As String
    DirName = Left(path, InStrRev(path, GetPathSeparator(), Len(path)) - 1)
End Function


' Function: FileExists
' Returns True if the file exists, False otherwise.
'
' Parameters:
' fileToTest - The path to file to test for existence.
'
' Returns:
' True if the file exists, False otherwise.
Public Function FileExists(ByVal fileToTest As String) As Boolean
    Dim result As Boolean
    result = False
    If Len(Dir(fileToTest)) > 0 Then
        result = True
    End If
    FileExists = result
End Function


' Sub: DeleteFileIfExists
' Deletes a file if it exists.
'
' Parameters:
' fileName - name of the file to be deleted
Public Sub DeleteFileIfExists(ByVal fileName As String)
    If FileExists(fileName) Then
        ' SetAttr fileName, vbNormal
        Kill fileName
    End If
End Sub


' Sub: WriteTextFile
' Writes content to a text file.
'
' Parameters:
' fileName - Name of the file to be written.
' content - The content to be written into the file.
Public Sub WriteTextFile(ByVal fileName As String, ByVal content As String)
    Dim fileHandle As Integer
    fileHandle = FreeFile
    Open fileName For Output As #fileHandle
    Print #fileHandle, content
    Close #fileHandle
End Sub


' Sub: ReadTextFile
' Returns all the contents of a text file.
'
' Parameters:
' fileName - Name of the file to be read.
'
' Returns:
' Content of the specified file.
Public Function ReadTextFile(ByVal fileName As String) As String
    Dim content As String
    Dim currentLine As String
    Dim fileHandle As Integer
    fileHandle = FreeFile
    Open fileName For Input As #fileHandle
    content = ""
    While Not EOF(fileHandle)
        Line Input #fileHandle, currentLine
        content = content & currentLine & vbCrLf
    Wend
    Close #fileHandle
    ReadTextFile = content
End Function


' Sub: WriteTextFileUTF8
' Writes content to a text file.
' SaveOptions:
' https://msdn.microsoft.com/en-us/library/ms676152%28v=vs.85%29.aspx
'
' Parameters:
' fileName - Name of the file to be written.
' content - The content to be written into the file.
Public Sub WriteTextFileUTF8(ByVal fileName As String, ByVal content As String)
    Const adSaveCreateNotExist = 1
    Const adSaveCreateOverWrite = 2
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.CharSet = "utf-8"
    objStream.Open
    objStream.WriteText content
    objStream.SaveToFile fileName, adSaveCreateOverWrite
End Sub


' Sub: ReadTextFileUTF8
' Returns all the contents of a text file.
'
' Parameters:
' fileName - Name of the file to be read.
'
' Returns:
' Content of the specified file.
Public Function ReadTextFileUTF8(ByVal fileName As String) As String
    Dim content As String
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    objStream.CharSet = "utf-8"
    objStream.Open
    objStream.LoadFromFile(fileName)
    content = objStream.ReadText()
    ReadTextFileUTF8 = content
End Function


' Function: DetermineEndOfLineChar
' Determines the end of line character used in the file.
' If more than one end of line character has been used in the file,
' the first one will be returned.
'
' Parameters:
' fileName - Name of the file to be read.
'
' Returns:
' The end of line character used in the file.
Private Function DetermineEndOfLineChar(ByVal fileName As String) As String
    Dim fileContent As String
    Dim fileHandle As Integer
    fileHandle = FreeFile
    Open fileName For Input As fileHandle
    fileContent = Input(LOF(fileHandle), fileHandle)
    Close fileHandle

    Dim iVbCrLf As Long
    Dim iVbCr As Long
    Dim iVbLf As Long

    iVbCrLf = InStr(1, fileContent, vbCrlf)
    iVbCr = InStr(1, fileContent, vbCr)
    iVbLf = InStr(1, fileContent, vbLf)

    If iVbCrLf = -1 Then
        iVbCrLf = Len(fileContent) + 1
    End If
    If iVbCr = -1 Then
        iVbCr = Len(fileContent) + 1
    End If
    If iVbLf = -1 Then
        iVbLf = Len(fileContent) + 1
    End If

    Dim eolChar As String

    If iVbCrLf > iVbCr And iVbCrLf > iVbLf Then
        eolChar = vbCrlf
    ElseIf iVbCr > iVbCrLf And iVbCr > iVbLf Then
        eolChar = vbCr
    ElseIf iVbLf > iVbCrLf And iVbLf > iVbCr Then
        eolChar = vbLf
    Else
        ' this is the default
        eolChar = vbCrlf
    End If

    DetermineEndOfLineChar = eolChar
End Function


' Sub: CreateFolderRecursively
' Creates folders recursively.
' It is compatible with Windows Shared folders.
' This function does not return throw error.
' It is advised to check the folder after calling this sub.
'
' Parameters:
' folderPath - the full path of the folder to be created
Public Sub CreateFolderRecursively(ByVal folderPath As String)
    On Error Resume Next
    Dim folderPathTemp As String
    folderPathTemp = folderPath
    
    Dim pathSeparator As String
    pathSeparator = GetPathSeparator()
    Dim windowsSharedPathStarter As String
    windowsSharedPathStarter = pathSeparator & pathSeparator
    
    Dim folderPathBase As String
    folderPathBase = ""
    If StartsWith(folderPath, windowsSharedPathStarter) Then
        folderPathBase = windowsSharedPathStarter
        folderPath = RemoveFromLeft(folderPath, Len(windowsSharedPathStarter))
    End If
    
    Dim items() As String
    items = Split(folderPath, pathSeparator)

    Dim newFolderPath As String
    newFolderPath = folderPathBase
    Dim i As Integer
    For i = LBound(items) To UBound(items)
        newFolderPath = newFolderPath + items(i) + pathSeparator
        MkDir (newFolderPath)
    Next i
    On Error GoTo 0
End Sub
