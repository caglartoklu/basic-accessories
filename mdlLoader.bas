' File: mdlLoader.bas
' mdlLoader is a special module that will not be removed or imported.
' For this reason, it duplicates some of the functions/subs in other modules on purpose.
' Please note that any changes to this file will not be directly
' imported to the .accdb file.
' Instead, one must manually keep this file and the module in the .accdb file in sync.
' Simply type:
' | Call mdlLoader.ImportModulesFromDisk()
' to load the modules to the file.


Option Compare Database
Option Explicit


' Constant: MODULES_FILE
' Holds the name of file that contains the list of modules to be
' imported to .accdb file.
Public Const MODULES_FILE = "basicaccessories_modules_to_import.txt"


' ---------- BEGIN functions/subs duplicated into this folder to remove the dependency to other modules

' These functions are not documented to prevent the duplicate
' generation of documentation tags since these are actually copies.
' To see the the corresponding documentation, please see the original one.


Private Function GetDatabasePath() As String
    ' original in: mdlDatabase.bas
    GetDatabasePath = CurrentProject.Path
End Function


Private Function StartsWith(ByVal haystack As String, ByVal needle As String) As Boolean
    ' original in: mdlStrings.bas
    Dim result As Boolean
    result = False
    If Left(haystack, Len(needle)) = needle Then
        result = True
    End If
    StartsWith = result
End Function


Private Function EndsWith(ByVal haystack As String, ByVal needle As String) As Boolean
    ' original in: mdlStrings.bas
    Dim result As Boolean
    result = False
    If Right(haystack, Len(needle)) = needle Then
        result = True
    End If
    EndsWith = result
End Function


Private Function IsComment(ByVal haystack As String) As Boolean
    ' original in: mdlStrings.bas
    Dim result As Boolean
    haystack = Trim(haystack)
    result = False
    If StartsWith(haystack, "'") = True Then
        result = True
    ElseIf StartsWith(haystack, "#") = True Then
        result = True
    ElseIf StartsWith(haystack, ";") = True Then
        result = True
    ElseIf StartsWith(haystack, "//") = True Then
        result = True
    End If
    IsComment = result
End Function


Private Function GetPathSeparator() As String
    ' original in: mdlFiles.bas
    GetPathSeparator = Chr(92) ' \
End Function


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

    iVbCrLf = InStr(1, fileContent, vbCrLf)
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
        eolChar = vbCrLf
    ElseIf iVbCr > iVbCrLf And iVbCr > iVbLf Then
        eolChar = vbCr
    ElseIf iVbLf > iVbCrLf And iVbLf > iVbCr Then
        eolChar = vbLf
    Else
        ' this is the default
        eolChar = vbCrLf
    End If

    DetermineEndOfLineChar = eolChar
End Function


Private Function PathCombine(ByVal path1 As String, ByVal path2 As String) As String
    ' original in: mdlFiles.bas
    Dim result As String
    result = ""
    If mdlLoader.EndsWith(path1, mdlLoader.GetPathSeparator()) = 1 Then
        result = path1 & path2
    Else
        result = path1 & mdlLoader.GetPathSeparator() & path2
    End If
    PathCombine = result
End Function


Private Sub FillStringArrayFromTextFile(ByRef arrContent() As String, ByVal fileName As String)
    ' TODO: copy this sub to a library module.
    Dim fileHandle As Integer
    Dim fileContent As String
    Dim eolChar As String
    eolChar = DetermineEndOfLineChar(fileName)

    fileHandle = FreeFile
    Open fileName For Input As fileHandle
    fileContent = Input(LOF(fileHandle), fileHandle)
    ' Debug.Print Len(fileContent)
    Close fileHandle
    arrContent = Split(fileContent, eolChar)
End Sub

' ---------- END functions/subs duplicated into this folder to remove the dependency to other modules


' Sub: FillModulesList
' Reads and fills the module list array.
'
' Parameters:
' arrModules: The array to be redimensioned and filled with the valid module names.
Public Sub FillModulesList(arrModules() As String)
    ChDir (mdlLoader.GetDatabasePath())
    Call FillStringArrayFromTextFile(arrModules, MODULES_FILE)
End Sub


' Function: FindVBComponent
' Returns the component specified by name as an object, Nothing otherwise.
'
' Parameters:
' componentName - The name of the component, such as a module.
'
' Returns:
' The component specified by name as an object, Nothing otherwise.
Public Function FindVBComponent(ByVal componentName As String) As Variant
    Dim component As Variant
    Dim result As Variant
    Set result = Nothing
    For Each component In VBE.ActiveVBProject.VBComponents
        If component.Name = componentName Then
            Set result = component
        End If
    Next
    Set FindVBComponent = result
End Function


' Sub: ImportModulesFromDisk
' Imports the modules specified by the <MODULES_FILE> to .accdb file.
' Already existing files are overwritten. New files are added.
' <mdlLoader.bas> is never imported.
'
' See also:
' <MODULES_FILE>
Public Sub ImportModulesFromDisk()
    Dim arrModules() As String
    Call FillModulesList(arrModules)

    Dim moduleName As String
    Dim onlyFileName As String
    Dim fullFileName As String
    Dim i As Integer
    Dim lastModule As Variant
    Dim component As Variant

    For i = LBound(arrModules) To UBound(arrModules)
        onlyFileName = arrModules(i)

        If Len(Trim(onlyFileName)) > 0 Then
            ' if there is a non empty line
            If IsComment(onlyFileName) = 0 Then
                On Error Resume Next
                ' if the line is not a comment
                moduleName = Left(onlyFileName, Len(onlyFileName) - Len(".bas"))
                fullFileName = mdlLoader.PathCombine(mdlLoader.GetDatabasePath(), onlyFileName)

                Set component = FindVBComponent(moduleName)
                If component Is Nothing Then
                    Debug.Print ("New module to add : " & moduleName)
                Else
                    Debug.Print ("Module to overwrite : " & moduleName)
                    VBE.ActiveVBProject.VBComponents.Remove (component)
                End If

                Set lastModule = VBE.ActiveVBProject.VBComponents.Import(fullFileName)
                lastModule.Name = moduleName
                DoCmd.Save acModule, moduleName
                If Err.Number <> 0 Then
                    Debug.Print ("Module Error in : " & moduleName & " : " & Err.Number & " : " & Err.Description)
                End If
                On Error GoTo 0
                Err.Clear
            End If
        End If
    Next i
End Sub


' Sub: RemoveUnregisteredModules
' Removes all the modules from .accdb that are not registered
' in the <MODULES_FILE> file.
' This function is never called by the software itself.
' It can only be called manually by the users,
' (by typing its name to Immediate window).
' Be careful about calling this sub.
'
' Parameters:
' some4LetterNumber - string including any four digit number.
' It is a protection to make sure that the sub can only be called
' by humans, and never automated software.
Public Sub RemoveUnregisteredModules(ByVal some4LetterNumber As String)
    If Len(some4LetterNumber) = 4 Then
        If IsNumeric(some4LetterNumber) Then
            ' TODO: Implement RemoveUnregisteredModules()
        End If
    End If
End Sub
