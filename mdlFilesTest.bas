' File: mdlFilesTest.bas
' Includes tests about mdlFiles.bas.


Option Compare Database
Option Explicit


' Sub: TestGetPathSeparator
' Tests <GetPathSeparator>
Public Sub TestGetPathSeparator()
    Dim testSubName As String
    testSubName = "TestGetPathSeparator"
    Dim actual As String
    actual = GetPathSeparator()
    Call AssertAreEqual(testSubName, Chr(92), actual)
End Sub


' Sub: TestPathCombine
' Tests <PathCombine>
Public Sub TestPathCombine()
    Dim testSubName As String
    testSubName = "TestPathCombine"
    ' TODO: Implement Sub TestPathCombine()
End Sub


' Sub: TestDirName
' Tests <DirName>
Public Sub TestDirName()
    Dim testSubName As String
    testSubName = "TestDirName"
    ' TODO: Implement Sub TestDirName()
End Sub


' Sub: TestFileExists
' Tests <FileExists>
Public Sub TestFileExists()
    Dim testSubName As String
    testSubName = "TestFileExists"
    ' TODO: Implement Sub TestFileExists()
End Sub


' Sub: TestDeleteFileIfExists
' Tests <DeleteFileIfExists>
Public Sub TestDeleteFileIfExists()
    Dim testSubName As String
    testSubName = "TestDeleteFileIfExists"
    ' TODO: Implement Sub TestDeleteFileIfExists()
End Sub


' Sub: TestWriteTextFile
' Tests <WriteTextFile>
Public Sub TestWriteTextFile()
    Dim testSubName As String
    testSubName = "TestWriteTextFile"
    ' TODO: Implement Sub TestWriteTextFile()
End Sub


' Sub: TestReadTextFile
' Tests <ReadTextFile>
Public Sub TestReadTextFile()
    Dim testSubName As String
    testSubName = "TestReadTextFile"
    ' TODO: Implement Sub TestReadTextFile()
End Sub


' Sub: RunAllMdlFilesTest
' Calls the tests in this module.
'
' See also:
' <RunAllUnitTests>
Public Sub RunAllMdlFilesTest()
    Call TestPathCombine()
    Call TestDirName()
    Call TestFileExists()
    Call TestDeleteFileIfExists()
    Call TestWriteTextFile()
    Call TestReadTextFile()
End Sub
