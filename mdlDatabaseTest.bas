' File: mdlDatabaseTest.bas
' Includes tests about mdlDatabase.bas.


Option Compare Database
Option Explicit


' Sub: TestGetDatabasePath
' Tests <GetDatabasePath>
' It ensures there at least one : and one \ characters
' in the returned path.
Public Sub TestGetDatabasePath()
    Dim testSubName As String
    testSubName = "TestGetDatabasePath"

    Dim databasePath As String
    databasePath = GetDatabasePath()
    Dim position As Integer
    position = InStr(databasePath, ":")
    Call AssertTrue(testSubName, position > 0)
    position = InStr(databasePath, GetPathSeparator())
    Call AssertTrue(testSubName, position > 0)
End Sub


' Sub: TestGetDatabaseName
' Tests <GetDatabaseName>
' It ensures that the returned value Is a file name without path info.
Public Sub TestGetDatabaseName()
    Dim testSubName As String
    testSubName = "TestGetDatabaseName"

    Dim databaseName As String
    databaseName = GetDatabaseName()

    Call AssertTrue(testSubName, EndsWith(LCase(databaseName), ".accdb"))
    Call AssertFalse(testSubName, StartsWith(LCase(databaseName), ".accdb"))
    Call AssertAreEqual(testSubName, 0, StrCount(databaseName, "/"))
    Call AssertAreEqual(testSubName, 0, StrCount(databaseName, "\"))
End Sub


' Sub: TestTableExists
' Tests <TableExists>
Public Sub TestTableExists()
    Dim testSubName As String
    testSubName = "TestTableExists"
    ' TODO: Implement Sub TestTableExists()
End Sub


' Sub: TestQueryExists
' Tests <QueryExists>
Public Sub TestQueryExists()
    Dim testSubName As String
    testSubName = "TestQueryExists"
    ' TODO: Implement Sub TestQueryExists()
End Sub


' Sub: TestCreateQueryObject
' Tests <CreateQueryObject>
Public Sub TestCreateQueryObject()
    Dim testSubName As String
    testSubName = "TestCreateQueryObject"
    ' TODO: Implement Sub TestCreateQueryObject()
End Sub


' Sub: TestDeleteQuery
' Tests <DeleteQuery>
Public Sub TestDeleteQuery()
    Dim testSubName As String
    testSubName = "TestDeleteQuery"
    ' TODO: Implement Sub TestDeleteQuery()
End Sub


' Sub: RunAllMdlDatabaseTest
' Calls the tests in this module.
'
' See also:
' <RunAllUnitTests>
Public Sub RunAllMdlDatabaseTest()
    Call TestGetDatabasePath
    Call TestGetDatabaseName
    Call TestTableExists
    Call TestQueryExists
    Call TestCreateQueryObject
    Call TestDeleteQuery
End Sub
