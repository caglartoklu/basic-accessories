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
    Call TestTableExists
    Call TestQueryExists
    Call TestCreateQueryObject
    Call TestDeleteQuery
End Sub
