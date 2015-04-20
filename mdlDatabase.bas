' File: mdlDatabase.bas
' Includes the functions and subroutines about database.
' While other modules are compatible with other components of
' Microsoft Office, this module is Microsoft Access specific.


Option Compare Database
Option Explicit


' Function: GetDatabasePath
' Returns the directory of the current Access file
'
' visit http://www.ammara.com/access_image_faq/get_mdb_database_path.html
' to get information about database paths for older versions.
'
' Returns:
' the directory of the current Access file
Public Function GetDatabasePath() As String
    GetDatabasePath = CurrentProject.path
End Function


' Function: TableExists
' Returns True if the table exists, False otherwise.
'
' Parameters:
' tableName - Name of the table
'
' Returns:
' True if the table exists, False otherwise.
Public Function TableExists(ByVal tableName As String) As Boolean
    Dim element As TableDef
    Dim result As Boolean
    result = False
    For Each element In CurrentDb.TableDefs
        If element.Name = tableName Then
            result = True
            Exit For
        End If
    Next
    TableExists = result
End Function


' Function: QueryExists
' Returns True if the query exists, False otherwise.
'
' Parameters:
' queryName - Name of the query
'
' Returns:
' True if the query exists, False otherwise.
Public Function QueryExists(ByVal queryName As String) As Boolean
    Dim element As QueryDef
    Dim result As Boolean
    result = False
    For Each element In CurrentDb.QueryDefs
        If element.Name = queryName Then
            result = True
            Exit For
        End If
    Next
    QueryExists = result
End Function


' Function: CreateQueryObject
' Creates a query and returns it as a QueryDef object.
'
' Parameters:
' queryName - Name of the query
' sql - SQL of the query
'
' Returns:
' Created QueryDef object
Public Function CreateQueryObject(ByVal queryDefName As String, ByVal sql As String) As QueryDef
    Set CreateQueryObject = CurrentDb.CreateQueryDef(queryDefName, sql)
End Function


' Sub: DeleteQuery
' Deletes the specified query if it exists.
'
' Parameters:
' queryName - Name of the query
Public Sub DeleteQuery(ByVal queryName As String)
    If QueryExists(queryName) Then
        DoCmd.DeleteObject acQuery, queryName
    End If
End Sub
