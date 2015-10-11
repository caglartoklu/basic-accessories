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


' Function: ParseToQueries
' Splits a string that contains multiple queries to an array.
'
' Parameters:
' queriesAsString - The string that includes multiple queries.
'
' Returns:
' An array of parsed queries.
Public Function ParseToQueries(ByVal queriesAsString As String) As String()
    ' TODO: 5 unit test ParseToQueries()
    Const separator = ";"
    queriesAsString = Trim(queriesAsString)
    If Not EndsWith(queriesAsString, separator) Then
        queriesAsString = queriesAsString & separator
    End If
    Const defaultQuerySize = 16
    ReDim queries(defaultQuerySize) As String
    Dim qCapacity As Integer
    Dim qIndex As Integer
    qCapacity = UBound(queries)
    qIndex = 0
    Dim inString As Boolean

    Dim buffer As String
    Dim i As Long
    Dim ch As String
    inString = False
    For i = 1 To Len(queriesAsString)
        ch = Mid(queriesAsString, i, 1)

        ' add this character anyway
        buffer = buffer + ch

        If ch = "'" Then
            Call ToggleBoolean(inString)
        End If

        If Not inString Then
            If ch = separator Then
                ' not in a string, and ; character is found.
                ' that means, this is the end of a query.
                If qIndex = qCapacity Then
                    ReDim Preserve queries(qCapacity + defaultQuerySize)
                    qCapacity = UBound(queries)
                End If
                qIndex = qIndex + 1
                queries(qIndex) = buffer
                buffer = ""
            End If
        End If
    Next i
    ParseToQueries = queries
End Function


' Sub: ExecuteNonQuery
' Executes a single non-query (INSERT, UPDATE, DELETE)
'
' Parameters:
' query - The query to be executed
Public Sub ExecuteNonQuery(ByVal query As String)
    ' TODO: 5 unit test ExecuteNonQuery()
    query = Strip(query)
    If Len(query) > 0 Then
        If Len(Replace(query, ";", "")) = 0 Then
            ' if all characters of the query are ;
            ' then do not execute the query.
        Else
            DoCmd.SetWarnings False
            DoCmd.RunSQL query
            DoCmd.SetWarnings True
            ' TODO: 6 DoCmd.SetWarnings True in a try-finally block
        End If
    End If
End Sub


' Sub: ExecuteNonQueriesFromFile
' Executes batch non-queries (INSERT, UPDATE, DELETE) from the provided file name.
' The queries are separated by ;
' Multi-line queries are supported.
'
' Parameters:
' fileName - Name of the file
Public Sub ExecuteNonQueriesFromFile(ByVal fileName As String)
    ' TODO: 5 unit test ExecuteNonQueriesFromFile()
    Dim data As String
    data = ReadTextFile(fileName)
    Call ExecuteNonQueriesFromString(data)
End Sub


' Sub: ExecuteNonQueriesFromString
' Executes batch non-queries (INSERT, UPDATE, DELETE) from the provided String.
' The queries are separated by ;
' Multi-line queries are supported.
'
' Parameters:
' data - A big String that includes INSERT, UPDATE, DELETE queries separated by ;
Public Sub ExecuteNonQueriesFromString(ByVal data As String)
    ' TODO: 5 unit test ExecuteNonQueriesFromString()
    Dim queries() As String
    queries = ParseToQueries(data)
    Dim i As Integer
    For i = LBound(queries) To UBound(queries)
        Dim query As String
        query = Strip(queries(i))
        If Len(query) > 0 Then
            Call ExecuteNonQuery(query)
        End If
    Next i
End Sub

