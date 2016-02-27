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


' Function: GetDatabaseName
'
' Returns:
' The name of the database, such as "MainProject.accdb"
Public Function GetDatabaseName() As String
    ' MainProject.accdb
    GetDatabaseName = CurrentProject.Name
End Function


' Function: GetBackendProjectName
' Returns a possible backend name for split database architecture.
' If suffix is "DataBE" then the file name would be:
' MyProjectDataBE.accdb
'
' Parameters:
' suffix - string, the suffix to be added to the current project name just before the extension.
'
' Returns:
' Only the file name of the backend database, such as "MyProjectDataBE.accdb"
Public Function GetBackendProjectName(ByVal suffix As String)
    ' TODO: 5 unit test for GetBackendProjectName
    If Strip(suffix) = "" Then
        ' do not let the the software to be without suffix.
        ' if there would be no suffix, the table would connect to front end database,
        ' which is not preferred.
        ' instead, provide a default suffix name.
        suffix = "DataBackEnd"
    End If

    Dim projectName As String
    projectName = CurrentProject.Name

    Dim positionOfExtension As Integer
    positionOfExtension = InStr(LCase(projectName), ".accdb")

    Dim fileNameWithoutExtension As String
    fileNameWithoutExtension = Left(projectName, positionOfExtension - 1)

    GetBackendProjectName = fileNameWithoutExtension & suffix & ".accdb"
End Function


' Function: GetBackendProjectPath
' Returns the full path for a possible backend file for split database architecture.
' If suffix is "DataBE" then the file name would be:
' MyProjectDataBE.accdb
'
' Parameters:
' suffix - string, the suffix to be added to the current project name just before the extension.
'
' Returns:
' Only the file name of the backend database, such as "MyProjectDataBE.accdb"
Public Function GetBackendProjectPath(ByVal suffix As String)
    ' TODO: 5 unit test for GetBackendProjectPath
    GetBackendProjectPath = GetDatabasePath() & "\" & GetBackendProjectName(suffix)
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
    ' TODO: 6 is it possible to make it faster directly using TableDefs("xxx")?
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


' Function: GetTableDefByName
' Returns the TableDef object for the specified table name.
'
' Parameters:
' tableName - Name of the table
'
' Returns:
' the TableDef object for the specified table name.
Public Function GetTableDefByName(ByVal tableName As String) As TableDef
    Set GetTableDefByName = CurrentDb.TableDefs(tableName)
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
    ' TODO: 6 is it possible to make it faster directly using QueryDefs("xxx")?
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


' Function: GetQueryDefByName
' Returns a QueryDef object matched by name.
'
' Parameters:
' queryDefName - Name of the query
'
' Returns:
' a QueryDef object matched by name.
Public Function GetQueryDefByName(ByVal queryDefName As String) As QueryDef
    ' TODO: 5 implement it LIKE GetTableDefByName, without a loop
    Dim element As QueryDef
    Dim result As QueryDef
    Set result = Nothing
    For Each element In CurrentDb.QueryDefs
        If element.Name = queryDefName Then
            Set result = element
            Exit For
        End If
    Next
    Set GetQueryDefByName = result
End Function


' Function: GetQueryOfQueryDef
' Returns the query of QueryDef object as a string.
'
' Parameters:
' queryDefName - Name of the query
'
' Returns:
' the query of QueryDef object as a string.
Public Function GetQueryOfQueryDef(ByVal queryDefName As String) As String
    Dim qDef As QueryDef
    Set qDef = GetQueryDefByName(queryDefName)
    ' TODO: 6 check if the query can not be found.
    GetQueryOfQueryDef = qDef.sql
End Function


' Function: SetQueryOfQueryDef
' Sets the query of QueryDef object from a string.
'
' Parameters:
' queryDefName - Name of the query
Public Sub SetQueryOfQueryDef(ByVal queryDefName As String, ByVal query As String)
    Dim qDef As QueryDef
    Set qDef = GetQueryDefByName(queryDefName)
    ' TODO: 6 check if the query can not be found.
    qDef.sql = query
End Sub


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


' Sub: DeleteFromTable
' DELETEs the data from the specified table.
'
' Parameters:
' tableName - name of the table
Public Sub DeleteFromTable(ByVal tableName As String)
    Dim query As String
    query = "DELETE FROM " & tableName
    DoCmd.SetWarnings False
    DoCmd.RunSQL query
    DoCmd.SetWarnings True
End Sub


' Sub: DebugPrintTableConnections
' Prints the table name and the database covering it to the immediate window.
' This function is for debugging purposes.
' Type the following in the immediate window to call this sub:
' | Call mdlDatabase.DebugPrintTableConnections()
Public Sub DebugPrintTableConnections()
    Dim currentTableDef As TableDef
    For Each currentTableDef In CurrentDb.TableDefs
        Dim tableConnection As String
        tableConnection = currentTableDef.Connect
        If Len(tableConnection) > 0 Then
            Debug.Print currentTableDef.SourceTableName & " | " & currentTableDef.Connect
            ' Table1 | ;DATABASE=C:\path\to\my\projects\myproject\DatabaseBackend.accdb
        End If
    Next
End Sub


' Sub: SetAllTableConnections
' Changes all table connections in the database, except system tables.
' Use it with caution.
'
' Parameters:
' newConnectionString - the new connection string to set for all tables.
' simulationMode - The operation will be effective if and only if this is True
'
' Type the following in the immediate window to call this sub:
' | Call mdlDatabase.SetAllTableConnections("new_connection_string_here")
Public Sub SetAllTableConnections(ByVal newConnectionString As String, ByVal simulationMode As Boolean)
    Dim currentTableDef As TableDef
    If Len(Strip(newConnectionString)) > 0 Then
        For Each currentTableDef In CurrentDb.TableDefs
            Dim tableConnection As String
            tableConnection = currentTableDef.Connect
            If Len(tableConnection) > 0 Then
                If Not simulationMode Then
                    currentTableDef.Connect = newConnectionString
                    Debug.Print "SetAllTableConnections : " & currentTableDef.SourceTableName & " | " & currentTableDef.Connect
                Else
                    Debug.Print "SetAllTableConnections [SiMULATED] : " & currentTableDef.SourceTableName & " | " & currentTableDef.Connect
                End If
                ' Table1 | ;DATABASE=C:\path\to\my\projects\myproject\DatabaseBackend.accdb
            End If
        Next
    Else
        Debug.Print "SetAllTableConnections : new connection string is empty, nothing to do."
    End If
End Sub


' Sub: ConnectTable
' Connects the table to another data source by changing its connection string.
' This sub is mostly used in split database architecture.
' It can be called from Autoexec macro or form load events.
' A possible usage is as follows:
' | Dim backendAccdbFilePath As String
' | backendAccdbFilePath = GetBackendProjectPath("DataBackEnd")
' | mdlDatabase.ConnectTable("Table1", backendAccdbFilePath)
'
' Parameters:
' tableName - name of the table
' backendAccdbFilePath - the full path to the access database file. This will be the data backend file for most cases.
' Note that if the full path is not specified, Access will check the user documents
' directory, which is rarely what is required.
Public Sub ConnectTable(ByVal tableName As String, ByVal backendAccdbFilePath As String)
    Dim currentTableDef As TableDef
    For Each currentTableDef In CurrentDb.TableDefs
        Dim currentTableName As String
        currentTableName = currentTableDef.SourceTableName
        If LCase(Strip(currentTableName)) = LCase(Strip(tableName)) Then
            Dim newConnectionString As String
            ' this will be the new connectionString:
            newConnectionString = ";DATABASE=" & backendAccdbFilePath
            ' Debug.Print (newConnectionString)
            ' ;DATABASE=C:\path\to\my\projects\myproject\MyProjectData.accdb
            If newConnectionString <> currentTableDef.Connect Then
                currentTableDef.Connect = newConnectionString
                currentTableDef.RefreshLink
            End If
        End If
    Next
End Sub


' Sub: ConnectAllTables
' Connects all the tables to another data source by changing its connection string.
' This sub is mostly used in split database architecture.
' It can be called from Autoexec macro or form load events.
' A possible usage is as follows:
' | Dim backendAccdbFilePath As String
' | backendAccdbFilePath = GetBackendProjectPath("DataBackEnd")
' | mdlDatabase.ConnectAllTables(backendAccdbFilePath)
'
' Parameters:
' backendAccdbFilePath - the full path to the access database file. This will be the data backend file for most cases.
' Note that if the full path is not specified, Access will check the user documents
' directory, which is rarely what is required.
Public Sub ConnectAllTables(ByVal backendAccdbFilePath As String)
    Dim currentTableDef As TableDef
    For Each currentTableDef In CurrentDb.TableDefs
        ' Note that Access has also its own system tables.
        ' Their connection link is empty string.
        ' To avoid relinking them, the following condition is checked.
        If Len(currentTableDef.Connect) > 0 Then
            Dim currentTableName As String
            currentTableName = currentTableDef.SourceTableName
            Dim newConnectionString As String
            ' this will be the new connectionString:
            newConnectionString = ";DATABASE=" & backendAccdbFilePath
            Debug.Print (newConnectionString)
            ' ;DATABASE=C:\path\to\my\projects\myproject\MyProjectData.accdb
            If newConnectionString <> currentTableDef.Connect Then
                currentTableDef.Connect = newConnectionString
                currentTableDef.RefreshLink
            End If
        End If
    Next
End Sub


' Sub: AddPrefixToAllTables
' Adds a prefix to all table names.
' If the table name has the prefix already, it will remain unchanged.
'
' Parameters:
' prefix - the string to be added to the beginning of the table name
' simulationMode - The operation will be effective if and only if this is True
Public Sub AddPrefixToAllTables(ByVal prefix As String, ByVal simulationMode As Boolean)
    Dim currentTableDef As TableDef
    For Each currentTableDef In CurrentDb.TableDefs
        ' Note that Access has also its own system tables.
        ' Their connection link is empty string.
        ' To avoid relinking them, the following condition is checked.
        If Len(currentTableDef.Connect) > 0 Then
            Dim currentTableName As String
            currentTableName = currentTableDef.name
            If Not StartsWith(currentTableName, prefix) Then
                Dim newTableName As String
                newTableName = prefix & currentTableName
                Debug.Print (currentTableName & " -> " & newTableName)
                If Not simulationMode Then
                    currentTableDef.name = newTableName
                End If
            End If
        End If
    Next
End Sub


' Sub: RemovePrefixFromAllTables
' Removes a prefix from all table names.
' If the table name does not have the prefix, it will remain unchanged.
' Can be used to remove automated names like "dbo_"
'
' Parameters:
' prefix - the string to be removed from the beginning of the table name
' simulationMode - The operation will be effective if and only if this is True
Public Sub RemovePrefixFromAllTables(ByVal prefix As String, ByVal simulationMode As Boolean)
    ' Call RemovePrefixFromAllTables("dbo_")
    Dim currentTableDef As TableDef
    For Each currentTableDef In CurrentDb.TableDefs
        ' Note that Access has also its own system tables.
        ' Their connection link is empty string.
        ' To avoid relinking them, the following condition is checked.
        If Len(currentTableDef.Connect) > 0 Then
            Dim currentTableName As String
            currentTableName = currentTableDef.name
            If StartsWith(currentTableName, prefix) Then
                Dim newTableName As String
                newTableName = RemoveFromLeftIfStartsWith(currentTableName, prefix)
                Debug.Print (currentTableName & " -> " & newTableName)
                If Not simulationMode Then
                    currentTableDef.name = newTableName
                End If
            End If
        End If
    Next
End Sub


' Sub: AddSuffixToAllTables
' Adds a suffix to all table names.
' If the table name has the suffix already, it will remain unchanged.
'
' Parameters:
' suffix - the string to be added to the end of the table name
' simulationMode - The operation will be effective if and only if this is True
Public Sub AddSuffixToAllTables(ByVal suffix As String, ByVal simulationMode As Boolean)
    Dim currentTableDef As TableDef
    For Each currentTableDef In CurrentDb.TableDefs
        ' Note that Access has also its own system tables.
        ' Their connection link is empty string.
        ' To avoid relinking them, the following condition is checked.
        If Len(currentTableDef.Connect) > 0 Then
            Dim currentTableName As String
            currentTableName = currentTableDef.name
            If Not EndsWith(currentTableName, suffix) Then
                Dim newTableName As String
                newTableName = currentTableName & suffix
                Debug.Print (currentTableName & " -> " & newTableName)
                If Not simulationMode Then
                    currentTableDef.name = newTableName
                End If
            End If
        End If
    Next
End Sub


' Sub: RemoveSuffixFromAllTables
' Removes a suffix to all table names.
' If the table name has the suffix already, it will remain unchanged.
'
' Parameters:
' suffix - the string to be added to the end of the table name
' simulationMode - The operation will be effective if and only if this is True
Public Sub RemoveSuffixFromAllTables(ByVal suffix As String, ByVal simulationMode As Boolean)
    Dim currentTableDef As TableDef
    For Each currentTableDef In CurrentDb.TableDefs
        ' Note that Access has also its own system tables.
        ' Their connection link is empty string.
        ' To avoid relinking them, the following condition is checked.
        If Len(currentTableDef.Connect) > 0 Then
            Dim currentTableName As String
            currentTableName = currentTableDef.name
            If EndsWith(currentTableName, suffix) Then
                Dim newTableName As String
                newTableName = RemoveFromRightIfEndsWith(currentTableName, suffix)
                Debug.Print (currentTableName & " -> " & newTableName)
                If Not simulationMode Then
                    currentTableDef.name = newTableName
                End If
            End If
        End If
    Next
End Sub
