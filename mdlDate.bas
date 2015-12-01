' File: mdlDate.bas
' Includes the functions and subroutines about date and time.


Option Compare Database
Option Explicit


' Function: FirstMomentOfTheDay
' Returns the first moment of a given date which is "00:00:00"
'
' Parameters:
' anyDate - Just a date. Its year, month and day information
' will be exactly copied to the returning value.
'
' Returns:
' The first moment of the provided anyDate.
' Its values of year, month and day information will not be changed,
' but hour, minute and second will be "00:00:00"
Public Function FirstMomentOfTheDay(ByVal anyDate As Date) As Date
    Dim result As Date
    result = DateSerial(Year(anyDate), Month(anyDate), Day(anyDate))
    ' No need for the following lines:
    ' result = DateAdd("h", 0, result) ' hour
    ' result = DateAdd("n", 0, result) ' minute
    ' result = DateAdd("s", 0, result) ' second
    FirstMomentOfTheDay = result
End Function


' Function: LastMomentOfTheDay
' Returns the last moment of a given date which is "23:59:59"
'
' Parameters:
' anyDate - Just a date. Its year, month and day information
' will be exactly copied to the returning value.
'
' Returns:
' The last moment of the provided anyDate.
' Its values of year, month and day information will not be changed,
' but hour, minute and second will be "23:59:59"
Public Function LastMomentOfTheDay(ByVal anyDate As Date) As Date
    Dim result As Date
    result = DateSerial(Year(anyDate), Month(anyDate), Day(anyDate))
    result = DateAdd("h", 23, result) ' hour
    result = DateAdd("n", 59, result) ' minute
    result = DateAdd("s", 59, result) ' second
    LastMomentOfTheDay = result
End Function


' Function: ProperDateTimeDetailed
' Returns the current time in the yyyy, mm, dd, hh, nn, ss format as a string.
' The format can be adjusted using separator and groupSeparator parameters.
'
' Parameters:
' separator - the little separator between year, month, day or hour, minute, second.
' groupSeparator - the big separator between date group and time group.
'
' Returns:
' The current time in the yyyy, mm, dd, hh, nn, ss format as a string.
Public Function ProperDateTimeDetailed(ByVal separator As String, ByVal groupSeparator As String) As String
    Dim result As String
    Dim formatString As String
    formatString = "yyyy-mm-dd_hh-nn-ss"
    formatString = Replace(formatString, "-", separator)
    formatString = Replace(formatString, "_", groupSeparator)
    result = Format(Now, formatString)
    ProperDateTimeDetailed = result
End Function


' Function: ProperDateTime
' Returns the current time in the yyyy, mm, dd, hh, nn, ss format as a string.
' The format is fixed and it is always yyyymmdd_hhnnss, such as 20151201_145414
' The returning value of this function can be used in file names.
'
' Returns:
' The current time in yyyymmdd_hhnnss format as a string.
Public Function ProperDateTime() As String
    Dim result As String
    Dim separator As String
    Dim groupSeparator As String
    separator = ""
    groupSeparator = "_"
    result = ProperDateTimeDetailed(separator:=separator, groupSeparator:=groupSeparator)
    ProperDateTime = result
End Function
