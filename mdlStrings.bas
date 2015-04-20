' File: mdlStrings.bas
' Includes the functions and subroutines about strings.

Option Compare Database
Option Explicit


' Function: StartsWith
' Returns True if the haystack starts with needle, False otherwise.
'
' Parameters:
' haystack - the big string
' needle - the small string
'
' Returns:
' True if the haystack starts with needle, False otherwise.
'
' See also:
' <EndsWith>
Public Function StartsWith(ByVal haystack As String, ByVal needle As String) As Boolean
    Dim result As Boolean
    result = 0
    If Left(haystack, Len(needle)) = needle Then
        result = 1
    End If
    StartsWith = result
End Function


' Function: EndsWith
' Returns True if the haystack ends with needle, False otherwise.
'
' Parameters:
' haystack - the big string
' needle - the small string
'
' Returns:
' True if the haystack ends with needle, False otherwise.
'
' See also:
' <StartsWith>
Public Function EndsWith(ByVal haystack As String, ByVal needle As String) As Boolean
    Dim result As Boolean
    result = False
    If Right(haystack, Len(needle)) = needle Then
        result = True
    End If
    EndsWith = result
End Function


' Function: IsComment
' Returns True if the haystack is a comment, False otherwise.
' Comment starter characters are as follows: ' # ; // '
'
' Parameters:
' haystack - Text to be checked if it is a comment or not.
'
' Returns:
' True if the haystack is a comment, False otherwise.
Public Function IsComment(ByVal haystack As String) As Boolean
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


' Function: RandomString
' Returns a random string with the specified length
' from the provided character set.
'
' Parameters:
' characterSet - A string containing candidate characters
' stringLength - The lenght of the produced random string
'
' Returns:
' a random string with the specified length
' from the provided character set.
Public Function RandomString(ByVal characterSet As String, ByVal stringLength As Integer) As String
    Dim result As String
    result = ""
    If stringLength >= 1 Then
        Dim i As Integer
        Dim lower As Integer
        Dim upper As Integer
        lower = 1
        upper = Len(characterSet)
        For i = 1 To stringLength
            Dim selectedChar As String
            Dim selectedIndex As Integer
            selectedIndex = Int((upper - lower + 1) * Rnd + lower)
            selectedChar = Mid(characterSet, selectedIndex, 1)
            result = result & selectedChar
        Next
    End If
    RandomString = result
End Function
