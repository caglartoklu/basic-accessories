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
' stringLength - The length of the produced random string
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


' Function: RemoveFromLeft
' Removes n characters from the left of
' the string and returns the remaining part.
'
' Parameters:
' haystack - The big string
' length - The length of the characters to be removed from left
'
' Returns:
' Haystack except some characters removed from left.
Public Function RemoveFromLeft(ByVal haystack As String, ByVal length As Integer) As String
    Dim result As String
    If length < 0 Then
        length = 0
    End If
    If length >= Len(haystack) Then
        result = ""
    Else
        result = Right(haystack, Len(haystack) - length)
    End If
    RemoveFromLeft = result
End Function


' Function: RemoveFromRight
' Removes n characters from the right of
' the string and returns the remaining part.
'
' Parameters:
' haystack - The big string
' length - The length of the characters to be removed from right
'
' Returns:
' Haystack except some characters removed from right.
Public Function RemoveFromRight(ByVal haystack As String, ByVal length As Integer) As String
    Dim result As String
    If length < 0 Then
        length = 0
    End If
    If length >= Len(haystack) Then
        result = ""
    Else
        result = Left(haystack, Len(haystack) - length)
    End If
    RemoveFromRight = result
End Function


' Function: PadLeft
' Pads strNeedle from left using strPadChar until the lenght becomes intMax
'
' Parameters:
' strNeedle - The string to be padded
' intMax - The maximum length of the padded string
' strPadChar - The character to be used for padding.
' If the length is longer than 1, only the first character will be used.
'
' Returns a copy of strNeedle padded with strPadChar from left.
Public Function PadLeft(ByVal strNeedle As String, ByVal intMax As Integer, ByVal strPadChar As String) As String
    Dim i As Integer
    Dim strResult As String
    strResult = strNeedle

    ' make sure that strPadChar is exactly 1 byte;
    ' nothing less:
    strPadChar = Trim(strPadChar)
    If Len(strPadChar) = 0 Then
        strPadChar = " "
    End If
    ' nothing more:
    strPadChar = Left(strPadChar, 1)

    Dim strMissing As String

    Dim intMissing As Integer
    intMissing = intMax - Len(strNeedle)

    If intMissing > 0 Then
        For i = 1 To intMissing
            strResult = strPadChar & strResult
        Next
    End IF

    PadLeft = strResult
End Function


' Function: PadRight
' Pads strNeedle from right using strPadChar until the lenght becomes intMax
'
' Parameters:
' strNeedle - The string to be padded
' intMax - The maximum length of the padded string
' strPadChar - The character to be used for padding.
' If the length is longer than 1, only the first character will be used.
'
' Returns a copy of strNeedle padded with strPadChar from right.
Public Function PadRight(ByVal strNeedle As String, ByVal intMax As Integer, ByVal strPadChar As String) As String
    Dim i As Integer
    Dim strResult As String
    strResult = strNeedle

    ' make sure that strPadChar is exactly 1 byte;
    ' nothing less:
    strPadChar = Trim(strPadChar)
    If Len(strPadChar) = 0 Then
        strPadChar = " "
    End If
    ' nothing more:
    strPadChar = Left(strPadChar, 1)

    Dim strMissing As String

    Dim intMissing As Integer
    intMissing = intMax - Len(strNeedle)

    If intMissing > 0 Then
        For i = 1 To intMissing
            strResult = strResult & strPadChar
        Next
    End IF

    PadRight = strResult
End Function
