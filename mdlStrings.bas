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


' Function: RemoveFromLeftIfStartsWith
' Removes the needle from the end of the haystack if haystack ends with needle,
' and returns the remaining part,
'
' Parameters:
' haystack - The big string
' needle - The needle to be cut out form the end of the haystack
'
' Returns:
' Haystack except the needle
Public Function RemoveFromLeftIfStartsWith(ByVal haystack As String, ByVal needle As String) As String
    Dim length As Integer
    Dim result As String
    length = Len(needle)
    result = haystack
    If length > 0 Then
        If StartsWith(haystack, needle) Then
            result = RemoveFromLeft(haystack, length)
        End If
    End If
    RemoveFromLeftIfStartsWith = result
End Function


' Function: RemoveFromRightIfEndsWith
' Removes the needle from the end of the haystack if haystack ends with needle,
' and returns the remaining part,
'
' Parameters:
' haystack - The big string
' needle - The needle to be cut out form the end of the haystack
'
' Returns:
' Haystack except the needle
Public Function RemoveFromRightIfEndsWith(ByVal haystack As String, ByVal needle As String) As String
    Dim length As Integer
    Dim result As String
    length = Len(needle)
    result = haystack
    If length > 0 Then
        If EndsWith(haystack, needle) Then
            result = RemoveFromRight(haystack, length)
        End If
    End If
    RemoveFromRightIfEndsWith = result
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
' Returns:
' a copy of strNeedle padded with strPadChar from left.
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
    End If

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
' Returns:
' a copy of strNeedle padded with strPadChar from right.
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
    End If

    PadRight = strResult
End Function


' Function: LStrip
' Strips all the leading whitespace characters from the string.
' In addition to built-in LTrim() function which removes CrLf,
' this function removes Cr, Lf, CrLf and Tab characters.
'
' Parameters:
' haystack - the string to be stripped
'
' Returns:
' A copy of haystack without leading whitespace characters.
Public Function LStrip(ByVal haystack As String) As String
    Dim result As String
    result = haystack
    Dim finished As Boolean
    finished = False
    While Not finished
        finished = True
        If Left(result, 1) = vbCr Then
            result = Right(result, Len(result) - 1)
            finished = False
        ElseIf Left(result, 1) = vbLf Then
            result = Right(result, Len(result) - 1)
            finished = False
        ElseIf Left(result, 1) = vbCrlf Then
            result = Right(result, Len(result) - 1)
            finished = False
        ElseIf Left(result, 1) = vbTab Then
            result = Right(result, Len(result) - 1)
            finished = False
        End If
        result = LTrim(result)
    Wend
    LStrip = result
End Function


' Function: LStrip
' Strips all the trailing whitespace characters from the string.
' In addition to built-in RTrim() function which removes CrLf,
' this function removes Cr, Lf, CrLf and Tab characters.
'
' Parameters:
' haystack - the string to be stripped
'
' Returns:
' A copy of haystack without trailing whitespace characters.
Public Function RStrip(ByVal haystack As String) As String
    Dim result As String
    result = haystack
    Dim finished As Boolean
    finished = False
    While Not finished
        finished = True
        If Right(result, 1) = vbCr Then
            result = Left(result, Len(result) - 1)
            finished = False
        ElseIf Right(result, 1) = vbLf Then
            result = Left(result, Len(result) - 1)
            finished = False
        ElseIf Right(result, 1) = vbCrlf Then
            result = Left(result, Len(result) - 1)
            finished = False
        ElseIf Right(result, 1) = vbTab Then
            result = Left(result, Len(result) - 1)
            finished = False
        End If
        result = RTrim(result)
    Wend
    RStrip = result
End Function


' Function: Strip
' Strips all the leading and trailing whitespace characters from the string.
' In addition to built-in Trim() function which removes CrLf,
' this function removes Cr, Lf, CrLf and Tab characters.
'
' Parameters:
' haystack - the string to be stripped
'
' Returns:
' A copy of haystack without leading and trailing whitespace characters.
Public Function Strip(ByVal haystack As String) As String
    Strip = LStrip(RStrip(haystack))
End Function



' Function StrPartRemove
' Removes the part of a string.
'
' Parameters:
' haystack - the big string
' iStart - the start index of the string to be removed
' iEnd - the end index of the string to be removed
'
' The original string does not change, a removed one is returned.
Public Function StrPartRemove(ByVal haystack As String, ByVal iStart As Long, ByVal iEnd As Long) As String
    Dim part1 As String
    Dim part2 As String
    Dim n As Long
    Dim result As String

    result = haystack

    If iStart > iEnd Then
        ' iStart must be smaller than iEnd
        Dim iTemp As Long
        iTemp = iStart
        iStart = iEnd
        iEnd = iTemp
    End If

    If iStart < 1 Then
        iStart = 1
    End If

    If iEnd < 1 Then
        iEnd = 1
    End If

    If iStart > Len(haystack) Then
        result = haystack
    Else
        If iEnd > Len(haystack) Then
            iEnd = Len(haystack)
        End If

        n = iStart - 1
        part1 = Left(haystack, n)

        n = Len(haystack) - iEnd
        part2 = Right(haystack, n)

        result = part1 & part2
    End If


    StrPartRemove = result
End Function 'StrPartRemove()


' Function StrCount
' Counts the number of occurrences of needle in haystack and returns the number.
' This function depends on the function StrPartRemove().
'
' Parameters:
' haystack - the big string
' needle - the small string
'
' Returns:
' the number of occurrences of needle in haystack and returns the number.
Public Function StrCount(ByVal haystack As String, ByVal needle As String) As Integer
    Dim h As Long
    Dim n As Long
    Dim posi As Long
    Dim occurenceCount As Long
    Dim haystackTemp As String
    Dim haystack2 As String
    Dim needle2 As String

    haystack2 = haystack
    needle2 = needle

    h = Len(haystack2)
    n = Len(needle2)
    occurenceCount = 0
    haystackTemp = haystack2

    posi = InStr(haystackTemp, needle2)
    While posi > 0
        occurenceCount = occurenceCount + 1
        haystackTemp = StrPartRemove(haystackTemp, posi, posi + n - 1)
        posi = InStr(haystackTemp, needle2)
    Wend

    StrCount = occurenceCount
End Function 'StrCount()


' Function SafeSql
' Replaces bad characters in the SQL statement.
'
' Parameters:
' sql - an sql statement with possibly bad characters
'
' Returns:
' A safe SQL statement without bad characters.
Public Function SafeSql(ByVal sql As String) As String
    sql = Replace(sql, "'", "''")
    SafeSql = sql
End Function 'SafeSql()

