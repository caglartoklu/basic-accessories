' File: mdlTestLib.bas
' Includes the functions and subroutines about unit tests.
' This is a simple module that displays the test results in the
' "Intermediate Window" of Microsoft Access.


Option Compare Database
Option Explicit


' Function: GetPassMessage
' Returns the pass message to be the same
' for all test subroutines.
'
' Returns:
' the pass message
Public Function GetPassMessage() As String
    GetPassMessage = "pass ... "
End Function


' Function: GetFailMessage
' Returns the fail message to be the same
' for all test subroutines.
'
' Returns:
' the fail message
Public Function GetFailMessage() As String
    GetFailMessage = "FAIL ... "
End Function


' Sub: AssertAreEqual
' Tests if two values are equal.
'
' Parameters:
' testSubName - name of the test subroutine
' expected - expected value of any type
' actual - actual value of any type
Public Sub AssertAreEqual(ByVal testSubName As String, ByVal expected As Variant, ByVal actual As Variant)
    Dim msg As String
    If expected = actual Then
        msg = GetPassMessage() & testSubName
    Else
        msg = GetFailMessage() & testSubName & " - "
        msg = msg & "  expected : " & expected & vbCrlf
        msg = msg & "  actual   : " & actual & vbCrlf
    End If
    Debug.Print(msg)
End Sub


' Sub: AssertAreNotEqual
' Tests if two values are not equal.
'
' Parameters:
' testSubName - name of the test subroutine
' expected - expected value of any type
' actual - actual value of any type
Public Sub AssertAreNotEqual(ByVal testSubName As String, ByVal expected As Variant, ByVal actual As Variant)
    Dim msg As String
    If expected = actual Then
        msg = GetFailMessage() & testSubName & " - "
        msg = msg & "  expected : " & expected & vbCrlf
        msg = msg & "  actual   : " & actual & vbCrlf
    Else
        msg = GetPassMessage() & testSubName
    End If
    Debug.Print(msg)
End Sub


' Sub: AssertTrue
' Tests if the specified value is True.
'
' Parameters:
' testSubName - name of the test subroutine
' actual - actual value
Public Sub AssertTrue(ByVal testSubName As String, ByVal actual As Boolean)
    Dim msg As String
    If actual = True Then
        msg = GetPassMessage() & testSubName
    Else
        msg = GetFailMessage() & testSubName & " - "
        msg = msg & "  expected : True" & vbCrlf
        msg = msg & "  actual   : " & actual & vbCrlf
    End If
    Debug.Print(msg)
End Sub


' Sub: AssertFalse
' Tests if the specified value is False.
'
' Parameters:
' testSubName - name of the test subroutine
' actual - actual value
Public Sub AssertFalse(ByVal testSubName As String, ByVal actual As Boolean)
    Dim msg As String
    If actual = False Then
        msg = GetPassMessage() & testSubName
    Else
        msg = GetFailMessage() & testSubName & " - "
        msg = msg & "  expected : False" & vbCrlf
        msg = msg & "  actual   : " & actual & vbCrlf
    End If
    Debug.Print(msg)
End Sub


' Sub: AssertAreArraysEqual
' Tests if the contents of two arrays are equal.
'
' Parameters:
' testSubName - name of the test subroutine
' arrExpected - expected array of any type
' arrActual - actual array of any type
Public Sub AssertAreArraysEqual(ByVal testSubName As String, ByRef arrExpected() As Variant, ByRef arrActual() As Variant)
    Dim msg As String
    If LBound(arrExpected) = LBound(arrActual) Then
        If UBound(arrExpected) = UBound(arrActual) Then
            Dim bolAllEqual As Boolean
            bolAllEqual = True
            Dim i As Long
            For i = LBound(arrExpected) To UBound(arrExpected)
                If arrExpected(i) <> arrActual(i) Then
                    ' same number of elements but at least one of them does not match
                    msg = GetFailMessage() & testSubName & " - "
                    msg = msg & "  expected(" & i & ") : " & arrExpected(i) & vbCrlf
                    msg = msg & "  actual(" & i & ")   : " & arrActual(i) & vbCrlf
                    bolAllEqual = False
                    Exit For
                End If
            Next
            If bolAllEqual = True Then
                ' same number of elements and all elements are equal
                msg = GetPassMessage() & testSubName
            End If
        Else
            ' UBound of array do not match
            msg = GetFailMessage() & testSubName & " - "
            msg = msg & "  expected UBound : " & UBound(arrExpected) & vbCrlf
            msg = msg & "  actual   UBound : " & UBound(arrActual) & vbCrlf
        End If
    Else
        ' LBound of array do not match
        msg = GetFailMessage() & testSubName & " - "
        msg = msg & "  expected LBound : " & LBound(arrExpected) & vbCrlf
        msg = msg & "  actual   LBound : " & LBound(arrActual) & vbCrlf
    End If
    Debug.Print(msg)
End Sub
