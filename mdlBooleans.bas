' File: mdlBooleans.bas
' Includes the functions and subroutines about strings.

Option Compare Database
Option Explicit


' Function: ToggleBoolean
' Returns True if the value is False, False otherwise.
'
' Parameters:
' booleanValue - the boolean value
'
' Returns:
' True if the value is False, False otherwise.
Public Sub ToggleBoolean(ByRef booleanValue As Boolean)
    If (booleanValue = True) Then
        booleanValue = False
    Else
        booleanValue = True
    End If
End Sub
