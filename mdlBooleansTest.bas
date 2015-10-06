' File: mdlBooleansTest.bas
' Includes tests about: mdlBooleans.bas


Option Compare Database
Option Explicit


' Sub: TestToggleBoolean
' Tests <ToggleBoolean>
Public Sub TestToggleBoolean()
    Dim testSubName As String
    testSubName = "TestToggleBoolean"

    Dim someValue As Boolean
    Call ToggleBoolean(someValue)
    Call AssertTrue(testSubName, someValue)
    Call ToggleBoolean(someValue)
    Call AssertFalse(testSubName, someValue)

    someValue = True
    Call ToggleBoolean(someValue)
    Call AssertFalse(testSubName, someValue)

    someValue = False
    Call ToggleBoolean(someValue)
    Call AssertTrue(testSubName, someValue)
End Sub


' Sub: RunAllMdlBooleansTest
' Calls the tests in this module.
'
' See also:
' <RunAllUnitTests>
Public Sub RunAllMdlBooleansTest()
    Call TestToggleBoolean
End Sub
