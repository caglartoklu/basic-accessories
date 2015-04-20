' File: mdlDateTest.bas
' Includes tests about mdlDate.bas.


Option Compare Database
Option Explicit


' Sub: TestFirstMomentOfTheDay
' Tests <FirstMomentOfTheDay>
Public Sub TestFirstMomentOfTheDay()
    Dim testSubName As String
    testSubName = "TestFirstMomentOfTheDay"
    ' TODO: Implement Sub TestFirstMomentOfTheDay()
End Sub


' Sub: TestLastMomentOfTheDay
' Tests <LastMomentOfTheDay>
Public Sub TestLastMomentOfTheDay()
    Dim testSubName As String
    testSubName = "TestLastMomentOfTheDay"
    ' TODO: Implement Sub TestLastMomentOfTheDay()
End Sub


' Sub: RunAllMdlDateTest
' Calls the tests in this module.
'
' See also:
' <RunAllUnitTests>
Public Sub RunAllMdlDateTest()
    Call TestFirstMomentOfTheDay()
    Call TestLastMomentOfTheDay()
End Sub
