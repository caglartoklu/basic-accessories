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


' Sub: TestProperDateTimeDetailed
' Tests <ProperDateTimeDetailed>
Public Sub TestProperDateTimeDetailed()
    Dim testSubName As String
    testSubName = "TestProperDateTimeDetailed"
    Dim actual As String
    actual = ProperDateTimeDetailed("-", "__")
    ' 2015-12-01__14-54-14
    Call AssertAreEqual(testSubName, 20, Len(actual))
    Call AssertAreEqual(testSubName, "-", Mid(actual, 5, 1))
    Call AssertAreEqual(testSubName, "-", Mid(actual, 8, 1))
    Call AssertAreEqual(testSubName, "__", Mid(actual, 11, 2))
End Sub


' Sub: TestProperDateTime
' Tests <ProperDateTime>
Public Sub TestProperDateTime()
    Dim testSubName As String
    testSubName = "TestProperDateTime"
    Dim actual As String
    actual = ProperDateTime()
    ' 20151201_145414
    Call AssertAreEqual(testSubName, 15, Len(actual))
    Call AssertAreEqual(testSubName, "_", Mid(actual, 9, 1))
End Sub


' Sub: RunAllMdlDateTest
' Calls the tests in this module.
'
' See also:
' <RunAllUnitTests>
Public Sub RunAllMdlDateTest()
    Call TestFirstMomentOfTheDay()
    Call TestLastMomentOfTheDay()
    Call TestProperDateTimeDetailed()
    Call TestProperDateTime()
End Sub
