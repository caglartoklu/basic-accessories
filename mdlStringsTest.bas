' File: mdlStringsTest.bas
' Includes tests about: mdlStrings.bas


Option Compare Database
Option Explicit


' Sub: TestStartsWith
' Tests <StartsWith>
Public Sub TestStartsWith()
    Dim testSubName As String
    testSubName = "TestStartsWith"
    Call AssertTrue(testSubName, StartsWith("OneTwoThree", "One"))
    Call AssertFalse(testSubName, StartsWith("OneTwoThree", "OneX"))
    Call AssertFalse(testSubName, StartsWith("OneTwoThree", "Two"))
    Call AssertFalse(testSubName, StartsWith("OneTwoThree", "Three"))
End Sub


' Sub: TestEndsWith
' Tests <EndsWith>
Public Sub TestEndsWith()
    Dim testSubName As String
    testSubName = "TestEndsWith"
    Call AssertTrue(testSubName, EndsWith("OneTwoThree", "Three"))
    Call AssertFalse(testSubName, EndsWith("OneTwoThree", "XThree"))
    Call AssertFalse(testSubName, EndsWith("OneTwoThree", "Two"))
    Call AssertFalse(testSubName, EndsWith("OneTwoThree", "One"))
End Sub


' Sub: TestIsComment
' Tests <IsComment>
Public Sub TestIsComment()
    Dim testSubName As String
    testSubName = "TestIsComment"
    Call AssertTrue(testSubName, IsComment(";one"))
    Call AssertTrue(testSubName, IsComment("//one"))
    Call AssertTrue(testSubName, IsComment("'one"))
    Call AssertTrue(testSubName, IsComment("#one"))
    Call AssertTrue(testSubName, IsComment(" ;one"))
    Call AssertTrue(testSubName, IsComment(" //one"))
    Call AssertTrue(testSubName, IsComment(" 'one"))
    Call AssertTrue(testSubName, IsComment(" #one"))
    Call AssertFalse(testSubName, IsComment("x;one"))
    Call AssertFalse(testSubName, IsComment("x//one"))
    Call AssertFalse(testSubName, IsComment("x'one"))
    Call AssertFalse(testSubName, IsComment("x#one"))
End Sub


' Sub: TestRandomString
' Tests <RandomString>
Public Sub TestRandomString()
    Dim testSubName As String
    testSubName = "TestRandomString"

    Dim actual As String
    Dim expected As String

    Dim characterSet As String
    Dim stringLength As Integer

    characterSet = "a"
    stringLength = 4
    expected = "aaaa"
    actual = RandomString(characterSet, stringLength)
    Call AssertAreEqual(testSubName, expected, actual)

    characterSet = ""
    stringLength = 4
    expected = ""
    actual = RandomString(characterSet, stringLength)
    Call AssertAreEqual(testSubName, expected, actual)
End Sub


' Sub: RunAllMdlStringsTest
' Calls the tests in this module.
'
' See also:
' <RunAllUnitTests>
Public Sub RunAllMdlStringsTest()
    Call TestStartsWith()
    Call TestEndsWith()
    Call TestIsComment()
    Call TestRandomString()
End Sub
