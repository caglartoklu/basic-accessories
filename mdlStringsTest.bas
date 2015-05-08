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


' Sub: TestRemoveFromLeft
' Tests <RemoveFromLeft>
Public Sub TestRemoveFromLeft()
    Dim testSubName As String
    testSubName = "TestRemoveFromLeft"
    Call AssertAreEqual(testSubName, "b123456", RemoveFromLeft("ab123456", 1))
    Call AssertAreEqual(testSubName, "123456", RemoveFromLeft("ab123456", 2))
    Call AssertAreEqual(testSubName, "56", RemoveFromLeft("ab123456", 6))
    Call AssertAreEqual(testSubName, "6", RemoveFromLeft("ab123456", 7))
    Call AssertAreEqual(testSubName, "", RemoveFromLeft("ab123456", 8))
    Call AssertAreEqual(testSubName, "", RemoveFromLeft("ab123456", 99))
    Call AssertAreEqual(testSubName, "ab123456", RemoveFromLeft("ab123456", -1))
End Sub


' Sub: TestRemoveFromRight
' Tests <RemoveFromRight>
Public Sub TestRemoveFromRight()
    Dim testSubName As String
    testSubName = "TestRemoveFromRight"
    Call AssertAreEqual(testSubName, "ab12345", RemoveFromRight("ab123456", 1))
    Call AssertAreEqual(testSubName, "ab1234", RemoveFromRight("ab123456", 2))
    Call AssertAreEqual(testSubName, "ab", RemoveFromRight("ab123456", 6))
    Call AssertAreEqual(testSubName, "a", RemoveFromRight("ab123456", 7))
    Call AssertAreEqual(testSubName, "", RemoveFromRight("ab123456", 8))
    Call AssertAreEqual(testSubName, "", RemoveFromRight("ab123456", 99))
    Call AssertAreEqual(testSubName, "ab123456", RemoveFromRight("ab123456", -1))
End Sub


' Sub: TestPadLeft
' Tests <PadLeft>
Public Sub TestPadLeft()
    Dim testSubName As String
    testSubName = "TestPadLeft"
    ' padding with 1 char
    Call AssertAreEqual(testSubName, "0aaaa", PadLeft("aaaa", 5, "0"))
    Call AssertAreEqual(testSubName, "0aaaa", PadLeft("aaaa", 5, "01"))

    ' padding with 2 chars
    Call AssertAreEqual(testSubName, "00aaaa", PadLeft("aaaa", 6, "0"))
    Call AssertAreEqual(testSubName, "00aaaa", PadLeft("aaaa", 6, "01"))

    ' same length
    Call AssertAreEqual(testSubName, "aaaa", PadLeft("aaaa", 4, "0"))
    Call AssertAreEqual(testSubName, "aaaa", PadLeft("aaaa", 4, "01"))

    ' pad char is empty string
    Call AssertAreEqual(testSubName, " aaaa", PadLeft("aaaa", 5, ""))

    ' inMax < Len(strNeedle)
    Call AssertAreEqual(testSubName, "aaaa", PadLeft("aaaa", 3, "0"))
End Sub


' Sub: TestPadRight
' Tests <PadRight>
Public Sub TestPadRight()
    Dim testSubName As String
    testSubName = "TestPadRight"
    ' padding with 1 char
    Call AssertAreEqual(testSubName, "aaaa0", PadRight("aaaa", 5, "0"))
    Call AssertAreEqual(testSubName, "aaaa0", PadRight("aaaa", 5, "01"))

    ' padding with 2 chars
    Call AssertAreEqual(testSubName, "aaaa00", PadRight("aaaa", 6, "0"))
    Call AssertAreEqual(testSubName, "aaaa00", PadRight("aaaa", 6, "01"))

    ' same length
    Call AssertAreEqual(testSubName, "aaaa", PadRight("aaaa", 4, "0"))
    Call AssertAreEqual(testSubName, "aaaa", PadRight("aaaa", 4, "01"))

    ' pad char is empty string
    Call AssertAreEqual(testSubName, "aaaa ", PadRight("aaaa", 5, ""))

    ' inMax < Len(strNeedle)
    Call AssertAreEqual(testSubName, "aaaa", PadRight("aaaa", 3, "0"))
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
    Call TestRemoveFromLeft()
    Call TestRemoveFromRight()
    Call TestPadLeft()
    Call TestPadRight()
End Sub
