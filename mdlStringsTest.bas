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


' Sub: TestStrPartRemove
' Tests <StrPartRemove>
Public Sub TestStrPartRemove()
    Dim testSubName As String
    testSubName = "TestStrPartRemove"
    Call AssertAreEqual(testSubName, "4567", StrPartRemove("1234567", 1, 3))
    Call AssertAreEqual(testSubName, "", StrPartRemove("1234567", 1, 7))
    Call AssertAreEqual(testSubName, "123", StrPartRemove("1234567", 4, 7))
    Call AssertAreEqual(testSubName, "1237", StrPartRemove("1234567", 4, 6))
    Call AssertAreEqual(testSubName, "1234567", StrPartRemove("1234567", 15, 33))
    Call AssertAreEqual(testSubName, "234567", StrPartRemove("1234567", 1, 1))
    Call AssertAreEqual(testSubName, "134567", StrPartRemove("1234567", 2, 2))
    Call AssertAreEqual(testSubName, "123456", StrPartRemove("1234567", 7, 7))
End Sub


' Sub: TestStrCount
' Tests <StrCount>
Public Sub TestStrCount()
    Dim testSubName As String
    testSubName = "TestStrCount"
    Call AssertAreEqual(testSubName, "1", CStr(StrCount("abcdefg", "cd")))
    Call AssertAreEqual(testSubName, "0", CStr(StrCount("aaa", "b")))
    Call AssertAreEqual(testSubName, "1", CStr(StrCount("aaa", "aa")))
    Call AssertAreEqual(testSubName, "3", CStr(StrCount("xaxaxa", "xa")))
    Call AssertAreEqual(testSubName, "2", CStr(StrCount("xaxax", "xa")))
    Call AssertAreEqual(testSubName, "2", CStr(StrCount("axaxa", "xa")))
    Call AssertAreEqual(testSubName, "3", CStr(StrCount("--ab--efg--", "--")))
    Call AssertAreEqual(testSubName, "6", CStr(StrCount("--ab--efg--", "-")))
End Sub


' Sub: TestLStrip
' Tests <LStrip>
Public Sub TestLStrip()
    Dim testSubName As String
    testSubName = "TestLStrip"
    Call AssertAreEqual(testSubName, "abc", LStrip("abc"))
    Call AssertAreEqual(testSubName, "abc", LStrip(" abc"))
    Call AssertAreEqual(testSubName, "abc ", LStrip(" abc "))
    Call AssertAreEqual(testSubName, "a b c ", LStrip(" a b c "))
    Call AssertAreEqual(testSubName, "a b c ", LStrip(vbCrlf & vbCr & vbLf & vbTab & " a b c "))
    Call AssertAreEqual(testSubName, "a " & vbTab & " b c ", LStrip(vbCrlf & vbCr & vbLf & vbTab & " a " & vbTab & " b c "))
End Sub


' Sub: TestRStrip
' Tests <RStrip>
Public Sub TestRStrip()
    Dim testSubName As String
    testSubName = "TestRStrip"
    Call AssertAreEqual(testSubName, "abc", RStrip("abc"))
    Call AssertAreEqual(testSubName, "abc", RStrip("abc "))
    Call AssertAreEqual(testSubName, " abc", RStrip(" abc "))
    Call AssertAreEqual(testSubName, " a b c", RStrip(" a b c "))
    Call AssertAreEqual(testSubName, " a b c", RStrip(" a b c " & vbCrlf & vbCr & vbLf & vbTab))
    Call AssertAreEqual(testSubName, " a " & vbTab & " b c", RStrip(" a " & vbTab & " b c " & vbCrlf & vbCr & vbLf & vbTab))
End Sub


' Sub: TestStrip
' Tests <Strip>
Public Sub TestStrip()
    Dim testSubName As String
    testSubName = "TestStrip"
    Call AssertAreEqual(testSubName, "abc", Strip("abc"))
    Call AssertAreEqual(testSubName, "abc", Strip("abc "))
    Call AssertAreEqual(testSubName, "abc", Strip(" abc "))
    Call AssertAreEqual(testSubName, "a b c", Strip(" a b c "))
    Call AssertAreEqual(testSubName, "a b c", Strip(" a b c " & vbCrlf & vbCr & vbLf & vbTab))
    Call AssertAreEqual(testSubName, "a " & vbTab & " b c", Strip(" a " & vbTab & " b c " & vbCrlf & vbCr & vbLf & vbTab))
End Sub


' Sub: TestSafeSql
' Tests <SafeSql>
Public Sub TestSafeSql()
    Dim testSubName As String
    testSubName = "TestSafeSql"
    Call AssertAreEqual(testSubName, "abc", SafeSql("abc"))
    Call AssertAreEqual(testSubName, "abc ", SafeSql("abc "))
    Call AssertAreEqual(testSubName, " a''b''''c ", SafeSql(" a'b''c "))
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
    Call TestStrPartRemove()
    Call TestStrCount()
    Call TestLStrip()
    Call TestRStrip()
    Call TestStrip()
    Call TestSafeSql()
End Sub
