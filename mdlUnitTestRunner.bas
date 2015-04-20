' File: mdlTestRunner.bas
' Includes the actual calls to unit tests of all modules.
' Simply type:
' | Call mdlUnitTestRunner.RunAllUnitTests()
' to run all the unit tests.


Option Compare Database
Option Explicit


' Sub: RunAllUnitTests
' Executes the unit test modules causing all the unit tests to be executed.
Public Sub RunAllUnitTests()
    Call RunAllMdlDatabaseTest()
    Call RunAllMdlDateTest()
    Call RunAllMdlFilesTest()
    Call RunAllMdlStringsTest()
End Sub
