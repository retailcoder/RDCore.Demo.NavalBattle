Attribute VB_Name = "TestRunner"
Attribute VB_Description = "Runs lists of tests: each test by name, with its module's initialize and cleanup procedures, and prints each outcome and a summary. A run with a failed test ends with an error."
'@Folder("NavalBattle.Tests.Framework")
'@ModuleDescription("Runs lists of tests: each test by name, with its module's initialize and cleanup procedures, and prints each outcome and a summary. A run with a failed test ends with an error.")
Option Explicit
Option Private Module

Private Type TRun
    List As ITestList
    CurrentModule As String
    Passed As Long
    Failed As Long
    Inconclusive As Long
    Ignored As Long
End Type

Private this As TRun

'@Description("Runs every test of the given lists, one list after the other, prints each outcome and one summary, and raises an error where a test failed.")
Public Sub RunAll(ByVal lists As Collection)
Attribute RunAll.VB_Description = "Runs every test of the given lists, one list after the other, prints each outcome and one summary, and raises an error where a test failed."
    this.Passed = 0
    this.Failed = 0
    this.Inconclusive = 0
    this.Ignored = 0

    Dim total As Long
    Dim item As Variant
    For Each item In lists
        Set this.List = item
        total = total + this.List.Count
        RunList
    Next
    Set this.List = Nothing

    Debug.Print total & " tests: " & this.Passed & " passed, " & this.Failed & " failed, " & this.Inconclusive & " inconclusive, " & this.Ignored & " ignored."
    If this.Failed > 0 Then Err.Raise vbObjectError + 1, "TestRunner", this.Failed & " of " & total & " tests failed."
End Sub

'@Description("Runs every test of the current list.")
Private Sub RunList()
Attribute RunList.VB_Description = "Runs every test of the current list."
    this.CurrentModule = vbNullString
    Dim index As Long
    Dim name As String
    For index = 1 To this.List.Count
        name = this.List.Test(index, run:=False)
        BeforeTest name
        On Error GoTo TestError
        this.List.Test index, run:=True
        On Error GoTo 0
        AfterTest name, TestAssert.Outcome, TestAssert.Reason
NextTest:
    Next
    EnterModule vbNullString
    Exit Sub

TestError:
    AfterTest name, TestFailed, "Error " & Err.Number & ": " & Err.Description
    Resume NextTest
End Sub

'@Description("Enters the test's module, where it is not the current one, and starts the test.")
Private Sub BeforeTest(ByVal name As String)
Attribute BeforeTest.VB_Description = "Enters the test's module, where it is not the current one, and starts the test."
    EnterModule Left$(name, InStr(name, ".") - 1)
    this.List.InitializeTest this.CurrentModule
    TestAssert.Reset
End Sub

'@Description("Ends the test, and prints and counts its outcome.")
Private Sub AfterTest(ByVal name As String, ByVal outcome As TestOutcome, ByVal reason As String)
Attribute AfterTest.VB_Description = "Ends the test, and prints and counts its outcome."
    this.List.CleanupTest this.CurrentModule
    Select Case outcome
        Case TestPassed
            this.Passed = this.Passed + 1
            Debug.Print "Passed        " & name
        Case TestFailed
            this.Failed = this.Failed + 1
            Debug.Print "Failed        " & name & ": " & reason
        Case TestInconclusive
            this.Inconclusive = this.Inconclusive + 1
            Debug.Print "Inconclusive  " & name & ": " & reason
        Case TestIgnored
            this.Ignored = this.Ignored + 1
            Debug.Print "Ignored       " & name & ": " & reason
    End Select
End Sub

'@Description("Cleans up the current module and initializes the next one, where the module changes; an empty name only cleans up.")
Private Sub EnterModule(ByVal moduleName As String)
Attribute EnterModule.VB_Description = "Cleans up the current module and initializes the next one, where the module changes; an empty name only cleans up."
    If moduleName = this.CurrentModule Then Exit Sub
    If this.CurrentModule <> vbNullString Then this.List.CleanupModule this.CurrentModule
    this.CurrentModule = moduleName
    If moduleName <> vbNullString Then this.List.InitializeModule moduleName
End Sub
