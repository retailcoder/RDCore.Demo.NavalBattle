VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "The assertions of the test modules: the members of legacy Rubberduck's AssertClass that the tests use. A failed assertion does not stop its test; the test's outcome is its first failure, or Inconclusive."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("The assertions of the test modules: the members of legacy Rubberduck's AssertClass that the tests use. A failed assertion does not stop its test; the test's outcome is its first failure, or Inconclusive.")
'@PredeclaredId
Option Explicit

Public Enum TestOutcome
    TestPassed
    TestFailed
    TestInconclusive
End Enum

Private Type TAssert
    Outcome As TestOutcome
    Reason As String
End Type

Private this As TAssert

'@Description("Starts a test: its outcome is Passed until an assertion says otherwise.")
Public Sub Reset()
Attribute Reset.VB_Description = "Starts a test: its outcome is Passed until an assertion says otherwise."
    this.Outcome = TestPassed
    this.Reason = vbNullString
End Sub

'@Description("The outcome of the test that is running, or that ran last.")
Public Property Get Outcome() As TestOutcome
Attribute Outcome.VB_Description = "The outcome of the test that is running, or that ran last."
    Outcome = this.Outcome
End Property

'@Description("Why the test failed or is inconclusive; empty when it passed.")
Public Property Get Reason() As String
Attribute Reason.VB_Description = "Why the test failed or is inconclusive; empty when it passed."
    Reason = this.Reason
End Property

'@Description("Fails the test where the condition is not True.")
Public Sub IsTrue(ByVal condition As Boolean, Optional ByVal Message As String = vbNullString)
Attribute IsTrue.VB_Description = "Fails the test where the condition is not True."
    If Not condition Then Record TestFailed, "IsTrue failed. " & Message
End Sub

'@Description("Fails the test where the condition is not False.")
Public Sub IsFalse(ByVal condition As Boolean, Optional ByVal Message As String = vbNullString)
Attribute IsFalse.VB_Description = "Fails the test where the condition is not False."
    If condition Then Record TestFailed, "IsFalse failed. " & Message
End Sub

'@Description("Fails the test where the two values are not equal.")
Public Sub AreEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal Message As String = vbNullString)
Attribute AreEqual.VB_Description = "Fails the test where the two values are not equal."
    If expected <> actual Then Record TestFailed, "AreEqual failed: expected " & expected & ", actual " & actual & ". " & Message
End Sub

'@Description("Fails the test where the value is not Nothing.")
Public Sub IsNothing(ByVal value As Object, Optional ByVal Message As String = vbNullString)
Attribute IsNothing.VB_Description = "Fails the test where the value is not Nothing."
    If Not value Is Nothing Then Record TestFailed, "IsNothing failed. " & Message
End Sub

'@Description("Fails the test where the value is Nothing.")
Public Sub IsNotNothing(ByVal value As Object, Optional ByVal Message As String = vbNullString)
Attribute IsNotNothing.VB_Description = "Fails the test where the value is Nothing."
    If value Is Nothing Then Record TestFailed, "IsNotNothing failed. " & Message
End Sub

'@Description("Fails the test.")
Public Sub Fail(Optional ByVal Message As String = vbNullString)
Attribute Fail.VB_Description = "Fails the test."
    Record TestFailed, "Fail. " & Message
End Sub

'@Description("Makes the test inconclusive: what it needs to begin with is not there.")
Public Sub Inconclusive(Optional ByVal Message As String = vbNullString)
Attribute Inconclusive.VB_Description = "Makes the test inconclusive: what it needs to begin with is not there."
    Record TestInconclusive, "Inconclusive. " & Message
End Sub

'@Description("Records an outcome: the first one that is not Passed is the test's.")
Private Sub Record(ByVal value As TestOutcome, ByVal text As String)
Attribute Record.VB_Description = "Records an outcome: the first one that is not Passed is the test's."
    If this.Outcome <> TestPassed Then Exit Sub
    this.Outcome = value
    this.Reason = Trim$(text)
End Sub
