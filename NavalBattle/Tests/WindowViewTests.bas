Attribute VB_Name = "WindowViewTests"
Attribute VB_Description = "Tests covering the NavalBattle.WindowView class: a window that fails before it opens, and one that fails once it is open."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.WindowView class: a window that fails before it opens, and one that fails once it is open.")
'@TestModule
Option Explicit
Option Private Module

' the error a forms host that cannot start raises on macOS (E_FAIL).
Private Const StartFailed As Long = &H80004005

Private Assert As TestAssert

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = TestAssert
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestMethod
Public Sub WindowFailingBeforeItOpens_TryPlayReturnsFalse_WithTheError()
    Dim window As FailingWindow
    Set window = New FailingWindow
    window.FailsWith StartFailed, opensFirst:=False
    Dim SUT As WindowView
    Set SUT = WindowView.Create(window)

    Assert.IsFalse SUT.TryPlay("Random", "Random"), "the game goes to the next view"
    Assert.AreEqual StartFailed, SUT.FailureNumber, "the error it failed with"
    Assert.AreEqual "The window failed.", SUT.FailureDescription
End Sub

'@TestMethod
Public Sub WindowFailingOnceOpen_RaisesTheGamesError()
    Dim window As FailingWindow
    Set window = New FailingWindow
    window.FailsWith 5, opensFirst:=True
    Dim SUT As WindowView
    Set SUT = WindowView.Create(window)

    Dim raised As Long
    On Error Resume Next
    SUT.TryPlay "Random", "Random"
    raised = Err.Number
    On Error GoTo 0

    Assert.AreEqual 5, raised, "an error in a game being played is not a window that cannot open"
    Assert.AreEqual 0, SUT.FailureNumber
End Sub
