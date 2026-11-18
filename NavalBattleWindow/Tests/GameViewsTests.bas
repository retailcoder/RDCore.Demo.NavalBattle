Attribute VB_Name = "GameViewsTests"
Attribute VB_Description = "Tests covering the window app's choice of window, Program.GameView."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the window app's choice of window, Program.GameView.")
'@TestModule
Option Explicit
Option Private Module

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
Public Sub NoViewSwitch_PlaysInTheRDFormsWindow()
    Assert.AreEqual ViewRDForms, GameView("--seed 42")
End Sub

'@TestMethod
Public Sub RDFormsSwitch_PlaysInTheRDFormsWindow()
    Assert.AreEqual ViewRDForms, GameView("--unattended --rdforms")
End Sub

'@TestMethod
Public Sub MSFormsSwitch_PlaysInTheMicrosoftFormsWindow()
    Assert.AreEqual ViewMicrosoftForms, GameView("--unattended --msforms")
End Sub

'@TestMethod
Public Sub ConsoleSwitch_IsNoView_TheConsoleAppPlaysThere()
    Assert.AreEqual ViewRDForms, GameView("--console")
    Assert.AreEqual ViewMicrosoftForms, GameView("--msforms --console")
End Sub
