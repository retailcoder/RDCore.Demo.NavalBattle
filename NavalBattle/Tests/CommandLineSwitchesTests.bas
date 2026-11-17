Attribute VB_Name = "CommandLineSwitchesTests"
Attribute VB_Description = "Tests covering the NavalBattle.CommandLineSwitches module."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.CommandLineSwitches module.")
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
Public Sub RuntimeSwitchBefore_TheSeedIsRead()
    Dim seed As String
    Assert.IsTrue TryGetArgument("--unattended --seed 42", "--seed", seed)
    Assert.AreEqual "42", seed
End Sub

'@TestMethod
Public Sub RuntimeSwitchAfter_TheSeedIsRead()
    Dim seed As String
    Assert.IsTrue TryGetArgument("--seed 42 --unattended", "--seed", seed)
    Assert.AreEqual "42", seed
End Sub

'@TestMethod
Public Sub RuntimeSwitchInTheValuesPlace_IsNoValue()
    Dim seed As String
    Assert.IsFalse TryGetArgument("--seed --unattended", "--seed", seed)
    Assert.AreEqual vbNullString, seed
End Sub

'@TestMethod
Public Sub SwitchesItIsNotAskedAbout_AreIgnored()
    Assert.IsTrue HasSwitch("--unattended --console --someday", "--console")
    Assert.IsFalse HasSwitch("--unattended --someday", "--test")
End Sub

'@TestMethod
Public Sub NoViewSwitch_TriesRDFormsThenTheConsole()
    Assert.AreEqual "RDForms,Console", Join(GameViews("--seed 42"), ",")
End Sub

'@TestMethod
Public Sub RDFormsSwitch_TriesRDFormsThenTheConsole()
    Assert.AreEqual "RDForms,Console", Join(GameViews("--unattended --rdforms"), ",")
End Sub

'@TestMethod
Public Sub MSFormsSwitch_TriesMicrosoftFormsThenTheConsole()
    Assert.AreEqual "MSForms,Console", Join(GameViews("--unattended --msforms"), ",")
End Sub

'@TestMethod
Public Sub ConsoleSwitch_PlaysInTheConsoleAlone_EvenWithAWindowSwitch()
    Assert.AreEqual "Console", Join(GameViews("--msforms --console"), ",")
    Assert.AreEqual "Console", Join(GameViews("--rdforms --console"), ",")
End Sub
