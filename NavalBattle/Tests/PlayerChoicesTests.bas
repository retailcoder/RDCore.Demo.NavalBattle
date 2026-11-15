Attribute VB_Name = "PlayerChoicesTests"
Attribute VB_Description = "Tests covering the NavalBattle.PlayerChoices module."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.PlayerChoices module.")
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
Public Sub FirstLetter_IsEnough_InEitherCase()
    Dim difficulty As AIDifficulty
    Assert.IsTrue TryParsePlayer("h", difficulty), "h"
    Assert.AreEqual Unspecific, difficulty, "h"
    Assert.IsTrue TryParsePlayer("Random", difficulty), "Random"
    Assert.AreEqual RandomAI, difficulty, "Random"
    Assert.IsTrue TryParsePlayer(" fair-play", difficulty), "fair-play"
    Assert.AreEqual FairplayAI, difficulty, "fair-play"
    Assert.IsTrue TryParsePlayer("M", difficulty), "M"
    Assert.AreEqual MercilessAI, difficulty, "M"
End Sub

'@TestMethod
Public Sub NotAPlayer_IsNotRead()
    Dim difficulty As AIDifficulty
    Assert.IsFalse TryParsePlayer("Xavier", difficulty), "Xavier"
    Assert.IsFalse TryParsePlayer(vbNullString, difficulty), "nothing"
End Sub

'@TestMethod
Public Sub EveryName_ReadsBackAsItsPlayer()
    Dim name As Variant
    For Each name In PlayerNames
        Dim difficulty As AIDifficulty
        Assert.IsTrue TryParsePlayer(name, difficulty), name
        Assert.AreEqual name, PlayerName(difficulty), name
    Next
End Sub

'@TestMethod
Public Sub TwoPeople_CannotPlayTogether()
    Assert.IsFalse CanPlayTogether(Unspecific, Unspecific), "two people"
    Assert.IsTrue CanPlayTogether(Unspecific, MercilessAI), "a person and the computer"
    Assert.IsTrue CanPlayTogether(RandomAI, Unspecific), "the computer and a person"
    Assert.IsTrue CanPlayTogether(FairplayAI, MercilessAI), "the computer on both grids"
End Sub
