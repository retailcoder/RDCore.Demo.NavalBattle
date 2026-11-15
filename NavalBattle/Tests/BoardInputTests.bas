Attribute VB_Name = "BoardInputTests"
Attribute VB_Description = "Tests covering the NavalBattle.BoardInput class."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.BoardInput class.")
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
Public Sub LockedGrid_TakesNoClicks()
    Dim SUT As BoardInput
    Dim window As BoardInputRecorder
    Set window = Watched(SUT)
    SUT.BeginShipPlacement PlayerGrid1
    SUT.LockGrid PlayerGrid1, True

    SUT.Hover PlayerGrid1, GridCoord.Create(3, 4)
    SUT.Click PlayerGrid1, GridCoord.Create(3, 4)
    SUT.RightClick PlayerGrid1, GridCoord.Create(3, 4)

    Assert.AreEqual 0, window.Calls.Count
End Sub

'@TestMethod
Public Sub ShipPlacementClick_PreviewsThenConfirms()
    Dim SUT As BoardInput
    Dim window As BoardInputRecorder
    Set window = Watched(SUT)
    SUT.BeginShipPlacement PlayerGrid1

    SUT.Click PlayerGrid1, GridCoord.Create(3, 4)

    Assert.AreEqual 2, window.Calls.Count, "the calls"
    Assert.AreEqual "PreviewShipPosition 1 (3,4)", window.Calls(1)
    Assert.AreEqual "ConfirmShipPosition 1 (3,4)", window.Calls(2)
End Sub

'@TestMethod
Public Sub ShipPlacementHover_PreviewsOncePerCell()
    Dim SUT As BoardInput
    Dim window As BoardInputRecorder
    Set window = Watched(SUT)
    SUT.BeginShipPlacement PlayerGrid1

    SUT.Hover PlayerGrid1, GridCoord.Create(2, 2)
    SUT.Hover PlayerGrid1, GridCoord.Create(2, 2)
    SUT.Hover PlayerGrid1, GridCoord.Create(2, 3)

    Assert.AreEqual 2, window.Calls.Count, "the calls"
    Assert.AreEqual "PreviewShipPosition 1 (2,2)", window.Calls(1)
    Assert.AreEqual "PreviewShipPosition 1 (2,3)", window.Calls(2)
End Sub

'@TestMethod
Public Sub ShipPlacementRightClick_TurnsTheShip()
    Dim SUT As BoardInput
    Dim window As BoardInputRecorder
    Set window = Watched(SUT)
    SUT.BeginShipPlacement PlayerGrid1

    SUT.RightClick PlayerGrid1, GridCoord.Create(5, 6)

    Assert.AreEqual 1, window.Calls.Count, "the calls"
    Assert.AreEqual "PreviewRotateShip 1 (5,6)", window.Calls(1)
End Sub

'@TestMethod
Public Sub Ready_ConfirmsThePlayers_AndMakesThem()
    Dim SUT As BoardInput
    Dim window As BoardInputRecorder
    Set window = Watched(SUT)

    SUT.ReadyClicked PlayerName(Unspecific), PlayerName(MercilessAI)

    Assert.IsTrue window.PlayersConfirmed, "the players are confirmed"
    Assert.IsFalse window.ReadyVisible, "Ready is hidden"
    Assert.IsFalse SUT.ComputerOnly, "a person plays"
    Assert.AreEqual 2, window.Calls.Count, "the calls"
    Assert.AreEqual "CreatePlayer 1 Human", window.Calls(1)
    Assert.AreEqual "CreatePlayer 2 Merciless AI", window.Calls(2)
End Sub

'@TestMethod
Public Sub Ready_TwoPeople_AreRefused()
    Dim SUT As BoardInput
    Dim window As BoardInputRecorder
    Set window = Watched(SUT)

    SUT.ReadyClicked PlayerName(Unspecific), PlayerName(Unspecific)

    Assert.IsFalse window.PlayersConfirmed, "the players are confirmed"
    Assert.AreEqual 0, window.Calls.Count, "the calls"
    Assert.AreEqual "Only one grid can be played by a person: pick an AI for the other one.", window.Status
End Sub

'@TestMethod
Public Sub Ready_TwoComputers_PlayTheWholeGame()
    Dim SUT As BoardInput
    Dim window As BoardInputRecorder
    Set window = Watched(SUT)

    SUT.ReadyClicked PlayerName(RandomAI), PlayerName(FairplayAI)

    Assert.IsTrue SUT.ComputerOnly, "the computer plays both grids"
    Assert.AreEqual "The computer plays both grids.", window.Status
    Assert.AreEqual 2, window.Calls.Count, "the calls"
End Sub

'@TestMethod
Public Sub PlayerPicked_SaysWhoPlays_UntilTheGameStarts()
    Dim SUT As BoardInput
    Dim window As BoardInputRecorder
    Set window = Watched(SUT)

    SUT.PlayerPicked PlayerGrid2, PlayerName(FairplayAI)
    Assert.AreEqual "Grid 2 is played by: Fair-play AI.", window.Status, "before the game"

    SUT.ReadyClicked PlayerName(Unspecific), PlayerName(FairplayAI)
    SUT.BeginShipPlacement PlayerGrid1
    SUT.PlayerPicked PlayerGrid2, PlayerName(RandomAI)
    Assert.AreEqual "Grid 2 is played by: Fair-play AI.", window.Status, "once it started"
End Sub

'@TestMethod
Public Sub ReadyOnceShipsArePlaced_StartsTheBattle()
    Dim SUT As BoardInput
    Dim window As BoardInputRecorder
    Set window = Watched(SUT)
    SUT.ReadyClicked PlayerName(Unspecific), PlayerName(MercilessAI)
    SUT.BeginShipPlacement PlayerGrid1
    SUT.BeginWaitForReady
    Assert.IsTrue window.ReadyVisible, "Ready shows"

    SUT.ReadyClicked PlayerName(Unspecific), PlayerName(MercilessAI)
    SUT.Click PlayerGrid2, GridCoord.Create(7, 8)

    Assert.IsFalse window.ReadyVisible, "Ready is hidden"
    Assert.AreEqual "Fire at grid 2: click a cell.", window.Status
    Assert.AreEqual 4, window.Calls.Count, "the calls"
    Assert.AreEqual "HumanPlayerReady", window.Calls(3)
    Assert.AreEqual "AttackPosition 2 (7,8)", window.Calls(4)
End Sub

'@TestMethod
Public Sub ComputerTurn_TakesNoClicks()
    Dim SUT As BoardInput
    Dim window As BoardInputRecorder
    Set window = Watched(SUT)
    SUT.ReadyClicked PlayerName(Unspecific), PlayerName(MercilessAI)
    SUT.BeginShipPlacement PlayerGrid1
    SUT.BeginWaitForReady
    SUT.ReadyClicked PlayerName(Unspecific), PlayerName(MercilessAI)

    SUT.BeginComputerTurn
    SUT.Click PlayerGrid2, GridCoord.Create(7, 8)
    Assert.AreEqual 3, window.Calls.Count, "during the computer's turn"

    SUT.EndComputerTurn
    SUT.Click PlayerGrid2, GridCoord.Create(7, 8)
    Assert.AreEqual 4, window.Calls.Count, "after it"
End Sub

'@TestMethod
Public Sub EndGame_HidesReady_AndTakesNoClicks()
    Dim SUT As BoardInput
    Dim window As BoardInputRecorder
    Set window = Watched(SUT)
    SUT.BeginShipPlacement PlayerGrid1
    SUT.BeginWaitForReady

    SUT.EndGame
    SUT.Click PlayerGrid1, GridCoord.Create(1, 1)
    SUT.Hover PlayerGrid1, GridCoord.Create(1, 2)

    Assert.IsFalse window.ReadyVisible, "Ready is hidden"
    Assert.AreEqual 0, window.Calls.Count, "the calls"
End Sub

' a new input, and the window that records it.
Private Function Watched(ByRef outInput As BoardInput) As BoardInputRecorder
    Set outInput = New BoardInput
    Dim result As BoardInputRecorder
    Set result = New BoardInputRecorder
    result.Watch outInput
    Set Watched = result
End Function
