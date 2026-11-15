Attribute VB_Name = "GameControllerTests"
Attribute VB_Description = "Tests covering the NavalBattle.StandardGameController class."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.StandardGameController class.")
'@TestModule
Option Explicit
Option Private Module

Private Assert As TestAssert
' the controller of a game a person plays: it outlives the test's first steps, since the person plays it a step at a time.
Private personGame As IGameController

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = TestAssert
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    Set personGame = Nothing
    Set Assert = Nothing
End Sub

'@TestMethod
Public Sub ComputerVsComputer_OnSinkNamesEachSunkShip()
    Dim view As SinkRecorderView
    Set view = New SinkRecorderView
    Dim players As ScriptedPlayerFactory
    Set players = New ScriptedPlayerFactory

    Dim controller As IGameController
    Set controller = StandardGameController.Create(GridViewAdapter.Create(view), New GameRandomizer, players)
    controller.NewGame
    view.ViewEvents.CreatePlayer PlayerGrid1, ComputerControlled, MercilessAI
    view.ViewEvents.CreatePlayer PlayerGrid2, ComputerControlled, MercilessAI

    ' both players place and fire alike, and grid 1 fires first: it sinks grid 2's fifth ship, and wins, a shot before grid 2 would sink
    ' its own fifth. The ship that ends the game is reported as the game over, so each grid reports 4 sinkings.
    Assert.AreEqual PlayerGrid1, view.WinningGrid, "the winner"
    Assert.AreEqual 8, view.SunkShips.Count, "the sinkings"

    Dim index As Long
    For index = 1 To view.SunkShips.Count
        Dim sunkShip As IShip
        Set sunkShip = view.SunkShips(index)
        Assert.IsTrue sunkShip.IsSunken, "sinking " & index & ": the " & sunkShip.Name & " is not sunk"
        Assert.IsTrue IsInFleet(sunkShip, players.Player(view.SunkGridIds(index)).Grid), "sinking " & index & ": the " & sunkShip.Name & " is not in grid " & view.SunkGridIds(index) & "'s fleet"
    Next
End Sub

'@TestMethod
Public Sub ComputerVsComputer_EachSinkingIsReportedWithTheShotThatSankTheShip()
    Dim view As SinkRecorderView
    Set view = New SinkRecorderView
    Dim players As ScriptedPlayerFactory
    Set players = New ScriptedPlayerFactory

    Dim controller As IGameController
    Set controller = StandardGameController.Create(GridViewAdapter.Create(view), New GameRandomizer, players)
    controller.NewGame
    view.ViewEvents.CreatePlayer PlayerGrid1, ComputerControlled, MercilessAI
    view.ViewEvents.CreatePlayer PlayerGrid2, ComputerControlled, MercilessAI

    Dim index As Long
    For index = 1 To view.SunkShips.Count
        Assert.IsTrue IsOnShip(view.SunkPositions(index), view.SunkShips(index)), "sinking " & index & ": the shot is not on the " & view.SunkShips(index).Name
    Next

    ' the game over is the fifth sinking of grid 2's fleet.
    Assert.IsNotNothing view.LastShip, "the game over's ship"
    Assert.IsTrue view.LastShip.IsSunken, "the game over's " & view.LastShip.Name & " is not sunk"
    Assert.IsTrue IsInFleet(view.LastShip, players.Player(PlayerGrid2).Grid), "the game over's " & view.LastShip.Name & " is not in grid 2's fleet"
    Assert.IsTrue IsOnShip(view.LastPosition, view.LastShip), "the game over's shot is not on the " & view.LastShip.Name
End Sub

'@TestMethod
Public Sub ComputerVsComputer_BothGridsStayLocked()
    Dim view As SinkRecorderView
    Set view = New SinkRecorderView
    Dim controller As IGameController
    Set controller = StandardGameController.Create(GridViewAdapter.Create(view), New GameRandomizer, New ScriptedPlayerFactory)
    controller.NewGame
    view.ViewEvents.CreatePlayer PlayerGrid1, ComputerControlled, MercilessAI
    view.ViewEvents.CreatePlayer PlayerGrid2, ComputerControlled, MercilessAI

    Assert.IsTrue view.IsLocked(PlayerGrid1), "grid 1"
    Assert.IsTrue view.IsLocked(PlayerGrid2), "grid 2"
    Assert.AreEqual 0&, view.Unlocks, "the unlocks"
End Sub

'@TestMethod
Public Sub PersonPlacingShips_OnlyTheirGridIsUnlocked()
    Dim view As SinkRecorderView
    Set view = StartPersonVsComputer

    Assert.IsFalse view.IsLocked(PlayerGrid1), "the person's grid"
    Assert.IsTrue view.IsLocked(PlayerGrid2), "the computer's grid"
End Sub

'@TestMethod
Public Sub PersonFiring_OnlyTheOtherGridIsUnlocked()
    Dim view As SinkRecorderView
    Set view = StartPersonVsComputer

    ' the person's five ships, across the grid, a row apart.
    Dim index As Long
    For index = 1 To PlayerGrid.ShipsPerGrid
        view.ViewEvents.PreviewShipPosition PlayerGrid1, GridCoord.Create(1, 2 * index - 1)
        view.ViewEvents.ConfirmShipPosition PlayerGrid1, GridCoord.Create(1, 2 * index - 1)
    Next

    Assert.IsTrue view.IsLocked(PlayerGrid1), "the person's grid"
    Assert.IsFalse view.IsLocked(PlayerGrid2), "the computer's grid"
End Sub

' a game of a person on grid 1 against the computer, with no pause before the computer's shots; the person has placed nothing yet.
Private Function StartPersonVsComputer() As SinkRecorderView
    Dim randomizer As IRandomizer
    Set randomizer = New GameRandomizer
    Dim view As SinkRecorderView
    Set view = New SinkRecorderView
    Dim controller As IGameController
    Set controller = StandardGameController.Create(GridViewAdapter.Create(view), randomizer, PlayerFactory.Create(randomizer, 0))
    controller.NewGame
    view.ViewEvents.CreatePlayer PlayerGrid1, HumanControlled, Unspecific
    view.ViewEvents.CreatePlayer PlayerGrid2, ComputerControlled, RandomAI
    ' nothing else holds the controller: it is kept here, so the game goes on once this returns.
    Set personGame = controller
    Set StartPersonVsComputer = view
End Function

Private Function IsOnShip(ByVal position As IGridCoord, ByVal item As IShip) As Boolean
    Dim offset As Long
    For offset = 0 To item.Size - 1
        If item.Orientation = Horizontal Then
            If position.X = item.GridPosition.X + offset And position.Y = item.GridPosition.Y Then IsOnShip = True
        Else
            If position.X = item.GridPosition.X And position.Y = item.GridPosition.Y + offset Then IsOnShip = True
        End If
    Next
End Function

Private Function IsInFleet(ByVal item As IShip, ByVal grid As PlayerGrid) As Boolean
    Dim fleetShip As IShip
    For Each fleetShip In grid.Fleet
        If fleetShip Is item Then
            IsInFleet = True
            Exit Function
        End If
    Next
End Function
