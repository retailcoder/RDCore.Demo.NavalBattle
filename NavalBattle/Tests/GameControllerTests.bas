Attribute VB_Name = "GameControllerTests"
Attribute VB_Description = "Tests covering the NavalBattle.StandardGameController class."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.StandardGameController class.")
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

Private Function IsInFleet(ByVal item As IShip, ByVal grid As PlayerGrid) As Boolean
    Dim fleetShip As IShip
    For Each fleetShip In grid.Fleet
        If fleetShip Is item Then
            IsInFleet = True
            Exit Function
        End If
    Next
End Function
