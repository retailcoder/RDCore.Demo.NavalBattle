Attribute VB_Name = "BoardFormTests"
Attribute VB_Description = "Tests covering the NavalBattle.BoardForm form, the RDForms window. Each loads the form, without showing it."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.BoardForm form, the RDForms window. Each loads the form, without showing it.")
'@TestModule
Option Explicit
Option Private Module

Private Assert As TestAssert
' the form places its controls by BoardLayout's points, 20 twips each.
Private Const TwipsPerPoint As Long = 20

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = TestAssert
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

' once loaded, the form has a cell for each square of both grids, each where BoardLayout places it: the cells of grid 1, row by row, then grid 2's.
'@TestMethod
Public Sub Cells_AreLoadedFromTheDesignedCell()
    Dim SUT As BoardForm
    Set SUT = New BoardForm
    Load SUT
    Dim cells As Long
    cells = PlayerGrid.Size * PlayerGrid.Size
    Assert.AreEqual 2 * cells, CLng(SUT.lblCell.Count), "the cells"
    Dim gridId As PlayGridId
    For gridId = PlayerGrid1 To PlayerGrid2
        Dim corner As Variant
        For Each corner In Array(Array(1, 1), Array(PlayerGrid.Size, 1), Array(1, PlayerGrid.Size), Array(PlayerGrid.Size, PlayerGrid.Size))
            Dim cellNumber As Long
            cellNumber = (gridId - PlayerGrid1) * cells + (corner(1) - 1) * PlayerGrid.Size + (corner(0) - 1)
            Assert.AreEqual BoardLayout.CellLeft(gridId, corner(0)) * TwipsPerPoint, CLng(SUT.lblCell(cellNumber).Left), "grid " & gridId & ", cell " & cellNumber & ": its left"
            Assert.AreEqual BoardLayout.CellTop(corner(1)) * TwipsPerPoint, CLng(SUT.lblCell(cellNumber).Top), "grid " & gridId & ", cell " & cellNumber & ": its top"
        Next
    Next
    Unload SUT
End Sub

' once loaded, each picker offers PlayerNames in order: grid 1 picks Human, and grid 2 Merciless AI.
'@TestMethod
Public Sub PlayerPickers_OfferEveryPlayer()
    Dim SUT As BoardForm
    Set SUT = New BoardForm
    Load SUT
    Dim names As Variant
    names = PlayerNames
    Dim picker As Long
    For picker = 0 To 1
        With SUT.cboPlayer(picker)
            Assert.AreEqual UBound(names) - LBound(names) + 1, CLng(.ListCount), "picker " & picker & ": its players"
            Dim index As Long
            For index = 0 To .ListCount - 1
                Assert.AreEqual names(LBound(names) + index), .List(index), "picker " & picker & ", player " & index
            Next
        End With
    Next
    Assert.AreEqual PlayerName(Unspecific), SUT.cboPlayer(0).Text, "grid 1"
    Assert.AreEqual PlayerName(MercilessAI), SUT.cboPlayer(1).Text, "grid 2"
    Unload SUT
End Sub

' a player picked from the command line shows in its grid's picker, read the way a person types it.
'@TestMethod
Public Sub PickedPlayer_ShowsInItsGridsPicker()
    Dim SUT As BoardForm
    Set SUT = New BoardForm
    Load SUT
    Dim window As IGameWindow
    Set window = SUT
    window.PickPlayer PlayerGrid1, "f"
    window.PickPlayer PlayerGrid2, "Random"
    Assert.AreEqual PlayerName(FairplayAI), SUT.cboPlayer(0).Text, "grid 1"
    Assert.AreEqual PlayerName(RandomAI), SUT.cboPlayer(1).Text, "grid 2"
    Unload SUT
End Sub

' once the ships are placed, the fleet status names each one, in the order they were placed, over a cell for each of its parts.
'@TestMethod
Public Sub FleetStatus_ShowsEveryShipPlaced()
    Dim grid As PlayerGrid
    Set grid = PlayerGrid.Create(PlayerGrid1)
    grid.AddShip Ship.Create(ShipType.Cruiser, Vertical, GridCoord.Create(5, 5))
    grid.AddShip Ship.Create(ShipType.Destroyer, Horizontal, GridCoord.Create(1, 1))
    Dim SUT As BoardForm
    Set SUT = New BoardForm
    Load SUT
    Dim window As IGameWindow
    Set window = SUT
    window.PaintGrid grid, True

    Assert.AreEqual 2, CLng(SUT.lblShip.Count), "a name for each ship"
    Dim shipNumber As Long
    For shipNumber = 1 To 2
        Dim item As IShip
        Set item = grid.Fleet(shipNumber)
        With SUT.lblShip(shipNumber - 1)
            Assert.AreEqual item.Name, .Caption, "ship " & shipNumber & ": its name"
            Assert.AreEqual BoardLayout.ShipNameTop(shipNumber) * TwipsPerPoint, CLng(.Top), "ship " & shipNumber & ": its name's top"
        End With
        Dim column As Long
        For column = 1 To item.Size
            Assert.IsTrue SUT.imgShipCell((shipNumber - 1) * PlayerGrid.Size + column - 1).Visible, "ship " & shipNumber & ", cell " & column
        Next
    Next
    Unload SUT
End Sub
