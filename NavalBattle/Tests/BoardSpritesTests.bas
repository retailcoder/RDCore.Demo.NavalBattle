Attribute VB_Name = "BoardSpritesTests"
Attribute VB_Description = "Tests covering the NavalBattle.BoardSprites class."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.BoardSprites class.")
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
Public Sub OpenWater_ShowsNoSprite()
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites
    Dim sprites As Variant
    sprites = SUT.CellSprites(PlayerGrid.Create(PlayerGrid1), True)

    Dim X As Long
    Dim Y As Long
    For X = 1 To PlayerGrid.Size
        For Y = 1 To PlayerGrid.Size
            If sprites(X, Y) <> vbNullString Then
                Assert.Fail "(" & X & "," & Y & ") shows " & sprites(X, Y)
                Exit Sub
            End If
        Next
    Next
    Assert.IsTrue True
End Sub

'@TestMethod
Public Sub HitsAndMisses_ShowWhereTheyWerePlayed()
    Dim grid As PlayerGrid
    Set grid = DestroyerGrid(Horizontal)
    grid.TryHit GridCoord.Create(2, 3)
    grid.TryHit GridCoord.Create(5, 5)

    Dim sprites As Variant
    sprites = ShownSprites(grid, True)

    Assert.AreEqual "duck-small-right-red", sprites(2, 3), "the hit, its duck red"
    Assert.AreEqual "miss-ripple", sprites(5, 5), "the miss"
    Assert.AreEqual "duck-big-right", sprites(3, 3), "the bow"
End Sub

'@TestMethod
Public Sub Ships_ShowTheirDucks_TheBigOneAtTheBow()
    Dim sprites As Variant
    sprites = ShownSprites(DestroyerGrid(Vertical), True)

    Assert.AreEqual "duck-small-down", sprites(2, 3), "the stern"
    Assert.AreEqual "duck-big-down", sprites(2, 4), "the bow"
End Sub

'@TestMethod
Public Sub HiddenShips_ShowOnlyTheirHits_AsTheDevilDuck()
    Dim grid As PlayerGrid
    Set grid = DestroyerGrid(Horizontal)
    grid.TryHit GridCoord.Create(2, 3)

    Dim sprites As Variant
    sprites = ShownSprites(grid, False)

    Assert.AreEqual "duck-devil", sprites(2, 3), "the hit, facing front"
    Assert.AreEqual vbNullString, sprites(3, 3), "the bow"
End Sub

'@TestMethod
Public Sub DizzyShip_ShowsItsBowDizzy_UntilItRecovers()
    Dim grid As PlayerGrid
    Set grid = DestroyerGrid(Horizontal)
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites

    SUT.DizzyShip PlayerGrid1, grid.Fleet(1)
    Assert.AreEqual "sunk-dizzy", ShownSprite(SUT, grid, True, 3, 3), "dizzy"

    SUT.DizzyShipRecovers
    Assert.AreEqual "duck-big-right", ShownSprite(SUT, grid, True, 3, 3), "recovered"
End Sub

'@TestMethod
Public Sub SunkShip_ShowsDizzy_ThenItsRedBigDuck_EvenHidden()
    Dim grid As PlayerGrid
    Set grid = DestroyerGrid(Horizontal)
    grid.TryHit GridCoord.Create(2, 3)
    grid.TryHit GridCoord.Create(3, 3)
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites

    Assert.AreEqual "sunk-dizzy", ShownSprite(SUT, grid, False, 3, 3), "while it sinks"
    Assert.AreEqual "duck-small-right-red", ShownSprite(SUT, grid, False, 2, 3), "the stern, shown once sunk"

    SUT.FinishSinking PlayerGrid1, grid.Fleet(1)
    SUT.FinishSinking PlayerGrid1, grid.Fleet(1)
    Assert.AreEqual "duck-big-right-red", ShownSprite(SUT, grid, False, 3, 3), "once it sank"
End Sub

'@TestMethod
Public Sub FallbackColours_TellWaterShipsHitsAndMissesApart()
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites

    Assert.AreEqual SUT.Water, SUT.FallbackColour(vbNullString), "open water"
    Assert.AreEqual SUT.FallbackColour("hit-squeak"), SUT.FallbackColour("sunk-dizzy"), "a dizzy duck is a hit"
    Assert.IsTrue SUT.FallbackColour("duck-big-right") <> SUT.Water, "a ship is not water"
    Assert.IsTrue SUT.FallbackColour("miss-ripple") <> SUT.FallbackColour("hit-squeak"), "a miss is not a hit"
    Assert.AreEqual SUT.FallbackColour("hit-squeak"), SUT.FallbackColour("duck-big-down-red"), "a red duck is a hit"
    Assert.AreEqual SUT.FallbackColour("hit-squeak"), SUT.FallbackColour("duck-devil"), "the devil duck is a hit"
    Assert.IsTrue SUT.IsHit("hit-squeak") And SUT.IsHit("duck-small-right-red") And SUT.IsHit("duck-devil"), "the hits write an X"
    Assert.IsFalse SUT.IsHit("sunk-dizzy") Or SUT.IsHit("duck-big-right"), "a dizzy duck and a duck do not"
    Assert.IsTrue SUT.IsMiss("miss-ripple"), "the ripple is a miss"
    Assert.IsFalse SUT.IsMiss("hit-squeak") Or SUT.IsMiss(vbNullString), "a hit and open water are not"
    Assert.IsTrue SUT.IsSqueak("hit-squeak") And Not SUT.IsSqueak("duck-devil"), "only the squeak squeaks"
End Sub

'@TestMethod
Public Sub EverySprite_IsAPngInTheWorkspaceAssets()
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites

    Assert.AreEqual 12, UBound(SUT.Names) - LBound(SUT.Names) + 1, "the sprites: four ducks and their four red ones, the devil duck, the squeak, the miss and the dizzy duck"
    Assert.AreEqual CurDir$ & "\assets\sprites\hit-squeak.png", SUT.Path("hit-squeak")
End Sub

'@TestMethod
Public Sub FleetGrid_IsTheFirstGridPaintedWithItsShips()
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites

    Assert.AreEqual 0, SUT.FleetGridId, "before any paint"
    Assert.IsFalse SUT.IsFleetGrid(PlayerGrid1, False), "grid 1, its ships hidden"
    Assert.IsTrue SUT.IsFleetGrid(PlayerGrid2, True), "grid 2, its ships shown"
    Assert.IsFalse SUT.IsFleetGrid(PlayerGrid1, True), "grid 1, its ships shown later"
    Assert.IsTrue SUT.IsFleetGrid(PlayerGrid2, False), "grid 2, its ships hidden later"
    Assert.AreEqual PlayerGrid2, SUT.FleetGridId, "after"
End Sub

'@TestMethod
Public Sub FleetSprites_ShowEachShipAcross_InTheOrderTheyWerePlaced()
    Dim grid As PlayerGrid
    Set grid = PlayerGrid.Create(PlayerGrid1)
    grid.AddShip Ship.Create(ShipType.Cruiser, Vertical, GridCoord.Create(5, 5))
    grid.AddShip Ship.Create(ShipType.Destroyer, Horizontal, GridCoord.Create(1, 1))
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites

    Assert.AreEqual 2, SUT.FleetSprites(grid).Count, "the ships"
    Dim cruiser As Variant
    cruiser = FleetShipSprites(SUT, grid, 1)
    Assert.AreEqual 3, UBound(cruiser), "the first ship placed, the cruiser"
    Assert.AreEqual "duck-small-right", cruiser(1), "its stern, across"
    Assert.AreEqual "duck-big-right", cruiser(3), "its bow, across"
    Dim destroyer As Variant
    destroyer = FleetShipSprites(SUT, grid, 2)
    Assert.AreEqual 2, UBound(destroyer), "the second ship placed, the destroyer"
End Sub

'@TestMethod
Public Sub FleetSprites_ShowTheHitDucksRed_AndNoOtherSprite()
    Dim grid As PlayerGrid
    Set grid = DestroyerGrid(Vertical)
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites
    Dim destroyer As Variant

    grid.TryHit GridCoord.Create(2, 3)
    SUT.DizzyShip PlayerGrid1, grid.Fleet(1)
    destroyer = FleetShipSprites(SUT, grid, 1)
    Assert.AreEqual "duck-small-right-red", destroyer(1), "the hit stern"
    Assert.AreEqual "duck-big-right", destroyer(2), "the bow, dizzy on the grid"

    grid.TryHit GridCoord.Create(2, 4)
    SUT.FinishSinking PlayerGrid1, grid.Fleet(1)
    destroyer = FleetShipSprites(SUT, grid, 1)
    Assert.AreEqual "duck-small-right-red", destroyer(1), "the sunk ship's stern"
    Assert.AreEqual "duck-big-right-red", destroyer(2), "its bow, red on the grid too"
End Sub

'@TestMethod
Public Sub Squeak_ShowsOnTheHitCell_OfItsGridOnly_UntilItStops()
    Dim grid As PlayerGrid
    Set grid = DestroyerGrid(Horizontal)
    grid.TryHit GridCoord.Create(2, 3)
    Dim otherGrid As PlayerGrid
    Set otherGrid = PlayerGrid.Create(PlayerGrid2)
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites

    SUT.Squeak PlayerGrid1, GridCoord.Create(2, 3)
    Assert.AreEqual "hit-squeak", ShownSprite(SUT, grid, False, 2, 3), "hidden"
    Assert.AreEqual "hit-squeak", ShownSprite(SUT, grid, True, 2, 3), "shown"
    Assert.AreEqual vbNullString, ShownSprite(SUT, otherGrid, True, 2, 3), "the other grid"

    SUT.SqueakStops
    Assert.AreEqual "duck-devil", ShownSprite(SUT, grid, False, 2, 3), "after"
End Sub

' what the given BoardSprites shows of a ship in the fleet status, by its place in the fleet.
Private Function FleetShipSprites(ByVal board As BoardSprites, ByVal grid As PlayerGrid, ByVal shipNumber As Long) As Variant
    Dim fleet As Collection
    Set fleet = board.FleetSprites(grid)
    FleetShipSprites = fleet(shipNumber)
End Function

' what a new BoardSprites shows on a grid.
Private Function ShownSprites(ByVal grid As PlayerGrid, ByVal showShips As Boolean) As Variant
    Dim board As BoardSprites
    Set board = New BoardSprites
    ShownSprites = board.CellSprites(grid, showShips)
End Function

' what the given BoardSprites shows on a cell of a grid.
Private Function ShownSprite(ByVal board As BoardSprites, ByVal grid As PlayerGrid, ByVal showShips As Boolean, ByVal X As Long, ByVal Y As Long) As String
    Dim result As Variant
    result = board.CellSprites(grid, showShips)
    ShownSprite = result(X, Y)
End Function

' grid 1, with a destroyer (two cells) from (2,3), across or down.
Private Function DestroyerGrid(ByVal direction As ShipOrientation) As PlayerGrid
    Dim result As PlayerGrid
    Set result = PlayerGrid.Create(PlayerGrid1)
    result.AddShip Ship.Create(ShipType.Destroyer, direction, GridCoord.Create(2, 3))
    Set DestroyerGrid = result
End Function
