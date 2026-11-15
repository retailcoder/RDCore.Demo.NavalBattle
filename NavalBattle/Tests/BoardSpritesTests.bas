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

    Assert.AreEqual "hit-squeak", sprites(2, 3), "the hit"
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
Public Sub HiddenShips_ShowOnlyTheirHits()
    Dim grid As PlayerGrid
    Set grid = DestroyerGrid(Horizontal)
    grid.TryHit GridCoord.Create(2, 3)

    Dim sprites As Variant
    sprites = ShownSprites(grid, False)

    Assert.AreEqual "hit-squeak", sprites(2, 3), "the hit"
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
Public Sub SunkShip_ShowsDizzy_ThenTheAnchor_EvenHidden()
    Dim grid As PlayerGrid
    Set grid = DestroyerGrid(Horizontal)
    grid.TryHit GridCoord.Create(2, 3)
    grid.TryHit GridCoord.Create(3, 3)
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites

    Assert.AreEqual "sunk-dizzy", ShownSprite(SUT, grid, False, 3, 3), "while it sinks"
    Assert.AreEqual "hit-squeak", ShownSprite(SUT, grid, False, 2, 3), "the stern"

    SUT.FinishSinking PlayerGrid1, grid.Fleet(1)
    SUT.FinishSinking PlayerGrid1, grid.Fleet(1)
    Assert.AreEqual "anchor-cracked", ShownSprite(SUT, grid, False, 3, 3), "once it sank"
End Sub

'@TestMethod
Public Sub FallbackColours_TellWaterShipsHitsAndMissesApart()
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites

    Assert.AreEqual SUT.Water, SUT.FallbackColour(vbNullString), "open water"
    Assert.AreEqual SUT.FallbackColour("hit-squeak"), SUT.FallbackColour("sunk-dizzy"), "a dizzy duck is a hit"
    Assert.AreEqual SUT.FallbackColour("hit-squeak"), SUT.FallbackColour("anchor-cracked"), "an anchor is a hit"
    Assert.IsTrue SUT.FallbackColour("duck-big-right") <> SUT.Water, "a ship is not water"
    Assert.IsTrue SUT.FallbackColour("miss-ripple") <> SUT.FallbackColour("hit-squeak"), "a miss is not a hit"
    Assert.IsTrue SUT.IsHit("hit-squeak") And Not SUT.IsHit("sunk-dizzy"), "only the hit writes an X"
End Sub

'@TestMethod
Public Sub EverySprite_IsAGifInTheWorkspaceAssets()
    Dim SUT As BoardSprites
    Set SUT = New BoardSprites

    Assert.AreEqual 8, UBound(SUT.Names) - LBound(SUT.Names) + 1, "the sprites"
    Assert.AreEqual CurDir$ & "\assets\sprites\hit-squeak.gif", SUT.Path("hit-squeak")
End Sub

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
