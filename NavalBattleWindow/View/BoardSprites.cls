VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoardSprites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "What each cell of a game window's grids shows: a duck, a hit duck in red, a miss or a dizzy duck, by its sprite's name, and the colour it shows instead where the sprites cannot be read. It knows which hit ship is dizzy, which cell squeaks, which sunk ships are done sinking, and which grid's fleet the fleet status shows."
'@Folder("NavalBattle.View.Window")
'@ModuleDescription("What each cell of a game window's grids shows: a duck, a hit duck in red, a miss or a dizzy duck, by its sprite's name, and the colour it shows instead where the sprites cannot be read. It knows which hit ship is dizzy, which cell squeaks, which sunk ships are done sinking, and which grid's fleet the fleet status shows.")
Option Explicit

' the sprites, PNG files in assets/sprites, see-through around the picture; the ducks are named by DuckSprite and RedDuckSprite.
Private Const SpriteFolder As String = "assets\sprites\"
Private Const SpriteSqueak As String = "hit-squeak"
' a hit on a hidden ship: a red duck facing front, which gives away neither the ship's direction nor its bow.
Private Const SpriteHiddenHit As String = "duck-devil"
Private Const RedSuffix As String = "-red"
Private Const SpriteMiss As String = "miss-ripple"
Private Const SpriteDizzy As String = "sunk-dizzy"

Private Const WaterColour As Long = &HF5E1C8
Private Const ShipColour As Long = &H707070
Private Const HitColour As Long = &H3C3CE6
Private Const MissColour As Long = &HFFFFFF
Private Const PreviewColour As Long = &H7CD67C
Private Const BadPreviewColour As Long = &H8080F0

Private Type TSprites
    SunkShipKeys As Collection
    DizzyShipKey As String
    FleetGridId As PlayGridId
    SqueakGridId As PlayGridId
    SqueakX As Long
    SqueakY As Long
    ' the cell of each grid that was hit last, by grid id: the shot that sank a ship, while it sinks.
    LastHitX(1 To 2) As Long
    LastHitY(1 To 2) As Long
End Type

Private this As TSprites

Private Sub Class_Initialize()
    Set this.SunkShipKeys = New Collection
End Sub

'@Description("The names of every sprite a cell can show, on a grid or in the fleet status.")
Public Function Names() As Variant
Attribute Names.VB_Description = "The names of every sprite a cell can show, on a grid or in the fleet status."
    Names = Array(DuckSprite(Horizontal, True), DuckSprite(Horizontal, False), DuckSprite(Vertical, True), DuckSprite(Vertical, False), _
        RedDuckSprite(Horizontal, True), RedDuckSprite(Horizontal, False), RedDuckSprite(Vertical, True), RedDuckSprite(Vertical, False), _
        SpriteHiddenHit, SpriteSqueak, SpriteMiss, SpriteDizzy)
End Function

'@Description("The file of a sprite, by its name: a PNG in the window app's assets/sprites folder, next to its project file (or its built program).")
Public Function Path(ByVal spriteName As String) As String
Attribute Path.VB_Description = "The file of a sprite, by its name: a PNG in the window app's assets/sprites folder, next to its project file (or its built program)."
    Path = App.Path & "\" & SpriteFolder & spriteName & ".png"
End Function

'@Description("The sprite each cell of a grid shows, by its name, in an array indexed by column and row from 1: empty for open water. A miss shows where it was played. A hit shows its duck in red where the ships are shown, or where its ship is sunk, and the devil duck facing front where its ship is hidden. The ships show a duck on each cell that is not hit, when they are shown, and a sunk ship's bow shows either way: a dizzy duck while it sinks, when the shot that sank it hit the bow, then its big duck in red. A bow hit by an earlier shot stays red. The big duck of the ship that was just hit is dizzy too, and the cell that was just hit squeaks.")
Public Function CellSprites(ByVal grid As PlayerGrid, ByVal showShips As Boolean) As Variant
Attribute CellSprites.VB_Description = "The sprite each cell of a grid shows, by its name, in an array indexed by column and row from 1: empty for open water. A miss shows where it was played. A hit shows its duck in red where the ships are shown, or where its ship is sunk, and the devil duck facing front where its ship is hidden. The ships show a duck on each cell that is not hit, when they are shown, and a sunk ship's bow shows either way: a dizzy duck while it sinks, when the shot that sank it hit the bow, then its big duck in red. A bow hit by an earlier shot stays red. The big duck of the ship that was just hit is dizzy too, and the cell that was just hit squeaks."
    Dim result() As String
    ReDim result(1 To PlayerGrid.Size, 1 To PlayerGrid.Size)
    Dim X As Long
    Dim Y As Long
    For X = 1 To PlayerGrid.Size
        For Y = 1 To PlayerGrid.Size
            If grid.State(GridCoord.Create(X, Y)) = GridState.PreviousMiss Then result(X, Y) = SpriteMiss
        Next
    Next

    Dim item As IShip
    For Each item In grid.Fleet
        Dim index As Long
        For index = 0 To item.Size - 1
            X = item.GridPosition.X + IIf(item.Orientation = Horizontal, index, 0)
            Y = item.GridPosition.Y + IIf(item.Orientation = Vertical, index, 0)
            Dim isBow As Boolean
            isBow = index = item.Size - 1
            If item.IsSunken And isBow Then
                result(X, Y) = IIf(IsSinkingDone(grid.gridId, item) Or Not IsLastHit(grid.gridId, X, Y), RedDuckSprite(item.Orientation, True), SpriteDizzy)
            ElseIf grid.State(GridCoord.Create(X, Y)) = GridState.PreviousHit Then
                result(X, Y) = IIf(showShips Or item.IsSunken, RedDuckSprite(item.Orientation, isBow), SpriteHiddenHit)
            ElseIf showShips Then
                If isBow And ShipKey(grid.gridId, item) = this.DizzyShipKey Then
                    result(X, Y) = SpriteDizzy
                Else
                    result(X, Y) = DuckSprite(item.Orientation, isBow)
                End If
            End If
        Next
    Next

    If this.SqueakGridId = grid.gridId Then result(this.SqueakX, this.SqueakY) = SpriteSqueak
    CellSprites = result
End Function

'@Description("The grid whose fleet the fleet status shows, or 0 before it is known: the first grid painted with its ships shown, which is the person's own when a person plays.")
Public Property Get FleetGridId() As PlayGridId
Attribute FleetGridId.VB_Description = "The grid whose fleet the fleet status shows, or 0 before it is known: the first grid painted with its ships shown, which is the person's own when a person plays."
    FleetGridId = this.FleetGridId
End Property

'@Description("Whether a grid painted with its ships shown, or not, is the one the fleet status shows. The first grid painted with its ships shown is, from then on.")
Public Function IsFleetGrid(ByVal gridId As PlayGridId, ByVal showShips As Boolean) As Boolean
Attribute IsFleetGrid.VB_Description = "Whether a grid painted with its ships shown, or not, is the one the fleet status shows. The first grid painted with its ships shown is, from then on."
    If this.FleetGridId = 0 And showShips Then this.FleetGridId = gridId
    IsFleetGrid = gridId = this.FleetGridId
End Function

'@Description("What the fleet status shows of a grid's ships, in the order they were placed: for each ship, an array of its cells' sprites from its stern, at 1, to its bow. Each ship lies across, and each of its cells shows its duck: yellow where it is not hit, red where it is, so a sunk ship is red all over.")
Public Function FleetSprites(ByVal grid As PlayerGrid) As Collection
Attribute FleetSprites.VB_Description = "What the fleet status shows of a grid's ships, in the order they were placed: for each ship, an array of its cells' sprites from its stern, at 1, to its bow. Each ship lies across, and each of its cells shows its duck: yellow where it is not hit, red where it is, so a sunk ship is red all over."
    Dim result As Collection
    Set result = New Collection
    Dim item As IShip
    For Each item In grid.Fleet
        Dim cells() As String
        ReDim cells(1 To item.Size)
        Dim index As Long
        For index = 1 To item.Size
            Dim X As Long
            X = item.GridPosition.X + IIf(item.Orientation = Horizontal, index - 1, 0)
            Dim Y As Long
            Y = item.GridPosition.Y + IIf(item.Orientation = Vertical, index - 1, 0)
            Dim isBow As Boolean
            isBow = index = item.Size
            If grid.State(GridCoord.Create(X, Y)) = GridState.PreviousHit Then
                cells(index) = RedDuckSprite(Horizontal, isBow)
            Else
                cells(index) = DuckSprite(Horizontal, isBow)
            End If
        Next
        result.Add cells
    Next
    Set FleetSprites = result
End Function

'@Description("The colour a cell shows in place of its sprite, where the sprites cannot be read: open water, a miss, a hit (a dizzy duck is a hit too), or a ship.")
Public Function FallbackColour(ByVal spriteName As String) As Long
Attribute FallbackColour.VB_Description = "The colour a cell shows in place of its sprite, where the sprites cannot be read: open water, a miss, a hit (a dizzy duck is a hit too), or a ship."
    If spriteName = vbNullString Then
        FallbackColour = WaterColour
    ElseIf spriteName = SpriteMiss Then
        FallbackColour = MissColour
    ElseIf IsHit(spriteName) Or spriteName = SpriteDizzy Then
        FallbackColour = HitColour
    Else
        FallbackColour = ShipColour
    End If
End Function

'@Description("The colour of a ship being placed: green where it fits, red where it does not.")
Public Function ShipPreviewColour(ByVal fits As Boolean) As Long
Attribute ShipPreviewColour.VB_Description = "The colour of a ship being placed: green where it fits, red where it does not."
    ShipPreviewColour = IIf(fits, PreviewColour, BadPreviewColour)
End Function

'@Description("The colour of open water.")
Public Property Get Water() As Long
Attribute Water.VB_Description = "The colour of open water."
    Water = WaterColour
End Property

'@Description("Whether a sprite marks a hit cell: a red duck, the devil duck, or the squeak. A window without its sprites writes an X on a hit cell.")
Public Function IsHit(ByVal spriteName As String) As Boolean
Attribute IsHit.VB_Description = "Whether a sprite marks a hit cell: a red duck, the devil duck, or the squeak. A window without its sprites writes an X on a hit cell."
    IsHit = spriteName = SpriteSqueak Or spriteName = SpriteHiddenHit Or Right$(spriteName, Len(RedSuffix)) = RedSuffix
End Function

'@Description("Whether a sprite marks a missed shot: the ripple.")
Public Function IsMiss(ByVal spriteName As String) As Boolean
Attribute IsMiss.VB_Description = "Whether a sprite marks a missed shot: the ripple."
    IsMiss = spriteName = SpriteMiss
End Function

'@Description("Whether a sprite is the squeak of a cell just hit.")
Public Function IsSqueak(ByVal spriteName As String) As Boolean
Attribute IsSqueak.VB_Description = "Whether a sprite is the squeak of a cell just hit."
    IsSqueak = spriteName = SpriteSqueak
End Function

'@Description("A cell of a grid was just hit: it squeaks, until SqueakStops.")
Public Sub Squeak(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
Attribute Squeak.VB_Description = "A cell of a grid was just hit: it squeaks, until SqueakStops."
    this.SqueakGridId = gridId
    this.SqueakX = position.X
    this.SqueakY = position.Y
    this.LastHitX(gridId) = position.X
    this.LastHitY(gridId) = position.Y
End Sub

'@Description("The cell that was just hit shows its hit again.")
Public Sub SqueakStops()
Attribute SqueakStops.VB_Description = "The cell that was just hit shows its hit again."
    this.SqueakGridId = 0
End Sub

'@Description("A ship was hit: its big duck is dizzy, until DizzyShipRecovers.")
Public Sub DizzyShip(ByVal gridId As PlayGridId, ByVal hitShip As IShip)
Attribute DizzyShip.VB_Description = "A ship was hit: its big duck is dizzy, until DizzyShipRecovers."
    this.DizzyShipKey = ShipKey(gridId, hitShip)
End Sub

'@Description("The dizzy big duck is itself again.")
Public Sub DizzyShipRecovers()
Attribute DizzyShipRecovers.VB_Description = "The dizzy big duck is itself again."
    this.DizzyShipKey = vbNullString
End Sub

'@Description("A ship's sinking is done: from now on, its bow shows its big duck in red.")
Public Sub FinishSinking(ByVal gridId As PlayGridId, ByVal sunkShip As IShip)
Attribute FinishSinking.VB_Description = "A ship's sinking is done: from now on, its bow shows its big duck in red."
    If IsSinkingDone(gridId, sunkShip) Then Exit Sub
    this.SunkShipKeys.Add ShipKey(gridId, sunkShip)
End Sub

Private Function IsSinkingDone(ByVal gridId As PlayGridId, ByVal sunkShip As IShip) As Boolean
    Dim key As String
    key = ShipKey(gridId, sunkShip)
    Dim item As Variant
    For Each item In this.SunkShipKeys
        If item = key Then
            IsSinkingDone = True
            Exit Function
        End If
    Next
End Function

Private Function IsLastHit(ByVal gridId As PlayGridId, ByVal X As Long, ByVal Y As Long) As Boolean
    IsLastHit = this.LastHitX(gridId) = X And this.LastHitY(gridId) = Y
End Function

' a ship is known by its grid, its position and its direction: no two ships of a grid share a bow.
Private Function ShipKey(ByVal gridId As PlayGridId, ByVal item As IShip) As String
    ShipKey = gridId & ":" & item.GridPosition.X & "," & item.GridPosition.Y & ":" & item.Orientation
End Function

' a ship's ducks face where it goes: right when it lies across, down when it stands; the big duck leads, at its bow, and the small ones follow.
Private Function DuckSprite(ByVal direction As ShipOrientation, ByVal isBow As Boolean) As String
    DuckSprite = "duck-" & IIf(isBow, "big", "small") & "-" & IIf(direction = Horizontal, "right", "down")
End Function

' a hit duck: the same duck, red, with horns.
Private Function RedDuckSprite(ByVal direction As ShipOrientation, ByVal isBow As Boolean) As String
    RedDuckSprite = DuckSprite(direction, isBow) & RedSuffix
End Function
