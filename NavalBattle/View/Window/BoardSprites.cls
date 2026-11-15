VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoardSprites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "What each cell of a game window's grids shows: a hit, a miss, a duck, a dizzy duck or an anchor, by its sprite's name, and the colour it shows instead where the sprites cannot be read. It knows which hit ship is dizzy, which sunk ships are done sinking, and which grid's fleet the fleet status shows."
'@Folder("NavalBattle.View.Window")
'@ModuleDescription("What each cell of a game window's grids shows: a hit, a miss, a duck, a dizzy duck or an anchor, by its sprite's name, and the colour it shows instead where the sprites cannot be read. It knows which hit ship is dizzy, which sunk ships are done sinking, and which grid's fleet the fleet status shows.")
Option Explicit

' the sprites, GIF files in assets/sprites with one transparent colour; the ducks are named by DuckSprite.
Private Const SpriteFolder As String = "assets\sprites\"
Private Const SpriteHit As String = "hit-squeak"
Private Const SpriteMiss As String = "miss-ripple"
Private Const SpriteDizzy As String = "sunk-dizzy"
Private Const SpriteAnchor As String = "anchor-cracked"

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
End Type

Private this As TSprites

Private Sub Class_Initialize()
    Set this.SunkShipKeys = New Collection
End Sub

'@Description("The names of every sprite a cell can show.")
Public Function Names() As Variant
Attribute Names.VB_Description = "The names of every sprite a cell can show."
    Names = Array(DuckSprite(Horizontal, True), DuckSprite(Horizontal, False), DuckSprite(Vertical, True), DuckSprite(Vertical, False), _
        SpriteHit, SpriteMiss, SpriteDizzy, SpriteAnchor)
End Function

'@Description("The file of a sprite, by its name: a GIF in the workspace's assets/sprites folder.")
Public Function Path(ByVal spriteName As String) As String
Attribute Path.VB_Description = "The file of a sprite, by its name: a GIF in the workspace's assets/sprites folder."
    Path = CurDir$ & "\" & SpriteFolder & spriteName & ".gif"
End Function

'@Description("The sprite each cell of a grid shows, by its name, in an array indexed by column and row from 1: empty for open water. A hit or a miss shows where it was played; the ships show a duck on each cell that is not hit, when they are shown, and a sunk ship shows either way, by its bow: a dizzy duck while it sinks, then an anchor. The big duck of the ship that was just hit is dizzy too.")
Public Function CellSprites(ByVal grid As PlayerGrid, ByVal showShips As Boolean) As Variant
Attribute CellSprites.VB_Description = "The sprite each cell of a grid shows, by its name, in an array indexed by column and row from 1: empty for open water. A hit or a miss shows where it was played; the ships show a duck on each cell that is not hit, when they are shown, and a sunk ship shows either way, by its bow: a dizzy duck while it sinks, then an anchor. The big duck of the ship that was just hit is dizzy too."
    Dim result() As String
    ReDim result(1 To PlayerGrid.Size, 1 To PlayerGrid.Size)
    Dim X As Long
    Dim Y As Long
    For X = 1 To PlayerGrid.Size
        For Y = 1 To PlayerGrid.Size
            Select Case grid.State(GridCoord.Create(X, Y))
                Case GridState.PreviousHit
                    result(X, Y) = SpriteHit
                Case GridState.PreviousMiss
                    result(X, Y) = SpriteMiss
            End Select
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
            If item.IsSunken Then
                If isBow Then result(X, Y) = IIf(IsSinkingDone(grid.gridId, item), SpriteAnchor, SpriteDizzy)
            ElseIf showShips And result(X, Y) = vbNullString Then
                If isBow And ShipKey(grid.gridId, item) = this.DizzyShipKey Then
                    result(X, Y) = SpriteDizzy
                Else
                    result(X, Y) = DuckSprite(item.Orientation, isBow)
                End If
            End If
        Next
    Next
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

'@Description("What the fleet status shows of a grid's ships, in the order they were placed: for each ship, an array of its cells' sprites from its stern, at 1, to its bow. Each ship lies across, and shows what its cells show on the grid: a duck where it is not hit, the hit where it is, and its bow dizzy, then the anchor.")
Public Function FleetSprites(ByVal grid As PlayerGrid) As Collection
Attribute FleetSprites.VB_Description = "What the fleet status shows of a grid's ships, in the order they were placed: for each ship, an array of its cells' sprites from its stern, at 1, to its bow. Each ship lies across, and shows what its cells show on the grid: a duck where it is not hit, the hit where it is, and its bow dizzy, then the anchor."
    Dim onGrid As Variant
    onGrid = CellSprites(grid, True)
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
            cells(index) = AcrossSprite(onGrid(X, Y))
        Next
        result.Add cells
    Next
    Set FleetSprites = result
End Function

' a standing ship's duck lies across in the fleet status; every other sprite is the same either way.
Private Function AcrossSprite(ByVal spriteName As String) As String
    If spriteName = DuckSprite(Vertical, True) Then
        AcrossSprite = DuckSprite(Horizontal, True)
    ElseIf spriteName = DuckSprite(Vertical, False) Then
        AcrossSprite = DuckSprite(Horizontal, False)
    Else
        AcrossSprite = spriteName
    End If
End Function

'@Description("The colour a cell shows in place of its sprite, where the sprites cannot be read: open water, a miss, a hit (a dizzy duck and an anchor are hits too), or a ship.")
Public Function FallbackColour(ByVal spriteName As String) As Long
Attribute FallbackColour.VB_Description = "The colour a cell shows in place of its sprite, where the sprites cannot be read: open water, a miss, a hit (a dizzy duck and an anchor are hits too), or a ship."
    Select Case spriteName
        Case vbNullString
            FallbackColour = WaterColour
        Case SpriteMiss
            FallbackColour = MissColour
        Case SpriteHit, SpriteDizzy, SpriteAnchor
            FallbackColour = HitColour
        Case Else
            FallbackColour = ShipColour
    End Select
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

'@Description("Whether a sprite is the hit's: a window without its sprites writes an X on a hit cell.")
Public Function IsHit(ByVal spriteName As String) As Boolean
Attribute IsHit.VB_Description = "Whether a sprite is the hit's: a window without its sprites writes an X on a hit cell."
    IsHit = spriteName = SpriteHit
End Function

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

'@Description("A ship's sinking is done: from now on, its bow shows the anchor.")
Public Sub FinishSinking(ByVal gridId As PlayGridId, ByVal sunkShip As IShip)
Attribute FinishSinking.VB_Description = "A ship's sinking is done: from now on, its bow shows the anchor."
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

' a ship is known by its grid, its position and its direction: no two ships of a grid share a bow.
Private Function ShipKey(ByVal gridId As PlayGridId, ByVal item As IShip) As String
    ShipKey = gridId & ":" & item.GridPosition.X & "," & item.GridPosition.Y & ":" & item.Orientation
End Function

' a ship's ducks face where it goes: right when it lies across, down when it stands; the big duck leads, at its bow, and the small ones follow.
Private Function DuckSprite(ByVal direction As ShipOrientation, ByVal isBow As Boolean) As String
    DuckSprite = "duck-" & IIf(isBow, "big", "small") & "-" & IIf(direction = Horizontal, "right", "down")
End Function
