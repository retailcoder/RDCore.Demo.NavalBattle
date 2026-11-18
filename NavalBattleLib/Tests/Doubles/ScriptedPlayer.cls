VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptedPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A computer player for the controller tests, with no randomness and no delay: it places each ship horizontally at column A, one ship every other row from row 1, and fires at every cell in turn, row by row from A1."
'@Folder("NavalBattle.Tests.Doubles")
'@ModuleDescription("A computer player for the controller tests, with no randomness and no delay: it places each ship horizontally at column A, one ship every other row from row 1, and fires at every cell in turn, row by row from A1.")
Option Explicit
Implements IPlayer

Private Type TPlayer
    Grid As PlayerGrid
    Shots As Long
End Type

Private this As TPlayer

'@Description("Gets/sets the player's own grid.")
Public Property Get Grid() As PlayerGrid
Attribute Grid.VB_Description = "Gets/sets the player's own grid."
    Set Grid = this.Grid
End Property

Public Property Set Grid(ByVal value As PlayerGrid)
    Set this.Grid = value
End Property

Private Property Get IPlayer_PlayerType() As PlayerType
    IPlayer_PlayerType = ComputerControlled
End Property

Private Property Get IPlayer_PlayGrid() As PlayerGrid
    Set IPlayer_PlayGrid = this.Grid
End Property

Private Sub IPlayer_PlaceShip(ByVal CurrentShip As IShip)
    this.Grid.AddShip Ship.Create(CurrentShip.ShipKind, Horizontal, GridCoord.Create(1, 2 * this.Grid.shipCount + 1))
End Sub

Private Function IPlayer_Play(ByVal enemyGrid As PlayerGrid) As IGridCoord
    Set IPlayer_Play = GridCoord.Create(this.Shots Mod enemyGrid.Size + 1, this.Shots \ enemyGrid.Size + 1)
    this.Shots = this.Shots + 1
End Function
