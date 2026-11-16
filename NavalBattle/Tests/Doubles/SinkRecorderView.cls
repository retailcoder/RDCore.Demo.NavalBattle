VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SinkRecorderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A view for the controller tests: it shows nothing, and records each sinking it is told of, which grids are locked, and the winner."
'@Folder("NavalBattle.Tests.Doubles")
'@ModuleDescription("A view for the controller tests: it shows nothing, and records each sinking it is told of, which grids are locked, and the winner.")
Option Explicit
Implements IGridViewCommands

Private Type TView
    ViewEvents As IGridViewEvents
    SunkGridIds As Collection
    SunkPositions As Collection
    SunkShips As Collection
    WinningGrid As PlayGridId
    LastPosition As IGridCoord
    LastShip As IShip
    Locked(PlayerGrid1 To PlayerGrid2) As Boolean
    Unlocks As Long
End Type

Private this As TView

Private Sub Class_Initialize()
    Set this.SunkGridIds = New Collection
    Set this.SunkPositions = New Collection
    Set this.SunkShips = New Collection
End Sub

'@Description("Gets whether the given grid is locked now.")
Public Property Get IsLocked(ByVal gridId As PlayGridId) As Boolean
Attribute IsLocked.VB_Description = "Gets whether the given grid is locked now."
    IsLocked = this.Locked(gridId)
End Property

'@Description("Gets how many times the view was told to unlock a grid.")
Public Property Get Unlocks() As Long
Attribute Unlocks.VB_Description = "Gets how many times the view was told to unlock a grid."
    Unlocks = this.Unlocks
End Property

'@Description("Gets the view's events, to play the game as a person would: the controller's adapter sets them.")
Public Property Get ViewEvents() As IGridViewEvents
Attribute ViewEvents.VB_Description = "Gets the view's events, to play the game as a person would: the controller's adapter sets them."
    Set ViewEvents = this.ViewEvents
End Property

'@Description("Gets the grid of each sinking the view was told of, in order.")
Public Property Get SunkGridIds() As Collection
Attribute SunkGridIds.VB_Description = "Gets the grid of each sinking the view was told of, in order."
    Set SunkGridIds = this.SunkGridIds
End Property

'@Description("Gets the position of the shot of each sinking the view was told of, in order.")
Public Property Get SunkPositions() As Collection
Attribute SunkPositions.VB_Description = "Gets the position of the shot of each sinking the view was told of, in order."
    Set SunkPositions = this.SunkPositions
End Property

'@Description("Gets the ship of each sinking the view was told of, in order.")
Public Property Get SunkShips() As Collection
Attribute SunkShips.VB_Description = "Gets the ship of each sinking the view was told of, in order."
    Set SunkShips = this.SunkShips
End Property

'@Description("Gets the grid that won, or 0 while the game is not over.")
Public Property Get WinningGrid() As PlayGridId
Attribute WinningGrid.VB_Description = "Gets the grid that won, or 0 while the game is not over."
    WinningGrid = this.WinningGrid
End Property

'@Description("Gets the position of the shot that ended the game, or Nothing while the game is not over.")
Public Property Get LastPosition() As IGridCoord
Attribute LastPosition.VB_Description = "Gets the position of the shot that ended the game, or Nothing while the game is not over."
    Set LastPosition = this.LastPosition
End Property

'@Description("Gets the ship whose sinking ended the game, or Nothing while the game is not over.")
Public Property Get LastShip() As IShip
Attribute LastShip.VB_Description = "Gets the ship whose sinking ended the game, or Nothing while the game is not over."
    Set LastShip = this.LastShip
End Property

Private Property Get IGridViewCommands_Events() As IGridViewEvents
    Set IGridViewCommands_Events = this.ViewEvents
End Property

Private Property Set IGridViewCommands_Events(ByVal value As IGridViewEvents)
    Set this.ViewEvents = value
End Property

Private Sub IGridViewCommands_OnSink(ByVal gridId As PlayGridId, ByVal position As IGridCoord, ByVal sunkShip As IShip)
    this.SunkGridIds.Add gridId
    this.SunkPositions.Add position
    this.SunkShips.Add sunkShip
End Sub

Private Sub IGridViewCommands_OnGameOver(ByVal winningGrid As PlayGridId, ByVal position As IGridCoord, ByVal sunkShip As IShip)
    this.WinningGrid = winningGrid
    Set this.LastPosition = position
    Set this.LastShip = sunkShip
End Sub

Private Sub IGridViewCommands_OnMiss(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
End Sub

Private Sub IGridViewCommands_OnHit(ByVal gridId As PlayGridId, ByVal position As IGridCoord, ByVal hitShip As IShip)
End Sub

Private Sub IGridViewCommands_OnLockGrid(ByVal gridId As PlayGridId)
    this.Locked(gridId) = True
End Sub

Private Sub IGridViewCommands_OnUnlockGrid(ByVal gridId As PlayGridId)
    this.Locked(gridId) = False
    this.Unlocks = this.Unlocks + 1
End Sub

Private Sub IGridViewCommands_OnNewGame()
End Sub

Private Sub IGridViewCommands_OnBeginShipPosition(ByVal CurrentShip As IShip, ByVal player As IPlayer)
End Sub

Private Sub IGridViewCommands_OnConfirmShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
End Sub

Private Sub IGridViewCommands_OnPreviewShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
End Sub

Private Sub IGridViewCommands_OnInvalidShipPosition()
End Sub

Private Sub IGridViewCommands_OnBeginAttack(ByVal currentPlayerGridId As PlayGridId)
End Sub

Private Sub IGridViewCommands_OnKnownPositionAttack()
End Sub

Private Sub IGridViewCommands_OnRefreshGrid(ByVal grid As PlayerGrid)
End Sub

Private Sub IGridViewCommands_OnBeginWaitForComputerPlayer()
End Sub

Private Sub IGridViewCommands_OnEndWaitForComputerPlayer()
End Sub
