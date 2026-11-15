VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SinkRecorderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A view for the controller tests: it shows nothing, and records each sinking it is told of, and the winner."
'@Folder("NavalBattle.Tests.Doubles")
'@ModuleDescription("A view for the controller tests: it shows nothing, and records each sinking it is told of, and the winner.")
Option Explicit
Implements IGridViewCommands

Private Type TView
    ViewEvents As IGridViewEvents
    SunkGridIds As Collection
    SunkShips As Collection
    WinningGrid As PlayGridId
End Type

Private this As TView

Private Sub Class_Initialize()
    Set this.SunkGridIds = New Collection
    Set this.SunkShips = New Collection
End Sub

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

Private Property Get IGridViewCommands_Events() As IGridViewEvents
    Set IGridViewCommands_Events = this.ViewEvents
End Property

Private Property Set IGridViewCommands_Events(ByVal value As IGridViewEvents)
    Set this.ViewEvents = value
End Property

Private Sub IGridViewCommands_OnSink(ByVal gridId As PlayGridId, ByVal sunkShip As IShip)
    this.SunkGridIds.Add gridId
    this.SunkShips.Add sunkShip
End Sub

Private Sub IGridViewCommands_OnGameOver(ByVal winningGrid As PlayGridId)
    this.WinningGrid = winningGrid
End Sub

Private Sub IGridViewCommands_OnMiss(ByVal gridId As PlayGridId)
End Sub

Private Sub IGridViewCommands_OnHit(ByVal gridId As PlayGridId)
End Sub

Private Sub IGridViewCommands_OnUpdateFleetStatus(ByVal player As IPlayer, ByVal hitShip As IShip, Optional ByVal showAIStatus As Boolean = False)
End Sub

Private Sub IGridViewCommands_OnSelectPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
End Sub

Private Sub IGridViewCommands_OnLockGrid(ByVal gridId As PlayGridId)
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
