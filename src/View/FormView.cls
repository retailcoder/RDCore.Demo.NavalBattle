VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An implementation of IGridViewCommands that plays the game in the game form: the commands paint its grids, fleet panels and status line, and what the person does in it is sent to the adapter."
'@Folder("NavalBattle.View")
'@ModuleDescription("An implementation of IGridViewCommands that plays the game in the game form: the commands paint its grids, fleet panels and status line, and what the person does in it is sent to the adapter.")
Option Explicit
Implements IGridViewCommands

Private Type TView
    ViewEvents As IGridViewEvents
    Closed As Boolean
    HumanGridId As PlayGridId
    LastShot As String
    LastHitShip As String
End Type

Private this As TView
Private WithEvents window As GameForm
Attribute window.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set window = New GameForm
End Sub

Private Sub window_Closed()
    this.Closed = True
End Sub

' once the person closes the window, a game that is still being played goes on without it, and is no longer shown.
Private Function IsOpen() As Boolean
    IsOpen = Not this.Closed
End Function

'@Description("Shows the game form, with the given players picked (Human, Random, Fair-play or Merciless; the first letter is enough), and returns when it is closed. A player that is not given keeps the form's own pick.")
Public Sub Play(Optional ByVal firstPlayer As String, Optional ByVal secondPlayer As String)
Attribute Play.VB_Description = "Shows the game form, with the given players picked (Human, Random, Fair-play or Merciless; the first letter is enough), and returns when it is closed. A player that is not given keeps the form's own pick."
    If IsOpen Then window.PickPlayer PlayerGrid1, firstPlayer
    If IsOpen Then window.PickPlayer PlayerGrid2, secondPlayer
    If IsOpen Then window.Show
End Sub

Private Property Get IGridViewCommands_Events() As IGridViewEvents
    Set IGridViewCommands_Events = this.ViewEvents
End Property

Private Property Set IGridViewCommands_Events(ByVal value As IGridViewEvents)
    Set this.ViewEvents = value
    If IsOpen Then Set window.ViewEvents = value
End Property

Private Sub IGridViewCommands_OnNewGame()
    Debug.Print "New game."
End Sub

Private Sub IGridViewCommands_OnBeginShipPosition(ByVal CurrentShip As IShip, ByVal player As IPlayer)
    this.HumanGridId = player.PlayGrid.gridId
    If IsOpen Then window.BeginShipPlacement this.HumanGridId
    If IsOpen Then window.PaintGrid player.PlayGrid, True
    If IsOpen Then window.PaintShipPreview player.PlayGrid, CurrentShip
    If IsOpen Then window.SetStatus "Grid " & this.HumanGridId & ": place your " & CurrentShip.Name & " (" & CurrentShip.Size & " cells). " _
        & "Point at your grid, right-click to turn it, click to place it."
End Sub

Private Sub IGridViewCommands_OnPreviewShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
    If IsOpen Then window.PaintGrid player.PlayGrid, True
    If IsOpen Then window.PaintShipPreview player.PlayGrid, newShip
End Sub

Private Sub IGridViewCommands_OnConfirmShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
    If IsOpen Then window.PaintGrid player.PlayGrid, True
    If IsOpen Then window.SetFleet player.PlayGrid.gridId, FleetText(player.PlayGrid)
    Debug.Print "Grid " & player.PlayGrid.gridId & " places its " & newShip.Name & "."
End Sub

Private Sub IGridViewCommands_OnInvalidShipPosition()
    If IsOpen Then window.SetStatus "A ship cannot go there: it must fit in the grid, and touch no other ship."
End Sub

Private Sub IGridViewCommands_OnBeginAttack(ByVal currentPlayerGridId As PlayGridId)
    If IsOpen Then window.BeginWaitForReady
    If IsOpen Then window.SetStatus "Your fleet is in place. Click Ready, then fire at grid " & OtherGrid(this.HumanGridId) & "."
End Sub

Private Sub IGridViewCommands_OnSelectPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    this.LastShot = Chr$(64 + position.X) & position.Y
End Sub

Private Sub IGridViewCommands_OnKnownPositionAttack()
    If IsOpen Then window.SetStatus "That cell was already played: pick another."
End Sub

Private Sub IGridViewCommands_OnUpdateFleetStatus(ByVal player As IPlayer, ByVal hitShip As IShip, Optional ByVal showAIStatus As Boolean = False)
    this.LastHitShip = hitShip.Name
    If IsOpen Then window.SetFleet player.PlayGrid.gridId, FleetText(player.PlayGrid)
End Sub

Private Sub IGridViewCommands_OnMiss(ByVal gridId As PlayGridId)
    Report gridId, "miss."
End Sub

Private Sub IGridViewCommands_OnHit(ByVal gridId As PlayGridId)
    Report gridId, "hit, the " & this.LastHitShip & "."
End Sub

Private Sub IGridViewCommands_OnSink(ByVal gridId As PlayGridId)
    Report gridId, "the " & this.LastHitShip & " is sunk."
End Sub

Private Sub IGridViewCommands_OnRefreshGrid(ByVal grid As PlayerGrid)
    If IsOpen Then window.PaintGrid grid, ShowsShips(grid.gridId)
    If IsOpen Then window.SetFleet grid.gridId, FleetText(grid)
End Sub

Private Sub IGridViewCommands_OnLockGrid(ByVal gridId As PlayGridId)
End Sub

Private Sub IGridViewCommands_OnBeginWaitForComputerPlayer()
    If IsOpen Then window.BeginComputerTurn
End Sub

Private Sub IGridViewCommands_OnEndWaitForComputerPlayer()
    If IsOpen Then window.EndComputerTurn
End Sub

Private Sub IGridViewCommands_OnGameOver(ByVal winningGrid As PlayGridId)
    Dim text As String
    text = "Game over: grid " & winningGrid & " wins."
    If this.HumanGridId <> 0 Then text = text & IIf(winningGrid = this.HumanGridId, " Well played!", " Better luck next time.")
    If IsOpen Then window.EndGame
    If IsOpen Then window.SetStatus text & " Close the window to leave."
    Debug.Print text
End Sub

' what a shot did, in the status line and in the output: the grid that fired is the other one.
Private Sub Report(ByVal gridId As PlayGridId, ByVal result As String)
    Dim text As String
    text = "Grid " & OtherGrid(gridId) & " fires at " & this.LastShot & ": " & result
    If IsOpen Then window.SetStatus text
    Debug.Print text
End Sub

' a person sees their own ships, and not the computer's; when the computer plays both grids, both are shown.
Private Function ShowsShips(ByVal gridId As PlayGridId) As Boolean
    ShowsShips = this.HumanGridId = 0 Or gridId = this.HumanGridId
End Function

Private Function OtherGrid(ByVal gridId As PlayGridId) As PlayGridId
    OtherGrid = IIf(gridId = PlayerGrid1, PlayerGrid2, PlayerGrid1)
End Function

' a line for each ship of the fleet: its name, and whether it is afloat or sunk.
Private Function FleetText(ByVal grid As PlayerGrid) As String
    Dim result As String
    Dim item As IShip
    For Each item In grid.Fleet
        If Len(result) > 0 Then result = result & vbCrLf
        result = result & item.Name & ": " & IIf(item.IsSunken, "sunk", "afloat")
    Next
    FleetText = result
End Function
