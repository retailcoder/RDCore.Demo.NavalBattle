VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An implementation of IGridViewCommands that plays the game in the console."
'@Folder("NavalBattle.View")
'@ModuleDescription("An implementation of IGridViewCommands that plays the game in the console.")
Option Explicit
Implements IGridViewCommands

Private Type TView
    ViewEvents As IGridViewEvents
    Grids(PlayerGrid1 To PlayerGrid2) As PlayerGrid
    LastShot As String
    LastHitShip As String
End Type

Private this As TView

Private Property Get IGridViewCommands_Events() As IGridViewEvents
    Set IGridViewCommands_Events = this.ViewEvents
End Property

Private Property Set IGridViewCommands_Events(ByVal value As IGridViewEvents)
    Set this.ViewEvents = value
End Property

Private Sub IGridViewCommands_OnNewGame()
    Debug.Print "New game."
End Sub

Private Sub IGridViewCommands_OnBeginShipPosition(ByVal CurrentShip As IShip, ByVal player As IPlayer)
End Sub

Private Sub IGridViewCommands_OnPreviewShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
End Sub

Private Sub IGridViewCommands_OnConfirmShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
    Debug.Print "Grid " & player.PlayGrid.gridId & " places its " & newShip.Name & "."
End Sub

Private Sub IGridViewCommands_OnInvalidShipPosition()
    Debug.Print "A ship cannot go there."
End Sub

Private Sub IGridViewCommands_OnBeginAttack(ByVal currentPlayerGridId As PlayGridId)
End Sub

Private Sub IGridViewCommands_OnSelectPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    this.LastShot = CellName(position.X, position.Y)
End Sub

Private Sub IGridViewCommands_OnKnownPositionAttack()
    Debug.Print "That position was already played."
End Sub

Private Sub IGridViewCommands_OnUpdateFleetStatus(ByVal player As IPlayer, ByVal hitShip As IShip, Optional ByVal showAIStatus As Boolean = False)
    this.LastHitShip = hitShip.Name
End Sub

Private Sub IGridViewCommands_OnMiss(ByVal gridId As PlayGridId)
    Debug.Print "Shot at grid " & gridId & ", " & this.LastShot & ": miss."
End Sub

Private Sub IGridViewCommands_OnHit(ByVal gridId As PlayGridId)
    Debug.Print "Shot at grid " & gridId & ", " & this.LastShot & ": hit, the " & this.LastHitShip & "."
End Sub

Private Sub IGridViewCommands_OnSink(ByVal gridId As PlayGridId)
    Debug.Print "Shot at grid " & gridId & ", " & this.LastShot & ": the " & this.LastHitShip & " is sunk."
End Sub

Private Sub IGridViewCommands_OnRefreshGrid(ByVal grid As PlayerGrid)
    Set this.Grids(grid.gridId) = grid
End Sub

Private Sub IGridViewCommands_OnLockGrid(ByVal gridId As PlayGridId)
End Sub

Private Sub IGridViewCommands_OnBeginWaitForComputerPlayer()
End Sub

Private Sub IGridViewCommands_OnEndWaitForComputerPlayer()
End Sub

Private Sub IGridViewCommands_OnGameOver(ByVal winningGrid As PlayGridId)
    Dim gridId As PlayGridId
    For gridId = PlayerGrid1 To PlayerGrid2
        PrintGrid gridId
    Next
    Debug.Print "Game over: grid " & winningGrid & " wins."
End Sub

'@Description("Prints a grid as text: # is a ship, X a hit, o a miss.")
Private Sub PrintGrid(ByVal gridId As PlayGridId)
Attribute PrintGrid.VB_Description = "Prints a grid as text: # is a ship, X a hit, o a miss."
    Dim grid As PlayerGrid
    Set grid = this.Grids(gridId)
    If grid Is Nothing Then Exit Sub

    Debug.Print
    Debug.Print "Grid " & gridId
    Dim header As String
    header = "    "
    Dim currentX As Long
    For currentX = 1 To grid.Size
        header = header & Chr$(64 + currentX) & " "
    Next
    Debug.Print header

    Dim currentY As Long
    For currentY = 1 To grid.Size
        Dim rowText As String
        rowText = Right$(" " & currentY, 2) & "  "
        For currentX = 1 To grid.Size
            rowText = rowText & CellText(grid.State(GridCoord.Create(currentX, currentY))) & " "
        Next
        Debug.Print RTrim$(rowText)
    Next
End Sub

Private Function CellText(ByVal value As GridState) As String
    Select Case value
        Case GridState.ShipPosition
            CellText = "#"
        Case GridState.PreviousHit
            CellText = "X"
        Case GridState.PreviousMiss
            CellText = "o"
        Case Else
            CellText = "."
    End Select
End Function

Private Function CellName(ByVal X As Long, ByVal Y As Long) As String
    CellName = Chr$(64 + X) & Y
End Function
