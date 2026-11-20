VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoardInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "What the person does in a game window, sent to the view's events as the game's phase allows it: picking the players, placing ships, the Ready button, and firing. A locked grid, and the computer's turn, take no clicks. The window shows what this says: its status text, the Ready button, and the players once they are confirmed."
'@Folder("NavalBattle.View.Window")
'@ModuleDescription("What the person does in a game window, sent to the view's events as the game's phase allows it: picking the players, placing ships, the Ready button, and firing. A locked grid, and the computer's turn, take no clicks. The window shows what this says: its status text, the Ready button, and the players once they are confirmed.")
Option Explicit

Private Enum BoardPhase
    PhaseSetup
    PhaseShipPlacement
    PhaseWaitForReady
    PhaseAttack
    PhaseOver
End Enum

Private Type TInput
    ViewEvents As IGridViewEvents
    Phase As BoardPhase
    HumanGridId As PlayGridId
    ComputerTurn As Boolean
    ComputerOnly As Boolean
    PreviewX As Long
    PreviewY As Long
    Locked(PlayerGrid1 To PlayerGrid2) As Boolean
    OtherButtonPressed As Boolean
End Type

Private this As TInput

' the window's status text says this.
Public Event StatusChanged(ByVal text As String)
' the window shows or hides its Ready button.
Public Event ReadyShown(ByVal visible As Boolean)
' the players are confirmed: the window shows them, in place of its player pickers.
Public Event PlayersConfirmed()

'@Description("Where what the person does is sent: the adapter's events. Nothing once the window is closed.")
Public Property Set Events(ByVal value As IGridViewEvents)
Attribute Events.VB_Description = "Where what the person does is sent: the adapter's events. Nothing once the window is closed."
    Set this.ViewEvents = value
End Property

Public Property Get Events() As IGridViewEvents
    Set Events = this.ViewEvents
End Property

'@Description("Whether the computer plays both grids: the window then draws after each of its shots, so that the game can be watched.")
Public Property Get ComputerOnly() As Boolean
Attribute ComputerOnly.VB_Description = "Whether the computer plays both grids: the window then draws after each of its shots, so that the game can be watched."
    ComputerOnly = this.ComputerOnly
End Property

'@Description("A grid's player was picked, before the game starts: the status says who plays it.")
Public Sub PlayerPicked(ByVal gridId As PlayGridId, ByVal playerName As String)
Attribute PlayerPicked.VB_Description = "A grid's player was picked, before the game starts: the status says who plays it."
    If this.Phase = PhaseSetup Then RaiseEvent StatusChanged("Grid " & gridId & " is played by: " & playerName & ".")
End Sub

'@Description("The person clicked Ready. Before the game, it confirms the players picked, and starts the game; once the person's ships are placed, it starts the battle.")
Public Sub ReadyClicked(ByVal firstPlayer As String, ByVal secondPlayer As String)
Attribute ReadyClicked.VB_Description = "The person clicked Ready. Before the game, it confirms the players picked, and starts the game; once the person's ships are placed, it starts the battle."
    If this.ViewEvents Is Nothing Then Exit Sub
    Select Case this.Phase
        Case PhaseSetup
            StartGame firstPlayer, secondPlayer
        Case PhaseWaitForReady
            this.Phase = PhaseAttack
            RaiseEvent ReadyShown(False)
            RaiseEvent StatusChanged("Fire at grid " & OtherGrid(this.HumanGridId) & ": click a cell.")
            this.ViewEvents.HumanPlayerReady
    End Select
End Sub

' the players picked are made, and the game starts: a person places their ships next, and when the computer plays both grids, the whole game is
' played now. So the window shows the players, and hides the Ready button, before the players are made.
Private Sub StartGame(ByVal firstPlayer As String, ByVal secondPlayer As String)
    Dim firstDifficulty As AIDifficulty
    Dim secondDifficulty As AIDifficulty
    If Not TryParsePlayer(firstPlayer, firstDifficulty) Or Not TryParsePlayer(secondPlayer, secondDifficulty) Then Exit Sub

    If Not CanPlayTogether(firstDifficulty, secondDifficulty) Then
        RaiseEvent StatusChanged("Only one grid can be played by a person: pick an AI for the other one.")
        Exit Sub
    End If

    RaiseEvent PlayersConfirmed
    RaiseEvent ReadyShown(False)
    this.ComputerOnly = PlayerTypeOf(firstDifficulty) = ComputerControlled And PlayerTypeOf(secondDifficulty) = ComputerControlled
    If this.ComputerOnly Then
        this.Phase = PhaseAttack
        RaiseEvent StatusChanged("The computer plays both grids.")
    End If

    this.ViewEvents.CreatePlayer PlayerGrid1, PlayerTypeOf(firstDifficulty), firstDifficulty
    this.ViewEvents.CreatePlayer PlayerGrid2, PlayerTypeOf(secondDifficulty), secondDifficulty
End Sub

'@Description("The person places their ships on the given grid: pointing at a cell moves the ship there, a right-click turns it, and a click places it.")
Public Sub BeginShipPlacement(ByVal gridId As PlayGridId)
Attribute BeginShipPlacement.VB_Description = "The person places their ships on the given grid: pointing at a cell moves the ship there, a right-click turns it, and a click places it."
    this.Phase = PhaseShipPlacement
    this.HumanGridId = gridId
    this.PreviewX = 0
    this.PreviewY = 0
End Sub

'@Description("The person's ships are placed: the Ready button shows, and its next click starts the battle.")
Public Sub BeginWaitForReady()
Attribute BeginWaitForReady.VB_Description = "The person's ships are placed: the Ready button shows, and its next click starts the battle."
    this.Phase = PhaseWaitForReady
    RaiseEvent ReadyShown(True)
End Sub

'@Description("The computer plays: the person's clicks are ignored until it is done.")
Public Sub BeginComputerTurn()
Attribute BeginComputerTurn.VB_Description = "The computer plays: the person's clicks are ignored until it is done."
    this.ComputerTurn = True
End Sub

'@Description("The computer has played.")
Public Sub EndComputerTurn()
Attribute EndComputerTurn.VB_Description = "The computer has played."
    this.ComputerTurn = False
End Sub

'@Description("The game is over: the grids take no more clicks, and the Ready button is gone.")
Public Sub EndGame()
Attribute EndGame.VB_Description = "The game is over: the grids take no more clicks, and the Ready button is gone."
    this.Phase = PhaseOver
    this.ComputerTurn = False
    RaiseEvent ReadyShown(False)
End Sub

'@Description("Locks or unlocks a grid: a locked grid's cells do nothing when they are pointed at or clicked.")
Public Sub LockGrid(ByVal gridId As PlayGridId, ByVal locked As Boolean)
Attribute LockGrid.VB_Description = "Locks or unlocks a grid: a locked grid's cells do nothing when they are pointed at or clicked."
    this.Locked(gridId) = locked
End Sub

'@Description("Whether a grid's cells take the person's clicks now: while they place their ships on it, or fire at it, on their turn.")
Public Function TakesClicks(ByVal gridId As PlayGridId) As Boolean
Attribute TakesClicks.VB_Description = "Whether a grid's cells take the person's clicks now: while they place their ships on it, or fire at it, on their turn."
    If this.Locked(gridId) Or this.ComputerTurn Or this.ViewEvents Is Nothing Then Exit Function
    TakesClicks = this.Phase = PhaseShipPlacement Or this.Phase = PhaseAttack
End Function

'@Description("The person pressed the given mouse button on a cell. A window raises Click for a right click too, after its MouseUp: the Click that follows a button other than the left one is ignored.")
Public Sub ButtonPressed(ByVal button As Integer)
Attribute ButtonPressed.VB_Description = "The person pressed the given mouse button on a cell. A window raises Click for a right click too, after its MouseUp: the Click that follows a button other than the left one is ignored."
    this.OtherButtonPressed = button <> vbLeftButton
End Sub

'@Description("The person clicked a cell: while they place their ships, the ship goes there; in the battle, it is their shot. Ignored when the button pressed was not the left one.")
Public Sub Click(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
Attribute Click.VB_Description = "The person clicked a cell: while they place their ships, the ship goes there; in the battle, it is their shot. Ignored when the button pressed was not the left one."
    If this.OtherButtonPressed Then Exit Sub
    If this.Locked(gridId) Or this.ComputerTurn Or this.ViewEvents Is Nothing Then Exit Sub
    Select Case this.Phase
        Case PhaseShipPlacement
            this.ViewEvents.PreviewShipPosition gridId, position
            this.ViewEvents.ConfirmShipPosition gridId, position
        Case PhaseAttack
            this.ViewEvents.AttackPosition gridId, position
    End Select
End Sub

'@Description("The person right-clicked a cell: while they place their ships, the ship turns there.")
Public Sub RightClick(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
Attribute RightClick.VB_Description = "The person right-clicked a cell: while they place their ships, the ship turns there."
    If this.Locked(gridId) Or this.Phase <> PhaseShipPlacement Or this.ViewEvents Is Nothing Then Exit Sub
    this.ViewEvents.PreviewRotateShip gridId, position
End Sub

'@Description("The person points at a cell: while they place their ships, the ship moves there, once for each cell pointed at.")
Public Sub Hover(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
Attribute Hover.VB_Description = "The person points at a cell: while they place their ships, the ship moves there, once for each cell pointed at."
    If this.Locked(gridId) Or this.Phase <> PhaseShipPlacement Or this.ViewEvents Is Nothing Then Exit Sub
    If position.X = this.PreviewX And position.Y = this.PreviewY Then Exit Sub

    this.PreviewX = position.X
    this.PreviewY = position.Y
    this.ViewEvents.PreviewShipPosition gridId, position
End Sub

Private Function OtherGrid(ByVal gridId As PlayGridId) As PlayGridId
    OtherGrid = IIf(gridId = PlayerGrid1, PlayerGrid2, PlayerGrid1)
End Function
