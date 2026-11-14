VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GameForm
   Caption         =   "Naval battle"
   ClientHeight    =   7840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13380
   OleObjectBlob   =   "GameForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "The game's window. It has no designed controls: it makes its grids, player pickers, fleet panels, status line and Ready button when it is loaded. FormView paints it, and what the person does in it is sent to the view's events."
'@Folder("NavalBattle.View")
'@ModuleDescription("The game's window. It has no designed controls: it makes its grids, player pickers, fleet panels, status line and Ready button when it is loaded. FormView paints it, and what the person does in it is sent to the view's events.")
Option Explicit

Private Const GridSize As Long = 10
Private Const CellSize As Long = 18
Private Const Margin As Long = 12
Private Const GridGap As Long = 24
Private Const LineHeight As Long = 18

Private Const PickerTop As Long = Margin
Private Const GridTop As Long = PickerTop + LineHeight + 10
Private Const FleetTop As Long = GridTop + (GridSize + 1) * CellSize + 8
Private Const FleetHeight As Long = 5 * 12
Private Const StatusTop As Long = FleetTop + FleetHeight + 8
Private Const StatusHeight As Long = 2 * LineHeight
Private Const ReadyTop As Long = StatusTop + StatusHeight + 6
Private Const ReadyWidth As Long = 72
Private Const ReadyHeight As Long = 24

Private Const WaterColour As Long = &HF5E1C8
Private Const ShipColour As Long = &H707070
Private Const HitColour As Long = &H3C3CE6
Private Const MissColour As Long = &HFFFFFF
Private Const PreviewColour As Long = &H7CD67C
Private Const BadPreviewColour As Long = &H8080F0
Private Const LineColour As Long = &H808080

Private Enum FormPhase
    PhaseSetup
    PhaseShipPlacement
    PhaseWaitForReady
    PhaseAttack
    PhaseOver
End Enum

Private Type TForm
    ViewEvents As IGridViewEvents
    Phase As FormPhase
    HumanGridId As PlayGridId
    ComputerTurn As Boolean
    ComputerOnly As Boolean
    PreviewX As Long
    PreviewY As Long
End Type

' the person closed the window.
Public Event Closed()

Private this As TForm
Private cells(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As GridCell
Private paintedColours(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As Long
Private paintedCaptions(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As String
Private fleetPanels(PlayerGrid1 To PlayerGrid2) As MSForms.Label
Private statusLine As MSForms.Label

Private WithEvents cellEvents As GridCellEvents
Attribute cellEvents.VB_VarHelpID = -1
Private WithEvents firstPlayerPicker As MSForms.ComboBox
Attribute firstPlayerPicker.VB_VarHelpID = -1
Private WithEvents secondPlayerPicker As MSForms.ComboBox
Attribute secondPlayerPicker.VB_VarHelpID = -1
Private WithEvents readyButton As MSForms.CommandButton
Attribute readyButton.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    Set cellEvents = New GridCellEvents

    ' the designer block sizes the window: the board on the left, 444 x 392 points, and on its right a picture area 225 points wide, where the form's
    ' picture (rubber ducks at sea, from GameForm.frx) is aligned at the bottom right.
    Dim boardWidth As Long
    boardWidth = GridLeft(PlayerGrid2) + GridWidth + Margin

    Set statusLine = AddLabel("Status", Margin, StatusTop, boardWidth - 2 * Margin, StatusHeight, "Pick the players, then click Ready.")
    statusLine.WordWrap = True

    Dim gridId As PlayGridId
    For gridId = PlayerGrid1 To PlayerGrid2
        AddGrid gridId
        Set fleetPanels(gridId) = AddLabel("Grid" & gridId & "Fleet", GridLeft(gridId), FleetTop, GridWidth, FleetHeight, "Fleet: no ship placed yet.")
    Next

    Set firstPlayerPicker = AddPlayerPicker(PlayerGrid1, "Human")
    Set secondPlayerPicker = AddPlayerPicker(PlayerGrid2, "Merciless AI")

    Set readyButton = Me.Controls.Add("Forms.CommandButton.1", "Ready", True)
    With readyButton
        .Caption = "Ready"
        .Left = boardWidth - Margin - ReadyWidth
        .Top = ReadyTop
        .Width = ReadyWidth
        .Height = ReadyHeight
    End With
End Sub

'@Description("Where what the person does in the form is sent: the adapter's events.")
Public Property Set ViewEvents(ByVal value As IGridViewEvents)
Attribute ViewEvents.VB_Description = "Where what the person does in the form is sent: the adapter's events."
    Set this.ViewEvents = value
End Property

'@Description("Picks the player of a grid: Human, Random, Fair-play or Merciless; the first letter is enough. Anything else keeps the pick as it is.")
Public Sub PickPlayer(ByVal gridId As PlayGridId, ByVal player As String)
Attribute PickPlayer.VB_Description = "Picks the player of a grid: Human, Random, Fair-play or Merciless; the first letter is enough. Anything else keeps the pick as it is."
    Dim picked As String
    Select Case UCase$(Left$(Trim$(player), 1))
        Case "H"
            picked = "Human"
        Case "R"
            picked = "Random AI"
        Case "F"
            picked = "Fair-play AI"
        Case "M"
            picked = "Merciless AI"
        Case Else
            Exit Sub
    End Select
    PickerOf(gridId).Value = picked
End Sub

'@Description("Sets the text of the status line.")
Public Sub SetStatus(ByVal text As String)
Attribute SetStatus.VB_Description = "Sets the text of the status line."
    statusLine.Caption = text
End Sub

'@Description("Sets the text of a grid's fleet panel.")
Public Sub SetFleet(ByVal gridId As PlayGridId, ByVal text As String)
Attribute SetFleet.VB_Description = "Sets the text of a grid's fleet panel."
    fleetPanels(gridId).Caption = text
End Sub

'@Description("The person places their ships on the given grid: pointing at a cell moves the ship there, a right-click turns it, and a click places it.")
Public Sub BeginShipPlacement(ByVal gridId As PlayGridId)
Attribute BeginShipPlacement.VB_Description = "The person places their ships on the given grid: pointing at a cell moves the ship there, a right-click turns it, and a click places it."
    this.Phase = PhaseShipPlacement
    this.HumanGridId = gridId
    this.PreviewX = 0
    this.PreviewY = 0
End Sub

'@Description("The person's ships are placed: the next click of Ready starts the battle.")
Public Sub BeginWaitForReady()
Attribute BeginWaitForReady.VB_Description = "The person's ships are placed: the next click of Ready starts the battle."
    this.Phase = PhaseWaitForReady
    readyButton.Enabled = True
End Sub

'@Description("The computer plays: the person's clicks are ignored until it is done.")
Public Sub BeginComputerTurn()
Attribute BeginComputerTurn.VB_Description = "The computer plays: the person's clicks are ignored until it is done."
    this.ComputerTurn = True
    Me.Repaint
End Sub

'@Description("The computer has played. When it plays both grids, the form is drawn again, so that the game can be watched.")
Public Sub EndComputerTurn()
Attribute EndComputerTurn.VB_Description = "The computer has played. When it plays both grids, the form is drawn again, so that the game can be watched."
    this.ComputerTurn = False
    If this.ComputerOnly Then DoEvents
End Sub

'@Description("The game is over: the grids and the Ready button do nothing more.")
Public Sub EndGame()
Attribute EndGame.VB_Description = "The game is over: the grids and the Ready button do nothing more."
    this.Phase = PhaseOver
    this.ComputerTurn = False
    readyButton.Enabled = False
End Sub

'@Description("Paints a grid as it stands: hits, misses, water, and the ships when they are to be shown.")
Public Sub PaintGrid(ByVal grid As PlayerGrid, ByVal showShips As Boolean)
Attribute PaintGrid.VB_Description = "Paints a grid as it stands: hits, misses, water, and the ships when they are to be shown."
    Dim X As Long
    For X = 1 To GridSize
        Dim Y As Long
        For Y = 1 To GridSize
            Select Case grid.State(GridCoord.Create(X, Y))
                Case GridState.PreviousHit
                    PaintCell grid.gridId, X, Y, HitColour, "X"
                Case GridState.PreviousMiss
                    PaintCell grid.gridId, X, Y, MissColour
                Case GridState.ShipPosition
                    PaintCell grid.gridId, X, Y, IIf(showShips, ShipColour, WaterColour)
                Case Else
                    PaintCell grid.gridId, X, Y, WaterColour
            End Select
        Next
    Next
End Sub

' a cell is painted only when it changes: each change is a call to Microsoft Forms, and a grid is painted after every shot.
Private Sub PaintCell(ByVal gridId As PlayGridId, ByVal X As Long, ByVal Y As Long, ByVal colour As Long, Optional ByVal text As String = vbNullString)
    If paintedColours(gridId, X, Y) <> colour Then
        cells(gridId, X, Y).Control.BackColor = colour
        paintedColours(gridId, X, Y) = colour
    End If
    If paintedCaptions(gridId, X, Y) <> text Then
        cells(gridId, X, Y).Control.Caption = text
        paintedCaptions(gridId, X, Y) = text
    End If
End Sub

'@Description("Paints where a ship would go, over the grid: green where it can go, red where it cannot.")
Public Sub PaintShipPreview(ByVal grid As PlayerGrid, ByVal previewShip As IShip)
Attribute PaintShipPreview.VB_Description = "Paints where a ship would go, over the grid: green where it can go, red where it cannot."
    Dim colour As Long
    colour = IIf(grid.CanAddShip(previewShip.GridPosition, previewShip.Orientation, previewShip.Size), PreviewColour, BadPreviewColour)

    Dim index As Long
    For index = 0 To previewShip.Size - 1
        Dim X As Long
        X = previewShip.GridPosition.X + IIf(previewShip.Orientation = Horizontal, index, 0)
        Dim Y As Long
        Y = previewShip.GridPosition.Y + IIf(previewShip.Orientation = Vertical, index, 0)
        If X <= GridSize And Y <= GridSize Then PaintCell grid.gridId, X, Y, colour
    Next
End Sub

Private Sub cellEvents_Clicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    If this.ComputerTurn Or this.ViewEvents Is Nothing Then Exit Sub
    Select Case this.Phase
        Case PhaseShipPlacement
            If gridId <> this.HumanGridId Then Exit Sub
            this.ViewEvents.PreviewShipPosition gridId, position
            this.ViewEvents.ConfirmShipPosition gridId, position
        Case PhaseAttack
            If gridId = this.HumanGridId Then Exit Sub
            this.ViewEvents.AttackPosition gridId, position
    End Select
End Sub

Private Sub cellEvents_RightClicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    If this.Phase <> PhaseShipPlacement Or gridId <> this.HumanGridId Then Exit Sub
    this.ViewEvents.PreviewRotateShip gridId, position
End Sub

Private Sub cellEvents_Hovered(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    If this.Phase <> PhaseShipPlacement Or gridId <> this.HumanGridId Then Exit Sub
    If position.X = this.PreviewX And position.Y = this.PreviewY Then Exit Sub

    this.PreviewX = position.X
    this.PreviewY = position.Y
    this.ViewEvents.PreviewShipPosition gridId, position
End Sub

Private Sub firstPlayerPicker_Change()
    If this.Phase = PhaseSetup Then SetStatus "Grid 1 is played by: " & firstPlayerPicker.Value & "."
End Sub

Private Sub secondPlayerPicker_Change()
    If this.Phase = PhaseSetup Then SetStatus "Grid 2 is played by: " & secondPlayerPicker.Value & "."
End Sub

Private Sub readyButton_Click()
    If this.ViewEvents Is Nothing Then Exit Sub
    Select Case this.Phase
        Case PhaseSetup
            StartGame
        Case PhaseWaitForReady
            this.Phase = PhaseAttack
            readyButton.Enabled = False
            SetStatus "Fire at grid " & IIf(this.HumanGridId = PlayerGrid1, PlayerGrid2, PlayerGrid1) & ": click a cell."
            this.ViewEvents.HumanPlayerReady
    End Select
End Sub

' the players picked are made, and the game starts: a person places their ships next, and when the computer plays both grids, the whole game is played now.
Private Sub StartGame()
    Dim firstDifficulty As AIDifficulty
    firstDifficulty = DifficultyOf(firstPlayerPicker.Value)
    Dim secondDifficulty As AIDifficulty
    secondDifficulty = DifficultyOf(secondPlayerPicker.Value)

    If firstDifficulty = Unspecific And secondDifficulty = Unspecific Then
        SetStatus "Only one grid can be played by a person: pick an AI for the other one."
        Exit Sub
    End If

    firstPlayerPicker.Enabled = False
    secondPlayerPicker.Enabled = False
    readyButton.Enabled = False
    this.ComputerOnly = firstDifficulty <> Unspecific And secondDifficulty <> Unspecific
    If this.ComputerOnly Then
        this.Phase = PhaseAttack
        SetStatus "The computer plays both grids."
    End If

    this.ViewEvents.CreatePlayer PlayerGrid1, PlayerTypeOf(firstDifficulty), firstDifficulty
    this.ViewEvents.CreatePlayer PlayerGrid2, PlayerTypeOf(secondDifficulty), secondDifficulty
End Sub

Private Function DifficultyOf(ByVal player As String) As AIDifficulty
    Select Case player
        Case "Random AI"
            DifficultyOf = RandomAI
        Case "Fair-play AI"
            DifficultyOf = FairplayAI
        Case "Merciless AI"
            DifficultyOf = MercilessAI
        Case Else
            DifficultyOf = Unspecific
    End Select
End Function

Private Function PlayerTypeOf(ByVal difficulty As AIDifficulty) As PlayerType
    PlayerTypeOf = IIf(difficulty = Unspecific, HumanControlled, ComputerControlled)
End Function

Private Function PickerOf(ByVal gridId As PlayGridId) As MSForms.ComboBox
    If gridId = PlayerGrid1 Then
        Set PickerOf = firstPlayerPicker
    Else
        Set PickerOf = secondPlayerPicker
    End If
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set this.ViewEvents = Nothing
    Erase cells
    Erase fleetPanels
    RaiseEvent Closed
End Sub

Private Sub AddGrid(ByVal gridId As PlayGridId)
    Dim leftEdge As Long
    leftEdge = GridLeft(gridId)

    Dim index As Long
    For index = 1 To GridSize
        AddLabel "Grid" & gridId & "Column" & index, leftEdge + index * CellSize, GridTop, CellSize, CellSize, Chr$(64 + index), fmTextAlignCenter
        AddLabel "Grid" & gridId & "Row" & index, leftEdge, GridTop + index * CellSize, CellSize, CellSize, CStr(index), fmTextAlignCenter
    Next

    Dim X As Long
    For X = 1 To GridSize
        Dim Y As Long
        For Y = 1 To GridSize
            Dim position As IGridCoord
            Set position = GridCoord.Create(X, Y)

            Dim cellLabel As MSForms.Label
            Set cellLabel = AddLabel("Grid" & gridId & "Cell" & CellName(position), leftEdge + X * CellSize, GridTop + Y * CellSize, CellSize, CellSize, vbNullString, fmTextAlignCenter)
            cellLabel.BackColor = WaterColour
            paintedColours(gridId, X, Y) = WaterColour
            cellLabel.BorderStyle = fmBorderStyleSingle
            cellLabel.BorderColor = LineColour

            Set cells(gridId, X, Y) = GridCell.Create(cellLabel, gridId, position, cellEvents)
        Next
    Next
End Sub

Private Function AddPlayerPicker(ByVal gridId As PlayGridId, ByVal initialPlayer As String) As MSForms.ComboBox
    Const CaptionWidth As Long = 48

    AddLabel "Grid" & gridId & "PlayerCaption", GridLeft(gridId), PickerTop + 3, CaptionWidth, LineHeight, "Grid " & gridId & ":"

    Dim picker As MSForms.ComboBox
    Set picker = Me.Controls.Add("Forms.ComboBox.1", "Grid" & gridId & "Player", True)
    With picker
        .Left = GridLeft(gridId) + CaptionWidth
        .Top = PickerTop
        .Width = GridWidth - CaptionWidth
        .Height = LineHeight
        .Style = fmStyleDropDownList
        .AddItem "Human"
        .AddItem "Random AI"
        .AddItem "Fair-play AI"
        .AddItem "Merciless AI"
        .Value = initialPlayer
    End With
    Set AddPlayerPicker = picker
End Function

Private Function AddLabel(ByVal controlName As String, ByVal leftEdge As Single, ByVal topEdge As Single, ByVal labelWidth As Single, ByVal labelHeight As Single, ByVal text As String, Optional ByVal alignment As fmTextAlign = fmTextAlignLeft) As MSForms.Label
    Dim result As MSForms.Label
    Set result = Me.Controls.Add("Forms.Label.1", controlName, True)
    With result
        .Left = leftEdge
        .Top = topEdge
        .Width = labelWidth
        .Height = labelHeight
        .Caption = text
        .TextAlign = alignment
    End With
    Set AddLabel = result
End Function

Private Function GridLeft(ByVal gridId As PlayGridId) As Long
    GridLeft = Margin + (gridId - PlayerGrid1) * (GridWidth + GridGap)
End Function

Private Property Get GridWidth() As Long
    GridWidth = (GridSize + 1) * CellSize
End Property

Private Function CellName(ByVal position As IGridCoord) As String
    CellName = Chr$(64 + position.X) & position.Y
End Function
