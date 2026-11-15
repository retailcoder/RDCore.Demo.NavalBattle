VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GameForm
   Caption         =   "Naval battle"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16320
   OleObjectBlob   =   "GameForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "The game's window. It has no designed controls: it makes its grids, player pickers, status line and Ready button when it is loaded. WindowView paints it, and what the person does in it is sent to the view's events."
'@Folder("NavalBattle.View.Window.MSForms")
'@ModuleDescription("The game's window. It has no designed controls: it makes its grids, player pickers, status line and Ready button when it is loaded. WindowView paints it, and what the person does in it is sent to the view's events.")
Option Explicit
Implements IGameWindow

Private Const GridSize As Long = 10
Private Const CellSize As Long = 26
Private Const Margin As Long = 12
Private Const GridGap As Long = 48
' the client area's width in points: the designer block's ClientWidth, in twips (20 to the point).
Private Const FormWidth As Long = 16320 \ 20
Private Const LineHeight As Long = 22

' the text of the form: the Windows UI font, at a size that reads at a glance; the status line is larger again, and bold.
Private Const TextFontName As String = "Segoe UI"
Private Const TextFontSize As Long = 11
Private Const PanelFontSize As Long = 12
' a text label's top, in a line or a cell taller than its text, so that the text sits in the middle.
Private Const TextOffset As Long = (CellSize - 16) \ 2

' from the top: the grids, each grid's player under it, then the status line and the Ready button on one line: the window is 392 points high.
Private Const GridTop As Long = Margin
Private Const PickerTop As Long = GridTop + (GridSize + 1) * CellSize + 8
Private Const StatusTop As Long = PickerTop + LineHeight + 8
Private Const StatusHeight As Long = 2 * LineHeight
Private Const ReadyTop As Long = StatusTop + 4
Private Const ReadyWidth As Long = 80
Private Const ReadyHeight As Long = 28

Private Const WaterColour As Long = &HF5E1C8
Private Const ShipColour As Long = &H707070
Private Const HitColour As Long = &H3C3CE6
Private Const MissColour As Long = &HFFFFFF
Private Const PreviewColour As Long = &H7CD67C
Private Const BadPreviewColour As Long = &H8080F0
Private Const LineColour As Long = &H808080

' the sprites, GIF files in assets/sprites with one transparent colour; the ducks are named by DuckSprite.
Private Const SpriteFolder As String = "assets\sprites\"
Private Const SpriteHit As String = "hit-squeak"
Private Const SpriteMiss As String = "miss-ripple"
Private Const SpriteDizzy As String = "sunk-dizzy"
Private Const SpriteAnchor As String = "anchor-cracked"

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
    Locked(PlayerGrid1 To PlayerGrid2) As Boolean
    Listener As IGameWindowEvents
End Type

Private this As TForm
Private cells(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As GridCell
Private paintedColours(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As Long
Private paintedCaptions(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As String
Private paintedSprites(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As String
Private statusLine As MSForms.Label
Private sprites As Collection
Private sunkShipKeys As Collection
Private dizzyShipKey As String

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

    ' the designer block sizes the window: 16320 x 7845 twips, 816 x 392.25 points, 1088 x 523 pixels at 96 dpi. The board is centred on it. The
    ' form's picture (GameForm.frx, from assets/game-form-background.jpg) is stretched over the whole of it, behind the board: the labels are
    ' transparent, and so are the cells of open water. The status line is text over the picture, in a bold, larger font to stay readable.
    ' Without a picture, the water colour shows.
    Me.BackColor = WaterColour
    Set sunkShipKeys = New Collection
    LoadSprites

    Dim boardLeft As Long
    boardLeft = GridLeft(PlayerGrid1)
    Dim boardRight As Long
    boardRight = GridLeft(PlayerGrid2) + GridWidth

    Set statusLine = AddPanel("Status", boardLeft, StatusTop, boardRight - boardLeft - ReadyWidth - Margin, StatusHeight, "Pick the players, then click Ready.")
    statusLine.WordWrap = True

    Dim gridId As PlayGridId
    For gridId = PlayerGrid1 To PlayerGrid2
        AddGrid gridId
    Next

    Set firstPlayerPicker = AddPlayerPicker(PlayerGrid1, PlayerName(Unspecific))
    Set secondPlayerPicker = AddPlayerPicker(PlayerGrid2, PlayerName(MercilessAI))

    Set readyButton = Me.Controls.Add("Forms.CommandButton.1", "Ready", True)
    With readyButton
        .Caption = "Ready"
        .Left = boardRight - ReadyWidth
        .Top = ReadyTop
        .Width = ReadyWidth
        .Height = ReadyHeight
        .Font.Name = TextFontName
        .Font.Size = TextFontSize
    End With
End Sub

Private Property Set IGameWindow_Events(ByVal value As IGridViewEvents)
    Set this.ViewEvents = value
End Property

Private Property Set IGameWindow_Listener(ByVal value As IGameWindowEvents)
    Set this.Listener = value
End Property

Private Sub IGameWindow_Show()
    Me.Show
End Sub

Private Sub IGameWindow_PickPlayer(ByVal gridId As PlayGridId, ByVal player As String)
    Dim difficulty As AIDifficulty
    If TryParsePlayer(player, difficulty) Then PickerOf(gridId).Value = PlayerName(difficulty)
End Sub

Private Sub IGameWindow_SetStatus(ByVal text As String)
    SetStatus text
End Sub

Private Sub SetStatus(ByVal text As String)
    statusLine.Caption = text
End Sub

Private Sub IGameWindow_BeginShipPlacement(ByVal gridId As PlayGridId)
    this.Phase = PhaseShipPlacement
    this.HumanGridId = gridId
    this.PreviewX = 0
    this.PreviewY = 0
End Sub

Private Sub IGameWindow_BeginWaitForReady()
    this.Phase = PhaseWaitForReady
    readyButton.Visible = True
End Sub

Private Sub IGameWindow_BeginComputerTurn()
    this.ComputerTurn = True
    ' DoEvents draws only what changed; Me.Repaint would draw the whole form again, every turn, and it flickers.
    DoEvents
End Sub

Private Sub IGameWindow_EndComputerTurn()
    this.ComputerTurn = False
    If this.ComputerOnly Then DoEvents
End Sub

Private Sub IGameWindow_EndGame()
    this.Phase = PhaseOver
    this.ComputerTurn = False
    readyButton.Visible = False
End Sub

Private Sub IGameWindow_PaintGrid(ByVal grid As PlayerGrid, ByVal showShips As Boolean)
    Dim names(1 To GridSize, 1 To GridSize) As String
    Dim X As Long
    Dim Y As Long
    For X = 1 To GridSize
        For Y = 1 To GridSize
            Select Case grid.State(GridCoord.Create(X, Y))
                Case GridState.PreviousHit
                    names(X, Y) = SpriteHit
                Case GridState.PreviousMiss
                    names(X, Y) = SpriteMiss
            End Select
        Next
    Next

    ' the ships: a duck on each cell that is not hit, when the ships are shown; a sunk ship is shown either way, its bow last.
    Dim item As IShip
    For Each item In grid.Fleet
        Dim index As Long
        For index = 0 To item.Size - 1
            X = item.GridPosition.X + IIf(item.Orientation = Horizontal, index, 0)
            Y = item.GridPosition.Y + IIf(item.Orientation = Vertical, index, 0)
            Dim isBow As Boolean
            isBow = index = item.Size - 1
            If item.IsSunken Then
                If isBow Then names(X, Y) = IIf(IsSinkingDone(grid.gridId, item), SpriteAnchor, SpriteDizzy)
            ElseIf showShips And names(X, Y) = vbNullString Then
                If isBow And ShipKey(grid.gridId, item) = dizzyShipKey Then
                    names(X, Y) = SpriteDizzy
                Else
                    names(X, Y) = DuckSprite(item.Orientation, isBow)
                End If
            End If
        Next
    Next

    For X = 1 To GridSize
        For Y = 1 To GridSize
            If sprites Is Nothing Then
                PaintCell grid.gridId, X, Y, FallbackColour(names(X, Y)), IIf(names(X, Y) = SpriteHit, "X", vbNullString)
            Else
                PaintCell grid.gridId, X, Y, WaterColour, vbNullString, names(X, Y)
            End If
        Next
    Next
End Sub

Private Sub IGameWindow_FinishSinking(ByVal gridId As PlayGridId, ByVal sunkShip As IShip)
    If IsSinkingDone(gridId, sunkShip) Then Exit Sub
    sunkShipKeys.Add ShipKey(gridId, sunkShip)
End Sub

Private Sub IGameWindow_DizzyShip(ByVal gridId As PlayGridId, ByVal hitShip As IShip)
    dizzyShipKey = ShipKey(gridId, hitShip)
End Sub

Private Sub IGameWindow_DizzyShipRecovers()
    dizzyShipKey = vbNullString
End Sub

Private Sub IGameWindow_Pause(ByVal seconds As Single)
    Dim startTime As Single
    startTime = Timer
    Dim elapsed As Single
    Do
        DoEvents
        elapsed = Timer - startTime
        If elapsed < 0 Then elapsed = elapsed + 86400 ' Timer went past midnight
    Loop While elapsed < seconds
End Sub

Private Function IsSinkingDone(ByVal gridId As PlayGridId, ByVal sunkShip As IShip) As Boolean
    Dim key As String
    key = ShipKey(gridId, sunkShip)
    Dim item As Variant
    For Each item In sunkShipKeys
        If item = key Then
            IsSinkingDone = True
            Exit Function
        End If
    Next
End Function

Private Function ShipKey(ByVal gridId As PlayGridId, ByVal item As IShip) As String
    ShipKey = gridId & ":" & item.GridPosition.X & "," & item.GridPosition.Y & ":" & item.Orientation
End Function

' the sprites are read once, from the workspace's assets/sprites folder. When one of them cannot be read, the form goes without them all, and
' paints colours instead.
Private Sub LoadSprites()
    Dim spriteNames As Variant
    spriteNames = Array(DuckSprite(Horizontal, True), DuckSprite(Horizontal, False), DuckSprite(Vertical, True), DuckSprite(Vertical, False), _
        SpriteHit, SpriteMiss, SpriteDizzy, SpriteAnchor)

    Dim path As String
    Set sprites = New Collection
    On Error GoTo NoSprites
    Dim spriteName As Variant
    For Each spriteName In spriteNames
        path = CurDir$ & "\" & SpriteFolder & spriteName & ".gif"
        sprites.Add LoadPicture(path), spriteName
    Next
    Exit Sub

NoSprites:
    Debug.Print "The game form paints colours: " & path & " cannot be read (" & Err.Description & ")."
    Set sprites = Nothing
End Sub

' a ship's ducks face where it goes: right when it lies across, down when it stands; the big duck leads, at its bow, and the small ones follow.
Private Function DuckSprite(ByVal direction As ShipOrientation, ByVal isBow As Boolean) As String
    DuckSprite = "duck-" & IIf(isBow, "big", "small") & "-" & IIf(direction = Horizontal, "right", "down")
End Function

' without its sprites, the form paints each cell with its colour.
Private Function FallbackColour(ByVal sprite As String) As Long
    Select Case sprite
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

' a cell is painted only when it changes: each change is a call to Microsoft Forms, and a grid is painted after every shot. A cell of open water is
' transparent, so the form's picture shows through it; a cell with a sprite is transparent too, and shows the sprite; any other cell is filled
' with its colour.
Private Sub PaintCell(ByVal gridId As PlayGridId, ByVal X As Long, ByVal Y As Long, ByVal colour As Long, Optional ByVal text As String = vbNullString, Optional ByVal sprite As String = vbNullString)
    If paintedSprites(gridId, X, Y) <> sprite Then
        With cells(gridId, X, Y).Control
            If sprite = vbNullString Then
                Set .Picture = Nothing
            Else
                Set .Picture = sprites(sprite)
            End If
        End With
        paintedSprites(gridId, X, Y) = sprite
    End If
    If paintedColours(gridId, X, Y) <> colour Then
        With cells(gridId, X, Y).Control
            If colour = WaterColour Then
                .BackStyle = fmBackStyleTransparent
            Else
                .BackColor = colour
                .BackStyle = fmBackStyleOpaque
            End If
        End With
        paintedColours(gridId, X, Y) = colour
    End If
    If paintedCaptions(gridId, X, Y) <> text Then
        cells(gridId, X, Y).Control.Caption = text
        paintedCaptions(gridId, X, Y) = text
    End If
End Sub

Private Sub IGameWindow_PaintShipPreview(ByVal grid As PlayerGrid, ByVal previewShip As IShip)
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

' the cells stay enabled, since a disabled label greys its picture: the cells' handlers ignore a locked grid instead.
Private Sub IGameWindow_LockGrid(ByVal gridId As PlayGridId, ByVal locked As Boolean)
    this.Locked(gridId) = locked
End Sub

Private Sub cellEvents_Clicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    If this.Locked(gridId) Or this.ComputerTurn Or this.ViewEvents Is Nothing Then Exit Sub
    Select Case this.Phase
        Case PhaseShipPlacement
            this.ViewEvents.PreviewShipPosition gridId, position
            this.ViewEvents.ConfirmShipPosition gridId, position
        Case PhaseAttack
            this.ViewEvents.AttackPosition gridId, position
    End Select
End Sub

Private Sub cellEvents_RightClicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    If this.Locked(gridId) Or this.Phase <> PhaseShipPlacement Then Exit Sub
    this.ViewEvents.PreviewRotateShip gridId, position
End Sub

Private Sub cellEvents_Hovered(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    If this.Locked(gridId) Or this.Phase <> PhaseShipPlacement Then Exit Sub
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
            readyButton.Visible = False
            SetStatus "Fire at grid " & IIf(this.HumanGridId = PlayerGrid1, PlayerGrid2, PlayerGrid1) & ": click a cell."
            this.ViewEvents.HumanPlayerReady
    End Select
End Sub

' the players picked are made, and the game starts: a person places their ships next, and when the computer plays both grids, the whole game is played now.
Private Sub StartGame()
    ' a picker only offers the players' names, so each one reads.
    Dim firstDifficulty As AIDifficulty
    TryParsePlayer firstPlayerPicker.Value, firstDifficulty
    Dim secondDifficulty As AIDifficulty
    TryParsePlayer secondPlayerPicker.Value, secondDifficulty

    If Not CanPlayTogether(firstDifficulty, secondDifficulty) Then
        SetStatus "Only one grid can be played by a person: pick an AI for the other one."
        Exit Sub
    End If

    ' the pickers and the Ready button are done with: the players' names stay, as plain text, and the water is clear.
    ShowPlayer PlayerGrid1, firstPlayerPicker
    ShowPlayer PlayerGrid2, secondPlayerPicker
    readyButton.Visible = False
    this.ComputerOnly = PlayerTypeOf(firstDifficulty) = ComputerControlled And PlayerTypeOf(secondDifficulty) = ComputerControlled
    If this.ComputerOnly Then
        this.Phase = PhaseAttack
        SetStatus "The computer plays both grids."
    End If

    this.ViewEvents.CreatePlayer PlayerGrid1, PlayerTypeOf(firstDifficulty), firstDifficulty
    this.ViewEvents.CreatePlayer PlayerGrid2, PlayerTypeOf(secondDifficulty), secondDifficulty
End Sub


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
    If Not this.Listener Is Nothing Then this.Listener.Closed
    Set this.Listener = Nothing
End Sub

' the top of a grid's row: row 1 is at the grid's top.
Private Function CellTop(ByVal row As Long) As Long
    CellTop = GridTop + (row - 1) * CellSize
End Function

Private Sub AddGrid(ByVal gridId As PlayGridId)
    Dim leftEdge As Long
    leftEdge = GridLeft(gridId)

    ' the row numbers on the left, and the column letters under the last row.
    Dim index As Long
    For index = 1 To GridSize
        AddLabel "Grid" & gridId & "Column" & index, leftEdge + index * CellSize, CellTop(GridSize + 1) + TextOffset, CellSize, CellSize - TextOffset, Chr$(64 + index), fmTextAlignCenter
        AddLabel "Grid" & gridId & "Row" & index, leftEdge, CellTop(index) + TextOffset, CellSize, CellSize - TextOffset, CStr(index), fmTextAlignCenter
    Next

    Dim X As Long
    For X = 1 To GridSize
        Dim Y As Long
        For Y = 1 To GridSize
            Dim position As IGridCoord
            Set position = GridCoord.Create(X, Y)

            Dim cellLabel As MSForms.Label
            Set cellLabel = AddLabel("Grid" & gridId & "Cell" & CellName(position), leftEdge + X * CellSize, CellTop(Y), CellSize, CellSize, vbNullString, fmTextAlignCenter)
            paintedColours(gridId, X, Y) = WaterColour
            cellLabel.BorderStyle = fmBorderStyleSingle
            cellLabel.BorderColor = LineColour
            cellLabel.PicturePosition = fmPicturePositionCenter

            Set cells(gridId, X, Y) = GridCell.Create(cellLabel, gridId, position, cellEvents)
        Next
    Next
End Sub

Private Function AddPlayerPicker(ByVal gridId As PlayGridId, ByVal initialPlayer As String) As MSForms.ComboBox
    Const CaptionWidth As Long = 56

    AddLabel "Grid" & gridId & "PlayerCaption", GridLeft(gridId), PickerTop + 2, CaptionWidth, LineHeight, "Grid " & gridId & ":"

    Dim picker As MSForms.ComboBox
    Set picker = Me.Controls.Add("Forms.ComboBox.1", "Grid" & gridId & "Player", True)
    With picker
        .Left = GridLeft(gridId) + CaptionWidth
        .Top = PickerTop
        .Width = GridWidth - CaptionWidth
        .Height = LineHeight
        .Font.Name = TextFontName
        .Font.Size = TextFontSize
        .Style = fmStyleDropDownList
        Dim offered As Variant
        For Each offered In PlayerNames
            .AddItem offered
        Next
        .Value = initialPlayer
    End With
    Set AddPlayerPicker = picker
End Function

' the players are confirmed: a grid's picker goes, and its caption names the player, in bold so it reads over the ducks.
Private Sub ShowPlayer(ByVal gridId As PlayGridId, ByVal picker As MSForms.ComboBox)
    Dim caption As MSForms.Label
    Set caption = Me.Controls("Grid" & gridId & "PlayerCaption")
    caption.Caption = "Grid " & gridId & ": " & picker.Value
    caption.Width = GridWidth
    caption.Font.Bold = True
    picker.Visible = False
End Sub

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
        .BackStyle = fmBackStyleTransparent
        .Font.Name = TextFontName
        .Font.Size = TextFontSize
    End With
    Set AddLabel = result
End Function

' a label of text over the picture, transparent like the others, in a bold, larger font so it stays readable over the water and the ducks.
Private Function AddPanel(ByVal controlName As String, ByVal leftEdge As Single, ByVal topEdge As Single, ByVal panelWidth As Single, ByVal panelHeight As Single, ByVal text As String) As MSForms.Label
    Dim result As MSForms.Label
    Set result = AddLabel(controlName, leftEdge, topEdge, panelWidth, panelHeight, text)
    result.Font.Bold = True
    result.Font.Size = PanelFontSize
    Set AddPanel = result
End Function

' the grids side by side, GridGap apart, centred on the form.
Private Function GridLeft(ByVal gridId As PlayGridId) As Long
    GridLeft = (FormWidth - 2 * GridWidth - GridGap) \ 2 + (gridId - PlayerGrid1) * (GridWidth + GridGap)
End Function

Private Property Get GridWidth() As Long
    GridWidth = (GridSize + 1) * CellSize
End Property

Private Function CellName(ByVal position As IGridCoord) As String
    CellName = Chr$(64 + position.X) & position.Y
End Function
