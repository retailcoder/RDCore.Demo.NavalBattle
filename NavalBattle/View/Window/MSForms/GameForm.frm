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
Attribute VB_Description = "The game's window. It has no designed controls: it makes its status line, its banner and the player pickers on it, its grids, its fleet status and its Ready button when it is loaded. WindowView paints it, and what the person does in it is sent to the view's events."
'@Folder("NavalBattle.View.Window.MSForms")
'@ModuleDescription("The game's window. It has no designed controls: it makes its status line, its banner and the player pickers on it, its grids, its fleet status and its Ready button when it is loaded. WindowView paints it, and what the person does in it is sent to the view's events.")
Option Explicit
Implements IGameWindow

' the text of the form: the Windows UI font, at a size that reads at a glance; the status line is larger again, and bold.
Private Const TextFontName As String = "Segoe UI"
Private Const TextFontSize As Long = 11
Private Const PanelFontSize As Long = 12
' a text label's top, in a line or a cell taller than its text, so that the text sits in the middle.
Private Const TextOffset As Long = (CellSize - 16) \ 2

' the parts go where BoardLayout says. The status bar is white, with a thin black border, of which only the bottom shows. Each column's banner
' is a light grey band, with a thin grey border: an empty label, behind its text and its player picker.
Private Const StatusColour As Long = &HFFFFFF
Private Const StatusBorderColour As Long = &H0
Private Const BannerColour As Long = &HE6E6E6

Private Const LineColour As Long = &H808080

' who is told when the person closes the window.
Private listener As IGameWindowEvents
Private cells(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As GridCell
Private paintedColours(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As Long
Private paintedCaptions(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As String
Private paintedSprites(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As String
Private statusLine As MSForms.Label
Private fleetCaption As MSForms.Label
' the fleet status's ships, in the order they were placed: each one's cells, from its stern, made when the ship first shows.
Private fleetCells As Collection
Private fleetNames As Collection
Private sprites As Collection
' what each cell shows, and what the person does in the window: the game's window logic, which this form only draws.
Private board As BoardSprites
Private WithEvents personInput As BoardInput
Attribute personInput.VB_VarHelpID = -1

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
    Set board = New BoardSprites
    Set personInput = New BoardInput

    ' the designer block sizes the window: 16320 x 7845 twips, 816 x 392.25 points, 1088 x 523 pixels at 96 dpi. The board is centred on it. The
    ' form's picture (GameForm.frx, from assets/game-form-background.jpg) is stretched over the whole of it, behind the board: the labels are
    ' transparent, and so are the cells of open water. The status line is bold, larger text on a white bar across the top of the window. A grey
    ' banner runs across the window under it, with each column's caption over it: the grids' player pickers, then the players' names, and the
    ' fleet status's. The Ready button starts the row of buttons at the bottom left, on open water.
    ' Without a picture, the water colour shows.
    Me.BackColor = board.Water
    LoadSprites

    With AddLabel("StatusBar", BoardLayout.StatusBarLeft, BoardLayout.StatusBarTop, BoardLayout.StatusBarWidth, BoardLayout.StatusBarHeight, vbNullString)
        .BackStyle = fmBackStyleOpaque
        .BackColor = StatusColour
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = StatusBorderColour
    End With
    Set statusLine = AddPanel("Status", BoardLayout.StatusTextLeft, BoardLayout.StatusTextTop, BoardLayout.StatusTextWidth, LineHeight, "Pick the players, then click Ready.")
    AddBanner

    Dim gridId As PlayGridId
    For gridId = PlayerGrid1 To PlayerGrid2
        AddGrid gridId
    Next

    Set firstPlayerPicker = AddPlayerPicker(PlayerGrid1, PlayerName(Unspecific))
    Set secondPlayerPicker = AddPlayerPicker(PlayerGrid2, PlayerName(MercilessAI))

    Set fleetCaption = AddPanel("FleetCaption", BoardLayout.FleetCaptionLeft, BoardLayout.BannerTextTop + 2, BoardLayout.CaptionWidth, LineHeight, "Fleet status")
    fleetCaption.TextAlign = fmTextAlignCenter
    Set fleetCells = New Collection
    Set fleetNames = New Collection

    Set readyButton = Me.Controls.Add("Forms.CommandButton.1", "Ready", True)
    With readyButton
        .Caption = "Ready"
        .Left = BoardLayout.ButtonLeft(1)
        .Top = BoardLayout.ButtonTop
        .Width = ButtonWidth
        .Height = ButtonHeight
        .Font.Name = TextFontName
        .Font.Size = TextFontSize
    End With
End Sub

Private Property Set IGameWindow_Events(ByVal value As IGridViewEvents)
    Set personInput.Events = value
End Property

Private Property Set IGameWindow_Listener(ByVal value As IGameWindowEvents)
    Set listener = value
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
    personInput.BeginShipPlacement gridId
End Sub

Private Sub IGameWindow_BeginWaitForReady()
    personInput.BeginWaitForReady
End Sub

Private Sub IGameWindow_BeginComputerTurn()
    personInput.BeginComputerTurn
    ' DoEvents draws only what changed; Me.Repaint would draw the whole form again, every turn, and it flickers.
    DoEvents
End Sub

Private Sub IGameWindow_EndComputerTurn()
    personInput.EndComputerTurn
    If personInput.ComputerOnly Then DoEvents
End Sub

Private Sub IGameWindow_EndGame()
    personInput.EndGame
End Sub

Private Sub IGameWindow_PaintGrid(ByVal grid As PlayerGrid, ByVal showShips As Boolean)
    Dim names As Variant
    names = board.CellSprites(grid, showShips)
    Dim X As Long
    Dim Y As Long
    For X = 1 To GridSize
        For Y = 1 To GridSize
            If sprites Is Nothing Then
                PaintCell grid.gridId, X, Y, board.FallbackColour(names(X, Y)), IIf(board.IsHit(names(X, Y)), "X", vbNullString)
            Else
                PaintCell grid.gridId, X, Y, board.Water, vbNullString, names(X, Y)
            End If
        Next
    Next
    If board.IsFleetGrid(grid.gridId, showShips) Then PaintFleet grid
End Sub

' the fleet status: each ship placed so far, by its name, lying across, with what its cells show on the grid. A sunk ship's name is struck out.
Private Sub PaintFleet(ByVal grid As PlayerGrid)
    fleetCaption.Caption = "Grid " & grid.gridId & " fleet"
    Dim ships As Collection
    Set ships = grid.Fleet
    Dim shipSprites As Collection
    Set shipSprites = board.FleetSprites(grid)
    Dim shipNumber As Long
    For shipNumber = 1 To ships.Count
        Dim item As IShip
        Set item = ships(shipNumber)
        If shipNumber > fleetCells.Count Then AddFleetShip shipNumber, item
        Dim shipName As MSForms.Label
        Set shipName = fleetNames(shipNumber)
        shipName.Font.Strikethrough = item.IsSunken

        Dim cells As Collection
        Set cells = fleetCells(shipNumber)
        Dim cellSprites As Variant
        cellSprites = shipSprites(shipNumber)
        Dim column As Long
        For column = 1 To item.Size
            PaintFleetCell cells(column), cellSprites(column)
        Next
    Next
End Sub

' a fleet cell is painted only when it changes, as a grid's is: its Tag holds the sprite it shows.
Private Sub PaintFleetCell(ByVal cell As MSForms.Label, ByVal sprite As String)
    If cell.Tag = sprite Then Exit Sub
    If sprites Is Nothing Then
        cell.BackColor = board.FallbackColour(sprite)
        cell.BackStyle = IIf(sprite = vbNullString, fmBackStyleTransparent, fmBackStyleOpaque)
        cell.Caption = IIf(board.IsHit(sprite), "X", vbNullString)
    ElseIf sprite = vbNullString Then
        Set cell.Picture = Nothing
    Else
        Set cell.Picture = sprites(sprite)
    End If
    cell.Tag = sprite
End Sub

Private Sub AddFleetShip(ByVal shipNumber As Long, ByVal item As IShip)
    Dim shipName As MSForms.Label
    Set shipName = AddPanel("Ship" & shipNumber & "Name", BoardLayout.ShipCellLeft(1), BoardLayout.ShipNameTop(shipNumber) + TextOffset, BoardLayout.CaptionWidth, LineHeight - TextOffset, item.Name)
    shipName.Font.Size = TextFontSize
    fleetNames.Add shipName

    Dim cells As Collection
    Set cells = New Collection
    Dim column As Long
    For column = 1 To item.Size
        Dim cell As MSForms.Label
        Set cell = AddLabel("Ship" & shipNumber & "Cell" & column, BoardLayout.ShipCellLeft(column), BoardLayout.ShipCellTop(shipNumber), CellSize, CellSize, vbNullString, fmTextAlignCenter)
        cell.BorderStyle = fmBorderStyleSingle
        cell.BorderColor = LineColour
        cell.PicturePosition = fmPicturePositionCenter
        cells.Add cell
    Next
    fleetCells.Add cells
End Sub

Private Sub IGameWindow_FinishSinking(ByVal gridId As PlayGridId, ByVal sunkShip As IShip)
    board.FinishSinking gridId, sunkShip
End Sub

Private Sub IGameWindow_DizzyShip(ByVal gridId As PlayGridId, ByVal hitShip As IShip)
    board.DizzyShip gridId, hitShip
End Sub

Private Sub IGameWindow_DizzyShipRecovers()
    board.DizzyShipRecovers
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

Private Sub LoadSprites()
    Dim path As String
    Set sprites = New Collection
    On Error GoTo NoSprites
    Dim spriteName As Variant
    For Each spriteName In board.Names
        path = board.Path(spriteName)
        sprites.Add LoadPicture(path), spriteName
    Next
    Exit Sub

NoSprites:
    Debug.Print "The game form paints colours: " & path & " cannot be read (" & Err.Description & ")."
    Set sprites = Nothing
End Sub

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
            If colour = board.Water Then
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
    colour = board.ShipPreviewColour(grid.CanAddShip(previewShip.GridPosition, previewShip.Orientation, previewShip.Size))

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
    personInput.LockGrid gridId, locked
End Sub

Private Sub cellEvents_Clicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    personInput.Click gridId, position
End Sub

Private Sub cellEvents_RightClicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    personInput.RightClick gridId, position
End Sub

Private Sub cellEvents_Hovered(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    personInput.Hover gridId, position
End Sub

Private Sub firstPlayerPicker_Change()
    personInput.PlayerPicked PlayerGrid1, firstPlayerPicker.Value
End Sub

Private Sub secondPlayerPicker_Change()
    personInput.PlayerPicked PlayerGrid2, secondPlayerPicker.Value
End Sub

Private Sub readyButton_Click()
    personInput.ReadyClicked firstPlayerPicker.Value, secondPlayerPicker.Value
End Sub

Private Sub personInput_StatusChanged(ByVal text As String)
    SetStatus text
End Sub

Private Sub personInput_ReadyShown(ByVal visible As Boolean)
    readyButton.Visible = visible
End Sub

' the pickers are done with: the players' names stay, as plain text, and the water is clear.
Private Sub personInput_PlayersConfirmed()
    ShowPlayer PlayerGrid1, firstPlayerPicker
    ShowPlayer PlayerGrid2, secondPlayerPicker
End Sub

Private Function PickerOf(ByVal gridId As PlayGridId) As MSForms.ComboBox
    If gridId = PlayerGrid1 Then
        Set PickerOf = firstPlayerPicker
    Else
        Set PickerOf = secondPlayerPicker
    End If
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set personInput.Events = Nothing
    Erase cells
    If Not listener Is Nothing Then listener.Closed
    Set listener = Nothing
End Sub

Private Sub AddGrid(ByVal gridId As PlayGridId)
    ' the row numbers on the left, and the column letters under the last row.
    Dim index As Long
    For index = 1 To GridSize
        AddLabel "Grid" & gridId & "Column" & index, BoardLayout.CellLeft(gridId, index), BoardLayout.CellTop(GridSize + 1) + TextOffset, CellSize, CellSize - TextOffset, Chr$(64 + index), fmTextAlignCenter
        AddLabel "Grid" & gridId & "Row" & index, BoardLayout.CellLeft(gridId, 0), BoardLayout.CellTop(index) + TextOffset, CellSize, CellSize - TextOffset, CStr(index), fmTextAlignCenter
    Next

    Dim X As Long
    For X = 1 To GridSize
        Dim Y As Long
        For Y = 1 To GridSize
            Dim position As IGridCoord
            Set position = GridCoord.Create(X, Y)

            Dim cellLabel As MSForms.Label
            Set cellLabel = AddLabel("Grid" & gridId & "Cell" & CellName(position), BoardLayout.CellLeft(gridId, X), BoardLayout.CellTop(Y), CellSize, CellSize, vbNullString, fmTextAlignCenter)
            paintedColours(gridId, X, Y) = board.Water
            cellLabel.BorderStyle = fmBorderStyleSingle
            cellLabel.BorderColor = LineColour
            cellLabel.PicturePosition = fmPicturePositionCenter

            Set cells(gridId, X, Y) = GridCell.Create(cellLabel, gridId, position, cellEvents)
        Next
    Next
End Sub

' a grid's caption on the banner, then its picker after it.
Private Function AddPlayerPicker(ByVal gridId As PlayGridId, ByVal initialPlayer As String) As MSForms.ComboBox
    AddLabel "Grid" & gridId & "PlayerCaption", BoardLayout.GridCaptionLeft(gridId) + BannerPadding, BoardLayout.BannerTextTop + 2, PlayerCaptionWidth - BannerPadding, LineHeight, "Grid " & gridId & ":"

    Dim picker As MSForms.ComboBox
    Set picker = Me.Controls.Add("Forms.ComboBox.1", "Grid" & gridId & "Player", True)
    With picker
        .Left = BoardLayout.PickerLeft(gridId)
        .Top = BoardLayout.BannerTextTop
        .Width = BoardLayout.PickerWidth
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

' the players are confirmed: a grid's picker goes, and its caption names the player, centred over the grid's cells and in bold.
Private Sub ShowPlayer(ByVal gridId As PlayGridId, ByVal picker As MSForms.ComboBox)
    Dim caption As MSForms.Label
    Set caption = Me.Controls("Grid" & gridId & "PlayerCaption")
    caption.Caption = "Grid " & gridId & ": " & picker.Value
    caption.Left = BoardLayout.GridCaptionLeft(gridId)
    caption.Width = BoardLayout.CaptionWidth
    caption.TextAlign = fmTextAlignCenter
    caption.Font.Bold = True
    caption.Font.Size = PanelFontSize
    picker.Visible = False
End Sub

' the banner across the window: an empty label, sent to the back, so that the captions and the pickers show over it whatever order they were made in.
Private Sub AddBanner()
    With AddLabel("Banner", BoardLayout.BannerLeft, BoardLayout.BannerTop, BoardLayout.BannerWidth, BoardLayout.BannerHeight, vbNullString)
        .BackStyle = fmBackStyleOpaque
        .BackColor = BannerColour
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = StatusBorderColour
        .ZOrder fmZOrderBack
    End With
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

Private Function CellName(ByVal position As IGridCoord) As String
    CellName = Chr$(64 + position.X) & position.Y
End Function
