VERSION 5.00
Begin VB.Form BoardForm
   Caption         =   "Naval battle"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16320
   ScaleHeight     =   9180
   ScaleWidth      =   16320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReady
      Caption         =   "Ready"
      Height          =   560
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1600
   End
   Begin VB.PictureBox picBanner
      Appearance      =   0  'Flat
      BackColor       =   &H00E6E6E6&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   16330
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   16360
      Begin VB.Label lblFleet
         BackStyle       =   0  'Transparent
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Fleet status"
         Height          =   440
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3600
      End
      Begin VB.ComboBox cboPlayer
         Height          =   315
         Index           =   1
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   0
         Width           =   2400
      End
      Begin VB.Label lblPlayer
         BackStyle       =   0  'Transparent
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Grid 2:"
         Height          =   440
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1040
      End
      Begin VB.ComboBox cboPlayer
         Height          =   315
         Index           =   0
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   0
         Width           =   2400
      End
      Begin VB.Label lblPlayer
         BackStyle       =   0  'Transparent
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Grid 1:"
         Height          =   440
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1040
      End
   End
   Begin VB.Image imgShipCell
      Height          =   360
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgBackground
      Height          =   7845
      Left            =   0
      Top             =   0
      Width           =   16320
   End
   Begin VB.PictureBox picCell
      Appearance      =   0  'Flat
      BackColor       =   &H00F5E1C8&
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Index           =   0
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   360
   End
   Begin VB.Label lblShip
      BackStyle       =   0  'Transparent
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   440
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Label lblStatus
      BackStyle       =   0  'Transparent
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Pick the players, then click Ready."
      Height          =   440
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   15840
   End
   Begin VB.Label lblStatusBar
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   780
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   16360
   End
End
Attribute VB_Name = "BoardForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "The game's RDForms window, a VB6 form: its status bar, banner, player pickers, Ready button, first cell, first ship name and first fleet cell are designed, and the other cells and names are loaded from the first ones. BoardLayout places them, BoardSprites says what each cell shows, and BoardInput sends what the person does to the view's events."
'@Folder("NavalBattle.View.Window.RDForms")
'@ModuleDescription("The game's RDForms window, a VB6 form: its status bar, banner, player pickers, Ready button, first cell, first ship name and first fleet cell are designed, and the other cells and names are loaded from the first ones. BoardLayout places them, BoardSprites says what each cell shows, and BoardInput sends what the person does to the view's events.")
Option Explicit
Implements IGameWindow

' BoardLayout measures in points, and a VB6 form in twips.
Private Const TwipsPerPoint As Long = 20
' the cells of a grid, in picCell: grid 1's are 0 to 99, and grid 2's follow.
Private Const CellsPerGrid As Long = GridSize * GridSize
Private Const BackgroundFile As String = "assets\game-form-background.jpg"
Private Const IconFile As String = "assets\ducky.ico"

' who is told when the person closes the window.
Private listener As IGameWindowEvents
Private paintedColours(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As Long
Private paintedCaptions(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As String
Private paintedSprites(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As String
' the ships the fleet status shows so far.
Private fleetShipCount As Long
Private sprites As Collection
' what each cell shows, and what the person does in the window: the game's window logic, which this form only draws.
Private board As BoardSprites
Private WithEvents personInput As BoardInput
Attribute personInput.VB_VarHelpID = -1

' the game talks to the window before it is loaded, when it is made: a VB6 form loads only once a control of it is used.
Private Sub Form_Initialize()
    Set board = New BoardSprites
    Set personInput = New BoardInput
End Sub

' the banner is a grey PictureBox across the window that holds the columns' captions and the player pickers: a label shows over a PictureBox
' only from inside it. The background picture is under the banner, whole, at its own size, in an image sent to the back, so the labels show
' over it; without it, the water colour shows. The Ready button starts the row of buttons at the bottom left, on open water.
Private Sub Form_Load()
    Me.BackColor = board.Water
    LoadWindowIcon
    LoadBackground
    LoadSprites

    Place lblStatusBar, BoardLayout.StatusBarLeft, BoardLayout.StatusBarTop, BoardLayout.StatusBarWidth, BoardLayout.StatusBarHeight
    Place lblStatus, BoardLayout.StatusTextLeft, BoardLayout.StatusTextTop, BoardLayout.StatusTextWidth, LineHeight
    Place picBanner, BoardLayout.BannerLeft, BoardLayout.BannerTop, BoardLayout.BannerWidth, BoardLayout.BannerHeight

    Dim gridId As PlayGridId
    For gridId = PlayerGrid1 To PlayerGrid2
        AddGrid gridId
        AddPlayerPicker gridId
    Next
    cboPlayer(PickerIndex(PlayerGrid1)).ListIndex = PlayerIndex(Unspecific)
    cboPlayer(PickerIndex(PlayerGrid2)).ListIndex = PlayerIndex(MercilessAI)

    PlaceInBanner lblFleet, BoardLayout.FleetCaptionLeft, BoardLayout.BannerTextTop + 2, BoardLayout.CaptionWidth, LineHeight

    Place cmdReady, BoardLayout.ButtonLeft(1), BoardLayout.ButtonTop, ButtonWidth, ButtonHeight
End Sub

Private Property Set IGameWindow_Events(ByVal value As IGridViewEvents)
    Set personInput.Events = value
End Property

Private Property Set IGameWindow_Listener(ByVal value As IGameWindowEvents)
    Set listener = value
End Property

Private Sub IGameWindow_Show()
    Me.Show vbModal
End Sub

Private Sub IGameWindow_PickPlayer(ByVal gridId As PlayGridId, ByVal player As String)
    Dim difficulty As AIDifficulty
    If TryParsePlayer(player, difficulty) Then cboPlayer(PickerIndex(gridId)).ListIndex = PlayerIndex(difficulty)
End Sub

Private Sub IGameWindow_SetStatus(ByVal text As String)
    lblStatus.Caption = text
End Sub

Private Sub IGameWindow_BeginShipPlacement(ByVal gridId As PlayGridId)
    personInput.BeginShipPlacement gridId
End Sub

Private Sub IGameWindow_BeginWaitForReady()
    personInput.BeginWaitForReady
End Sub

Private Sub IGameWindow_BeginComputerTurn()
    personInput.BeginComputerTurn
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

Private Sub IGameWindow_FinishSinking(ByVal gridId As PlayGridId, ByVal sunkShip As IShip)
    board.FinishSinking gridId, sunkShip
End Sub

Private Sub IGameWindow_DizzyShip(ByVal gridId As PlayGridId, ByVal hitShip As IShip)
    board.DizzyShip gridId, hitShip
End Sub

Private Sub IGameWindow_DizzyShipRecovers()
    board.DizzyShipRecovers
End Sub

Private Sub IGameWindow_Squeak(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    board.Squeak gridId, position
End Sub

Private Sub IGameWindow_SqueakStops()
    board.SqueakStops
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

Private Sub IGameWindow_LockGrid(ByVal gridId As PlayGridId, ByVal locked As Boolean)
    personInput.LockGrid gridId, locked
End Sub

Private Sub picCell_Click(Index As Integer)
    personInput.Click GridOf(Index), PositionOf(Index)
End Sub

Private Sub picCell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    personInput.Hover GridOf(Index), PositionOf(Index)
End Sub

Private Sub picCell_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then personInput.RightClick GridOf(Index), PositionOf(Index)
End Sub

' a Style 2 list raises Click when a player is picked, by the person or by PickPlayer.
Private Sub cboPlayer_Click(Index As Integer)
    personInput.PlayerPicked GridOfPicker(Index), cboPlayer(Index).Text
End Sub

Private Sub cmdReady_Click()
    personInput.ReadyClicked cboPlayer(PickerIndex(PlayerGrid1)).Text, cboPlayer(PickerIndex(PlayerGrid2)).Text
End Sub

Private Sub personInput_StatusChanged(ByVal text As String)
    lblStatus.Caption = text
End Sub

Private Sub personInput_ReadyShown(ByVal visible As Boolean)
    cmdReady.Visible = visible
End Sub

' the pickers are done with: each grid's caption names its player, centred over the grid's cells.
Private Sub personInput_PlayersConfirmed()
    Dim gridId As PlayGridId
    For gridId = PlayerGrid1 To PlayerGrid2
        PlaceInBanner lblPlayer(PickerIndex(gridId)), BoardLayout.GridCaptionLeft(gridId), BoardLayout.BannerTextTop + 2, BoardLayout.CaptionWidth, LineHeight
        With lblPlayer(PickerIndex(gridId))
            .Caption = "Grid " & gridId & ": " & cboPlayer(PickerIndex(gridId)).Text
            .Alignment = vbCenter
        End With
        cboPlayer(PickerIndex(gridId)).Visible = False
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set personInput.Events = Nothing
    If Not listener Is Nothing Then listener.Closed
    Set listener = Nothing
End Sub

' the row numbers on the left, and the column letters under the last row; then the cells, loaded from the designed one.
Private Sub AddGrid(ByVal gridId As PlayGridId)
    Dim index As Long
    For index = 1 To GridSize
        AddCaption "Grid" & gridId & "Column" & index, BoardLayout.CellLeft(gridId, index), BoardLayout.CellTop(GridSize + 1), Chr$(64 + index)
        AddCaption "Grid" & gridId & "Row" & index, BoardLayout.CellLeft(gridId, 0), BoardLayout.CellTop(index), CStr(index)
    Next

    Dim X As Long
    For X = 1 To GridSize
        Dim Y As Long
        For Y = 1 To GridSize
            Dim cellNumber As Long
            cellNumber = CellIndex(gridId, X, Y)
            If cellNumber > 0 Then Load picCell(cellNumber)
            Place picCell(cellNumber), BoardLayout.CellLeft(gridId, X), BoardLayout.CellTop(Y), CellSize, CellSize
            picCell(cellNumber).Visible = True
            paintedColours(gridId, X, Y) = board.Water
        Next
    Next
End Sub

' a grid's caption on the banner, then its picker after it.
Private Sub AddPlayerPicker(ByVal gridId As PlayGridId)
    PlaceInBanner lblPlayer(PickerIndex(gridId)), BoardLayout.GridCaptionLeft(gridId) + BannerPadding, BoardLayout.BannerTextTop + 2, PlayerCaptionWidth - BannerPadding, LineHeight
    With cboPlayer(PickerIndex(gridId))
        .Left = (BoardLayout.PickerLeft(gridId) - BoardLayout.BannerLeft) * TwipsPerPoint
        .Top = (BoardLayout.BannerTextTop - BoardLayout.BannerTop) * TwipsPerPoint
        .Width = BoardLayout.PickerWidth * TwipsPerPoint
        Dim offered As Variant
        For Each offered In PlayerNames
            .AddItem offered
        Next
    End With
End Sub

' the fleet status: each ship placed so far, by its name, lying across, with its ducks red where they are hit. A sunk ship's name is
' struck out.
Private Sub PaintFleet(ByVal grid As PlayerGrid)
    lblFleet.Caption = "Grid " & grid.gridId & " fleet"
    Dim ships As Collection
    Set ships = grid.Fleet
    Dim shipSprites As Collection
    Set shipSprites = board.FleetSprites(grid)
    Dim shipNumber As Long
    For shipNumber = 1 To ships.Count
        Dim item As IShip
        Set item = ships(shipNumber)
        If shipNumber > fleetShipCount Then AddFleetShip shipNumber, item
        lblShip(shipNumber - 1).FontStrikethru = item.IsSunken

        Dim cellSprites As Variant
        cellSprites = shipSprites(shipNumber)
        Dim column As Long
        For column = 1 To item.Size
            PaintFleetCell shipNumber, column, cellSprites(column)
        Next
    Next
End Sub

' a ship's name and cells, loaded from the designed ones.
Private Sub AddFleetShip(ByVal shipNumber As Long, ByVal item As IShip)
    If shipNumber > 1 Then Load lblShip(shipNumber - 1)
    Place lblShip(shipNumber - 1), BoardLayout.ShipCellLeft(1), BoardLayout.ShipNameTop(shipNumber), BoardLayout.CaptionWidth, LineHeight
    lblShip(shipNumber - 1).Caption = item.Name
    lblShip(shipNumber - 1).Visible = True

    Dim column As Long
    For column = 1 To item.Size
        Dim cellNumber As Long
        cellNumber = FleetCellIndex(shipNumber, column)
        If cellNumber > 0 Then Load imgShipCell(cellNumber)
        imgShipCell(cellNumber).Visible = True
    Next
    fleetShipCount = shipNumber
End Sub

' a fleet cell is an image, with no border, so the sky shows around its duck. It is painted only when it changes, as a grid's cell is: its Tag
' holds the sprite it shows. An image takes its picture's size, so it is centred on the cell's place again. Without the sprites, the fleet
' status shows the ships' names only: an image cannot be filled with a colour.
Private Sub PaintFleetCell(ByVal shipNumber As Long, ByVal column As Long, ByVal sprite As String)
    Dim cell As VB.Image
    Set cell = imgShipCell(FleetCellIndex(shipNumber, column))
    If cell.Tag = sprite Or sprites Is Nothing Then Exit Sub
    Set cell.Picture = sprites(sprite)
    cell.Move BoardLayout.ShipCellLeft(column) * TwipsPerPoint + (CellSize * TwipsPerPoint - cell.Width) \ 2, _
        BoardLayout.ShipCellTop(shipNumber) * TwipsPerPoint + (CellSize * TwipsPerPoint - cell.Height) \ 2
    cell.Tag = sprite
End Sub

Private Sub AddCaption(ByVal controlName As String, ByVal leftEdge As Long, ByVal topEdge As Long, ByVal text As String)
    Dim result As VB.Label
    Set result = Me.Controls.Add("VB.Label", controlName)
    Place result, leftEdge, topEdge, CellSize, CellSize
    With result
        .Caption = text
        .Alignment = vbCenter
        .BackStyle = 0 ' transparent
        .Visible = True
    End With
End Sub

' places a control by BoardLayout's points.
Private Sub Place(ByVal target As Object, ByVal leftEdge As Long, ByVal topEdge As Long, ByVal controlWidth As Long, ByVal controlHeight As Long)
    target.Move leftEdge * TwipsPerPoint, topEdge * TwipsPerPoint, controlWidth * TwipsPerPoint, controlHeight * TwipsPerPoint
End Sub

' places a control inside the banner by BoardLayout's points: the banner's controls are placed from its own top left corner.
Private Sub PlaceInBanner(ByVal target As Object, ByVal leftEdge As Long, ByVal topEdge As Long, ByVal controlWidth As Long, ByVal controlHeight As Long)
    Place target, leftEdge - BoardLayout.BannerLeft, topEdge - BoardLayout.BannerTop, controlWidth, controlHeight
End Sub

' a cell is painted only when it changes, as in the Microsoft Forms window. A PictureBox cannot be transparent, so a cell of open water shows
' the water colour, not the background picture under it.
Private Sub PaintCell(ByVal gridId As PlayGridId, ByVal X As Long, ByVal Y As Long, ByVal colour As Long, Optional ByVal text As String = vbNullString, Optional ByVal sprite As String = vbNullString)
    Dim cell As VB.PictureBox
    Set cell = picCell(CellIndex(gridId, X, Y))
    If paintedSprites(gridId, X, Y) <> sprite Then
        If sprite = vbNullString Then
            Set cell.Picture = Nothing
        Else
            Set cell.Picture = sprites(sprite)
        End If
        paintedSprites(gridId, X, Y) = sprite
    End If
    If paintedColours(gridId, X, Y) <> colour Then
        cell.BackColor = colour
        paintedColours(gridId, X, Y) = colour
    End If
    If paintedCaptions(gridId, X, Y) <> text Then
        cell.Cls
        If text <> vbNullString Then cell.Print text
        paintedCaptions(gridId, X, Y) = text
    End If
End Sub

' the ducky on the window's title bar. Without it, the window keeps the icon it has.
Private Sub LoadWindowIcon()
    On Error GoTo NoIcon
    Set Me.Icon = LoadPicture(CurDir$ & "\" & IconFile)
    Exit Sub

NoIcon:
    Debug.Print "The RDForms window keeps its own icon: " & IconFile & " cannot be read (" & Err.Description & ")."
End Sub

Private Sub LoadBackground()
    Place imgBackground, 0, BoardLayout.BackgroundTop, WindowWidth, BackgroundHeight
    imgBackground.ZOrder vbSendToBack
    On Error GoTo NoBackground
    Set imgBackground.Picture = LoadPicture(CurDir$ & "\" & BackgroundFile)
    Exit Sub

NoBackground:
    Debug.Print "The RDForms window shows the water colour: " & BackgroundFile & " cannot be read (" & Err.Description & ")."
End Sub

' the sprites are read once. When one of them cannot be read, the window goes without them all, and paints colours instead.
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
    Debug.Print "The RDForms window paints colours: " & path & " cannot be read (" & Err.Description & ")."
    Set sprites = Nothing
End Sub

Private Function CellIndex(ByVal gridId As PlayGridId, ByVal X As Long, ByVal Y As Long) As Long
    CellIndex = (gridId - PlayerGrid1) * CellsPerGrid + (Y - 1) * GridSize + (X - 1)
End Function

' the cells of the fleet status, in imgShipCell: each ship's take a row of GridSize, longer than any ship, and the first ship's start at 0.
Private Function FleetCellIndex(ByVal shipNumber As Long, ByVal column As Long) As Long
    FleetCellIndex = (shipNumber - 1) * GridSize + (column - 1)
End Function

Private Function GridOf(ByVal cellNumber As Integer) As PlayGridId
    GridOf = PlayerGrid1 + cellNumber \ CellsPerGrid
End Function

Private Function PositionOf(ByVal cellNumber As Integer) As IGridCoord
    Dim inGrid As Long
    inGrid = cellNumber Mod CellsPerGrid
    Set PositionOf = GridCoord.Create(inGrid Mod GridSize + 1, inGrid \ GridSize + 1)
End Function

Private Function PickerIndex(ByVal gridId As PlayGridId) As Integer
    PickerIndex = gridId - PlayerGrid1
End Function

Private Function GridOfPicker(ByVal pickerNumber As Integer) As PlayGridId
    GridOfPicker = PlayerGrid1 + pickerNumber
End Function

' a player's place in the pickers' lists, which offer PlayerNames in order.
Private Function PlayerIndex(ByVal difficulty As AIDifficulty) As Long
    Dim names As Variant
    names = PlayerNames
    Dim index As Long
    For index = LBound(names) To UBound(names)
        If names(index) = PlayerName(difficulty) Then
            PlayerIndex = index - LBound(names)
            Exit Function
        End If
    Next
End Function
