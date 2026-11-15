VERSION 5.00
Begin VB.Form BoardForm
   Caption         =   "Naval battle"
   ClientHeight    =   7840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16320
   ScaleHeight     =   7840
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
   Begin VB.ComboBox cboPlayer
      Height          =   315
      Index           =   1
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   2400
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
   Begin VB.Label lblPlayer
      BackStyle       =   0  'Transparent
      Caption         =   "Grid 2:"
      Height          =   440
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1120
   End
   Begin VB.Label lblPlayer
      BackStyle       =   0  'Transparent
      Caption         =   "Grid 1:"
      Height          =   440
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1120
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
Attribute VB_Description = "The game's RDForms window, a VB6 form: its status bar, player pickers, Ready button and first cell are designed, and the other cells are loaded from the first one. BoardLayout places them, BoardSprites says what each cell shows, and BoardInput sends what the person does to the view's events."
'@Folder("NavalBattle.View.Window.RDForms")
'@ModuleDescription("The game's RDForms window, a VB6 form: its status bar, player pickers, Ready button and first cell are designed, and the other cells are loaded from the first one. BoardLayout places them, BoardSprites says what each cell shows, and BoardInput sends what the person does to the view's events.")
Option Explicit
Implements IGameWindow

' BoardLayout measures in points, and a VB6 form in twips.
Private Const TwipsPerPoint As Long = 20
' the cells of a grid, in picCell: grid 1's are 0 to 99, and grid 2's follow.
Private Const CellsPerGrid As Long = GridSize * GridSize
Private Const BackgroundFile As String = "assets\game-form-background.jpg"

' who is told when the person closes the window.
Private listener As IGameWindowEvents
Private paintedColours(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As Long
Private paintedCaptions(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As String
Private paintedSprites(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As String
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

' the background picture is the window's own size, 1088 x 523 pixels at 96 dpi, so it is drawn as it is. Without it, the water colour shows.
Private Sub Form_Load()
    Me.BackColor = board.Water
    LoadBackground
    LoadSprites

    Place lblStatusBar, BoardLayout.StatusBarLeft, BoardLayout.StatusBarTop, BoardLayout.StatusBarWidth, BoardLayout.StatusBarHeight
    Place lblStatus, BoardLayout.StatusTextLeft, BoardLayout.StatusTextTop, BoardLayout.StatusTextWidth, LineHeight

    Dim gridId As PlayGridId
    For gridId = PlayerGrid1 To PlayerGrid2
        AddGrid gridId
        AddPlayerPicker gridId
    Next
    cboPlayer(PickerIndex(PlayerGrid1)).ListIndex = PlayerIndex(Unspecific)
    cboPlayer(PickerIndex(PlayerGrid2)).ListIndex = PlayerIndex(MercilessAI)

    Place cmdReady, BoardLayout.ReadyLeft, BoardLayout.ReadyTop, ReadyWidth, ReadyHeight
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

' the pickers are done with: each grid's caption names its player, in bold so it reads over the ducks.
Private Sub personInput_PlayersConfirmed()
    Dim gridId As PlayGridId
    For gridId = PlayerGrid1 To PlayerGrid2
        With lblPlayer(PickerIndex(gridId))
            .Caption = "Grid " & gridId & ": " & cboPlayer(PickerIndex(gridId)).Text
            .Width = BoardLayout.GridWidth * TwipsPerPoint
            .FontBold = True
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

' a grid's caption, then its picker after it.
Private Sub AddPlayerPicker(ByVal gridId As PlayGridId)
    Place lblPlayer(PickerIndex(gridId)), BoardLayout.GridLeft(gridId), BoardLayout.PlayerTop + 2, PlayerCaptionWidth, LineHeight
    With cboPlayer(PickerIndex(gridId))
        .Left = (BoardLayout.GridLeft(gridId) + PlayerCaptionWidth) * TwipsPerPoint
        .Top = BoardLayout.PlayerTop * TwipsPerPoint
        .Width = (BoardLayout.GridWidth - PlayerCaptionWidth) * TwipsPerPoint
        Dim offered As Variant
        For Each offered In PlayerNames
            .AddItem offered
        Next
    End With
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

Private Sub LoadBackground()
    On Error GoTo NoBackground
    Set Me.Picture = LoadPicture(CurDir$ & "\" & BackgroundFile)
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
