VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GameForm
   Caption         =   "Naval battle"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "The game's window. It has no designed controls: it makes its grids, player pickers, fleet panels, status line and Ready button when it is loaded."
'@Folder("NavalBattle.View")
'@ModuleDescription("The game's window. It has no designed controls: it makes its grids, player pickers, fleet panels, status line and Ready button when it is loaded.")
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
Private Const ReadyTop As Long = StatusTop + LineHeight + 6
Private Const ReadyWidth As Long = 72
Private Const ReadyHeight As Long = 24

Private Const WaterColour As Long = &HF5E1C8
Private Const HoverColour As Long = &HA0F0FF
Private Const LineColour As Long = &H808080

Private cells(PlayerGrid1 To PlayerGrid2, 1 To GridSize, 1 To GridSize) As GridCell
Private hoveredCell As GridCell
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

    Dim formWidth As Long
    formWidth = GridLeft(PlayerGrid2) + GridWidth + Margin

    Set statusLine = AddLabel("Status", Margin, StatusTop, formWidth - 2 * Margin, LineHeight, "Pick the players, then click Ready.")

    Dim gridId As PlayGridId
    For gridId = PlayerGrid1 To PlayerGrid2
        AddGrid gridId
        AddFleetPanel gridId
    Next

    Set firstPlayerPicker = AddPlayerPicker(PlayerGrid1, "Human")
    Set secondPlayerPicker = AddPlayerPicker(PlayerGrid2, "Merciless AI")

    Set readyButton = Me.Controls.Add("Forms.CommandButton.1", "Ready", True)
    With readyButton
        .Caption = "Ready"
        .Left = formWidth - Margin - ReadyWidth
        .Top = ReadyTop
        .Width = ReadyWidth
        .Height = ReadyHeight
    End With

    Me.Width = formWidth + (Me.Width - Me.InsideWidth)
    Me.Height = ReadyTop + ReadyHeight + Margin + (Me.Height - Me.InsideHeight)
End Sub

'@Description("Sets the text of the status line.")
Public Sub SetStatus(ByVal text As String)
Attribute SetStatus.VB_Description = "Sets the text of the status line."
    statusLine.Caption = text
End Sub

Private Sub cellEvents_Clicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    SetStatus "Grid " & gridId & ", " & CellName(position) & ": clicked."
End Sub

Private Sub cellEvents_RightClicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    SetStatus "Grid " & gridId & ", " & CellName(position) & ": right-clicked."
End Sub

Private Sub cellEvents_Hovered(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    Dim cell As GridCell
    Set cell = cells(gridId, position.X, position.Y)
    If cell Is hoveredCell Then Exit Sub

    If Not hoveredCell Is Nothing Then hoveredCell.Control.BackColor = WaterColour
    cell.Control.BackColor = HoverColour
    Set hoveredCell = cell
End Sub

Private Sub firstPlayerPicker_Change()
    SetStatus "Grid 1 is played by: " & firstPlayerPicker.Value & "."
End Sub

Private Sub secondPlayerPicker_Change()
    SetStatus "Grid 2 is played by: " & secondPlayerPicker.Value & "."
End Sub

Private Sub readyButton_Click()
    SetStatus "Ready."
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set hoveredCell = Nothing
    Erase cells
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
            Set cellLabel = AddLabel("Grid" & gridId & "Cell" & CellName(position), leftEdge + X * CellSize, GridTop + Y * CellSize, CellSize, CellSize, vbNullString)
            cellLabel.BackColor = WaterColour
            cellLabel.BorderStyle = fmBorderStyleSingle
            cellLabel.BorderColor = LineColour

            Set cells(gridId, X, Y) = GridCell.Create(cellLabel, gridId, position, cellEvents)
        Next
    Next
End Sub

Private Sub AddFleetPanel(ByVal gridId As PlayGridId)
    AddLabel "Grid" & gridId & "Fleet", GridLeft(gridId), FleetTop, GridWidth, FleetHeight, "Fleet: no ship placed yet."
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
