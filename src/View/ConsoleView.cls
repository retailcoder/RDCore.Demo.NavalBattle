VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An implementation of IGridViewCommands that plays the game in the console: it prints the grids as text, and asks a person for their moves with InputBox."
'@Folder("NavalBattle.View")
'@ModuleDescription("An implementation of IGridViewCommands that plays the game in the console: it prints the grids as text, and asks a person for their moves with InputBox.")
Option Explicit
Implements IGridViewCommands

Private Const Title As String = "Naval battle"

Private Enum ConsolePhase
    PhaseSetup
    PhaseShipPlacement
    PhaseAttack
    PhaseOver
End Enum

Private Type TView
    ViewEvents As IGridViewEvents
    Grids(PlayerGrid1 To PlayerGrid2) As PlayerGrid
    Phase As ConsolePhase
    HumanGridId As PlayGridId
    HumanIsReady As Boolean
    PreviewShip As IShip
    LastShot As String
    LastHitShip As String
End Type

Private this As TView

'@Description("Plays a game: asks who plays each grid when the argument is empty, then asks the person, if one plays, for their ships and shots until the game is over. A player is Human, Random, Fair-play or Merciless (the first letter is enough). An empty answer ends the game.")
Public Sub Play(Optional ByVal firstPlayer As String, Optional ByVal secondPlayer As String)
Attribute Play.VB_Description = "Plays a game: asks who plays each grid when the argument is empty, then asks the person, if one plays, for their ships and shots until the game is over. A player is Human, Random, Fair-play or Merciless (the first letter is enough). An empty answer ends the game."
    Dim firstDifficulty As AIDifficulty
    If Not TryGetPlayer(PlayerGrid1, firstPlayer, False, firstDifficulty) Then Exit Sub

    Dim secondDifficulty As AIDifficulty
    If Not TryGetPlayer(PlayerGrid2, secondPlayer, firstDifficulty = Unspecific, secondDifficulty) Then Exit Sub

    this.ViewEvents.CreatePlayer PlayerGrid1, PlayerTypeOf(firstDifficulty), firstDifficulty
    this.ViewEvents.CreatePlayer PlayerGrid2, PlayerTypeOf(secondDifficulty), secondDifficulty

    Do
        Select Case this.Phase
            Case PhaseShipPlacement
                If Not TryPlaceShip Then Exit Do
            Case PhaseAttack
                If Not TryAttack Then Exit Do
            Case Else
                Exit Do
        End Select
    Loop
End Sub

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
    this.Phase = PhaseShipPlacement
    this.HumanGridId = player.PlayGrid.gridId
    Set this.Grids(this.HumanGridId) = player.PlayGrid
    Set this.PreviewShip = CurrentShip
End Sub

Private Sub IGridViewCommands_OnPreviewShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
    Set this.PreviewShip = newShip
End Sub

Private Sub IGridViewCommands_OnConfirmShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
    Debug.Print "Grid " & player.PlayGrid.gridId & " places its " & newShip.Name & "."
End Sub

Private Sub IGridViewCommands_OnInvalidShipPosition()
    Debug.Print "A ship cannot go there."
End Sub

Private Sub IGridViewCommands_OnBeginAttack(ByVal currentPlayerGridId As PlayGridId)
    this.Phase = PhaseAttack
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
    this.Phase = PhaseOver
    PrintGrids True, True
    Debug.Print "Game over: grid " & winningGrid & " wins."
End Sub

'@Description("Gets the player for a grid from the argument, or asks for it when the argument is empty or not a player. Returns False when the answer is empty.")
Private Function TryGetPlayer(ByVal gridId As PlayGridId, ByVal answer As String, ByVal humanIsTaken As Boolean, ByRef outDifficulty As AIDifficulty) As Boolean
Attribute TryGetPlayer.VB_Description = "Gets the player for a grid from the argument, or asks for it when the argument is empty or not a player. Returns False when the answer is empty."
    Dim isPlayer As Boolean
    Do
        If answer = vbNullString Then
            answer = Ask("Who plays grid " & gridId & "? Human, Random, Fair-play or Merciless (the first letter is enough).")
            If answer = vbNullString Then
                TryGetPlayer = False
                Exit Function
            End If
        End If

        isPlayer = True
        Select Case UCase$(Left$(answer, 1))
            Case "H"
                If humanIsTaken Then
                    Debug.Print "Only one grid can be played by a person: pick an AI for grid " & gridId & "."
                    isPlayer = False
                Else
                    outDifficulty = Unspecific
                End If
            Case "R"
                outDifficulty = RandomAI
            Case "F"
                outDifficulty = FairplayAI
            Case "M"
                outDifficulty = MercilessAI
            Case Else
                Debug.Print "'" & answer & "' is not a player."
                isPlayer = False
        End Select
        answer = vbNullString
    Loop Until isPlayer
    TryGetPlayer = True
End Function

Private Function PlayerTypeOf(ByVal difficulty As AIDifficulty) As PlayerType
    PlayerTypeOf = IIf(difficulty = Unspecific, HumanControlled, ComputerControlled)
End Function

'@Description("Asks the person where the next ship goes, and places it there. Returns False when the answer is empty.")
Private Function TryPlaceShip() As Boolean
Attribute TryPlaceShip.VB_Description = "Asks the person where the next ship goes, and places it there. Returns False when the answer is empty."
    PrintGrids this.HumanGridId = PlayerGrid1, this.HumanGridId = PlayerGrid2

    Dim answer As String
    answer = Ask("Grid " & this.HumanGridId & ": where does your " & this.PreviewShip.Name & " (" & this.PreviewShip.Size & " cells) go? " _
        & "Its first cell, like B3, goes across the grid; add V to go down: B3 V.")
    If answer = vbNullString Then
        Abandon
        TryPlaceShip = False
        Exit Function
    End If
    TryPlaceShip = True

    Dim position As IGridCoord
    Dim direction As ShipOrientation
    If Not TryParsePlacement(answer, position, direction) Then
        Debug.Print "'" & answer & "' is not a cell."
        Exit Function
    End If

    this.ViewEvents.PreviewShipPosition this.HumanGridId, position
    If this.PreviewShip.Orientation <> direction Then this.ViewEvents.PreviewRotateShip this.HumanGridId, position

    With this.PreviewShip
        If .GridPosition.X <> position.X Or .GridPosition.Y <> position.Y Or .Orientation <> direction Then
            Debug.Print "The " & .Name & " does not fit there."
            Exit Function
        End If
    End With
    this.ViewEvents.ConfirmShipPosition this.HumanGridId, position
End Function

'@Description("Asks the person for their next shot, and fires it. Returns False when the answer is empty.")
Private Function TryAttack() As Boolean
Attribute TryAttack.VB_Description = "Asks the person for their next shot, and fires it. Returns False when the answer is empty."
    If Not this.HumanIsReady Then
        this.HumanIsReady = True
        this.ViewEvents.HumanPlayerReady
    End If
    PrintGrids this.HumanGridId = PlayerGrid1, this.HumanGridId = PlayerGrid2

    Dim targetGridId As PlayGridId
    targetGridId = IIf(this.HumanGridId = PlayerGrid1, PlayerGrid2, PlayerGrid1)

    Dim answer As String
    answer = Ask("Your shot at grid " & targetGridId & ": a cell, like B3.")
    If answer = vbNullString Then
        Abandon
        TryAttack = False
        Exit Function
    End If
    TryAttack = True

    Dim position As IGridCoord
    If Not TryParseCell(answer, position) Then
        Debug.Print "'" & answer & "' is not a cell."
        Exit Function
    End If
    this.ViewEvents.AttackPosition targetGridId, position
End Function

Private Sub Abandon()
    this.Phase = PhaseOver
    Debug.Print "Game abandoned."
End Sub

Private Function Ask(ByVal prompt As String) As String
    Ask = Trim$(InputBox(prompt, Title))
End Function

'@Description("Reads a ship's first cell and direction, like B3 (across) or B3 V (down).")
Private Function TryParsePlacement(ByVal text As String, ByRef outPosition As IGridCoord, ByRef outDirection As ShipOrientation) As Boolean
Attribute TryParsePlacement.VB_Description = "Reads a ship's first cell and direction, like B3 (across) or B3 V (down)."
    text = UCase$(Replace(text, " ", vbNullString))
    outDirection = Horizontal
    Select Case Right$(text, 1)
        Case "V"
            outDirection = Vertical
            text = Left$(text, Len(text) - 1)
        Case "H"
            text = Left$(text, Len(text) - 1)
    End Select
    TryParsePlacement = TryParseCell(text, outPosition)
End Function

'@Description("Reads a cell, like B3: a column letter, then a row number.")
Private Function TryParseCell(ByVal text As String, ByRef outPosition As IGridCoord) As Boolean
Attribute TryParseCell.VB_Description = "Reads a cell, like B3: a column letter, then a row number."
    TryParseCell = False
    text = UCase$(Replace(text, " ", vbNullString))
    If Len(text) < 2 Then Exit Function

    Dim column As Long
    column = Asc(text) - 64
    If column < 1 Or column > PlayerGrid.Size Then Exit Function

    Dim row As String
    row = Mid$(text, 2)
    If Not (row Like "#" Or row Like "##") Then Exit Function
    If CLng(row) < 1 Or CLng(row) > PlayerGrid.Size Then Exit Function

    Set outPosition = GridCoord.Create(column, CLng(row))
    TryParseCell = True
End Function

'@Description("Prints both grids side by side: # is a ship (when it is shown), X a hit, o a miss.")
Private Sub PrintGrids(ByVal showFirstShips As Boolean, ByVal showSecondShips As Boolean)
Attribute PrintGrids.VB_Description = "Prints both grids side by side: # is a ship (when it is shown), X a hit, o a miss."
    Const Gap As String = "      "
    Dim gridWidth As Long
    gridWidth = 4 + 2 * PlayerGrid.Size

    Debug.Print
    Debug.Print Left$(GridTitle(PlayerGrid1) & Space$(gridWidth), gridWidth) & Gap & GridTitle(PlayerGrid2)

    Dim header As String
    header = "    "
    Dim currentX As Long
    For currentX = 1 To PlayerGrid.Size
        header = header & Chr$(64 + currentX) & " "
    Next
    Debug.Print header & Gap & RTrim$(header)

    Dim currentY As Long
    For currentY = 1 To PlayerGrid.Size
        Debug.Print RowText(PlayerGrid1, currentY, showFirstShips) & Gap & RTrim$(RowText(PlayerGrid2, currentY, showSecondShips))
    Next
    Debug.Print
End Sub

Private Function GridTitle(ByVal gridId As PlayGridId) As String
    Dim result As String
    result = "Grid " & gridId
    If this.Phase <> PhaseOver And this.HumanGridId = gridId Then result = result & " (yours)"
    GridTitle = result
End Function

Private Function RowText(ByVal gridId As PlayGridId, ByVal currentY As Long, ByVal showShips As Boolean) As String
    Dim result As String
    result = Right$(" " & currentY, 2) & "  "

    Dim currentX As Long
    For currentX = 1 To PlayerGrid.Size
        Dim value As GridState
        value = GridState.Unknown
        If Not this.Grids(gridId) Is Nothing Then value = this.Grids(gridId).State(GridCoord.Create(currentX, currentY))
        result = result & CellText(value, showShips) & " "
    Next
    RowText = result
End Function

Private Function CellText(ByVal value As GridState, ByVal showShips As Boolean) As String
    Select Case value
        Case GridState.ShipPosition
            CellText = IIf(showShips, "#", ".")
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
