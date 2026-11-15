VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AIPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An implementation of IPlayer that is AI-controlled."
'@Folder("NavalBattle.Model.Players")
'@ModuleDescription("An implementation of IPlayer that is AI-controlled.")
'@PredeclaredId
Option Explicit
Implements IPlayer

#If Win32 Or Win64 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Type TPlayer
    PlayerType As PlayerType
    PlayGrid As PlayerGrid
    Strategy As IGameStrategy
    PauseMilliseconds As Long
End Type

Private this As TPlayer

'@Description("Creates a computer player. It waits the given number of milliseconds before each shot, so a person can watch it play; 0 plays at once.")
Public Function Create(ByVal grid As PlayerGrid, ByVal GameStrategy As IGameStrategy, ByVal pauseMilliseconds As Long) As IPlayer
Attribute Create.VB_Description = "Creates a computer player. It waits the given number of milliseconds before each shot, so a person can watch it play; 0 plays at once."
    With New AIPlayer
        .PlayerType = ComputerControlled
        Set .Strategy = GameStrategy
        Set .PlayGrid = grid
        .PauseMilliseconds = pauseMilliseconds
        Set Create = .Self
    End With
End Function

Public Property Get PauseMilliseconds() As Long
    PauseMilliseconds = this.PauseMilliseconds
End Property

Public Property Let PauseMilliseconds(ByVal value As Long)
    this.PauseMilliseconds = value
End Property

Public Property Get Self() As IPlayer
    Set Self = Me
End Property

Public Property Get Strategy() As IGameStrategy
    Set Strategy = this.Strategy
End Property

Public Property Set Strategy(ByVal value As IGameStrategy)
    Set this.Strategy = value
End Property

Public Property Get PlayGrid() As PlayerGrid
    Set PlayGrid = this.PlayGrid
End Property

Public Property Set PlayGrid(ByVal value As PlayerGrid)
    Set this.PlayGrid = value
End Property

Public Property Get PlayerType() As PlayerType
    PlayerType = this.PlayerType
End Property

Public Property Let PlayerType(ByVal value As PlayerType)
    this.PlayerType = value
End Property

Private Property Get IPlayer_PlayGrid() As PlayerGrid
    Set IPlayer_PlayGrid = this.PlayGrid
End Property

Private Sub IPlayer_PlaceShip(ByVal CurrentShip As IShip)
    this.Strategy.PlaceShip this.PlayGrid, CurrentShip
End Sub

Private Function IPlayer_Play(ByVal enemyGrid As PlayerGrid) As IGridCoord
    
    If this.PauseMilliseconds > 0 Then Pause this.PauseMilliseconds
    Set IPlayer_Play = this.Strategy.Play(enemyGrid)
    
    Static shots As Long
    shots = shots + 1
    Debug.Print "AI Player " & this.PlayGrid.gridId & "(" & TypeName(this.Strategy) & ") has played " & shots & " turns"
    
End Function

Private Property Get IPlayer_PlayerType() As PlayerType
    IPlayer_PlayerType = this.PlayerType
End Property

'@Description("Waits the given number of milliseconds, so a person can watch the AI play.")
Private Sub Pause(ByVal milliseconds As Long)
Attribute Pause.VB_Description = "Waits the given number of milliseconds, so a person can watch the AI play."
#If Win32 Or Win64 Then
    Sleep milliseconds
#Else
    Dim startTime As Single
    startTime = Timer
    Dim elapsed As Single
    Do
        DoEvents
        elapsed = Timer - startTime
        If elapsed < 0 Then elapsed = elapsed + 86400 ' Timer went past midnight
    Loop While elapsed * 1000 < milliseconds
#End If
End Sub
