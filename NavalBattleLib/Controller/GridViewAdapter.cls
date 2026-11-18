VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GridViewAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("NavalBattle.Controller")
Option Explicit
Implements IGridViewCommands
Implements IGridViewEvents

Public Event OnCreatePlayer(ByVal gridId As PlayGridId, ByVal pt As PlayerType, ByVal difficulty As AIDifficulty)

Public Event OnPreviewCurrentShipPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
Public Event OnRotateCurrentShipPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
Public Event OnConfirmCurrentShipPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)

Public Event OnPlayerReady()

Public Event OnAttackPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)

Public Event OnViewClosed()

Private Type TAdapter
    ShipsToPosition As Byte
    GridView As IGridViewCommands
End Type
Private this As TAdapter

Private Sub Class_Initialize()
    this.ShipsToPosition = PlayerGrid.ShipsPerGrid
End Sub

Public Function Create(ByVal View As IGridViewCommands) As GridViewAdapter
    With New GridViewAdapter
        Set .GridView = View
        Set View.Events = .Self
        Set Create = .Self
    End With
End Function

Public Property Get Self() As GridViewAdapter
    Set Self = Me
End Property

'@Description("Gets/sets a reference that exposes commands to send to the view.")
Public Property Get GridView() As IGridViewCommands
Attribute GridView.VB_Description = "Gets/sets a reference that exposes commands to send to the view."
    Set GridView = this.GridView
End Property

Public Property Set GridView(ByVal value As IGridViewCommands)
    Set this.GridView = value
End Property


Private Sub Class_Terminate()
    Debug.Print TypeName(Me) & " is terminating"
End Sub

':IGridViewEvents
':Messages sent from the view
':***************************

Private Sub IGridViewEvents_AttackPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    RaiseEvent OnAttackPosition(gridId, position)
End Sub

Private Sub IGridViewEvents_ConfirmShipPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    RaiseEvent OnConfirmCurrentShipPosition(gridId, position)
End Sub

Private Sub IGridViewEvents_CreatePlayer(ByVal gridId As PlayGridId, ByVal pt As PlayerType, ByVal difficulty As AIDifficulty)
    RaiseEvent OnCreatePlayer(gridId, pt, difficulty)
End Sub

Private Sub IGridViewEvents_HumanPlayerReady()
    RaiseEvent OnPlayerReady
End Sub

Private Sub IGridViewEvents_PreviewRotateShip(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    RaiseEvent OnRotateCurrentShipPosition(gridId, position)
End Sub

Private Sub IGridViewEvents_PreviewShipPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    RaiseEvent OnPreviewCurrentShipPosition(gridId, position)
End Sub

Private Sub IGridViewEvents_ViewClosed()
    RaiseEvent OnViewClosed
End Sub


':IGridViewCommands
':Messages sent from the controller
':*********************************

Private Property Set IGridViewCommands_Events(ByVal value As IGridViewEvents)
    Err.Raise 5, TypeName(Me), "Invalid use of property"
End Property

Private Property Get IGridViewCommands_Events() As IGridViewEvents
    Set IGridViewCommands_Events = Me
End Property

' the person fires at the other grid now: their own grid is locked, and the other one unlocked.
Private Sub IGridViewCommands_OnBeginAttack(ByVal currentPlayerGridId As PlayGridId)
    this.GridView.OnLockGrid currentPlayerGridId
    this.GridView.OnUnlockGrid OtherGrid(currentPlayerGridId)
    this.GridView.OnBeginAttack currentPlayerGridId
End Sub

' the person places a ship on their own grid: it is unlocked, and the other one locked.
Private Sub IGridViewCommands_OnBeginShipPosition(ByVal CurrentShip As IShip, ByVal player As IPlayer)
    this.GridView.OnLockGrid OtherGrid(player.PlayGrid.gridId)
    this.GridView.OnUnlockGrid player.PlayGrid.gridId
    this.GridView.OnBeginShipPosition CurrentShip, player
End Sub

Private Function OtherGrid(ByVal gridId As PlayGridId) As PlayGridId
    OtherGrid = IIf(gridId = PlayerGrid1, PlayerGrid2, PlayerGrid1)
End Function

Private Sub IGridViewCommands_OnConfirmShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
    If player.PlayerType = ComputerControlled Then Exit Sub
    this.GridView.OnConfirmShipPosition player, newShip
    this.ShipsToPosition = this.ShipsToPosition - 1
End Sub

Private Sub IGridViewCommands_OnGameOver(ByVal winningGrid As PlayGridId, ByVal position As IGridCoord, ByVal sunkShip As IShip)
    this.GridView.OnLockGrid PlayerGrid1
    this.GridView.OnLockGrid PlayerGrid2
    this.GridView.OnGameOver winningGrid, position, sunkShip
    Set this.GridView.Events = Nothing
End Sub

Private Sub IGridViewCommands_OnHit(ByVal gridId As PlayGridId, ByVal position As IGridCoord, ByVal hitShip As IShip)
    this.GridView.OnHit gridId, position, hitShip
End Sub

Private Sub IGridViewCommands_OnInvalidShipPosition()
    this.GridView.OnInvalidShipPosition
End Sub

Private Sub IGridViewCommands_OnKnownPositionAttack()
    this.GridView.OnKnownPositionAttack
End Sub

Private Sub IGridViewCommands_OnLockGrid(ByVal gridId As PlayGridId)
    this.GridView.OnLockGrid gridId
End Sub

Private Sub IGridViewCommands_OnUnlockGrid(ByVal gridId As PlayGridId)
    this.GridView.OnUnlockGrid gridId
End Sub

Private Sub IGridViewCommands_OnMiss(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    this.GridView.OnMiss gridId, position
End Sub

Private Sub IGridViewCommands_OnNewGame()
    this.GridView.OnLockGrid PlayerGrid1
    this.GridView.OnLockGrid PlayerGrid2
    this.GridView.OnNewGame
End Sub

Private Sub IGridViewCommands_OnPreviewShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
    If player.PlayerType = ComputerControlled Then Exit Sub
    this.GridView.OnPreviewShipPosition player, newShip
End Sub

Private Sub IGridViewCommands_OnRefreshGrid(ByVal grid As PlayerGrid)
    this.GridView.OnRefreshGrid grid
End Sub

Private Sub IGridViewCommands_OnSink(ByVal gridId As PlayGridId, ByVal position As IGridCoord, ByVal sunkShip As IShip)
    this.GridView.OnSink gridId, position, sunkShip
End Sub

Private Sub IGridViewCommands_OnBeginWaitForComputerPlayer()
    this.GridView.OnBeginWaitForComputerPlayer
End Sub

Private Sub IGridViewCommands_OnEndWaitForComputerPlayer()
    this.GridView.OnEndWaitForComputerPlayer
End Sub


