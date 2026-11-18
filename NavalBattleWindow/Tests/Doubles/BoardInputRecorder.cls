VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoardInputRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A window for the BoardInput tests: it shows nothing, and records, in order, what the input sends to the view's events and what it tells the window."
'@Folder("NavalBattle.Tests.Doubles")
'@ModuleDescription("A window for the BoardInput tests: it shows nothing, and records, in order, what the input sends to the view's events and what it tells the window.")
Option Explicit
Implements IGridViewEvents

Private Type TRecorder
    Calls As Collection
    Status As String
    ReadyVisible As Boolean
    PlayersConfirmed As Boolean
End Type

Private this As TRecorder
Private WithEvents watched As BoardInput
Attribute watched.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set this.Calls = New Collection
End Sub

'@Description("Records what the given input sends and says: the recorder is the input's events from now on.")
Public Sub Watch(ByVal source As BoardInput)
Attribute Watch.VB_Description = "Records what the given input sends and says: the recorder is the input's events from now on."
    Set watched = source
    Set source.Events = Me
End Sub

'@Description("Gets what the input sent to the view's events, in order: each call's name, then its grid and position, as in ""PreviewShipPosition 1 (3,4)"".")
Public Property Get Calls() As Collection
Attribute Calls.VB_Description = "Gets what the input sent to the view's events, in order: each call's name, then its grid and position, as in ""PreviewShipPosition 1 (3,4)""."
    Set Calls = this.Calls
End Property

'@Description("Gets the status text the input last said, or nothing.")
Public Property Get Status() As String
Attribute Status.VB_Description = "Gets the status text the input last said, or nothing."
    Status = this.Status
End Property

'@Description("Gets whether the input last showed the Ready button.")
Public Property Get ReadyVisible() As Boolean
Attribute ReadyVisible.VB_Description = "Gets whether the input last showed the Ready button."
    ReadyVisible = this.ReadyVisible
End Property

'@Description("Gets whether the input confirmed the players.")
Public Property Get PlayersConfirmed() As Boolean
Attribute PlayersConfirmed.VB_Description = "Gets whether the input confirmed the players."
    PlayersConfirmed = this.PlayersConfirmed
End Property

Private Sub watched_StatusChanged(ByVal text As String)
    this.Status = text
End Sub

Private Sub watched_ReadyShown(ByVal visible As Boolean)
    this.ReadyVisible = visible
End Sub

Private Sub watched_PlayersConfirmed()
    this.PlayersConfirmed = True
End Sub

Private Sub Record(ByVal callName As String, ByVal gridId As PlayGridId, Optional ByVal position As IGridCoord)
    If position Is Nothing Then
        this.Calls.Add callName & " " & gridId
    Else
        this.Calls.Add callName & " " & gridId & " (" & position.X & "," & position.Y & ")"
    End If
End Sub

Private Sub IGridViewEvents_CreatePlayer(ByVal gridId As PlayGridId, ByVal pt As PlayerType, ByVal difficulty As AIDifficulty)
    this.Calls.Add "CreatePlayer " & gridId & " " & PlayerName(difficulty)
End Sub

Private Sub IGridViewEvents_PreviewRotateShip(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    Record "PreviewRotateShip", gridId, position
End Sub

Private Sub IGridViewEvents_PreviewShipPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    Record "PreviewShipPosition", gridId, position
End Sub

Private Sub IGridViewEvents_ConfirmShipPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    Record "ConfirmShipPosition", gridId, position
End Sub

Private Sub IGridViewEvents_AttackPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    Record "AttackPosition", gridId, position
End Sub

Private Sub IGridViewEvents_HumanPlayerReady()
    this.Calls.Add "HumanPlayerReady"
End Sub

Private Sub IGridViewEvents_ViewClosed()
    this.Calls.Add "ViewClosed"
End Sub
