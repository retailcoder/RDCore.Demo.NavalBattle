VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FailingWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A window for the WindowView tests: it shows nothing, and its Show raises the given error, before or after it tells its listener that it opened."
'@Folder("NavalBattle.Tests.Doubles")
'@ModuleDescription("A window for the WindowView tests: it shows nothing, and its Show raises the given error, before or after it tells its listener that it opened.")
Option Explicit
Implements IGameWindow

Private Type TWindow
    Listener As IGameWindowEvents
    ErrorNumber As Long
    OpensFirst As Boolean
End Type

Private this As TWindow

'@Description("The error Show raises, and whether the window opens before it raises it.")
Public Sub FailsWith(ByVal errorNumber As Long, ByVal opensFirst As Boolean)
Attribute FailsWith.VB_Description = "The error Show raises, and whether the window opens before it raises it."
    this.ErrorNumber = errorNumber
    this.OpensFirst = opensFirst
End Sub

Private Property Set IGameWindow_Events(ByVal value As IGridViewEvents)
End Property

Private Property Set IGameWindow_Listener(ByVal value As IGameWindowEvents)
    Set this.Listener = value
End Property

Private Sub IGameWindow_Show()
    If this.OpensFirst Then this.Listener.Opened
    Err.Raise this.ErrorNumber, "FailingWindow", "The window failed."
End Sub

Private Property Get IGameWindow_WantsNewGame() As Boolean
End Property

Private Sub IGameWindow_PickPlayer(ByVal gridId As PlayGridId, ByVal player As String)
End Sub

Private Sub IGameWindow_SetStatus(ByVal text As String)
End Sub

Private Sub IGameWindow_BeginShipPlacement(ByVal gridId As PlayGridId)
End Sub

Private Sub IGameWindow_BeginWaitForReady()
End Sub

Private Sub IGameWindow_BeginComputerTurn()
End Sub

Private Sub IGameWindow_EndComputerTurn()
End Sub

Private Sub IGameWindow_EndGame()
End Sub

Private Sub IGameWindow_PaintGrid(ByVal grid As PlayerGrid, ByVal showShips As Boolean)
End Sub

Private Sub IGameWindow_PaintShipPreview(ByVal grid As PlayerGrid, ByVal previewShip As IShip)
End Sub

Private Sub IGameWindow_FinishSinking(ByVal gridId As PlayGridId, ByVal sunkShip As IShip)
End Sub

Private Sub IGameWindow_DizzyShip(ByVal gridId As PlayGridId, ByVal hitShip As IShip)
End Sub

Private Sub IGameWindow_DizzyShipRecovers()
End Sub

Private Sub IGameWindow_Squeak(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
End Sub

Private Sub IGameWindow_SqueakStops()
End Sub

Private Sub IGameWindow_Pause(ByVal seconds As Single)
End Sub

Private Sub IGameWindow_LockGrid(ByVal gridId As PlayGridId, ByVal locked As Boolean)
End Sub
