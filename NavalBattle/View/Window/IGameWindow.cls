VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A window the game is played in: its two grids, a player picker under each, a status line and a Ready button. WindowView paints it, and what the person does in it is sent to the view's events."
'@Folder("NavalBattle.View.Window")
'@ModuleDescription("A window the game is played in: its two grids, a player picker under each, a status line and a Ready button. WindowView paints it, and what the person does in it is sent to the view's events.")
'@Interface
Option Explicit

'@Description("Where what the person does in the window is sent: the adapter's events.")
Public Property Set Events(ByVal value As IGridViewEvents)
Attribute Events.VB_Description = "Where what the person does in the window is sent: the adapter's events."
End Property

'@Description("Who is told when the person closes the window.")
Public Property Set Listener(ByVal value As IGameWindowEvents)
Attribute Listener.VB_Description = "Who is told when the person closes the window."
End Property

'@Description("Shows the window, and returns once it is closed.")
Public Sub Show()
Attribute Show.VB_Description = "Shows the window, and returns once it is closed."
End Sub

'@Description("Picks the player of a grid: Human, Random, Fair-play or Merciless; the first letter is enough. Anything else keeps the pick as it is.")
Public Sub PickPlayer(ByVal gridId As PlayGridId, ByVal player As String)
Attribute PickPlayer.VB_Description = "Picks the player of a grid: Human, Random, Fair-play or Merciless; the first letter is enough. Anything else keeps the pick as it is."
End Sub

'@Description("Sets the text of the status line.")
Public Sub SetStatus(ByVal text As String)
Attribute SetStatus.VB_Description = "Sets the text of the status line."
End Sub

'@Description("The person places their ships on the given grid: pointing at a cell moves the ship there, a right-click turns it, and a click places it.")
Public Sub BeginShipPlacement(ByVal gridId As PlayGridId)
Attribute BeginShipPlacement.VB_Description = "The person places their ships on the given grid: pointing at a cell moves the ship there, a right-click turns it, and a click places it."
End Sub

'@Description("The person's ships are placed: the next click of Ready starts the battle.")
Public Sub BeginWaitForReady()
Attribute BeginWaitForReady.VB_Description = "The person's ships are placed: the next click of Ready starts the battle."
End Sub

'@Description("The computer plays: the person's clicks are ignored until it is done. The cells that changed are drawn first, so the last shot shows while the computer thinks.")
Public Sub BeginComputerTurn()
Attribute BeginComputerTurn.VB_Description = "The computer plays: the person's clicks are ignored until it is done. The cells that changed are drawn first, so the last shot shows while the computer thinks."
End Sub

'@Description("The computer has played. When it plays both grids, the cells that changed are drawn, so that the game can be watched.")
Public Sub EndComputerTurn()
Attribute EndComputerTurn.VB_Description = "The computer has played. When it plays both grids, the cells that changed are drawn, so that the game can be watched."
End Sub

'@Description("The game is over: the grids do nothing more, and the Ready button is gone.")
Public Sub EndGame()
Attribute EndGame.VB_Description = "The game is over: the grids do nothing more, and the Ready button is gone."
End Sub

'@Description("Paints a grid as it stands: hits, misses, water, and the ships when they are to be shown. A sunk ship shows its anchor once its sinking is done, and the dizzy duck until then; the big duck of a ship that was just hit is dizzy too.")
Public Sub PaintGrid(ByVal grid As PlayerGrid, ByVal showShips As Boolean)
Attribute PaintGrid.VB_Description = "Paints a grid as it stands: hits, misses, water, and the ships when they are to be shown. A sunk ship shows its anchor once its sinking is done, and the dizzy duck until then; the big duck of a ship that was just hit is dizzy too."
End Sub

'@Description("Paints where a ship would go, over the grid: green where it can go, red where it cannot.")
Public Sub PaintShipPreview(ByVal grid As PlayerGrid, ByVal previewShip As IShip)
Attribute PaintShipPreview.VB_Description = "Paints where a ship would go, over the grid: green where it can go, red where it cannot."
End Sub

'@Description("A ship's sinking is done: from now on, its bow shows the anchor.")
Public Sub FinishSinking(ByVal gridId As PlayGridId, ByVal sunkShip As IShip)
Attribute FinishSinking.VB_Description = "A ship's sinking is done: from now on, its bow shows the anchor."
End Sub

'@Description("A ship was hit: its big duck is dizzy, until DizzyShipRecovers. Nothing shows where the ships are not shown.")
Public Sub DizzyShip(ByVal gridId As PlayGridId, ByVal hitShip As IShip)
Attribute DizzyShip.VB_Description = "A ship was hit: its big duck is dizzy, until DizzyShipRecovers. Nothing shows where the ships are not shown."
End Sub

'@Description("The dizzy big duck is itself again.")
Public Sub DizzyShipRecovers()
Attribute DizzyShipRecovers.VB_Description = "The dizzy big duck is itself again."
End Sub

'@Description("A cell of a grid was just hit: it squeaks, until SqueakStops, whether the ships are shown or not.")
Public Sub Squeak(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
Attribute Squeak.VB_Description = "A cell of a grid was just hit: it squeaks, until SqueakStops, whether the ships are shown or not."
End Sub

'@Description("The cell that was just hit shows its hit again.")
Public Sub SqueakStops()
Attribute SqueakStops.VB_Description = "The cell that was just hit shows its hit again."
End Sub

'@Description("Waits the given number of seconds, while the window keeps drawing and answering.")
Public Sub Pause(ByVal seconds As Single)
Attribute Pause.VB_Description = "Waits the given number of seconds, while the window keeps drawing and answering."
End Sub

'@Description("Locks or unlocks a grid: a locked grid's cells do nothing when they are pointed at or clicked.")
Public Sub LockGrid(ByVal gridId As PlayGridId, ByVal locked As Boolean)
Attribute LockGrid.VB_Description = "Locks or unlocks a grid: a locked grid's cells do nothing when they are pointed at or clicked."
End Sub
