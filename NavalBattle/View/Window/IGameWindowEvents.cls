VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGameWindowEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "What a game window tells the view that shows it. An interface cannot declare events, so the window calls its listener instead."
'@Folder("NavalBattle.View.Window")
'@ModuleDescription("What a game window tells the view that shows it. An interface cannot declare events, so the window calls its listener instead.")
'@Interface
Option Explicit

'@Description("The person closed the window. The window lets go of its listener once it has told it.")
Public Sub Closed()
Attribute Closed.VB_Description = "The person closed the window. The window lets go of its listener once it has told it."
End Sub
