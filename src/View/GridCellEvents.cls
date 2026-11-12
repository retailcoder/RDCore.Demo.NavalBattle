VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GridCellEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Raises the events of every cell of the game form's grids, so that one WithEvents variable handles them all."
'@Folder("NavalBattle.View")
'@ModuleDescription("Raises the events of every cell of the game form's grids, so that one WithEvents variable handles them all.")
Option Explicit

'@Description("A cell was clicked.")
Public Event Clicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
Attribute Clicked.VB_Description = "A cell was clicked."

'@Description("A cell was clicked with the right mouse button.")
Public Event RightClicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
Attribute RightClicked.VB_Description = "A cell was clicked with the right mouse button."

'@Description("The mouse moved over a cell.")
Public Event Hovered(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
Attribute Hovered.VB_Description = "The mouse moved over a cell."

Friend Sub OnClicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    RaiseEvent Clicked(gridId, position)
End Sub

Friend Sub OnRightClicked(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    RaiseEvent RightClicked(gridId, position)
End Sub

Friend Sub OnHovered(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    RaiseEvent Hovered(gridId, position)
End Sub
