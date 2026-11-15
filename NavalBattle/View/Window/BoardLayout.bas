Attribute VB_Name = "BoardLayout"
Attribute VB_Description = "Where the parts of a game window go, in points, from the top: the status bar, the two grids side by side, each grid's player under it, then the Ready button. Every game window lays itself out by it; a window that measures in twips multiplies by 20."
'@Folder("NavalBattle.View.Window")
'@ModuleDescription("Where the parts of a game window go, in points, from the top: the status bar, the two grids side by side, each grid's player under it, then the Ready button. Every game window lays itself out by it; a window that measures in twips multiplies by 20.")
Option Explicit

' the number of rows and of columns of a grid.
Public Const GridSize As Long = 10
' the side of a grid's cell.
Public Const CellSize As Long = 18
' the window's client area: 16320 x 7845 twips, 1088 x 523 pixels at 96 dpi.
Public Const WindowWidth As Long = 816
Public Const WindowHeight As Long = 392
' the height of a line of text: a player's line, and the status text.
Public Const LineHeight As Long = 22
' the space between the window's edges and what it shows.
Public Const Margin As Long = 12
' how far the status bar starts past the window's top, left and right edges, so that only its bottom border shows.
Public Const StatusBleed As Long = 1
' the space between the status bar's edges and its text.
Public Const StatusPadding As Long = 8
' the width of the "Grid n:" caption before a grid's player picker.
Public Const PlayerCaptionWidth As Long = 56
Public Const ReadyWidth As Long = 80
Public Const ReadyHeight As Long = 28

' the space between the two grids.
Private Const GridGap As Long = 48
' the space between a grid's column letters and its player's line, and between that line and the Ready button.
Private Const LineGap As Long = 8

'@Description("The status bar's left edge: past the window's.")
Public Function StatusBarLeft() As Long
Attribute StatusBarLeft.VB_Description = "The status bar's left edge: past the window's."
    StatusBarLeft = -StatusBleed
End Function

'@Description("The status bar's top edge: past the window's.")
Public Function StatusBarTop() As Long
Attribute StatusBarTop.VB_Description = "The status bar's top edge: past the window's."
    StatusBarTop = -StatusBleed
End Function

'@Description("The status bar's width: past the window's left and right edges.")
Public Function StatusBarWidth() As Long
Attribute StatusBarWidth.VB_Description = "The status bar's width: past the window's left and right edges."
    StatusBarWidth = WindowWidth + 2 * StatusBleed
End Function

'@Description("The status bar's height: a line of text, padded above and below.")
Public Function StatusBarHeight() As Long
Attribute StatusBarHeight.VB_Description = "The status bar's height: a line of text, padded above and below."
    StatusBarHeight = LineHeight + 2 * StatusPadding + StatusBleed
End Function

'@Description("The status text's left edge, inside the bar.")
Public Function StatusTextLeft() As Long
Attribute StatusTextLeft.VB_Description = "The status text's left edge, inside the bar."
    StatusTextLeft = Margin
End Function

'@Description("The status text's top edge, inside the bar.")
Public Function StatusTextTop() As Long
Attribute StatusTextTop.VB_Description = "The status text's top edge, inside the bar."
    StatusTextTop = StatusBarTop + StatusBleed + StatusPadding
End Function

'@Description("The status text's width, inside the bar.")
Public Function StatusTextWidth() As Long
Attribute StatusTextWidth.VB_Description = "The status text's width, inside the bar."
    StatusTextWidth = WindowWidth - 2 * Margin
End Function

'@Description("A grid's width: its row numbers, then its cells.")
Public Function GridWidth() As Long
Attribute GridWidth.VB_Description = "A grid's width: its row numbers, then its cells."
    GridWidth = (GridSize + 1) * CellSize
End Function

'@Description("A grid's left edge, where its row numbers are: the two grids are side by side, centred on the window.")
Public Function GridLeft(ByVal gridId As PlayGridId) As Long
Attribute GridLeft.VB_Description = "A grid's left edge, where its row numbers are: the two grids are side by side, centred on the window."
    GridLeft = (WindowWidth - 2 * GridWidth - GridGap) \ 2 + (gridId - PlayerGrid1) * (GridWidth + GridGap)
End Function

'@Description("The grids' top edge, under the status bar.")
Public Function GridTop() As Long
Attribute GridTop.VB_Description = "The grids' top edge, under the status bar."
    GridTop = StatusBarTop + StatusBarHeight + Margin
End Function

'@Description("A column's left edge in a grid: column 0 holds the row numbers, columns 1 to GridSize the cells.")
Public Function CellLeft(ByVal gridId As PlayGridId, ByVal column As Long) As Long
Attribute CellLeft.VB_Description = "A column's left edge in a grid: column 0 holds the row numbers, columns 1 to GridSize the cells."
    CellLeft = GridLeft(gridId) + column * CellSize
End Function

'@Description("A row's top edge in a grid: rows 1 to GridSize hold the cells, and the row after them the column letters.")
Public Function CellTop(ByVal row As Long) As Long
Attribute CellTop.VB_Description = "A row's top edge in a grid: rows 1 to GridSize hold the cells, and the row after them the column letters."
    CellTop = GridTop + (row - 1) * CellSize
End Function

'@Description("The top edge of each grid's player line, under its column letters.")
Public Function PlayerTop() As Long
Attribute PlayerTop.VB_Description = "The top edge of each grid's player line, under its column letters."
    PlayerTop = CellTop(GridSize + 2) + LineGap
End Function

'@Description("The Ready button's left edge, under the right of grid 2.")
Public Function ReadyLeft() As Long
Attribute ReadyLeft.VB_Description = "The Ready button's left edge, under the right of grid 2."
    ReadyLeft = GridLeft(PlayerGrid2) + GridWidth - ReadyWidth
End Function

'@Description("The Ready button's top edge, under the players' line.")
Public Function ReadyTop() As Long
Attribute ReadyTop.VB_Description = "The Ready button's top edge, under the players' line."
    ReadyTop = PlayerTop + LineHeight + LineGap
End Function
