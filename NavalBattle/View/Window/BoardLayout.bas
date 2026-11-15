Attribute VB_Name = "BoardLayout"
Attribute VB_Description = "Where the parts of a game window go, in points, from the top: the status bar, then a banner across the window with a caption over each column, then the three columns side by side: the two grids, and the fleet status. A row of buttons starts at the bottom left. Every game window lays itself out by it; a window that measures in twips multiplies by 20."
'@Folder("NavalBattle.View.Window")
'@ModuleDescription("Where the parts of a game window go, in points, from the top: the status bar, then a banner across the window with a caption over each column, then the three columns side by side: the two grids, and the fleet status. A row of buttons starts at the bottom left. Every game window lays itself out by it; a window that measures in twips multiplies by 20.")
Option Explicit

' the number of rows and of columns of a grid.
Public Const GridSize As Long = 10
' the side of a grid's cell.
Public Const CellSize As Long = 18
' the window's client area: 16320 x 7845 twips, 1088 x 523 pixels at 96 dpi.
Public Const WindowWidth As Long = 816
Public Const WindowHeight As Long = 392
' the height of a line of text: a player's line, a ship's name, and the status text.
Public Const LineHeight As Long = 22
' the space between the window's edges and what it shows.
Public Const Margin As Long = 12
' how far the status bar starts past the window's top, left and right edges, so that only its bottom border shows.
Public Const StatusBleed As Long = 1
' the space between the status bar's edges and its text.
Public Const StatusPadding As Long = 8
' the space between the banner's edges and what it holds.
Public Const BannerPadding As Long = 4
' the width of the "Grid n:" caption before a grid's player picker, in the banner.
Public Const PlayerCaptionWidth As Long = 56
' the size of a button in the row at the bottom left.
Public Const ButtonWidth As Long = 80
Public Const ButtonHeight As Long = 28

' the space between two columns.
Private Const ColumnGap As Long = 48
' the space between the banner and the columns under it.
Private Const LineGap As Long = 8
' the space between two buttons in the row.
Private Const ButtonGap As Long = 10
' the columns: the two grids, then the fleet status.
Private Const ColumnCount As Long = 3

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

'@Description("A column's width, a grid's: its row numbers, then its cells. The fleet status is as wide.")
Public Function GridWidth() As Long
Attribute GridWidth.VB_Description = "A column's width, a grid's: its row numbers, then its cells. The fleet status is as wide."
    GridWidth = (GridSize + 1) * CellSize
End Function

'@Description("A grid's left edge, where its row numbers are: the grids are the first two of the three columns, which are centred on the window.")
Public Function GridLeft(ByVal gridId As PlayGridId) As Long
Attribute GridLeft.VB_Description = "A grid's left edge, where its row numbers are: the grids are the first two of the three columns, which are centred on the window."
    GridLeft = ColumnLeft(gridId - PlayerGrid1)
End Function

'@Description("The fleet status's left edge: the third column, after the grids.")
Public Function FleetLeft() As Long
Attribute FleetLeft.VB_Description = "The fleet status's left edge: the third column, after the grids."
    FleetLeft = ColumnLeft(ColumnCount - 1)
End Function

Private Function ColumnLeft(ByVal column As Long) As Long
    ColumnLeft = (WindowWidth - ColumnCount * GridWidth - (ColumnCount - 1) * ColumnGap) \ 2 + column * (GridWidth + ColumnGap)
End Function

'@Description("The banner's left edge: past the window's, like the status bar's.")
Public Function BannerLeft() As Long
Attribute BannerLeft.VB_Description = "The banner's left edge: past the window's, like the status bar's."
    BannerLeft = StatusBarLeft
End Function

'@Description("The banner's top edge: on the status bar's bottom border, so that the two share it.")
Public Function BannerTop() As Long
Attribute BannerTop.VB_Description = "The banner's top edge: on the status bar's bottom border, so that the two share it."
    BannerTop = StatusBarTop + StatusBarHeight - StatusBleed
End Function

'@Description("The banner's width: past the window's left and right edges, like the status bar's.")
Public Function BannerWidth() As Long
Attribute BannerWidth.VB_Description = "The banner's width: past the window's left and right edges, like the status bar's."
    BannerWidth = StatusBarWidth
End Function

'@Description("The banner's height: a line of text, padded above and below.")
Public Function BannerHeight() As Long
Attribute BannerHeight.VB_Description = "The banner's height: a line of text, padded above and below."
    BannerHeight = LineHeight + 2 * BannerPadding
End Function

'@Description("A column's caption's width, on the banner: as wide as a grid's cells.")
Public Function CaptionWidth() As Long
Attribute CaptionWidth.VB_Description = "A column's caption's width, on the banner: as wide as a grid's cells."
    CaptionWidth = GridSize * CellSize
End Function

'@Description("A grid's caption's left edge, on the banner: over its cells, past its row numbers.")
Public Function GridCaptionLeft(ByVal gridId As PlayGridId) As Long
Attribute GridCaptionLeft.VB_Description = "A grid's caption's left edge, on the banner: over its cells, past its row numbers."
    GridCaptionLeft = CellLeft(gridId, 1)
End Function

'@Description("The fleet status's caption's left edge, on the banner: over its ships' cells.")
Public Function FleetCaptionLeft() As Long
Attribute FleetCaptionLeft.VB_Description = "The fleet status's caption's left edge, on the banner: over its ships' cells."
    FleetCaptionLeft = ShipCellLeft(1)
End Function

'@Description("The top edge of what the banner holds: a caption's text, or a player picker.")
Public Function BannerTextTop() As Long
Attribute BannerTextTop.VB_Description = "The top edge of what the banner holds: a caption's text, or a player picker."
    BannerTextTop = BannerTop + BannerPadding
End Function

'@Description("A grid's player picker's left edge, on the banner, after the grid's caption.")
Public Function PickerLeft(ByVal gridId As PlayGridId) As Long
Attribute PickerLeft.VB_Description = "A grid's player picker's left edge, on the banner, after the grid's caption."
    PickerLeft = GridCaptionLeft(gridId) + PlayerCaptionWidth
End Function

'@Description("A player picker's width: the rest of its grid's caption.")
Public Function PickerWidth() As Long
Attribute PickerWidth.VB_Description = "A player picker's width: the rest of its grid's caption."
    PickerWidth = CaptionWidth - PlayerCaptionWidth - BannerPadding
End Function

'@Description("The grids' top edge, under the banner.")
Public Function GridTop() As Long
Attribute GridTop.VB_Description = "The grids' top edge, under the banner."
    GridTop = BannerTop + BannerHeight + LineGap
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

'@Description("The height of a ship in the fleet status: its name, then its cells.")
Public Function ShipHeight() As Long
Attribute ShipHeight.VB_Description = "The height of a ship in the fleet status: its name, then its cells."
    ShipHeight = LineHeight + CellSize
End Function

'@Description("The top edge of a ship's name in the fleet status, by the ship's place in the fleet, from 1.")
Public Function ShipNameTop(ByVal shipNumber As Long) As Long
Attribute ShipNameTop.VB_Description = "The top edge of a ship's name in the fleet status, by the ship's place in the fleet, from 1."
    ShipNameTop = GridTop + (shipNumber - 1) * ShipHeight
End Function

'@Description("The top edge of a ship's cells in the fleet status, under its name.")
Public Function ShipCellTop(ByVal shipNumber As Long) As Long
Attribute ShipCellTop.VB_Description = "The top edge of a ship's cells in the fleet status, under its name."
    ShipCellTop = ShipNameTop(shipNumber) + LineHeight
End Function

'@Description("The left edge of a ship's cell in the fleet status, from 1 at its stern: the ships lie across, under their caption, as a grid's cells do.")
Public Function ShipCellLeft(ByVal column As Long) As Long
Attribute ShipCellLeft.VB_Description = "The left edge of a ship's cell in the fleet status, from 1 at its stern: the ships lie across, under their caption, as a grid's cells do."
    ShipCellLeft = FleetLeft + column * CellSize
End Function

'@Description("A button's left edge in the row at the bottom left, by its place in the row, from 1: the Ready button is the first.")
Public Function ButtonLeft(ByVal buttonNumber As Long) As Long
Attribute ButtonLeft.VB_Description = "A button's left edge in the row at the bottom left, by its place in the row, from 1: the Ready button is the first."
    ButtonLeft = Margin + (buttonNumber - 1) * (ButtonWidth + ButtonGap)
End Function

'@Description("The top edge of the row of buttons, at the bottom of the window.")
Public Function ButtonTop() As Long
Attribute ButtonTop.VB_Description = "The top edge of the row of buttons, at the bottom of the window."
    ButtonTop = WindowHeight - Margin - ButtonHeight
End Function
