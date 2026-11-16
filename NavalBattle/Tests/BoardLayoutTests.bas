Attribute VB_Name = "BoardLayoutTests"
Attribute VB_Description = "Tests covering the NavalBattle.BoardLayout module."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.BoardLayout module.")
'@TestModule
Option Explicit
Option Private Module

Private Assert As TestAssert

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = TestAssert
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestMethod
Public Sub StatusBar_ShowsOnlyItsBottomBorder()
    ' the bar starts past the window's top, left and right edges, and ends inside it.
    Assert.IsTrue BoardLayout.StatusBarLeft < 0, "left"
    Assert.IsTrue BoardLayout.StatusBarTop < 0, "top"
    Assert.IsTrue BoardLayout.StatusBarLeft + BoardLayout.StatusBarWidth > BoardLayout.WindowWidth, "right"
    Assert.IsTrue BoardLayout.StatusBarTop + BoardLayout.StatusBarHeight > 0, "bottom"
End Sub

'@TestMethod
Public Sub StatusText_IsInsideTheBar()
    Assert.IsTrue BoardLayout.StatusTextTop > 0, "top"
    Assert.IsTrue BoardLayout.StatusTextTop + BoardLayout.LineHeight < BoardLayout.StatusBarTop + BoardLayout.StatusBarHeight, "bottom"
    Assert.IsTrue BoardLayout.StatusTextLeft > 0, "left"
    Assert.IsTrue BoardLayout.StatusTextLeft + BoardLayout.StatusTextWidth < BoardLayout.WindowWidth, "right"
End Sub

'@TestMethod
Public Sub Columns_AreSideBySide_CentredUnderTheBanner()
    Dim leftSpace As Long
    leftSpace = BoardLayout.GridLeft(PlayerGrid1)
    Dim rightSpace As Long
    rightSpace = BoardLayout.WindowWidth - (BoardLayout.FleetLeft + BoardLayout.GridWidth)

    Assert.IsTrue BoardLayout.GridLeft(PlayerGrid1) + BoardLayout.GridWidth < BoardLayout.GridLeft(PlayerGrid2), "the grids overlap"
    Assert.IsTrue BoardLayout.GridLeft(PlayerGrid2) + BoardLayout.GridWidth < BoardLayout.FleetLeft, "grid 2 and the fleet status overlap"
    Assert.IsTrue Abs(leftSpace - rightSpace) <= 1, "not centred: " & leftSpace & " left, " & rightSpace & " right"
    Assert.IsTrue BoardLayout.GridTop >= BoardLayout.BannerTop + BoardLayout.BannerHeight, "under the banner"
End Sub

'@TestMethod
Public Sub Banner_IsAcrossTheWindow_OnTheStatusBarsBottomBorder()
    Assert.AreEqual BoardLayout.StatusBarLeft, BoardLayout.BannerLeft, "its left"
    Assert.AreEqual BoardLayout.StatusBarWidth, BoardLayout.BannerWidth, "its width"
    Assert.AreEqual BoardLayout.StatusBarTop + BoardLayout.StatusBarHeight - BoardLayout.StatusBleed, BoardLayout.BannerTop, "its top, on the status bar's bottom border"
End Sub

'@TestMethod
Public Sub Background_IsUnderTheBanner_AndTheWindowEndsWithIt()
    Assert.AreEqual BoardLayout.BannerTop + BoardLayout.BannerHeight, BoardLayout.BackgroundTop, "under the banner"
    Assert.AreEqual BoardLayout.BackgroundTop + BoardLayout.BackgroundHeight, BoardLayout.WindowHeight, "the window's bottom"
    Assert.IsTrue BoardLayout.GridTop > BoardLayout.BackgroundTop, "the grids, on the picture"
End Sub

'@TestMethod
Public Sub Cells_FollowTheRowNumbers_AndTheLettersFollowTheLastRow()
    Assert.AreEqual BoardLayout.GridLeft(PlayerGrid2), BoardLayout.CellLeft(PlayerGrid2, 0), "the row numbers"
    Assert.AreEqual BoardLayout.CellLeft(PlayerGrid2, 0) + BoardLayout.CellSize, BoardLayout.CellLeft(PlayerGrid2, 1), "column A"
    Assert.AreEqual BoardLayout.GridTop, BoardLayout.CellTop(1), "row 1"
    Assert.AreEqual BoardLayout.CellTop(BoardLayout.GridSize) + BoardLayout.CellSize, BoardLayout.CellTop(BoardLayout.GridSize + 1), "the letters"
End Sub

'@TestMethod
Public Sub Captions_AreOverTheirCells_OnTheBanner_AndHoldThePickers()
    Assert.AreEqual BoardLayout.CellLeft(PlayerGrid2, 1), BoardLayout.GridCaptionLeft(PlayerGrid2), "over grid 2's cells"
    Assert.AreEqual BoardLayout.CellLeft(PlayerGrid2, BoardLayout.GridSize) + BoardLayout.CellSize, BoardLayout.GridCaptionLeft(PlayerGrid2) + BoardLayout.CaptionWidth, "as wide as its cells"
    Assert.AreEqual BoardLayout.ShipCellLeft(1), BoardLayout.FleetCaptionLeft, "over the fleet status's cells"
    Assert.IsTrue BoardLayout.BannerTextTop > BoardLayout.BannerTop, "a picker's top, on the banner"
    Assert.IsTrue BoardLayout.BannerTextTop + BoardLayout.LineHeight < BoardLayout.BannerTop + BoardLayout.BannerHeight, "a picker's bottom, on the banner"
    Assert.IsTrue BoardLayout.PickerLeft(PlayerGrid2) > BoardLayout.GridCaptionLeft(PlayerGrid2), "a picker, after its caption"
    Assert.IsTrue BoardLayout.PickerLeft(PlayerGrid2) + BoardLayout.PickerWidth < BoardLayout.GridCaptionLeft(PlayerGrid2) + BoardLayout.CaptionWidth, "a picker's right, inside its caption"
End Sub

'@TestMethod
Public Sub FleetShips_AreUnderTheirBanner_OneUnderTheOther_InsideTheWindow()
    ' the carrier, five cells, is the longest ship.
    Const LongestShip As Long = 5
    Assert.AreEqual BoardLayout.GridTop, BoardLayout.ShipNameTop(1), "the first ship's name"
    Assert.AreEqual BoardLayout.ShipNameTop(1) + BoardLayout.LineHeight, BoardLayout.ShipCellTop(1), "its cells, under its name"
    Assert.IsTrue BoardLayout.ShipNameTop(2) >= BoardLayout.ShipCellTop(1) + BoardLayout.CellSize, "the next ship, under the first one's cells"
    Assert.IsTrue BoardLayout.ShipCellTop(PlayerGrid.ShipsPerGrid) + BoardLayout.CellSize <= BoardLayout.WindowHeight - BoardLayout.Margin, "the window's height"
    Assert.IsTrue BoardLayout.ShipCellLeft(LongestShip) + BoardLayout.CellSize <= BoardLayout.FleetLeft + BoardLayout.GridWidth, "the longest ship, inside its column"
End Sub

'@TestMethod
Public Sub Buttons_AreARow_AtTheBottomLeft_UnderTheGrids()
    Dim gap As Long
    gap = BoardLayout.ButtonLeft(2) - (BoardLayout.ButtonLeft(1) + BoardLayout.ButtonWidth)

    Assert.AreEqual BoardLayout.Margin, BoardLayout.ButtonLeft(1), "the first button, at the left margin"
    Assert.AreEqual BoardLayout.WindowHeight - BoardLayout.Margin, BoardLayout.ButtonTop + BoardLayout.ButtonHeight, "the row, at the bottom margin"
    Assert.IsTrue BoardLayout.ButtonTop >= BoardLayout.CellTop(BoardLayout.GridSize + 2), "under the letters"
    Assert.IsTrue gap >= 8 And gap <= 12, "the gap between two buttons: " & gap
End Sub
