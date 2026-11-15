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
Public Sub Grids_AreSideBySide_CentredUnderTheStatusBar()
    Dim leftSpace As Long
    leftSpace = BoardLayout.GridLeft(PlayerGrid1)
    Dim rightSpace As Long
    rightSpace = BoardLayout.WindowWidth - (BoardLayout.GridLeft(PlayerGrid2) + BoardLayout.GridWidth)

    Assert.IsTrue BoardLayout.GridLeft(PlayerGrid1) + BoardLayout.GridWidth < BoardLayout.GridLeft(PlayerGrid2), "the grids overlap"
    Assert.IsTrue Abs(leftSpace - rightSpace) <= 1, "not centred: " & leftSpace & " left, " & rightSpace & " right"
    Assert.IsTrue BoardLayout.GridTop > BoardLayout.StatusBarTop + BoardLayout.StatusBarHeight, "under the status bar"
End Sub

'@TestMethod
Public Sub Cells_FollowTheRowNumbers_AndTheLettersFollowTheLastRow()
    Assert.AreEqual BoardLayout.GridLeft(PlayerGrid2), BoardLayout.CellLeft(PlayerGrid2, 0), "the row numbers"
    Assert.AreEqual BoardLayout.CellLeft(PlayerGrid2, 0) + BoardLayout.CellSize, BoardLayout.CellLeft(PlayerGrid2, 1), "column A"
    Assert.AreEqual BoardLayout.GridTop, BoardLayout.CellTop(1), "row 1"
    Assert.AreEqual BoardLayout.CellTop(BoardLayout.GridSize) + BoardLayout.CellSize, BoardLayout.CellTop(BoardLayout.GridSize + 1), "the letters"
End Sub

'@TestMethod
Public Sub PlayersAndReady_AreUnderTheGrids_InsideTheWindow()
    Assert.IsTrue BoardLayout.PlayerTop >= BoardLayout.CellTop(BoardLayout.GridSize + 2), "the players' line"
    Assert.IsTrue BoardLayout.ReadyTop >= BoardLayout.PlayerTop + BoardLayout.LineHeight, "Ready"
    Assert.IsTrue BoardLayout.ReadyTop + BoardLayout.ReadyHeight <= BoardLayout.WindowHeight - BoardLayout.Margin, "the window's height"
    Assert.AreEqual BoardLayout.GridLeft(PlayerGrid2) + BoardLayout.GridWidth, BoardLayout.ReadyLeft + BoardLayout.ReadyWidth, "Ready ends with grid 2"
End Sub
