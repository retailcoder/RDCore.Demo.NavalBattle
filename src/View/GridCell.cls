VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GridCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A cell of a grid of the game form: a label made at run time, whose events it passes on with the cell's grid and position."
'@Folder("NavalBattle.View")
'@ModuleDescription("A cell of a grid of the game form: a label made at run time, whose events it passes on with the cell's grid and position.")
'@PredeclaredId
Option Explicit

Private Const RightMouseButton As Long = 2

Private Type TCell
    Events As GridCellEvents
    gridId As PlayGridId
    Position As IGridCoord
End Type

Private this As TCell
Private WithEvents cellLabel As MSForms.Label
Attribute cellLabel.VB_VarHelpID = -1

'@Description("Creates a cell over a label, for the given grid and position; its events are raised by the given GridCellEvents.")
Public Function Create(ByVal cellControl As MSForms.Label, ByVal cellGridId As PlayGridId, ByVal cellPosition As IGridCoord, ByVal events As GridCellEvents) As GridCell
Attribute Create.VB_Description = "Creates a cell over a label, for the given grid and position; its events are raised by the given GridCellEvents."
    Dim result As GridCell
    Set result = New GridCell
    result.Initialize cellControl, cellGridId, cellPosition, events
    Set Create = result
End Function

Friend Sub Initialize(ByVal cellControl As MSForms.Label, ByVal cellGridId As PlayGridId, ByVal cellPosition As IGridCoord, ByVal events As GridCellEvents)
    Set cellLabel = cellControl
    this.gridId = cellGridId
    Set this.Position = cellPosition
    Set this.Events = events
End Sub

'@Description("The label that draws the cell.")
Public Property Get Control() As MSForms.Label
Attribute Control.VB_Description = "The label that draws the cell."
    Set Control = cellLabel
End Property

'@Description("The grid the cell belongs to.")
Public Property Get gridId() As PlayGridId
Attribute gridId.VB_Description = "The grid the cell belongs to."
    gridId = this.gridId
End Property

'@Description("The cell's position in its grid.")
Public Property Get Position() As IGridCoord
Attribute Position.VB_Description = "The cell's position in its grid."
    Set Position = this.Position
End Property

Private Sub cellLabel_Click()
    this.Events.OnClicked this.gridId, this.Position
End Sub

Private Sub cellLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    this.Events.OnHovered this.gridId, this.Position
End Sub

Private Sub cellLabel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = RightMouseButton Then this.Events.OnRightClicked this.gridId, this.Position
End Sub
