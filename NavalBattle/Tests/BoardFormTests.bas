Attribute VB_Name = "BoardFormTests"
Attribute VB_Description = "Tests covering the NavalBattle.BoardForm form, the RDForms window. They are ignored while the runtime's forms host lacks what the form needs; each one's reason names what is missing."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.BoardForm form, the RDForms window. They are ignored while the runtime's forms host lacks what the form needs; each one's reason names what is missing.")
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

' once loaded, the form has 200 cells, lblCell(0) to lblCell(199), each where BoardLayout places it.
'@TestMethod
Public Sub Cells_AreLoadedFromTheDesignedCell()
    Assert.Ignore "control arrays: Load lblCell(n), and every member of an element, raise 438; For Each over the array raises 92"
End Sub

' once loaded, each picker offers PlayerNames in order: grid 1 picks Human, and grid 2 Merciless AI.
'@TestMethod
Public Sub PlayerPickers_OfferEveryPlayer()
    Assert.Ignore "list members: ComboBox AddItem, ListIndex and ListCount raise 438"
End Sub

' once the players are confirmed, each grid's caption names its player in bold.
'@TestMethod
Public Sub ConfirmedPlayers_ShowInBold()
    Assert.Ignore "fonts: a control's Font cannot be given to the program, and FontBold raises 438"
End Sub

' once the ships are placed, the fleet status names each one, in the order they were placed, over a cell for each of its parts.
'@TestMethod
Public Sub FleetStatus_ShowsEveryShipPlaced()
    Assert.Ignore "control arrays: Load lblShip(n) and Load imgShipCell(n), like every control array member, raise 438"
End Sub
