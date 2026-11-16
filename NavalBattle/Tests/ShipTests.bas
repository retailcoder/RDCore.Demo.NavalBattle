Attribute VB_Name = "ShipTests"
Attribute VB_Description = "Tests covering the NavalBattle.Ship class."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.Ship class.")
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

'@TestInitialize
Public Sub TestInitialize()
End Sub

'@TestCleanup
Public Sub TestCleanup()
End Sub

'@TestMethod
Public Sub CreatesShipOfSpecifiedType()
    Dim expected As ShipType
    expected = ShipType.Dreadnought
    
    Dim SUT As IShip
    Set SUT = Ship.Create(expected, Horizontal, GridCoord.Default)
    
    Dim actual As ShipType
    actual = SUT.ShipKind
    
    Assert.AreEqual expected, actual
End Sub

'@TestMethod
Public Sub SuccessfulHit_SetsStateToTrue()
    Dim position As IGridCoord
    Set position = GridCoord.Default
    
    Dim obj As Ship
    Set obj = Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    If Not position.Equals(GridCoord.Default) Then Assert.Inconclusive
    If obj.State(position.ToString) Then Assert.Inconclusive
    
    Dim SUT As IShip
    Set SUT = obj
    SUT.Hit position
    
    Assert.IsTrue obj.State(position.ToString)
End Sub

'@TestMethod
Public Sub SunkenShip_IsSunken()
    Dim position As IGridCoord
    Set position = GridCoord.Default
    
    Dim obj As Ship
    Set obj = Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    If Not position.Equals(GridCoord.Default) Then Assert.Inconclusive
    
    Dim SUT As IShip
    Set SUT = obj
    
    Dim current As Variant
    For Each current In obj.State.Keys
        SUT.Hit GridCoord.FromString(current)
    Next
    
    Assert.IsTrue SUT.IsSunken
End Sub

'@TestMethod
Public Sub MissedHit_ReturnsFalse()
    Dim position As IGridCoord
    Set position = GridCoord.Default
    
    Dim obj As Ship
    Set obj = Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    Dim SUT As IShip
    Set SUT = obj
    
    Dim target As GridCoord
    Set target = position.Offset( _
        xOffset:=IIf(SUT.Orientation = Horizontal, 0, 1), _
        yOffset:=IIf(SUT.Orientation = Vertical, 0, 1))
    
    Assert.IsFalse SUT.Hit(target)
End Sub

'@TestMethod
Public Sub SuccessfulHit_ReturnsTrue()
    Dim position As GridCoord
    Set position = GridCoord.Default
    
    Dim obj As Ship
    Set obj = Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    Dim SUT As IShip
    Set SUT = obj
    
    Assert.IsTrue SUT.Hit(position)
End Sub

'@TestMethod
Public Sub ShipOrientationInvalidEnumValue_Throws()
    
    Const ExpectedError As Long = 5
    Const badValue = 42
    
    On Error GoTo TestFail
    
    If badValue = ShipOrientation.Horizontal Or _
       badValue = ShipOrientation.Vertical _
    Then
        Assert.Inconclusive
    End If
    
    Dim SUT As Ship
    Set SUT = Ship.Create(ShipType.Dreadnought, badValue, GridCoord.Default)
    
Assert:
    Assert.Fail "Expected error was not raised. A Ship instance could be created with an unknown orientation."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod
Public Sub IntersectingShip_IntersectReturnsGridCoord()
    Dim Ship1 As IShip
    Set Ship1 = Ship.Create(ShipType.Dreadnought, Horizontal, GridCoord.Default)
    
    Dim Ship2 As IShip
    Set Ship2 = Ship.Create(ShipType.Dreadnought, Vertical, GridCoord.Default)
    
    Assert.IsNotNothing Ship1.Intersects(Ship2.Size, Ship2.Orientation, Ship2.GridPosition)
End Sub

'@TestMethod
Public Sub NonIntersectingShip_IntersectReturnsNothing()
    Dim Ship1 As IShip
    Set Ship1 = Ship.Create(ShipType.Dreadnought, Horizontal, GridCoord.Default)
    
    Dim Ship2 As IShip
    Set Ship2 = Ship.Create(ShipType.Dreadnought, Horizontal, GridCoord.Default.Offset(yOffset:=1))
    
    Assert.IsNothing Ship1.Intersects(Ship2.Size, Ship2.Orientation, Ship2.GridPosition)
End Sub

'@TestMethod
Public Sub CreateInside_PastTheRightEdge_AnchorsWhereTheShipFits()
    ' a carrier is 5 cells: at J5 across, it would end past the grid; it starts at F5 instead, and ends at J5.
    Dim SUT As IShip
    Set SUT = Ship.CreateInside(ShipType.Carrier, Horizontal, GridCoord.Create(10, 5))
    Assert.AreEqual 6&, SUT.GridPosition.X, "X"
    Assert.AreEqual 5&, SUT.GridPosition.Y, "Y"
End Sub

'@TestMethod
Public Sub CreateInside_PastTheBottomEdge_AnchorsWhereTheShipFits()
    Dim SUT As IShip
    Set SUT = Ship.CreateInside(ShipType.Carrier, Vertical, GridCoord.Create(3, 9))
    Assert.AreEqual 3&, SUT.GridPosition.X, "X"
    Assert.AreEqual 6&, SUT.GridPosition.Y, "Y"
End Sub

'@TestMethod
Public Sub CreateInside_WhereTheShipFits_KeepsThePosition()
    Dim SUT As IShip
    Set SUT = Ship.CreateInside(ShipType.Destroyer, Horizontal, GridCoord.Create(9, 10))
    Assert.AreEqual 9&, SUT.GridPosition.X, "X"
    Assert.AreEqual 10&, SUT.GridPosition.Y, "Y"
End Sub


'@TestMethod
Public Sub NameChoices_AreFiveAKind_NoNameTwice()
    Dim seen As Collection
    Set seen = New Collection
    Dim kind As Variant
    For Each kind In Ship.ShipKinds
        Dim choices As Variant
        choices = Ship.NameChoices(kind)
        Assert.AreEqual 5, UBound(choices) - LBound(choices) + 1, "kind " & kind
        Dim choice As Variant
        For Each choice In choices
            On Error Resume Next
            seen.Add choice, choice
            Assert.AreEqual 0, Err.Number, "named twice: " & choice
            On Error GoTo 0
        Next
    Next
End Sub

'@TestMethod
Public Sub PickNames_NamesTheShipsCreatedAfterIt_AmongTheirKindsChoices()
    Dim randomizer As EdgeRandomizer
    Set randomizer = New EdgeRandomizer
    Dim choices As Variant
    choices = Ship.NameChoices(ShipType.Destroyer)

    randomizer.PicksHighest = True
    Ship.PickNames randomizer
    Assert.AreEqual choices(UBound(choices)), Ship.Create(ShipType.Destroyer, Horizontal, GridCoord.Default).Name, "the highest pick"

    ' the first choices again, as the other tests expect.
    randomizer.PicksHighest = False
    Ship.PickNames randomizer
    Assert.AreEqual choices(LBound(choices)), Ship.Create(ShipType.Destroyer, Horizontal, GridCoord.Default).Name, "the lowest pick"
End Sub
