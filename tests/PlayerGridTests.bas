Attribute VB_Name = "PlayerGridTests"
Attribute VB_Description = "Tests covering the NavalBattle.PlayerGrid class."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Tests covering the NavalBattle.PlayerGrid class.")
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
Public Sub CanAddShipInsideGridBoundaries_ReturnsTrue()
    Dim position As GridCoord
    Set position = GridCoord.Create(1, 1)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    Assert.IsTrue SUT.CanAddShip(position, Horizontal, 2)
End Sub

'@TestMethod
Public Sub CanAddShipAtPositionZeroZero_ReturnsFalse()
'i.e. PlayerGrid coordinates are 1-based
    Dim position As GridCoord
    Set position = GridCoord.Create(0, 0)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    Assert.IsFalse SUT.CanAddShip(position, Horizontal, 2)
End Sub

'@TestMethod
Public Sub CanAddShipGivenInterectingShips_ReturnsFalse()
    Dim Ship1 As IShip
    Set Ship1 = Ship.Create(ShipType.Dreadnought, Horizontal, GridCoord.Create(1, 1))
    
    Dim Ship2 As IShip
    Set Ship2 = Ship.Create(ShipType.Dreadnought, Vertical, GridCoord.Create(2, 1))
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    SUT.AddShip Ship1
    Assert.IsFalse SUT.CanAddShip(Ship2.GridPosition, Ship2.Orientation, Ship2.Size)
End Sub

'@TestMethod
Public Sub AddingSameShipTypeTwice_Throws()
    Const ExpectedError As Long = 457 ' "This key is already associated with an element of this collection"
    On Error GoTo TestFail
    
    Dim Ship1 As IShip
    Set Ship1 = Ship.Create(ShipType.Dreadnought, Horizontal, GridCoord.Create(1, 1))
    
    Dim Ship2 As IShip
    Set Ship2 = Ship.Create(ShipType.Dreadnought, Horizontal, GridCoord.Create(5, 5))
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    SUT.AddShip Ship1
    SUT.AddShip Ship2

Assert:
    Assert.Fail "Expected error was not raised."

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
Public Sub AddingShipOutsideGridBoundaries_Throws()
    Const ExpectedError As Long = PlayerGridErrors.CannotAddShipAtPosition
    On Error GoTo TestFail
    
    Dim Ship1 As IShip
    Set Ship1 = Ship.Create(ShipType.Dreadnought, Horizontal, GridCoord.Create(0, 0))
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    SUT.AddShip Ship1

Assert:
    Assert.Fail "Expected error was not raised."

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
Public Sub TryHitKnownState_Throws()
    Const ExpectedError As Long = PlayerGridErrors.KnownGridStateError
    On Error GoTo TestFail
    
    Dim position As GridCoord
    Set position = GridCoord.Create(1, 1)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    SUT.AddShip Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    SUT.TryHit position
    SUT.TryHit position

Assert:
    Assert.Fail "Expected error was not raised."
    
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
Public Sub TryHitMiss_SetsPreviousMissState()
    Const expected = GridState.PreviousMiss
    
    Dim position As IGridCoord
    Set position = GridCoord.Create(1, 1)
    
    Dim badPosition As GridCoord
    Set badPosition = position.Offset(5, 5)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    SUT.AddShip Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    SUT.TryHit badPosition
    Dim actual As GridState
    actual = SUT.State(badPosition)
    Assert.AreEqual expected, actual
End Sub

'@TestMethod
Public Sub TryHitSuccess_SetsPreviousHitState()
    Const expected = GridState.PreviousHit
    
    Dim position As GridCoord
    Set position = GridCoord.Create(1, 1)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    SUT.AddShip Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    SUT.TryHit position
    Dim actual As GridState
    actual = SUT.State(position)
    Assert.AreEqual expected, actual
End Sub

'@TestMethod
Public Sub TryHitSuccess_ReturnsHit()
    Dim position As GridCoord
    Set position = GridCoord.Create(1, 1)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    SUT.AddShip Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    Assert.AreEqual AttackResult.Hit, SUT.TryHit(position)
End Sub

'@TestMethod
Public Sub TryHitMisses_ReturnsMiss()
    Dim position As IGridCoord
    Set position = GridCoord.Create(1, 1)
    
    Dim badPosition As IGridCoord
    Set badPosition = position.Offset(5, 5)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    SUT.AddShip Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    Assert.AreEqual AttackResult.Miss, SUT.TryHit(badPosition)
End Sub

'@TestMethod
Public Sub GridInitialState_UnknownState()
    Const expected = GridState.Unknown
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    Dim actual As GridState
    actual = SUT.State(GridCoord.Create(1, 1))
    
    Assert.AreEqual expected, actual
End Sub

'@TestMethod
Public Sub GivenAdjacentShip_HasRightAdjacentShipReturnsTrue()
    Dim position As GridCoord
    Set position = GridCoord.Create(2, 2)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    SUT.AddShip Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    Assert.IsTrue SUT.HasAdjacentShip(GridCoord.Create(1, 2), Vertical, 3)
End Sub

'@TestMethod
Public Sub GivenAdjacentShip_HasLeftAdjacentShipReturnsTrue()
    Dim position As GridCoord
    Set position = GridCoord.Create(2, 1)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    SUT.AddShip Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    Assert.IsTrue SUT.HasAdjacentShip(GridCoord.Create(1, 1), Vertical, 3)
End Sub

'@TestMethod
Public Sub GivenAdjacentShip_HasDownAdjacentShipReturnsTrue()
    Dim position As GridCoord
    Set position = GridCoord.Create(2, 2)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    SUT.AddShip Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    Assert.IsTrue SUT.HasAdjacentShip(GridCoord.Create(1, 3), Horizontal, 3)
End Sub

'@TestMethod
Public Sub GivenAdjacentShip_HasUpAdjacentShipReturnsTrue()
    Dim position As GridCoord
    Set position = GridCoord.Create(2, 2)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    SUT.AddShip Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    Assert.IsTrue SUT.HasAdjacentShip(GridCoord.Create(1, 1), Horizontal, 3)
End Sub

'@TestMethod
Public Sub GivenAdjacentShipAtHorizontalTipEnd_ReturnsTrue()
    Dim position As GridCoord
    Set position = GridCoord.Create(10, 4)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    SUT.AddShip Ship.Create(ShipType.Carrier, Vertical, position)
    
    Assert.IsTrue SUT.HasAdjacentShip(GridCoord.Create(6, 7), Horizontal, 4)
End Sub

'@TestMethod
Public Sub GivenAdjacentShipAtVerticalTipEnd_ReturnsTrue()
    Dim position As GridCoord
    Set position = GridCoord.Create(6, 7)
    
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    SUT.AddShip Ship.Create(ShipType.Dreadnought, Horizontal, position)
    
    Assert.IsTrue SUT.HasAdjacentShip(GridCoord.Create(10, 4), Vertical, 5)
End Sub

'@TestMethod
Public Sub GivenTwoSideBySideHits_GetHitAreaReturnsTwoItems()

    Const expected As Long = 2

    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    
    SUT.AddShip Ship.Create(ShipType.Carrier, Horizontal, GridCoord.Create(1, 1))
    SUT.TryHit GridCoord.Create(1, 1)
    SUT.TryHit GridCoord.Create(2, 1)
    
    Dim area As Collection
    Set area = SUT.FindHitArea
    
    Dim actual As Long
    actual = area.Count

    Assert.AreEqual expected, actual
End Sub

'@TestMethod
Public Sub GivenSunkHorizontalShip_IsSunkenOnEveryCell()
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    SUT.AddShip Ship.Create(ShipType.Dreadnought, Horizontal, GridCoord.Create(2, 3))

    Dim currentX As Long
    For currentX = 2 To 5
        SUT.TryHit GridCoord.Create(currentX, 3)
    Next

    For currentX = 2 To 5
        Assert.IsTrue SUT.IsSunken(GridCoord.Create(currentX, 3)), "Not sunken at " & currentX & ",3."
    Next
End Sub

'@TestMethod
Public Sub GivenSunkVerticalShip_IsSunkenOnEveryCell()
    Dim SUT As PlayerGrid
    Set SUT = New PlayerGrid
    SUT.AddShip Ship.Create(ShipType.Cruiser, Vertical, GridCoord.Create(6, 4))

    Dim currentY As Long
    For currentY = 4 To 6
        SUT.TryHit GridCoord.Create(6, currentY)
    Next

    For currentY = 4 To 6
        Assert.IsTrue SUT.IsSunken(GridCoord.Create(6, currentY)), "Not sunken at 6," & currentY & "."
    Next
End Sub
