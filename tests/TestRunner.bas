Attribute VB_Name = "TestRunner"
Attribute VB_Description = "Runs the test modules: each test by name, with its module's initialize and cleanup procedures, and prints each outcome and a summary. A run with a failed test ends with an error."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("Runs the test modules: each test by name, with its module's initialize and cleanup procedures, and prints each outcome and a summary. A run with a failed test ends with an error.")
Option Explicit

Private Type TRun
    CurrentModule As String
    Passed As Long
    Failed As Long
    Inconclusive As Long
End Type

Private this As TRun

'@Description("Runs every test, prints each outcome and a summary, and raises an error where a test failed.")
Public Sub RunAll()
Attribute RunAll.VB_Description = "Runs every test, prints each outcome and a summary, and raises an error where a test failed."
    this.CurrentModule = vbNullString
    this.Passed = 0
    this.Failed = 0
    this.Inconclusive = 0

    Dim index As Long
    Dim name As String
    For index = 1 To TestCount
        name = Test(index, run:=False)
        BeforeTest name
        On Error GoTo TestError
        Test index, run:=True
        On Error GoTo 0
        AfterTest name, TestAssert.Outcome, TestAssert.Reason
NextTest:
    Next
    EnterModule vbNullString

    Debug.Print TestCount & " tests: " & this.Passed & " passed, " & this.Failed & " failed, " & this.Inconclusive & " inconclusive."
    If this.Failed > 0 Then Err.Raise vbObjectError + 1, "TestRunner", this.Failed & " of " & TestCount & " tests failed."
    Exit Sub

TestError:
    AfterTest name, TestFailed, "Error " & Err.Number & ": " & Err.Description
    Resume NextTest
End Sub

'@Description("The number of tests.")
Private Property Get TestCount() As Long
Attribute TestCount.VB_Description = "The number of tests."
    TestCount = 32
End Property

'@Description("The name of a test, by its number, and runs it where run is True.")
Private Function Test(ByVal index As Long, ByVal run As Boolean) As String
Attribute Test.VB_Description = "The name of a test, by its number, and runs it where run is True."
    Select Case index
        Case 1: Test = "PlayerGridTests.CanAddShipInsideGridBoundaries_ReturnsTrue": If run Then PlayerGridTests.CanAddShipInsideGridBoundaries_ReturnsTrue
        Case 2: Test = "PlayerGridTests.CanAddShipAtPositionZeroZero_ReturnsFalse": If run Then PlayerGridTests.CanAddShipAtPositionZeroZero_ReturnsFalse
        Case 3: Test = "PlayerGridTests.CanAddShipGivenInterectingShips_ReturnsFalse": If run Then PlayerGridTests.CanAddShipGivenInterectingShips_ReturnsFalse
        Case 4: Test = "PlayerGridTests.AddingSameShipTypeTwice_Throws": If run Then PlayerGridTests.AddingSameShipTypeTwice_Throws
        Case 5: Test = "PlayerGridTests.AddingShipOutsideGridBoundaries_Throws": If run Then PlayerGridTests.AddingShipOutsideGridBoundaries_Throws
        Case 6: Test = "PlayerGridTests.TryHitKnownState_Throws": If run Then PlayerGridTests.TryHitKnownState_Throws
        Case 7: Test = "PlayerGridTests.TryHitMiss_SetsPreviousMissState": If run Then PlayerGridTests.TryHitMiss_SetsPreviousMissState
        Case 8: Test = "PlayerGridTests.TryHitSuccess_SetsPreviousHitState": If run Then PlayerGridTests.TryHitSuccess_SetsPreviousHitState
        Case 9: Test = "PlayerGridTests.TryHitSuccess_ReturnsHit": If run Then PlayerGridTests.TryHitSuccess_ReturnsHit
        Case 10: Test = "PlayerGridTests.TryHitMisses_ReturnsMiss": If run Then PlayerGridTests.TryHitMisses_ReturnsMiss
        Case 11: Test = "PlayerGridTests.GridInitialState_UnknownState": If run Then PlayerGridTests.GridInitialState_UnknownState
        Case 12: Test = "PlayerGridTests.GivenAdjacentShip_HasRightAdjacentShipReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShip_HasRightAdjacentShipReturnsTrue
        Case 13: Test = "PlayerGridTests.GivenAdjacentShip_HasLeftAdjacentShipReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShip_HasLeftAdjacentShipReturnsTrue
        Case 14: Test = "PlayerGridTests.GivenAdjacentShip_HasDownAdjacentShipReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShip_HasDownAdjacentShipReturnsTrue
        Case 15: Test = "PlayerGridTests.GivenAdjacentShip_HasUpAdjacentShipReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShip_HasUpAdjacentShipReturnsTrue
        Case 16: Test = "PlayerGridTests.GivenAdjacentShipAtHorizontalTipEnd_ReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShipAtHorizontalTipEnd_ReturnsTrue
        Case 17: Test = "PlayerGridTests.GivenAdjacentShipAtVerticalTipEnd_ReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShipAtVerticalTipEnd_ReturnsTrue
        Case 18: Test = "PlayerGridTests.GivenTwoSideBySideHits_GetHitAreaReturnsTwoItems": If run Then PlayerGridTests.GivenTwoSideBySideHits_GetHitAreaReturnsTwoItems
        Case 19: Test = "PlayerGridTests.GivenSunkHorizontalShip_IsSunkenOnEveryCell": If run Then PlayerGridTests.GivenSunkHorizontalShip_IsSunkenOnEveryCell
        Case 20: Test = "PlayerGridTests.GivenSunkVerticalShip_IsSunkenOnEveryCell": If run Then PlayerGridTests.GivenSunkVerticalShip_IsSunkenOnEveryCell
        Case 21: Test = "ShipTests.CreatesShipOfSpecifiedType": If run Then ShipTests.CreatesShipOfSpecifiedType
        Case 22: Test = "ShipTests.SuccessfulHit_SetsStateToTrue": If run Then ShipTests.SuccessfulHit_SetsStateToTrue
        Case 23: Test = "ShipTests.SunkenShip_IsSunken": If run Then ShipTests.SunkenShip_IsSunken
        Case 24: Test = "ShipTests.MissedHit_ReturnsFalse": If run Then ShipTests.MissedHit_ReturnsFalse
        Case 25: Test = "ShipTests.SuccessfulHit_ReturnsTrue": If run Then ShipTests.SuccessfulHit_ReturnsTrue
        Case 26: Test = "ShipTests.ShipOrientationInvalidEnumValue_Throws": If run Then ShipTests.ShipOrientationInvalidEnumValue_Throws
        Case 27: Test = "ShipTests.IntersectingShip_IntersectReturnsGridCoord": If run Then ShipTests.IntersectingShip_IntersectReturnsGridCoord
        Case 28: Test = "ShipTests.NonIntersectingShip_IntersectReturnsNothing": If run Then ShipTests.NonIntersectingShip_IntersectReturnsNothing
        Case 29: Test = "CommandLineSwitchesTests.RuntimeSwitchBefore_TheSeedIsRead": If run Then CommandLineSwitchesTests.RuntimeSwitchBefore_TheSeedIsRead
        Case 30: Test = "CommandLineSwitchesTests.RuntimeSwitchAfter_TheSeedIsRead": If run Then CommandLineSwitchesTests.RuntimeSwitchAfter_TheSeedIsRead
        Case 31: Test = "CommandLineSwitchesTests.RuntimeSwitchInTheValuesPlace_IsNoValue": If run Then CommandLineSwitchesTests.RuntimeSwitchInTheValuesPlace_IsNoValue
        Case 32: Test = "CommandLineSwitchesTests.SwitchesItIsNotAskedAbout_AreIgnored": If run Then CommandLineSwitchesTests.SwitchesItIsNotAskedAbout_AreIgnored
        Case Else: Test = vbNullString
    End Select
End Function

'@Description("Enters the test's module, where it is not the current one, and starts the test.")
Private Sub BeforeTest(ByVal name As String)
Attribute BeforeTest.VB_Description = "Enters the test's module, where it is not the current one, and starts the test."
    EnterModule Left$(name, InStr(name, ".") - 1)
    If this.CurrentModule = "ShipTests" Then ShipTests.TestInitialize
    TestAssert.Reset
End Sub

'@Description("Ends the test, and prints and counts its outcome.")
Private Sub AfterTest(ByVal name As String, ByVal outcome As TestOutcome, ByVal reason As String)
Attribute AfterTest.VB_Description = "Ends the test, and prints and counts its outcome."
    If this.CurrentModule = "ShipTests" Then ShipTests.TestCleanup
    Select Case outcome
        Case TestPassed
            this.Passed = this.Passed + 1
            Debug.Print "Passed        " & name
        Case TestFailed
            this.Failed = this.Failed + 1
            Debug.Print "Failed        " & name & ": " & reason
        Case TestInconclusive
            this.Inconclusive = this.Inconclusive + 1
            Debug.Print "Inconclusive  " & name & ": " & reason
    End Select
End Sub

'@Description("Cleans up the current module and initializes the next one, where the module changes; an empty name only cleans up.")
Private Sub EnterModule(ByVal moduleName As String)
Attribute EnterModule.VB_Description = "Cleans up the current module and initializes the next one, where the module changes; an empty name only cleans up."
    If moduleName = this.CurrentModule Then Exit Sub
    Select Case this.CurrentModule
        Case "PlayerGridTests": PlayerGridTests.ModuleCleanup
        Case "ShipTests": ShipTests.ModuleCleanup
        Case "CommandLineSwitchesTests": CommandLineSwitchesTests.ModuleCleanup
    End Select
    this.CurrentModule = moduleName
    Select Case moduleName
        Case "PlayerGridTests": PlayerGridTests.ModuleInitialize
        Case "ShipTests": ShipTests.ModuleInitialize
        Case "CommandLineSwitchesTests": CommandLineSwitchesTests.ModuleInitialize
    End Select
End Sub
