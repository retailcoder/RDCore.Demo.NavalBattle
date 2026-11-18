VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibraryTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "The library's tests: the model, the controller and the command line."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("The library's tests: the model, the controller and the command line.")
Option Explicit
Implements ITestList

Private Property Get ITestList_Count() As Long
    ITestList_Count = 47
End Property

Private Function ITestList_Test(ByVal index As Long, ByVal run As Boolean) As String
    Select Case index
        Case 1: ITestList_Test = "PlayerGridTests.CanAddShipInsideGridBoundaries_ReturnsTrue": If run Then PlayerGridTests.CanAddShipInsideGridBoundaries_ReturnsTrue
        Case 2: ITestList_Test = "PlayerGridTests.CanAddShipAtPositionZeroZero_ReturnsFalse": If run Then PlayerGridTests.CanAddShipAtPositionZeroZero_ReturnsFalse
        Case 3: ITestList_Test = "PlayerGridTests.CanAddShipGivenInterectingShips_ReturnsFalse": If run Then PlayerGridTests.CanAddShipGivenInterectingShips_ReturnsFalse
        Case 4: ITestList_Test = "PlayerGridTests.AddingSameShipTypeTwice_Throws": If run Then PlayerGridTests.AddingSameShipTypeTwice_Throws
        Case 5: ITestList_Test = "PlayerGridTests.AddingShipOutsideGridBoundaries_Throws": If run Then PlayerGridTests.AddingShipOutsideGridBoundaries_Throws
        Case 6: ITestList_Test = "PlayerGridTests.TryHitKnownState_Throws": If run Then PlayerGridTests.TryHitKnownState_Throws
        Case 7: ITestList_Test = "PlayerGridTests.TryHitMiss_SetsPreviousMissState": If run Then PlayerGridTests.TryHitMiss_SetsPreviousMissState
        Case 8: ITestList_Test = "PlayerGridTests.TryHitSuccess_SetsPreviousHitState": If run Then PlayerGridTests.TryHitSuccess_SetsPreviousHitState
        Case 9: ITestList_Test = "PlayerGridTests.TryHitSuccess_ReturnsHit": If run Then PlayerGridTests.TryHitSuccess_ReturnsHit
        Case 10: ITestList_Test = "PlayerGridTests.TryHitMisses_ReturnsMiss": If run Then PlayerGridTests.TryHitMisses_ReturnsMiss
        Case 11: ITestList_Test = "PlayerGridTests.GridInitialState_UnknownState": If run Then PlayerGridTests.GridInitialState_UnknownState
        Case 12: ITestList_Test = "PlayerGridTests.GivenAdjacentShip_HasRightAdjacentShipReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShip_HasRightAdjacentShipReturnsTrue
        Case 13: ITestList_Test = "PlayerGridTests.GivenAdjacentShip_HasLeftAdjacentShipReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShip_HasLeftAdjacentShipReturnsTrue
        Case 14: ITestList_Test = "PlayerGridTests.GivenAdjacentShip_HasDownAdjacentShipReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShip_HasDownAdjacentShipReturnsTrue
        Case 15: ITestList_Test = "PlayerGridTests.GivenAdjacentShip_HasUpAdjacentShipReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShip_HasUpAdjacentShipReturnsTrue
        Case 16: ITestList_Test = "PlayerGridTests.GivenAdjacentShipAtHorizontalTipEnd_ReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShipAtHorizontalTipEnd_ReturnsTrue
        Case 17: ITestList_Test = "PlayerGridTests.GivenAdjacentShipAtVerticalTipEnd_ReturnsTrue": If run Then PlayerGridTests.GivenAdjacentShipAtVerticalTipEnd_ReturnsTrue
        Case 18: ITestList_Test = "PlayerGridTests.GivenTwoSideBySideHits_GetHitAreaReturnsTwoItems": If run Then PlayerGridTests.GivenTwoSideBySideHits_GetHitAreaReturnsTwoItems
        Case 19: ITestList_Test = "PlayerGridTests.GivenSunkHorizontalShip_IsSunkenOnEveryCell": If run Then PlayerGridTests.GivenSunkHorizontalShip_IsSunkenOnEveryCell
        Case 20: ITestList_Test = "PlayerGridTests.GivenSunkVerticalShip_IsSunkenOnEveryCell": If run Then PlayerGridTests.GivenSunkVerticalShip_IsSunkenOnEveryCell
        Case 21: ITestList_Test = "ShipTests.CreatesShipOfSpecifiedType": If run Then ShipTests.CreatesShipOfSpecifiedType
        Case 22: ITestList_Test = "ShipTests.SuccessfulHit_SetsStateToTrue": If run Then ShipTests.SuccessfulHit_SetsStateToTrue
        Case 23: ITestList_Test = "ShipTests.SunkenShip_IsSunken": If run Then ShipTests.SunkenShip_IsSunken
        Case 24: ITestList_Test = "ShipTests.MissedHit_ReturnsFalse": If run Then ShipTests.MissedHit_ReturnsFalse
        Case 25: ITestList_Test = "ShipTests.SuccessfulHit_ReturnsTrue": If run Then ShipTests.SuccessfulHit_ReturnsTrue
        Case 26: ITestList_Test = "ShipTests.ShipOrientationInvalidEnumValue_Throws": If run Then ShipTests.ShipOrientationInvalidEnumValue_Throws
        Case 27: ITestList_Test = "ShipTests.IntersectingShip_IntersectReturnsGridCoord": If run Then ShipTests.IntersectingShip_IntersectReturnsGridCoord
        Case 28: ITestList_Test = "ShipTests.NonIntersectingShip_IntersectReturnsNothing": If run Then ShipTests.NonIntersectingShip_IntersectReturnsNothing
        Case 29: ITestList_Test = "CommandLineSwitchesTests.RuntimeSwitchBefore_TheSeedIsRead": If run Then CommandLineSwitchesTests.RuntimeSwitchBefore_TheSeedIsRead
        Case 30: ITestList_Test = "CommandLineSwitchesTests.RuntimeSwitchAfter_TheSeedIsRead": If run Then CommandLineSwitchesTests.RuntimeSwitchAfter_TheSeedIsRead
        Case 31: ITestList_Test = "CommandLineSwitchesTests.RuntimeSwitchInTheValuesPlace_IsNoValue": If run Then CommandLineSwitchesTests.RuntimeSwitchInTheValuesPlace_IsNoValue
        Case 32: ITestList_Test = "CommandLineSwitchesTests.SwitchesItIsNotAskedAbout_AreIgnored": If run Then CommandLineSwitchesTests.SwitchesItIsNotAskedAbout_AreIgnored
        Case 33: ITestList_Test = "GameControllerTests.ComputerVsComputer_OnSinkNamesEachSunkShip": If run Then GameControllerTests.ComputerVsComputer_OnSinkNamesEachSunkShip
        Case 34: ITestList_Test = "GameControllerTests.ComputerVsComputer_EachSinkingIsReportedWithTheShotThatSankTheShip": If run Then GameControllerTests.ComputerVsComputer_EachSinkingIsReportedWithTheShotThatSankTheShip
        Case 35: ITestList_Test = "PlayerChoicesTests.FirstLetter_IsEnough_InEitherCase": If run Then PlayerChoicesTests.FirstLetter_IsEnough_InEitherCase
        Case 36: ITestList_Test = "PlayerChoicesTests.NotAPlayer_IsNotRead": If run Then PlayerChoicesTests.NotAPlayer_IsNotRead
        Case 37: ITestList_Test = "PlayerChoicesTests.EveryName_ReadsBackAsItsPlayer": If run Then PlayerChoicesTests.EveryName_ReadsBackAsItsPlayer
        Case 38: ITestList_Test = "PlayerChoicesTests.TwoPeople_CannotPlayTogether": If run Then PlayerChoicesTests.TwoPeople_CannotPlayTogether
        Case 39: ITestList_Test = "GameControllerTests.ComputerVsComputer_BothGridsStayLocked": If run Then GameControllerTests.ComputerVsComputer_BothGridsStayLocked
        Case 40: ITestList_Test = "GameControllerTests.PersonPlacingShips_OnlyTheirGridIsUnlocked": If run Then GameControllerTests.PersonPlacingShips_OnlyTheirGridIsUnlocked
        Case 41: ITestList_Test = "GameControllerTests.PersonFiring_OnlyTheOtherGridIsUnlocked": If run Then GameControllerTests.PersonFiring_OnlyTheOtherGridIsUnlocked
        Case 42: ITestList_Test = "ShipTests.CreateInside_PastTheRightEdge_AnchorsWhereTheShipFits": If run Then ShipTests.CreateInside_PastTheRightEdge_AnchorsWhereTheShipFits
        Case 43: ITestList_Test = "ShipTests.CreateInside_PastTheBottomEdge_AnchorsWhereTheShipFits": If run Then ShipTests.CreateInside_PastTheBottomEdge_AnchorsWhereTheShipFits
        Case 44: ITestList_Test = "ShipTests.CreateInside_WhereTheShipFits_KeepsThePosition": If run Then ShipTests.CreateInside_WhereTheShipFits_KeepsThePosition
        Case 45: ITestList_Test = "ShipTests.DefaultNameChoices_AreFiveAKind_NoNameTwice": If run Then ShipTests.DefaultNameChoices_AreFiveAKind_NoNameTwice
        Case 46: ITestList_Test = "ShipTests.PickNames_NamesTheShipsCreatedAfterIt_AmongTheirKindsChoices": If run Then ShipTests.PickNames_NamesTheShipsCreatedAfterIt_AmongTheirKindsChoices
        Case 47: ITestList_Test = "ShipTests.KindCaption_NamesTheKind_AndItsDucks": If run Then ShipTests.KindCaption_NamesTheKind_AndItsDucks
        Case Else: ITestList_Test = vbNullString
    End Select
End Function

Private Sub ITestList_InitializeModule(ByVal moduleName As String)
    Select Case moduleName
        Case "PlayerGridTests": PlayerGridTests.ModuleInitialize
        Case "ShipTests": ShipTests.ModuleInitialize
        Case "CommandLineSwitchesTests": CommandLineSwitchesTests.ModuleInitialize
        Case "GameControllerTests": GameControllerTests.ModuleInitialize
        Case "PlayerChoicesTests": PlayerChoicesTests.ModuleInitialize
    End Select
End Sub

Private Sub ITestList_CleanupModule(ByVal moduleName As String)
    Select Case moduleName
        Case "PlayerGridTests": PlayerGridTests.ModuleCleanup
        Case "ShipTests": ShipTests.ModuleCleanup
        Case "CommandLineSwitchesTests": CommandLineSwitchesTests.ModuleCleanup
        Case "GameControllerTests": GameControllerTests.ModuleCleanup
        Case "PlayerChoicesTests": PlayerChoicesTests.ModuleCleanup
    End Select
End Sub

Private Sub ITestList_InitializeTest(ByVal moduleName As String)
    Select Case moduleName
        Case "ShipTests": ShipTests.TestInitialize
    End Select
End Sub

Private Sub ITestList_CleanupTest(ByVal moduleName As String)
    Select Case moduleName
        Case "ShipTests": ShipTests.TestCleanup
    End Select
End Sub
