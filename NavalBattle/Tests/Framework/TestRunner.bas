Attribute VB_Name = "TestRunner"
Attribute VB_Description = "Runs the test modules: each test by name, with its module's initialize and cleanup procedures, and prints each outcome and a summary. A run with a failed test ends with an error."
'@Folder("NavalBattle.Tests.Framework")
'@ModuleDescription("Runs the test modules: each test by name, with its module's initialize and cleanup procedures, and prints each outcome and a summary. A run with a failed test ends with an error.")
Option Explicit

Private Type TRun
    CurrentModule As String
    Passed As Long
    Failed As Long
    Inconclusive As Long
    Ignored As Long
End Type

Private this As TRun

'@Description("Runs every test, prints each outcome and a summary, and raises an error where a test failed.")
Public Sub RunAll()
Attribute RunAll.VB_Description = "Runs every test, prints each outcome and a summary, and raises an error where a test failed."
    this.CurrentModule = vbNullString
    this.Passed = 0
    this.Failed = 0
    this.Inconclusive = 0
    this.Ignored = 0

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

    Debug.Print TestCount & " tests: " & this.Passed & " passed, " & this.Failed & " failed, " & this.Inconclusive & " inconclusive, " & this.Ignored & " ignored."
    If this.Failed > 0 Then Err.Raise vbObjectError + 1, "TestRunner", this.Failed & " of " & TestCount & " tests failed."
    Exit Sub

TestError:
    AfterTest name, TestFailed, "Error " & Err.Number & ": " & Err.Description
    Resume NextTest
End Sub

'@Description("The number of tests.")
Private Property Get TestCount() As Long
Attribute TestCount.VB_Description = "The number of tests."
    TestCount = 92
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
        Case 33: Test = "CommandLineSwitchesTests.NoViewSwitch_TriesRDFormsThenTheConsole": If run Then CommandLineSwitchesTests.NoViewSwitch_TriesRDFormsThenTheConsole
        Case 34: Test = "CommandLineSwitchesTests.RDFormsSwitch_TriesRDFormsThenTheConsole": If run Then CommandLineSwitchesTests.RDFormsSwitch_TriesRDFormsThenTheConsole
        Case 35: Test = "CommandLineSwitchesTests.MSFormsSwitch_TriesMicrosoftFormsThenTheConsole": If run Then CommandLineSwitchesTests.MSFormsSwitch_TriesMicrosoftFormsThenTheConsole
        Case 36: Test = "CommandLineSwitchesTests.ConsoleSwitch_PlaysInTheConsoleAlone_EvenWithAWindowSwitch": If run Then CommandLineSwitchesTests.ConsoleSwitch_PlaysInTheConsoleAlone_EvenWithAWindowSwitch
        Case 37: Test = "GameControllerTests.ComputerVsComputer_OnSinkNamesEachSunkShip": If run Then GameControllerTests.ComputerVsComputer_OnSinkNamesEachSunkShip
        Case 38: Test = "GameControllerTests.ComputerVsComputer_EachSinkingIsReportedWithTheShotThatSankTheShip": If run Then GameControllerTests.ComputerVsComputer_EachSinkingIsReportedWithTheShotThatSankTheShip
        Case 39: Test = "PlayerChoicesTests.FirstLetter_IsEnough_InEitherCase": If run Then PlayerChoicesTests.FirstLetter_IsEnough_InEitherCase
        Case 40: Test = "PlayerChoicesTests.NotAPlayer_IsNotRead": If run Then PlayerChoicesTests.NotAPlayer_IsNotRead
        Case 41: Test = "PlayerChoicesTests.EveryName_ReadsBackAsItsPlayer": If run Then PlayerChoicesTests.EveryName_ReadsBackAsItsPlayer
        Case 42: Test = "PlayerChoicesTests.TwoPeople_CannotPlayTogether": If run Then PlayerChoicesTests.TwoPeople_CannotPlayTogether
        Case 43: Test = "GameControllerTests.ComputerVsComputer_BothGridsStayLocked": If run Then GameControllerTests.ComputerVsComputer_BothGridsStayLocked
        Case 44: Test = "GameControllerTests.PersonPlacingShips_OnlyTheirGridIsUnlocked": If run Then GameControllerTests.PersonPlacingShips_OnlyTheirGridIsUnlocked
        Case 45: Test = "GameControllerTests.PersonFiring_OnlyTheOtherGridIsUnlocked": If run Then GameControllerTests.PersonFiring_OnlyTheOtherGridIsUnlocked
        Case 46: Test = "ShipTests.CreateInside_PastTheRightEdge_AnchorsWhereTheShipFits": If run Then ShipTests.CreateInside_PastTheRightEdge_AnchorsWhereTheShipFits
        Case 47: Test = "ShipTests.CreateInside_PastTheBottomEdge_AnchorsWhereTheShipFits": If run Then ShipTests.CreateInside_PastTheBottomEdge_AnchorsWhereTheShipFits
        Case 48: Test = "ShipTests.CreateInside_WhereTheShipFits_KeepsThePosition": If run Then ShipTests.CreateInside_WhereTheShipFits_KeepsThePosition
        Case 49: Test = "BoardLayoutTests.StatusBar_ShowsOnlyItsBottomBorder": If run Then BoardLayoutTests.StatusBar_ShowsOnlyItsBottomBorder
        Case 50: Test = "BoardLayoutTests.StatusText_IsInsideTheBar": If run Then BoardLayoutTests.StatusText_IsInsideTheBar
        Case 51: Test = "BoardLayoutTests.Columns_AreSideBySide_CentredUnderTheBanner": If run Then BoardLayoutTests.Columns_AreSideBySide_CentredUnderTheBanner
        Case 52: Test = "BoardLayoutTests.Cells_FollowTheRowNumbers_AndTheLettersFollowTheLastRow": If run Then BoardLayoutTests.Cells_FollowTheRowNumbers_AndTheLettersFollowTheLastRow
        Case 53: Test = "BoardLayoutTests.Captions_AreOverTheirCells_OnTheBanner_AndHoldThePickers": If run Then BoardLayoutTests.Captions_AreOverTheirCells_OnTheBanner_AndHoldThePickers
        Case 54: Test = "BoardInputTests.LockedGrid_TakesNoClicks": If run Then BoardInputTests.LockedGrid_TakesNoClicks
        Case 55: Test = "BoardInputTests.ShipPlacementClick_PreviewsThenConfirms": If run Then BoardInputTests.ShipPlacementClick_PreviewsThenConfirms
        Case 56: Test = "BoardInputTests.ShipPlacementHover_PreviewsOncePerCell": If run Then BoardInputTests.ShipPlacementHover_PreviewsOncePerCell
        Case 57: Test = "BoardInputTests.ShipPlacementRightClick_TurnsTheShip": If run Then BoardInputTests.ShipPlacementRightClick_TurnsTheShip
        Case 58: Test = "BoardInputTests.Ready_ConfirmsThePlayers_AndMakesThem": If run Then BoardInputTests.Ready_ConfirmsThePlayers_AndMakesThem
        Case 59: Test = "BoardInputTests.Ready_TwoPeople_AreRefused": If run Then BoardInputTests.Ready_TwoPeople_AreRefused
        Case 60: Test = "BoardInputTests.Ready_TwoComputers_PlayTheWholeGame": If run Then BoardInputTests.Ready_TwoComputers_PlayTheWholeGame
        Case 61: Test = "BoardInputTests.PlayerPicked_SaysWhoPlays_UntilTheGameStarts": If run Then BoardInputTests.PlayerPicked_SaysWhoPlays_UntilTheGameStarts
        Case 62: Test = "BoardInputTests.ReadyOnceShipsArePlaced_StartsTheBattle": If run Then BoardInputTests.ReadyOnceShipsArePlaced_StartsTheBattle
        Case 63: Test = "BoardInputTests.ComputerTurn_TakesNoClicks": If run Then BoardInputTests.ComputerTurn_TakesNoClicks
        Case 64: Test = "BoardInputTests.EndGame_HidesReady_AndTakesNoClicks": If run Then BoardInputTests.EndGame_HidesReady_AndTakesNoClicks
        Case 65: Test = "BoardSpritesTests.OpenWater_ShowsNoSprite": If run Then BoardSpritesTests.OpenWater_ShowsNoSprite
        Case 66: Test = "BoardSpritesTests.HitsAndMisses_ShowWhereTheyWerePlayed": If run Then BoardSpritesTests.HitsAndMisses_ShowWhereTheyWerePlayed
        Case 67: Test = "BoardSpritesTests.Ships_ShowTheirDucks_TheBigOneAtTheBow": If run Then BoardSpritesTests.Ships_ShowTheirDucks_TheBigOneAtTheBow
        Case 68: Test = "BoardSpritesTests.HiddenShips_ShowOnlyTheirHits_AsTheDevilDuck": If run Then BoardSpritesTests.HiddenShips_ShowOnlyTheirHits_AsTheDevilDuck
        Case 69: Test = "BoardSpritesTests.DizzyShip_ShowsItsBowDizzy_UntilItRecovers": If run Then BoardSpritesTests.DizzyShip_ShowsItsBowDizzy_UntilItRecovers
        Case 70: Test = "BoardSpritesTests.SunkShip_ShowsDizzy_ThenItsRedBigDuck_EvenHidden": If run Then BoardSpritesTests.SunkShip_ShowsDizzy_ThenItsRedBigDuck_EvenHidden
        Case 71: Test = "BoardSpritesTests.FallbackColours_TellWaterShipsHitsAndMissesApart": If run Then BoardSpritesTests.FallbackColours_TellWaterShipsHitsAndMissesApart
        Case 72: Test = "BoardSpritesTests.EverySprite_IsAPngInTheWorkspaceAssets": If run Then BoardSpritesTests.EverySprite_IsAPngInTheWorkspaceAssets
        Case 73: Test = "BoardFormTests.Cells_AreLoadedFromTheDesignedCell": If run Then BoardFormTests.Cells_AreLoadedFromTheDesignedCell
        Case 74: Test = "BoardFormTests.PlayerPickers_OfferEveryPlayer": If run Then BoardFormTests.PlayerPickers_OfferEveryPlayer
        Case 75: Test = "BoardFormTests.PickedPlayer_ShowsInItsGridsPicker": If run Then BoardFormTests.PickedPlayer_ShowsInItsGridsPicker
        Case 76: Test = "BoardLayoutTests.FleetShips_AreUnderTheirBanner_OneUnderTheOther_InsideTheWindow": If run Then BoardLayoutTests.FleetShips_AreUnderTheirBanner_OneUnderTheOther_InsideTheWindow
        Case 77: Test = "BoardLayoutTests.Buttons_AreARow_AtTheBottomLeft_UnderTheGrids": If run Then BoardLayoutTests.Buttons_AreARow_AtTheBottomLeft_UnderTheGrids
        Case 78: Test = "BoardSpritesTests.FleetGrid_IsTheFirstGridPaintedWithItsShips": If run Then BoardSpritesTests.FleetGrid_IsTheFirstGridPaintedWithItsShips
        Case 79: Test = "BoardSpritesTests.FleetSprites_ShowEachShipAcross_InTheOrderTheyWerePlaced": If run Then BoardSpritesTests.FleetSprites_ShowEachShipAcross_InTheOrderTheyWerePlaced
        Case 80: Test = "BoardSpritesTests.FleetSprites_ShowTheHitDucksRed_AndNoOtherSprite": If run Then BoardSpritesTests.FleetSprites_ShowTheHitDucksRed_AndNoOtherSprite
        Case 81: Test = "BoardFormTests.FleetStatus_ShowsEveryShipPlaced": If run Then BoardFormTests.FleetStatus_ShowsEveryShipPlaced
        Case 82: Test = "BoardLayoutTests.Banner_IsAcrossTheWindow_OnTheStatusBarsBottomBorder": If run Then BoardLayoutTests.Banner_IsAcrossTheWindow_OnTheStatusBarsBottomBorder
        Case 83: Test = "BoardLayoutTests.Background_IsUnderTheBanner_AndTheWindowEndsWithIt": If run Then BoardLayoutTests.Background_IsUnderTheBanner_AndTheWindowEndsWithIt
        Case 84: Test = "ShipTests.DefaultNameChoices_AreFiveAKind_NoNameTwice": If run Then ShipTests.DefaultNameChoices_AreFiveAKind_NoNameTwice
        Case 85: Test = "ShipTests.PickNames_NamesTheShipsCreatedAfterIt_AmongTheirKindsChoices": If run Then ShipTests.PickNames_NamesTheShipsCreatedAfterIt_AmongTheirKindsChoices
        Case 86: Test = "BoardSpritesTests.Squeak_ShowsOnTheHitCell_OfItsGridOnly_UntilItStops": If run Then BoardSpritesTests.Squeak_ShowsOnTheHitCell_OfItsGridOnly_UntilItStops
        Case 87: Test = "BoardInputTests.TakesClicks_WhereAClickDoesSomething": If run Then BoardInputTests.TakesClicks_WhereAClickDoesSomething
        Case 88: Test = "ShipTests.KindCaption_NamesTheKind_AndItsDucks": If run Then ShipTests.KindCaption_NamesTheKind_AndItsDucks
        Case 89: Test = "BoardSpritesTests.SunkShip_ItsBowHitBefore_StaysRed_NotDizzy": If run Then BoardSpritesTests.SunkShip_ItsBowHitBefore_StaysRed_NotDizzy
        Case 90: Test = "BoardSpritesTests.DizzyShip_ADucklingHit_TheHelmIsDizzy_OnlyWhileItIsYellow": If run Then BoardSpritesTests.DizzyShip_ADucklingHit_TheHelmIsDizzy_OnlyWhileItIsYellow
        Case 91: Test = "WindowViewTests.WindowFailingBeforeItOpens_TryPlayReturnsFalse_WithTheError": If run Then WindowViewTests.WindowFailingBeforeItOpens_TryPlayReturnsFalse_WithTheError
        Case 92: Test = "WindowViewTests.WindowFailingOnceOpen_RaisesTheGamesError": If run Then WindowViewTests.WindowFailingOnceOpen_RaisesTheGamesError
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
        Case TestIgnored
            this.Ignored = this.Ignored + 1
            Debug.Print "Ignored       " & name & ": " & reason
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
        Case "GameControllerTests": GameControllerTests.ModuleCleanup
        Case "PlayerChoicesTests": PlayerChoicesTests.ModuleCleanup
        Case "BoardLayoutTests": BoardLayoutTests.ModuleCleanup
        Case "BoardInputTests": BoardInputTests.ModuleCleanup
        Case "BoardSpritesTests": BoardSpritesTests.ModuleCleanup
        Case "BoardFormTests": BoardFormTests.ModuleCleanup
        Case "WindowViewTests": WindowViewTests.ModuleCleanup
    End Select
    this.CurrentModule = moduleName
    Select Case moduleName
        Case "PlayerGridTests": PlayerGridTests.ModuleInitialize
        Case "ShipTests": ShipTests.ModuleInitialize
        Case "CommandLineSwitchesTests": CommandLineSwitchesTests.ModuleInitialize
        Case "GameControllerTests": GameControllerTests.ModuleInitialize
        Case "PlayerChoicesTests": PlayerChoicesTests.ModuleInitialize
        Case "BoardLayoutTests": BoardLayoutTests.ModuleInitialize
        Case "BoardInputTests": BoardInputTests.ModuleInitialize
        Case "BoardSpritesTests": BoardSpritesTests.ModuleInitialize
        Case "BoardFormTests": BoardFormTests.ModuleInitialize
        Case "WindowViewTests": WindowViewTests.ModuleInitialize
    End Select
End Sub
