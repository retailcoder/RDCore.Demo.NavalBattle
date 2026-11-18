VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WindowTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "The window app's own tests: the window's layout, input and sprites, its RDForms form, and the view it plays in."
'@Folder("NavalBattle.Tests")
'@ModuleDescription("The window app's own tests: the window's layout, input and sprites, its RDForms form, and the view it plays in.")
Option Explicit
Implements ITestList

Private Property Get ITestList_Count() As Long
    ITestList_Count = 43
End Property

Private Function ITestList_Test(ByVal index As Long, ByVal run As Boolean) As String
    Select Case index
        Case 1: ITestList_Test = "BoardLayoutTests.StatusBar_ShowsOnlyItsBottomBorder": If run Then BoardLayoutTests.StatusBar_ShowsOnlyItsBottomBorder
        Case 2: ITestList_Test = "BoardLayoutTests.StatusText_IsInsideTheBar": If run Then BoardLayoutTests.StatusText_IsInsideTheBar
        Case 3: ITestList_Test = "BoardLayoutTests.Columns_AreSideBySide_CentredUnderTheBanner": If run Then BoardLayoutTests.Columns_AreSideBySide_CentredUnderTheBanner
        Case 4: ITestList_Test = "BoardLayoutTests.Cells_FollowTheRowNumbers_AndTheLettersFollowTheLastRow": If run Then BoardLayoutTests.Cells_FollowTheRowNumbers_AndTheLettersFollowTheLastRow
        Case 5: ITestList_Test = "BoardLayoutTests.Captions_AreOverTheirCells_OnTheBanner_AndHoldThePickers": If run Then BoardLayoutTests.Captions_AreOverTheirCells_OnTheBanner_AndHoldThePickers
        Case 6: ITestList_Test = "BoardInputTests.LockedGrid_TakesNoClicks": If run Then BoardInputTests.LockedGrid_TakesNoClicks
        Case 7: ITestList_Test = "BoardInputTests.ShipPlacementClick_PreviewsThenConfirms": If run Then BoardInputTests.ShipPlacementClick_PreviewsThenConfirms
        Case 8: ITestList_Test = "BoardInputTests.ShipPlacementHover_PreviewsOncePerCell": If run Then BoardInputTests.ShipPlacementHover_PreviewsOncePerCell
        Case 9: ITestList_Test = "BoardInputTests.ShipPlacementRightClick_TurnsTheShip": If run Then BoardInputTests.ShipPlacementRightClick_TurnsTheShip
        Case 10: ITestList_Test = "BoardInputTests.Ready_ConfirmsThePlayers_AndMakesThem": If run Then BoardInputTests.Ready_ConfirmsThePlayers_AndMakesThem
        Case 11: ITestList_Test = "BoardInputTests.Ready_TwoPeople_AreRefused": If run Then BoardInputTests.Ready_TwoPeople_AreRefused
        Case 12: ITestList_Test = "BoardInputTests.Ready_TwoComputers_PlayTheWholeGame": If run Then BoardInputTests.Ready_TwoComputers_PlayTheWholeGame
        Case 13: ITestList_Test = "BoardInputTests.PlayerPicked_SaysWhoPlays_UntilTheGameStarts": If run Then BoardInputTests.PlayerPicked_SaysWhoPlays_UntilTheGameStarts
        Case 14: ITestList_Test = "BoardInputTests.ReadyOnceShipsArePlaced_StartsTheBattle": If run Then BoardInputTests.ReadyOnceShipsArePlaced_StartsTheBattle
        Case 15: ITestList_Test = "BoardInputTests.ComputerTurn_TakesNoClicks": If run Then BoardInputTests.ComputerTurn_TakesNoClicks
        Case 16: ITestList_Test = "BoardInputTests.EndGame_HidesReady_AndTakesNoClicks": If run Then BoardInputTests.EndGame_HidesReady_AndTakesNoClicks
        Case 17: ITestList_Test = "BoardSpritesTests.OpenWater_ShowsNoSprite": If run Then BoardSpritesTests.OpenWater_ShowsNoSprite
        Case 18: ITestList_Test = "BoardSpritesTests.HitsAndMisses_ShowWhereTheyWerePlayed": If run Then BoardSpritesTests.HitsAndMisses_ShowWhereTheyWerePlayed
        Case 19: ITestList_Test = "BoardSpritesTests.Ships_ShowTheirDucks_TheBigOneAtTheBow": If run Then BoardSpritesTests.Ships_ShowTheirDucks_TheBigOneAtTheBow
        Case 20: ITestList_Test = "BoardSpritesTests.HiddenShips_ShowOnlyTheirHits_AsTheDevilDuck": If run Then BoardSpritesTests.HiddenShips_ShowOnlyTheirHits_AsTheDevilDuck
        Case 21: ITestList_Test = "BoardSpritesTests.DizzyShip_ShowsItsBowDizzy_UntilItRecovers": If run Then BoardSpritesTests.DizzyShip_ShowsItsBowDizzy_UntilItRecovers
        Case 22: ITestList_Test = "BoardSpritesTests.SunkShip_ShowsDizzy_ThenItsRedBigDuck_EvenHidden": If run Then BoardSpritesTests.SunkShip_ShowsDizzy_ThenItsRedBigDuck_EvenHidden
        Case 23: ITestList_Test = "BoardSpritesTests.FallbackColours_TellWaterShipsHitsAndMissesApart": If run Then BoardSpritesTests.FallbackColours_TellWaterShipsHitsAndMissesApart
        Case 24: ITestList_Test = "BoardSpritesTests.EverySprite_IsAPngInTheAppsAssets": If run Then BoardSpritesTests.EverySprite_IsAPngInTheAppsAssets
        Case 25: ITestList_Test = "BoardFormTests.Cells_AreLoadedFromTheDesignedCell": If run Then BoardFormTests.Cells_AreLoadedFromTheDesignedCell
        Case 26: ITestList_Test = "BoardFormTests.PlayerPickers_OfferEveryPlayer": If run Then BoardFormTests.PlayerPickers_OfferEveryPlayer
        Case 27: ITestList_Test = "BoardFormTests.PickedPlayer_ShowsInItsGridsPicker": If run Then BoardFormTests.PickedPlayer_ShowsInItsGridsPicker
        Case 28: ITestList_Test = "BoardLayoutTests.FleetShips_AreUnderTheirBanner_OneUnderTheOther_InsideTheWindow": If run Then BoardLayoutTests.FleetShips_AreUnderTheirBanner_OneUnderTheOther_InsideTheWindow
        Case 29: ITestList_Test = "BoardLayoutTests.Buttons_AreARow_AtTheBottomLeft_UnderTheGrids": If run Then BoardLayoutTests.Buttons_AreARow_AtTheBottomLeft_UnderTheGrids
        Case 30: ITestList_Test = "BoardSpritesTests.FleetGrid_IsTheFirstGridPaintedWithItsShips": If run Then BoardSpritesTests.FleetGrid_IsTheFirstGridPaintedWithItsShips
        Case 31: ITestList_Test = "BoardSpritesTests.FleetSprites_ShowEachShipAcross_InTheOrderTheyWerePlaced": If run Then BoardSpritesTests.FleetSprites_ShowEachShipAcross_InTheOrderTheyWerePlaced
        Case 32: ITestList_Test = "BoardSpritesTests.FleetSprites_ShowTheHitDucksRed_AndNoOtherSprite": If run Then BoardSpritesTests.FleetSprites_ShowTheHitDucksRed_AndNoOtherSprite
        Case 33: ITestList_Test = "BoardFormTests.FleetStatus_ShowsEveryShipPlaced": If run Then BoardFormTests.FleetStatus_ShowsEveryShipPlaced
        Case 34: ITestList_Test = "BoardLayoutTests.Banner_IsAcrossTheWindow_OnTheStatusBarsBottomBorder": If run Then BoardLayoutTests.Banner_IsAcrossTheWindow_OnTheStatusBarsBottomBorder
        Case 35: ITestList_Test = "BoardLayoutTests.Background_IsUnderTheBanner_AndTheWindowEndsWithIt": If run Then BoardLayoutTests.Background_IsUnderTheBanner_AndTheWindowEndsWithIt
        Case 36: ITestList_Test = "BoardSpritesTests.Squeak_ShowsOnTheHitCell_OfItsGridOnly_UntilItStops": If run Then BoardSpritesTests.Squeak_ShowsOnTheHitCell_OfItsGridOnly_UntilItStops
        Case 37: ITestList_Test = "BoardInputTests.TakesClicks_WhereAClickDoesSomething": If run Then BoardInputTests.TakesClicks_WhereAClickDoesSomething
        Case 38: ITestList_Test = "BoardSpritesTests.SunkShip_ItsBowHitBefore_StaysRed_NotDizzy": If run Then BoardSpritesTests.SunkShip_ItsBowHitBefore_StaysRed_NotDizzy
        Case 39: ITestList_Test = "BoardSpritesTests.DizzyShip_ADucklingHit_TheHelmIsDizzy_OnlyWhileItIsYellow": If run Then BoardSpritesTests.DizzyShip_ADucklingHit_TheHelmIsDizzy_OnlyWhileItIsYellow
        Case 40: ITestList_Test = "GameViewsTests.NoViewSwitch_PlaysInTheRDFormsWindow": If run Then GameViewsTests.NoViewSwitch_PlaysInTheRDFormsWindow
        Case 41: ITestList_Test = "GameViewsTests.RDFormsSwitch_PlaysInTheRDFormsWindow": If run Then GameViewsTests.RDFormsSwitch_PlaysInTheRDFormsWindow
        Case 42: ITestList_Test = "GameViewsTests.MSFormsSwitch_PlaysInTheMicrosoftFormsWindow": If run Then GameViewsTests.MSFormsSwitch_PlaysInTheMicrosoftFormsWindow
        Case 43: ITestList_Test = "GameViewsTests.ConsoleSwitch_IsNoView_TheConsoleAppPlaysThere": If run Then GameViewsTests.ConsoleSwitch_IsNoView_TheConsoleAppPlaysThere
        Case Else: ITestList_Test = vbNullString
    End Select
End Function

Private Sub ITestList_InitializeModule(ByVal moduleName As String)
    Select Case moduleName
        Case "BoardLayoutTests": BoardLayoutTests.ModuleInitialize
        Case "BoardInputTests": BoardInputTests.ModuleInitialize
        Case "BoardSpritesTests": BoardSpritesTests.ModuleInitialize
        Case "BoardFormTests": BoardFormTests.ModuleInitialize
        Case "GameViewsTests": GameViewsTests.ModuleInitialize
    End Select
End Sub

Private Sub ITestList_CleanupModule(ByVal moduleName As String)
    Select Case moduleName
        Case "BoardLayoutTests": BoardLayoutTests.ModuleCleanup
        Case "BoardInputTests": BoardInputTests.ModuleCleanup
        Case "BoardSpritesTests": BoardSpritesTests.ModuleCleanup
        Case "BoardFormTests": BoardFormTests.ModuleCleanup
        Case "GameViewsTests": GameViewsTests.ModuleCleanup
    End Select
End Sub

Private Sub ITestList_InitializeTest(ByVal moduleName As String)
End Sub

Private Sub ITestList_CleanupTest(ByVal moduleName As String)
End Sub
