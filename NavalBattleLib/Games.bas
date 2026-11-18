Attribute VB_Name = "Games"
Attribute VB_Description = "What the apps that reference the library start a game with, and run its tests with. A game is played in a view of the app's own, through IGridViewCommands."
'@Folder("NavalBattle")
'@ModuleDescription("What the apps that reference the library start a game with, and run its tests with. A game is played in a view of the app's own, through IGridViewCommands.")
Option Explicit

Private Const WatchableAIPause As Long = 800

'@Description("Seeds the random numbers from the command line, so that a game can be replayed: --seed <number>. With no seed, the game is a new one each time.")
Public Sub SeedRandomNumbers(ByVal commandLine As String)
Attribute SeedRandomNumbers.VB_Description = "Seeds the random numbers from the command line, so that a game can be replayed: --seed <number>. With no seed, the game is a new one each time."
    Dim seed As String
    If TryGetArgument(commandLine, "--seed", seed) And IsNumeric(seed) Then
        Rnd -1
        Randomize CLng(seed)
    Else
        Randomize
    End If
End Sub

'@Description("Starts a new game in the given view, and returns its controller, which the caller keeps for as long as the game is played. Under --unattended, nobody watches: the computer plays without its pause before each shot.")
Public Function StartGame(ByVal view As IGridViewCommands) As IGameController
Attribute StartGame.VB_Description = "Starts a new game in the given view, and returns its controller, which the caller keeps for as long as the game is played. Under --unattended, nobody watches: the computer plays without its pause before each shot."
    Dim randomizer As IRandomizer
    Set randomizer = New GameRandomizer

    ' nobody watches an unattended run: the computer plays at once.
    Dim aiPause As Long
    If Not HasSwitch(Command$, "--unattended") Then aiPause = WatchableAIPause

    Dim controller As IGameController
    Set controller = StandardGameController.Create(GridViewAdapter.Create(view), randomizer, PlayerFactory.Create(randomizer, aiPause))
    controller.NewGame
    Set StartGame = controller
End Function

'@Description("Runs the library's tests, then the given app's own tests where there are any, prints each outcome and one summary, and raises an error where a test failed.")
Public Sub RunTests(Optional ByVal appTests As ITestList)
Attribute RunTests.VB_Description = "Runs the library's tests, then the given app's own tests where there are any, prints each outcome and one summary, and raises an error where a test failed."
    Dim lists As Collection
    Set lists = New Collection
    lists.Add New LibraryTests
    If Not appTests Is Nothing Then lists.Add appTests
    TestRunner.RunAll lists
End Sub
