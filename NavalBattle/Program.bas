Attribute VB_Name = "Program"
Attribute VB_Description = "The program's entry point."
'@Folder("NavalBattle")
'@ModuleDescription("The program's entry point.")
Option Explicit

Private Const ClassNotRegistered As Long = 429
Private Const WatchableAIPause As Long = 800

Private controller As IGameController

'@Description("Plays a game: in the RDForms window, else in the console, on every platform; --msforms plays in the Microsoft Forms window instead, to compare the two windows, and --console plays in the console alone. The command line can name the players, --player1 <player> --player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: --seed <number>. The form opens with the players it names already picked; the console asks for a player it does not name. --test runs the tests instead of a game, in the console. Under --unattended, nobody watches: the computer plays without its pause before each shot.")
Public Sub Main()
Attribute Main.VB_Description = "Plays a game: in the RDForms window, else in the console, on every platform; --msforms plays in the Microsoft Forms window instead, to compare the two windows, and --console plays in the console alone. The command line can name the players, --player1 <player> --player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: --seed <number>. The form opens with the players it names already picked; the console asks for a player it does not name. --test runs the tests instead of a game, in the console. Under --unattended, nobody watches: the computer plays without its pause before each shot."
    If HasSwitch(Command$, "--test") Then
        TestRunner.RunAll
        Exit Sub
    End If

    Dim seed As String
    If TryGetArgument(Command$, "--seed", seed) And IsNumeric(seed) Then
        Rnd -1
        Randomize CLng(seed)
    Else
        Randomize
    End If

    Dim firstPlayer As String
    TryGetArgument Command$, "--player1", firstPlayer
    Dim secondPlayer As String
    TryGetArgument Command$, "--player2", secondPlayer

    Dim viewName As Variant
    For Each viewName In GameViews(Command$)
        If viewName = ViewConsole Then
            PlayInConsole firstPlayer, secondPlayer
            Exit Sub
        End If
        If TryPlayInWindow(viewName, firstPlayer, secondPlayer) Then Exit Sub
    Next
End Sub

'@Description("Plays games in the given window, Microsoft Forms' or RDForms', one after the other while the person asks for a new one, and returns True when it is closed. Returns False when the window cannot be made or fails before it opens, whatever the error, and says why: the game is then played in the next view. An error once the window is open is the game's, and is raised.")
Private Function TryPlayInWindow(ByVal viewName As String, ByVal firstPlayer As String, ByVal secondPlayer As String) As Boolean
Attribute TryPlayInWindow.VB_Description = "Plays games in the given window, Microsoft Forms' or RDForms', one after the other while the person asks for a new one, and returns True when it is closed. Returns False when the window cannot be made or fails before it opens, whatever the error, and says why: the game is then played in the next view. An error once the window is open is the game's, and is raised."
    Dim view As WindowView
    Do
        If Not TryMakeWindow(viewName, view) Then Exit Function

        StartGame GridViewAdapter.Create(view)
        If Not view.TryPlay(firstPlayer, secondPlayer) Then
            ReportNoWindow viewName, view.FailureNumber, view.FailureDescription
            Exit Function
        End If
    Loop While view.WantsNewGame
    TryPlayInWindow = True
End Function

'@Description("Makes the view of the given window, and returns True; returns False, and says why, when the window cannot be made, whatever the error: Microsoft Forms is missing as soon as its form is made.")
Private Function TryMakeWindow(ByVal viewName As String, ByRef outView As WindowView) As Boolean
Attribute TryMakeWindow.VB_Description = "Makes the view of the given window, and returns True; returns False, and says why, when the window cannot be made, whatever the error: Microsoft Forms is missing as soon as its form is made."
    On Error GoTo NoWindow
    If viewName = ViewMicrosoftForms Then
        Set outView = WindowView.Create(New GameForm)
    Else
        Set outView = WindowView.Create(New BoardForm)
    End If
    TryMakeWindow = True
    Exit Function

NoWindow:
    ReportNoWindow viewName, Err.Number, Err.Description
    TryMakeWindow = False
End Function

'@Description("Says why the given window could not be opened here, before the game is played in the next view.")
Private Sub ReportNoWindow(ByVal viewName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
Attribute ReportNoWindow.VB_Description = "Says why the given window could not be opened here, before the game is played in the next view."
    If errorNumber = ClassNotRegistered And viewName = ViewMicrosoftForms Then
        Debug.Print "The game form needs Microsoft Forms, which is not installed here."
    ElseIf errorNumber = ClassNotRegistered Then
        Debug.Print "The RDForms window cannot be opened here: this runtime has no forms host, or its windows cannot start."
    Else
        Debug.Print "The " & IIf(viewName = ViewMicrosoftForms, "Microsoft Forms", "RDForms") & " window could not be opened here: error " & errorNumber & ", " & errorDescription
    End If
End Sub

'@Description("Plays a game in the console.")
Private Sub PlayInConsole(ByVal firstPlayer As String, ByVal secondPlayer As String)
Attribute PlayInConsole.VB_Description = "Plays a game in the console."
    Dim view As ConsoleView
    Set view = New ConsoleView

    StartGame GridViewAdapter.Create(view)
    view.Play firstPlayer, secondPlayer
End Sub

'@Description("Makes the controller of a new game, over the given view's adapter.")
Private Sub StartGame(ByVal adapter As GridViewAdapter)
Attribute StartGame.VB_Description = "Makes the controller of a new game, over the given view's adapter."
    Dim randomizer As IRandomizer
    Set randomizer = New GameRandomizer

    ' nobody watches an unattended run: the computer plays at once.
    Dim aiPause As Long
    If Not HasSwitch(Command$, "--unattended") Then aiPause = WatchableAIPause

    Set controller = StandardGameController.Create(adapter, randomizer, PlayerFactory.Create(randomizer, aiPause))
    controller.NewGame
End Sub
