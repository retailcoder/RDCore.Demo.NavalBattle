Attribute VB_Name = "Program"
Attribute VB_Description = "The program's entry point."
'@Folder("NavalBattle")
'@ModuleDescription("The program's entry point.")
Option Explicit

Private Const ClassNotRegistered As Long = 429
Private Const WatchableAIPause As Long = 800

Private controller As IGameController

'@Description("Plays a game: in the Microsoft Forms window on Windows, else in the console; --console plays in the console alone. The command line can name the players, --player1 <player> --player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: --seed <number>. The form opens with the players it names already picked; the console asks for a player it does not name. --test runs the tests instead of a game, in the console. Under --unattended, nobody watches: the computer plays without its pause before each shot.")
Public Sub Main()
Attribute Main.VB_Description = "Plays a game: in the Microsoft Forms window on Windows, else in the console; --console plays in the console alone. The command line can name the players, --player1 <player> --player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: --seed <number>. The form opens with the players it names already picked; the console asks for a player it does not name. --test runs the tests instead of a game, in the console. Under --unattended, nobody watches: the computer plays without its pause before each shot."
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

    Dim onWindows As Boolean
#If Win32 Or Win64 Then
    onWindows = True
#End If
    Dim viewName As Variant
    For Each viewName In GameViews(Command$, onWindows)
        If viewName = ViewConsole Then
            PlayInConsole firstPlayer, secondPlayer
            Exit Sub
        End If
        If TryPlayInWindow(viewName, firstPlayer, secondPlayer) Then Exit Sub
    Next
End Sub

'@Description("Plays a game in the Microsoft Forms window, and returns True when it is closed; returns False when Microsoft Forms is not installed here.")
Private Function TryPlayInWindow(ByVal viewName As String, ByVal firstPlayer As String, ByVal secondPlayer As String) As Boolean
Attribute TryPlayInWindow.VB_Description = "Plays a game in the Microsoft Forms window, and returns True when it is closed; returns False when Microsoft Forms is not installed here."
    Dim view As WindowView
    ' Microsoft Forms is missing as soon as its form is made.
    On Error GoTo NoWindow
    Set view = WindowView.Create(New GameForm)

    StartGame GridViewAdapter.Create(view)
    view.Play firstPlayer, secondPlayer
    TryPlayInWindow = True
    Exit Function

NoWindow:
    If Err.Number <> ClassNotRegistered Then Err.Raise Err.Number, Err.Source, Err.Description
    Debug.Print "The game form needs Microsoft Forms, which is not installed here."
    TryPlayInWindow = False
End Function

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
