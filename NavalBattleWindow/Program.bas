Attribute VB_Name = "Program"
Attribute VB_Description = "The window app's entry point."
'@Folder("NavalBattle")
'@ModuleDescription("The window app's entry point.")
Option Explicit

Private Const ClassNotRegistered As Long = 429

' the windows a game can be played in, as GameView names them.
Public Const ViewMicrosoftForms As String = "MSForms"
Public Const ViewRDForms As String = "RDForms"

Private controller As IGameController

'@Description("Plays games in the RDForms window, on every platform, one after the other while the person asks for a new one; --msforms plays in the Microsoft Forms window instead, to compare the two windows. The command line can name the players, --player1 <player> --player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: --seed <number>. The window opens with the players it names already picked. --test runs the library's tests and the window's own instead of a game, in the console. Under --unattended, nobody watches: the computer plays without its pause before each shot.")
Public Sub Main()
Attribute Main.VB_Description = "Plays games in the RDForms window, on every platform, one after the other while the person asks for a new one; --msforms plays in the Microsoft Forms window instead, to compare the two windows. The command line can name the players, --player1 <player> --player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: --seed <number>. The window opens with the players it names already picked. --test runs the library's tests and the window's own instead of a game, in the console. Under --unattended, nobody watches: the computer plays without its pause before each shot."
    If HasSwitch(Command$, "--test") Then
        RunTests New WindowTests
        Exit Sub
    End If

    SeedRandomNumbers Command$
    Dim firstPlayer As String
    TryGetArgument Command$, "--player1", firstPlayer
    Dim secondPlayer As String
    TryGetArgument Command$, "--player2", secondPlayer

    Dim viewName As String
    viewName = GameView(Command$)
    Dim view As WindowView
    ' Microsoft Forms is missing as soon as its form is made; a missing forms host shows when the RDForms window is first shown, or its player picked.
    On Error GoTo NoWindow
    Do
        If viewName = ViewMicrosoftForms Then
            Set view = WindowView.Create(New GameForm)
        Else
            Set view = WindowView.Create(New BoardForm)
        End If

        Set controller = StartGame(view)
        view.Play firstPlayer, secondPlayer
    Loop While view.WantsNewGame
    Exit Sub

NoWindow:
    If Err.Number <> ClassNotRegistered Then Err.Raise Err.Number, Err.Source, Err.Description
    If viewName = ViewMicrosoftForms Then
        Debug.Print "The game form needs Microsoft Forms, which is not installed here. The console app plays in the console."
    Else
        Debug.Print "The RDForms window needs a forms host, which this runtime does not have. The console app plays in the console."
    End If
End Sub

'@Description("The window to play a game in: the RDForms window, on every platform, or the Microsoft Forms window with --msforms, to compare the two windows.")
Public Function GameView(ByVal commandLine As String) As String
Attribute GameView.VB_Description = "The window to play a game in: the RDForms window, on every platform, or the Microsoft Forms window with --msforms, to compare the two windows."
    If HasSwitch(commandLine, "--msforms") Then
        GameView = ViewMicrosoftForms
    Else
        GameView = ViewRDForms
    End If
End Function
