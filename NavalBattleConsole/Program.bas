Attribute VB_Name = "Program"
Attribute VB_Description = "The console app's entry point."
'@Folder("NavalBattle")
'@ModuleDescription("The console app's entry point.")
Option Explicit

Private controller As IGameController

'@Description("Plays a game in the console. The command line can name the players, --player1 <player> --player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: --seed <number>. The console asks for a player it does not name. --test runs the library's tests instead of a game. Under --unattended, nobody watches: the computer plays without its pause before each shot.")
Public Sub Main()
Attribute Main.VB_Description = "Plays a game in the console. The command line can name the players, --player1 <player> --player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: --seed <number>. The console asks for a player it does not name. --test runs the library's tests instead of a game. Under --unattended, nobody watches: the computer plays without its pause before each shot."
    If HasSwitch(Command$, "--test") Then
        RunTests
        Exit Sub
    End If

    SeedRandomNumbers Command$
    Dim firstPlayer As String
    TryGetArgument Command$, "--player1", firstPlayer
    Dim secondPlayer As String
    TryGetArgument Command$, "--player2", secondPlayer

    Dim view As ConsoleView
    Set view = New ConsoleView
    Set controller = StartGame(view)
    view.Play firstPlayer, secondPlayer
End Sub
