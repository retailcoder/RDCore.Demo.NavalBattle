Attribute VB_Name = "Program"
Attribute VB_Description = "The program's entry point."
'@Folder("NavalBattle")
'@ModuleDescription("The program's entry point.")
Option Explicit

Private controller As IGameController

'@Description("Plays a game in the console. The command line can name the players, /player1 <player> /player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: /seed <number>. A player it does not name is asked for.")
Public Sub Main()
Attribute Main.VB_Description = "Plays a game in the console. The command line can name the players, /player1 <player> /player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: /seed <number>. A player it does not name is asked for."
    Dim seed As String
    If TryGetArgument(Command$, "/seed", seed) And IsNumeric(seed) Then
        Rnd -1
        Randomize CLng(seed)
    Else
        Randomize
    End If

    Dim randomizer As IRandomizer
    Set randomizer = New GameRandomizer

    Dim view As ConsoleView
    Set view = New ConsoleView

    Dim adapter As GridViewAdapter
    Set adapter = GridViewAdapter.Create(view)

    Set controller = StandardGameController.Create(adapter, randomizer, PlayerFactory.Create(randomizer))
    controller.NewGame

    Dim firstPlayer As String
    TryGetArgument Command$, "/player1", firstPlayer
    Dim secondPlayer As String
    TryGetArgument Command$, "/player2", secondPlayer
    view.Play firstPlayer, secondPlayer
End Sub

'@Description("Shows the game form. It does not play a game yet: its status line says which cell was clicked.")
Public Sub ShowGameForm()
Attribute ShowGameForm.VB_Description = "Shows the game form. It does not play a game yet: its status line says which cell was clicked."
    GameForm.Show
End Sub

'@Description("Gets the value that follows a /name argument on the command line.")
Private Function TryGetArgument(ByVal commandLine As String, ByVal argumentName As String, ByRef outValue As String) As Boolean
Attribute TryGetArgument.VB_Description = "Gets the value that follows a /name argument on the command line."
    Dim arguments As Variant
    arguments = Split(Trim$(commandLine), " ")
    Dim index As Long
    For index = LBound(arguments) To UBound(arguments) - 1
        If StrComp(arguments(index), argumentName, vbTextCompare) = 0 Then
            outValue = arguments(index + 1)
            TryGetArgument = True
            Exit Function
        End If
    Next
    TryGetArgument = False
End Function
