Attribute VB_Name = "Program"
Attribute VB_Description = "The program's entry point."
'@Folder("NavalBattle")
'@ModuleDescription("The program's entry point.")
Option Explicit

Private Const ClassNotRegistered As Long = 429

Private controller As IGameController

'@Description("Plays a game: in the game form on Windows, and in the console elsewhere, or when /console is on the command line. The command line can name the players, /player1 <player> /player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: /seed <number>. The form opens with the players it names already picked; the console asks for a player it does not name.")
Public Sub Main()
Attribute Main.VB_Description = "Plays a game: in the game form on Windows, and in the console elsewhere, or when /console is on the command line. The command line can name the players, /player1 <player> /player2 <player> (Human, Random, Fair-play or Merciless), and fix the random seed so a game can be replayed: /seed <number>. The form opens with the players it names already picked; the console asks for a player it does not name."
    Dim seed As String
    If TryGetArgument(Command$, "/seed", seed) And IsNumeric(seed) Then
        Rnd -1
        Randomize CLng(seed)
    Else
        Randomize
    End If

    Dim firstPlayer As String
    TryGetArgument Command$, "/player1", firstPlayer
    Dim secondPlayer As String
    TryGetArgument Command$, "/player2", secondPlayer

#If Win32 Or Win64 Then
    If Not HasSwitch(Command$, "/console") Then
        If TryPlayInForm(firstPlayer, secondPlayer) Then Exit Sub
    End If
#End If
    PlayInConsole firstPlayer, secondPlayer
End Sub

'@Description("Plays a game in the game form, and returns True when the form is closed; returns False at once when Microsoft Forms is not installed.")
Private Function TryPlayInForm(ByVal firstPlayer As String, ByVal secondPlayer As String) As Boolean
Attribute TryPlayInForm.VB_Description = "Plays a game in the game form, and returns True when the form is closed; returns False at once when Microsoft Forms is not installed."
    Dim view As FormView
    Dim adapter As GridViewAdapter

    On Error GoTo NoForms
    Set view = New FormView
    Set adapter = GridViewAdapter.Create(view)
    On Error GoTo 0

    StartGame adapter
    view.Play firstPlayer, secondPlayer
    TryPlayInForm = True
    Exit Function

NoForms:
    If Err.Number <> ClassNotRegistered Then Err.Raise Err.Number, Err.Source, Err.Description
    Debug.Print "The game form needs Microsoft Forms, which is not installed here: the game is played in the console."
    TryPlayInForm = False
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

    Set controller = StandardGameController.Create(adapter, randomizer, PlayerFactory.Create(randomizer))
    controller.NewGame
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

'@Description("Whether a /name switch is on the command line.")
Private Function HasSwitch(ByVal commandLine As String, ByVal switchName As String) As Boolean
Attribute HasSwitch.VB_Description = "Whether a /name switch is on the command line."
    Dim arguments As Variant
    arguments = Split(Trim$(commandLine), " ")
    Dim index As Long
    For index = LBound(arguments) To UBound(arguments)
        If StrComp(arguments(index), switchName, vbTextCompare) = 0 Then
            HasSwitch = True
            Exit Function
        End If
    Next
    HasSwitch = False
End Function
