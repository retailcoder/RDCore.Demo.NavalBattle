Attribute VB_Name = "Program"
Attribute VB_Description = "The program's entry point."
'@Folder("NavalBattle")
'@ModuleDescription("The program's entry point.")
Option Explicit

Private controller As IGameController

'@Description("Plays a game. The command line can fix the random seed, so a game can be replayed: /seed <number>.")
Public Sub Main()
Attribute Main.VB_Description = "Plays a game. The command line can fix the random seed, so a game can be replayed: /seed <number>."
    Dim seed As Long
    If TryGetSeed(Command$, seed) Then
        Rnd -1
        Randomize seed
    Else
        Randomize
    End If

    Dim randomizer As IRandomizer
    Set randomizer = New GameRandomizer

    Dim adapter As GridViewAdapter
    Set adapter = GridViewAdapter.Create(New ConsoleView)

    Set controller = StandardGameController.Create(adapter, randomizer, PlayerFactory.Create(randomizer))
    controller.NewGame

    Dim viewEvents As IGridViewEvents
    Set viewEvents = adapter
    viewEvents.CreatePlayer PlayerGrid1, ComputerControlled, MercilessAI
    viewEvents.CreatePlayer PlayerGrid2, ComputerControlled, FairplayAI
End Sub

Private Function TryGetSeed(ByVal commandLine As String, ByRef outSeed As Long) As Boolean
    Dim arguments As Variant
    arguments = Split(Trim$(commandLine), " ")
    Dim index As Long
    For index = LBound(arguments) To UBound(arguments) - 1
        If StrComp(arguments(index), "/seed", vbTextCompare) = 0 And IsNumeric(arguments(index + 1)) Then
            outSeed = CLng(arguments(index + 1))
            TryGetSeed = True
            Exit Function
        End If
    Next
End Function
