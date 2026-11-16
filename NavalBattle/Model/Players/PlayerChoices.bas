Attribute VB_Name = "PlayerChoices"
Attribute VB_Description = "The players a grid can be played by, as people name them: Human, Random, Fair-play or Merciless, where the first letter is enough. Every view reads and shows them here, and asks here whether two players can play a game."
'@Folder("NavalBattle.Model.Players")
'@ModuleDescription("The players a grid can be played by, as people name them: Human, Random, Fair-play or Merciless, where the first letter is enough. Every view reads and shows them here, and asks here whether two players can play a game.")
Option Explicit

'@Description("Reads a player from what a person typed or picked: its first letter, in either case, is H (a person), R, F or M (the computer, at that difficulty). Returns False for anything else, and for nothing.")
Public Function TryParsePlayer(ByVal text As String, ByRef outDifficulty As AIDifficulty) As Boolean
Attribute TryParsePlayer.VB_Description = "Reads a player from what a person typed or picked: its first letter, in either case, is H (a person), R, F or M (the computer, at that difficulty). Returns False for anything else, and for nothing."
    TryParsePlayer = True
    Select Case UCase$(Left$(Trim$(text), 1))
        Case "H"
            outDifficulty = Unspecific
        Case "R"
            outDifficulty = RandomAI
        Case "F"
            outDifficulty = FairplayAI
        Case "M"
            outDifficulty = MercilessAI
        Case Else
            TryParsePlayer = False
    End Select
End Function

'@Description("The name a view shows for a player: Human, Random AI, Fair-play AI or Merciless AI.")
Public Function PlayerName(ByVal difficulty As AIDifficulty) As String
Attribute PlayerName.VB_Description = "The name a view shows for a player: Human, Random AI, Fair-play AI or Merciless AI."
    Select Case difficulty
        Case RandomAI
            PlayerName = "Random AI"
        Case FairplayAI
            PlayerName = "Fair-play AI"
        Case MercilessAI
            PlayerName = "Merciless AI"
        Case Else
            PlayerName = "Human"
    End Select
End Function

'@Description("The names of every player, in the order a view offers them: Human first, then the computer from the easiest to the hardest.")
Public Function PlayerNames() As Variant
Attribute PlayerNames.VB_Description = "The names of every player, in the order a view offers them: Human first, then the computer from the easiest to the hardest."
    PlayerNames = Array(PlayerName(Unspecific), PlayerName(RandomAI), PlayerName(FairplayAI), PlayerName(MercilessAI))
End Function

'@Description("Whether a player is a person, or the computer.")
Public Function PlayerTypeOf(ByVal difficulty As AIDifficulty) As PlayerType
Attribute PlayerTypeOf.VB_Description = "Whether a player is a person, or the computer."
    PlayerTypeOf = IIf(difficulty = Unspecific, HumanControlled, ComputerControlled)
End Function

'@Description("Whether two players can play a game: one person at most, since a person sees their own ships and not the other grid's.")
Public Function CanPlayTogether(ByVal firstDifficulty As AIDifficulty, ByVal secondDifficulty As AIDifficulty) As Boolean
Attribute CanPlayTogether.VB_Description = "Whether two players can play a game: one person at most, since a person sees their own ships and not the other grid's."
    CanPlayTogether = Not (PlayerTypeOf(firstDifficulty) = HumanControlled And PlayerTypeOf(secondDifficulty) = HumanControlled)
End Function
