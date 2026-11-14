VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlayerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A concrete implementation for the IPlayer abstract factory."
'@Folder("NavalBattle.Model")
'@ModuleDescription("A concrete implementation for the IPlayer abstract factory.")
'@PredeclaredId
Option Explicit

Private Type TFactory
    Random As IRandomizer
    AIPauseMilliseconds As Long
End Type

Private this As TFactory
Implements IPlayerFactory

'@Description("Creates the factory. Its computer players wait the given number of milliseconds before each shot; 0 plays at once.")
Public Function Create(ByVal randomizer As IRandomizer, ByVal aiPauseMilliseconds As Long) As IPlayerFactory
Attribute Create.VB_Description = "Creates the factory. Its computer players wait the given number of milliseconds before each shot; 0 plays at once."
    With New PlayerFactory
        Set .Random = randomizer
        .AIPauseMilliseconds = aiPauseMilliseconds
        Set Create = .Self
    End With
End Function

Public Property Get AIPauseMilliseconds() As Long
    AIPauseMilliseconds = this.AIPauseMilliseconds
End Property

Public Property Let AIPauseMilliseconds(ByVal value As Long)
    this.AIPauseMilliseconds = value
End Property

Public Property Get Self() As IPlayerFactory
    Set Self = Me
End Property

Public Property Get Random() As IRandomizer
    Set Random = this.Random
End Property

Public Property Set Random(ByVal value As IRandomizer)
    Set this.Random = value
End Property

Private Function IPlayerFactory_Create(ByVal gridId As PlayGridId, Optional ByVal difficulty As AIDifficulty = AIDifficulty.Unspecific) As IPlayer
    
    Dim grid As PlayerGrid
    Set grid = PlayerGrid.Create(gridId)
    
    If difficulty = Unspecific Then
        Set IPlayerFactory_Create = CreateHumanPlayer(grid)
    Else
        Set IPlayerFactory_Create = CreateAIPlayer(grid, difficulty)
    End If

End Function

Private Function CreateHumanPlayer(ByVal grid As PlayerGrid) As IPlayer
    Set CreateHumanPlayer = HumanPlayer.Create(grid)
End Function

Private Function CreateAIPlayer(ByVal grid As PlayerGrid, ByVal difficulty As AIDifficulty) As IPlayer
    Select Case difficulty
        Case AIDifficulty.RandomAI
            Set CreateAIPlayer = AIPlayer.Create(grid, RandomShotStrategy.Create(this.Random), this.AIPauseMilliseconds)
            
        Case AIDifficulty.FairplayAI
            Set CreateAIPlayer = AIPlayer.Create(grid, FairPlayStrategy.Create(this.Random), this.AIPauseMilliseconds)
            
        Case AIDifficulty.MercilessAI
            Set CreateAIPlayer = AIPlayer.Create(grid, MercilessStrategy.Create(this.Random), this.AIPauseMilliseconds)
            
        Case Else
            Debug.Assert False
    End Select
End Function
