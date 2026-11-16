VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptedPlayerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A player factory for the controller tests: it makes a ScriptedPlayer for each grid, whatever the difficulty, and keeps it so the test can look at its grid."
'@Folder("NavalBattle.Tests.Doubles")
'@ModuleDescription("A player factory for the controller tests: it makes a ScriptedPlayer for each grid, whatever the difficulty, and keeps it so the test can look at its grid.")
Option Explicit
Implements IPlayerFactory

Private players(PlayerGrid1 To PlayerGrid2) As ScriptedPlayer

'@Description("Gets the player made for the specified grid, or Nothing.")
Public Function Player(ByVal gridId As PlayGridId) As ScriptedPlayer
Attribute Player.VB_Description = "Gets the player made for the specified grid, or Nothing."
    Set Player = players(gridId)
End Function

Private Function IPlayerFactory_Create(ByVal gridId As PlayGridId, Optional ByVal difficulty As AIDifficulty = AIDifficulty.Unspecific) As IPlayer
    Set players(gridId) = New ScriptedPlayer
    Set players(gridId).Grid = PlayerGrid.Create(gridId)
    Set IPlayerFactory_Create = players(gridId)
End Function
