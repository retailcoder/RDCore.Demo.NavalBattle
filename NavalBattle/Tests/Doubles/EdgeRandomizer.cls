VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EdgeRandomizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A randomizer for the tests, with no randomness: Between gives the lowest of its range, or the highest once PicksHighest is set; Maybe and NextSingle give False and 0."
'@Folder("NavalBattle.Tests.Doubles")
'@ModuleDescription("A randomizer for the tests, with no randomness: Between gives the lowest of its range, or the highest once PicksHighest is set; Maybe and NextSingle give False and 0.")
Option Explicit
Implements IRandomizer

Private picksHighestValue As Boolean

'@Description("Gets/sets whether Between gives the highest of its range, not the lowest.")
Public Property Get PicksHighest() As Boolean
Attribute PicksHighest.VB_Description = "Gets/sets whether Between gives the highest of its range, not the lowest."
    PicksHighest = picksHighestValue
End Property

Public Property Let PicksHighest(ByVal value As Boolean)
    picksHighestValue = value
End Property

Private Function IRandomizer_Maybe(ByVal likelihood As Probability) As Boolean
    IRandomizer_Maybe = False
End Function

Private Function IRandomizer_NextSingle() As Single
    IRandomizer_NextSingle = 0
End Function

Private Function IRandomizer_Between(ByVal lower As Long, ByVal upper As Long) As Long
    IRandomizer_Between = IIf(picksHighestValue, upper, lower)
End Function
