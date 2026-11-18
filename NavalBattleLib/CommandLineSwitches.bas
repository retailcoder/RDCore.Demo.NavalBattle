Attribute VB_Name = "CommandLineSwitches"
Attribute VB_Description = "Reads the program's --name switches on its command line. A switch it is not asked about is ignored, never refused: the runtime leaves its own switches, such as --unattended, on the command line."
'@Folder("NavalBattle")
'@ModuleDescription("Reads the program's --name switches on its command line. A switch it is not asked about is ignored, never refused: the runtime leaves its own switches, such as --unattended, on the command line.")
Option Explicit

'@Description("Gets the value that follows a --name switch on the command line. Another switch is not a value: --seed --unattended has no seed.")
Public Function TryGetArgument(ByVal commandLine As String, ByVal argumentName As String, ByRef outValue As String) As Boolean
Attribute TryGetArgument.VB_Description = "Gets the value that follows a --name switch on the command line. Another switch is not a value: --seed --unattended has no seed."
    Dim arguments As Variant
    arguments = Split(Trim$(commandLine), " ")
    Dim index As Long
    For index = LBound(arguments) To UBound(arguments) - 1
        If StrComp(arguments(index), argumentName, vbTextCompare) = 0 Then
            If Left$(arguments(index + 1), 2) = "--" Then Exit For
            outValue = arguments(index + 1)
            TryGetArgument = True
            Exit Function
        End If
    Next
    TryGetArgument = False
End Function

'@Description("Whether a --name switch is on the command line.")
Public Function HasSwitch(ByVal commandLine As String, ByVal switchName As String) As Boolean
Attribute HasSwitch.VB_Description = "Whether a --name switch is on the command line."
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
