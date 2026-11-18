VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A project's list of tests, for the TestRunner: each test by its number, and what its module does around it. Each project has its own, since only a project's own code can call its test modules."
'@Exposed
'@Folder("NavalBattle.Tests.Framework")
'@ModuleDescription("A project's list of tests, for the TestRunner: each test by its number, and what its module does around it. Each project has its own, since only a project's own code can call its test modules.")
'@Interface
Option Explicit

'@Description("The number of tests in the list.")
Public Property Get Count() As Long
Attribute Count.VB_Description = "The number of tests in the list."
End Property

'@Description("The name of a test, by its number from 1, as Module.Procedure; runs it where run is True.")
Public Function Test(ByVal index As Long, ByVal run As Boolean) As String
Attribute Test.VB_Description = "The name of a test, by its number from 1, as Module.Procedure; runs it where run is True."
End Function

'@Description("Runs the given test module's initialize procedure, where it has one.")
Public Sub InitializeModule(ByVal moduleName As String)
Attribute InitializeModule.VB_Description = "Runs the given test module's initialize procedure, where it has one."
End Sub

'@Description("Runs the given test module's cleanup procedure, where it has one.")
Public Sub CleanupModule(ByVal moduleName As String)
Attribute CleanupModule.VB_Description = "Runs the given test module's cleanup procedure, where it has one."
End Sub

'@Description("Runs the given test module's procedure that comes before each of its tests, where it has one.")
Public Sub InitializeTest(ByVal moduleName As String)
Attribute InitializeTest.VB_Description = "Runs the given test module's procedure that comes before each of its tests, where it has one."
End Sub

'@Description("Runs the given test module's procedure that comes after each of its tests, where it has one.")
Public Sub CleanupTest(ByVal moduleName As String)
Attribute CleanupTest.VB_Description = "Runs the given test module's procedure that comes after each of its tests, where it has one."
End Sub
