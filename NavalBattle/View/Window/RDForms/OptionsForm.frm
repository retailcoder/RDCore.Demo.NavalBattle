VERSION 5.00
Begin VB.Form OptionsForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   BeginProperty Font
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12360
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   12360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDefaults
      Caption         =   "&Default names"
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   3540
      Width           =   1800
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   9600
      TabIndex        =   27
      Top             =   3540
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10920
      TabIndex        =   28
      Top             =   3540
      Width           =   1200
   End
   Begin VB.TextBox txtName
      Height          =   330
      Index           =   0
      Left            =   2400
      MaxLength       =   24
      TabIndex        =   1
      Top             =   960
      Width           =   1860
   End
   Begin VB.Label lblIntro
      Caption         =   "Each formation of ducks goes by one of these five names, picked at random for each game. A new game uses the names given here; a box left empty takes its default name."
      Height          =   600
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11880
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblKind
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   29
      Top             =   990
      Width           =   2040
   End
End
Attribute VB_Name = "OptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("NavalBattle.View.Window.RDForms")
Option Explicit

' the boxes: a row for each kind of ship, with a box for each of its names, loaded from the designed ones.
Private Const NamesPerKind As Long = 5
Private Const RowHeight As Long = 480
Private Const NameStep As Long = 1950

Private Sub Form_Load()
    Dim kind As ShipType
    For kind = ShipType.Carrier To ShipType.Destroyer
        If kind > 0 Then Load lblKind(kind)
        With lblKind(kind)
            .Top = lblKind(0).Top + kind * RowHeight
            .Caption = Ship.KindCaption(kind)
            .Visible = True
        End With
        Dim choices As Variant
        choices = Ship.NameChoices(kind)
        Dim position As Long
        For position = 0 To NamesPerKind - 1
            Dim index As Long
            index = kind * NamesPerKind + position
            If index > 0 Then Load txtName(index)
            With txtName(index)
                .Move txtName(0).Left + position * NameStep, txtName(0).Top + kind * RowHeight
                .Text = choices(LBound(choices) + position)
                .TabIndex = index + 1
                .Visible = True
            End With
        Next
    Next
End Sub

Private Sub cmdDefaults_Click()
    Dim kind As ShipType
    For kind = ShipType.Carrier To ShipType.Destroyer
        Dim choices As Variant
        choices = Ship.DefaultNameChoices(kind)
        Dim position As Long
        For position = 0 To NamesPerKind - 1
            txtName(kind * NamesPerKind + position).Text = choices(LBound(choices) + position)
        Next
    Next
End Sub

Private Sub cmdOK_Click()
    Dim kind As ShipType
    For kind = ShipType.Carrier To ShipType.Destroyer
        Dim names(0 To NamesPerKind - 1) As String
        Dim position As Long
        For position = 0 To NamesPerKind - 1
            names(position) = txtName(kind * NamesPerKind + position).Text
        Next
        Ship.SaveNameChoices kind, names
    Next
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
