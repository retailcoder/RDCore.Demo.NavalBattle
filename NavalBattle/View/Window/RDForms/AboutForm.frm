VERSION 5.00
Begin VB.Form AboutForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About NavalBattle"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6600
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2700
      TabIndex        =   0
      Top             =   9600
      Width           =   1200
   End
   Begin VB.PictureBox picLogo
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   4200
      Left            =   240
      ScaleHeight     =   4170
      ScaleWidth      =   6090
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   6120
      Begin VB.Image imgLogo
         Height          =   3840
         Left            =   1125
         Top             =   165
         Width           =   3840
      End
   End
   Begin VB.Label lblProduct
      Alignment       =   2  'Center
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   16
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NavalBattle"
      Height          =   480
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   6120
   End
   Begin VB.Label lblCompany
      Alignment       =   2  'Center
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "© 2026 9562-7303 Québec inc."
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   5040
      Width           =   6120
   End
   Begin VB.Label lblThanksTitle
      Alignment       =   2  'Center
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Special thanks to:"
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   5460
      Width           =   6120
   End
   Begin VB.Label lblThanks
      Alignment       =   2  'Center
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   240
      TabIndex        =   5
      Top             =   5760
      Width           =   6120
   End
   Begin VB.Label lblTrademarksTitle
      Alignment       =   2  'Center
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Trademarks"
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   6660
      Width           =   6120
   End
   Begin VB.Label lblTrademarks
      Alignment       =   2  'Center
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   240
      TabIndex        =   7
      Top             =   6960
      Width           =   6120
   End
   Begin VB.Label lblSignature
      Alignment       =   2  'Center
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   8
      Top             =   8960
      Width           =   6120
   End
   Begin VB.Line lineAboveSignature
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   6360
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Line lineUnderSignature
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   6360
      Y1              =   9400
      Y2              =   9400
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("NavalBattle.View.Window.RDForms")
Option Explicit

' the RDCore ducky, centred in its white box.
Private Const LogoFile As String = "assets\rdcore-ducky.png"
' the width of the logo box's border, inside its edges.
Private Const LogoBorder As Long = 15

Private Sub Form_Load()
    LoadLogo
    lblThanks.Caption = "Anthropic" & vbCrLf & "HexIDE" & vbCrLf & "ThunderFrame"
    lblTrademarks.Caption = _
        """RDCore"" and ""VIVAT CUCUMIS"" are trademarks claimed by 9562-7303 Québec inc. (pending)." & vbCrLf & _
        "BATTLESHIP is a trademark of Hasbro, Inc. NavalBattle is not affiliated with, sponsored by or endorsed by Hasbro, Inc." & vbCrLf & _
        "NavalBattle is not a Microsoft product and is not affiliated with Microsoft, directly or indirectly." & vbCrLf & _
        "All trademarks belong to their respective owners."
    ' the signature's heart is not in the form file's code page, so the caption is written here.
    lblSignature.Caption = "V I V A T  " & ChrW$(&H2764) & ChrW$(&HFE0F) & "  C U C U M I S " & ChrW$(&H2122)
End Sub

Private Sub LoadLogo()
    On Error GoTo NoLogo
    Set imgLogo.Picture = LoadPicture(CurDir$ & "\" & LogoFile)
    imgLogo.Move (picLogo.Width - 2 * LogoBorder - imgLogo.Width) / 2, (picLogo.Height - 2 * LogoBorder - imgLogo.Height) / 2
    Exit Sub

NoLogo:
    Debug.Print "The About box goes without its logo: " & LogoFile & " cannot be read (" & Err.Description & ")."
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
