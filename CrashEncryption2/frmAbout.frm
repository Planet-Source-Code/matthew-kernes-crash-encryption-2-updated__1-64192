VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   7620
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5715
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5259.459
   ScaleMode       =   0  'User
   ScaleWidth      =   5366.681
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   4245
      TabIndex        =   0
      Top             =   6720
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   4420.845
      Y2              =   4420.845
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   5415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   4431.198
      Y2              =   4431.198
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":012E
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   6600
      Width           =   3855
   End
   Begin VB.Image imgLogo 
      Height          =   5145
      Left            =   120
      Picture         =   "frmAbout.frx":01C1
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   5535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

