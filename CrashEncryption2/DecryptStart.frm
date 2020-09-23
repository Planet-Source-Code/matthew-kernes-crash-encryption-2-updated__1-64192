VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form DecryptStart 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6270
   ControlBox      =   0   'False
   Icon            =   "DecryptStart.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   4575
   End
   Begin MSComDlg.CommonDialog CDBrowse 
      Left            =   3480
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Input decryption key file:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Input file to decrypt:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "DecryptStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'|||||||||||||||||||||||||||||||||||Disclaimer|||||||||||||||||||||||||||||||||||||'
' This code is written by Matthew Kernes. It is not intended for commercial use.   '
' It has no warranty nor is Matthew Kernes responsible for any damage it does to   '
' any computer it runs on. Any changes made to this software after the day October '
' 24th, 2005 by anyone other then Matthew Kernes is liabile for the changes and    '
' Matthew Kernes is not responsible for those changes or the problems they may     '
' cause.                                                                           '
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||'

' I wrote this code because I wanted to see if I could make an unhackable encryption.
' I know there is no such thing as "unhackable" encryption. But I'd like to see how fast
' someone can break the code. Originally, I created a program that uses the Ceaser Cypher,
' but I didn't know it at the time. I was pretty new to encryption and I feel I still am.
' But I lay the challenge out now to anyone who might want it. Can you break this encryption?

' I wrote the software with 2 ideas in mind.
'    - How many varaibles does it take to make it so you can't solve for x?
'    - How do you make it so the user cannot control the key or password?

' With this in mind, I set out to make what is now "Crash Encryption".
' I dubbed it "Crash Encryption" because this program is somewhat a resource hog,
' as well, my nick-name is "Crash".

' If you like what you see and have comments or questions, please feel free
' to email me at compiano@socal.rr.com. Voting is not necessary on my code.

' Thanks for the view,
'                   Matthew Kernes (Crash)



' This entire form is just so simple that
' I'm not really going to comment it too much.

Public decrypt_Title As String
Public Decrypt_Path As String
Option Explicit

Private Sub Command1_Click()
CDBrowse.Filter = "Encrypted Files (*.encrypt) | *.Encrypt" ' Let's find that file type I so cleverly created.
CDBrowse.DialogTitle = "Open File to Decrypt"
CDBrowse.ShowOpen
If CDBrowse.FileName = "" Then Exit Sub
decrypt_Title = CDBrowse.FileTitle
Decrypt_Path = Mid(CDBrowse.FileName, 1, Len(CDBrowse.FileName) - Len(CDBrowse.FileTitle) - 1)
Text1 = CDBrowse.FileName
End Sub

Private Sub Command2_Click()
Encrypt.Show
Me.Hide
Encrypt.Decrypt Text1, Text2
End Sub

Private Sub Command3_Click()
Me.Hide
StartFrm.Show
End Sub

Private Sub Command4_Click()
CDBrowse.Filter = "Decryption Keys (*.DecKey) | *.DecKey" ' Yet another awesome file extention.
CDBrowse.DialogTitle = "Open Decryption Key"
CDBrowse.ShowOpen
If CDBrowse.FileName = "" Then Exit Sub
Text2 = CDBrowse.FileName
End Sub
