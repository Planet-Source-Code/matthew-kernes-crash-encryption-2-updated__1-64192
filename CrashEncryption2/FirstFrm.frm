VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form StartFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crash Encryption 2"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   ControlBox      =   0   'False
   Icon            =   "FirstFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   8040
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CDBrowse 
      Left            =   240
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Encrypt A File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   17
      Top             =   240
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Decrypt A File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   16
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Height          =   4815
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   7215
      Begin VB.CommandButton Command5 
         Caption         =   "Change Save Location"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4800
         TabIndex        =   21
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Browse"
         Height          =   255
         Left            =   5760
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   5415
      End
      Begin MSComctlLib.Slider Complexity 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   3840
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   10
         Min             =   10
         Max             =   50
         SelStart        =   25
         TickFrequency   =   2
         Value           =   25
      End
      Begin VB.Line Line5 
         X1              =   1920
         X2              =   4680
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Key Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Potential File Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "File Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label KeySize 
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label AdjSize 
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label FlSize 
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Line Line4 
         X1              =   4680
         X2              =   4680
         Y1              =   1920
         Y2              =   1320
      End
      Begin VB.Line Line3 
         X1              =   4680
         X2              =   7080
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Input file to be encrypted:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "File name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Encrypted file name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Encryption Complexity:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   12
         ToolTipText     =   $"FirstFrm.frx":0442
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Decryption Key:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1920
         TabIndex        =   9
         Top             =   2040
         Width           =   5175
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Top             =   2640
         Width           =   5175
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "<<< Less Secure                                                        More Secure >>>"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   7
         Top             =   3600
         Width           =   5175
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "<<< Smaller File                                                              Larger File >>>"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   6
         Top             =   4320
         Width           =   5175
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   $"FirstFrm.frx":053C
      Height          =   855
      Left            =   1080
      TabIndex        =   20
      Top             =   600
      Width           =   7215
   End
   Begin VB.Label Label9 
      Caption         =   $"FirstFrm.frx":0686
      Height          =   615
      Left            =   1080
      TabIndex        =   19
      Top             =   6840
      Width           =   7215
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "FirstFrm.frx":0761
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FirstFrm.frx":0BA3
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   8280
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line2 
      DrawMode        =   8  'Xor Pen
      X1              =   720
      X2              =   8275
      Y1              =   7815
      Y2              =   7815
   End
End
Attribute VB_Name = "StartFrm"
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


' I'd like to thank Derek Haas for the great I/O module. It's saved me a LOT of time and
' it's probably the easiest I/O mod to use that I've seen.



Option Explicit

Private Sub Command1_Click()
' OK. This is the imporant part of the form.

If Option1(1).value = True Then
    Me.Hide
    DecryptStart.Show ' This is the easy part. Show the decryption form.

Else ' But here's the complicated part. I just had to make things a pain in the ass with this moving form.
        If Text1 = "" Then MsgBox "Before continuing, make sure you've selected a file to encrypt.": Exit Sub
                
        Encrypt.Encrypt Text1, Label7, Label6, Label8, Complexity.value ' OK... Now send it encryption routine because the startfrm is lazy.
        
End If

End Sub

Private Sub Command2_Click()
End ' Leaving so soon? Oh well.
End Sub


Private Sub Command3_Click()
With CDBrowse
.Filter = "All Files (*.*) | *.*"
.DialogTitle = "Open File to Encrypt"
.ShowOpen
If .FileName = "" Then Exit Sub
FlSize = FileLen(.FileName) & " Bytes"
AdjSize = (Complexity.value * FileLen(.FileName)) / 8 & " Bytes"
KeySize = Int((FileLen(.FileName) / (51 - Complexity))) * 9 & " Bytes"
If .FileName = "" Then Exit Sub
Text1 = .FileName
Label6 = .FileTitle
Label7 = .FileName & ".Encrypt" ' Aren't I snazzy with my cool file extentions? ;)
Label8 = .FileName & ".DecKey"
End With
End Sub

Private Sub Command4_Click()
frmAbout.Show , Me
End Sub


Private Sub Command5_Click()
    
    Encrypt.UseFolderPath = "" ' Blah blah blah.
    Folders.Show , Me ' blah blah blah... again.
    Folders.FlNm = Label6 ' OK! We should tell them what the file is called.
    
    Do While Encrypt.UseFolderPath = "" ' Waiting for Christmas.
        DoEvents
    Loop
    
    If Encrypt.UseFolderPath = "Cancel" Then Exit Sub
    
    Label7 = Encrypt.UseFolderPath & "\" & Label6 & ".Encrypt"  ' Aren't I snazzy with my cool file extentions? ;)
    Label8 = Encrypt.UseFolderPath & "\" & Label6 & ".DecKey"
    
End Sub

Private Sub Complexity_Change()
If Text1 = "" Then Exit Sub
AdjSize = (Complexity.value * FileLen(CDBrowse.FileName)) / 8 & " Bytes"
KeySize = Int((FileLen(CDBrowse.FileName) / (51 - Complexity))) * 9 & " Bytes"
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(1).value = True Then Frame3.Enabled = False Else Frame3.Enabled = True
End Sub

Private Sub Text1_Change()
If Text1 <> "" Then Command5.Enabled = True Else Command5.Enabled = False
End Sub
