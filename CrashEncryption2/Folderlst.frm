VERSION 5.00
Begin VB.Form Folders 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select A Folder"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4440
   ControlBox      =   0   'False
   Icon            =   "Folderlst.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select Folder"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Folder"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label FlNm 
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "File will be saved as:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Folders"
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



' This entire form is just so way simple that
' I'm not really going to comment it too much.


Option Explicit

Private Sub Cancel_Click()
Encrypt.UseFolderPath = "Cancel"
Me.Hide
End Sub

Private Sub Command1_Click()
' I figured that they should be able to put the file in a new folder.
' It's easier to create it in the program then to go all the way to the path to do it.
Dim h As String
On Error Resume Next
ChDir Dir1.Path
h = InputBox("New folder name:", "Create New Folder", "New Folder")
If h = "" Then Exit Sub
MkDir h
Dir1.Refresh
Dir1.Path = Dir1.Path & "\" & h
End Sub

Private Sub Command2_Click()
Encrypt.UseFolderPath = Dir1.Path
Me.Hide
End Sub

Private Sub Drive1_Change()
On Error GoTo errorr
Dir1.Path = Drive1.Drive
Exit Sub
errorr:
If Err = 68 Then
MsgBox (Drive1.Drive & ": Drive not ready.")
Exit Sub
Else
MsgBox "Unknown error: " & Err.Description
End If
End Sub
