VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Encrypt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crash Encryption (Processing...)"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4230
   ControlBox      =   0   'False
   Icon            =   "Encrypt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prog 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label StatusLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-----+++++-----"
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
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Encrypt"
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

Public UseFolderPath  As String
Option Explicit

Public Function Encrypt(OpenFile As String, SaveFile As String, FileTitle As String, KeyFile As String, complex As Integer)
Me.Show ' Tah Dah! Here I am.
StartFrm.Hide ' Hide the prepping form.

encryptFile OpenFile, SaveFile, KeyFile, FileTitle, complex, prog, StatusLbl

Me.Hide
MsgBox "Encryption complete." ' What good program doesn't have some kind of good news to tell you after it just slaved for you?
StartFrm.Show ' OK... Back to the start.
End Function


Public Function Decrypt(EncFile As String, DecFile As String)
Dim FullFileTitle As String

    Me.Show ' Tah Dah! Here I am, again...

    FullFileTitle = decryptFile(EncFile, DecFile, App.Path & "\temp.file", prog, StatusLbl)
    
UseFolders:

    UseFolderPath = "" ' Blah blah blah.
    Folders.Show , Me ' blah blah blah... again.
    Folders.FlNm = FullFileTitle ' OK! We should tell them what the file is called.
    
    Do While UseFolderPath = "" ' Waiting for Christmas.
        DoEvents
    Loop
    
    If UseFolderPath = "Cancel" Then
        If MsgBox("Are you sure you want to cancel?" & vbCrLf & vbCrLf & "If you cancel now, your decryption will not be saved." & vbCrLf, vbYesNo, "Confirm Cancel") = vbNo Then
            GoTo UseFolders
          Else
            Kill App.Path & "\temp.file"
            MsgBox "Decryption Canceled"
            Me.Hide
            StartFrm.Show ' It'd be nice to let them back to the begining.
            Exit Function
        End If
    End If
            
    
    'Create/use the directory. Copy the new file over. Kill the temp file.
    FileCopy App.Path & "\temp.file", UseFolderPath & "\" & FullFileTitle
    Kill App.Path & "\temp.file"
    

Me.Hide

MsgBox "Decryption complete." ' See?

StartFrm.Show ' It'd be nice to let them back to the begining.

End Function


