Attribute VB_Name = "CryptAction"
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


' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' THIS IS THE HEART OF MY ENCRYPTION. IT IS THE SINGLE MOST IMPORTANT PART OF THE
' PROGRAM. I'VE COMMENTED IT PRETTY HIGHLY AND HOPE YOU CAN UNDERSTAND IT AS WELL
' AS I CAN. IF YOU HAVE QUESTIONS, YOU CAN EMAIL ME.


Private Type KeyType ' Create a key type.
        bitloc(7) As Integer ' 8 bits in a byte. This is the location in our string.
        bitsLen As Integer ' The full length of the bin string.
End Type


Option Explicit
Public Const CrashID As String = "<|-Crash-Encryption-2-|>"  ' This is going to be our identifier.

Public Function encryptFile(File2Encrypt As String, SaveDataFileAs As String, SaveKeyFileAs As String, _
                            FileTitle As String, Complexity As Integer, Optional Progbar As ProgressBar, _
                            Optional StatusLabel As Label)

' IMPORANT! - The complexity must be between 8 and 50. At 8, there are NO random bits inserted. I use a minimum of 10.

Dim EncFile As BitFile ' File we're getting data from (i.e. picture, document, zip, etc.)
Dim KeyFile As BitFile ' File we're saving the key to.
Dim SaveFile As BitFile ' File we're saving the encrypted data to.
Dim oBits(7) As Integer ' This is going to be 0 - 8 of our bits in our byte.
Dim nBits(50) As Integer ' This is going to be up to 50 possible new bits for our byte.
Dim ByteKey As KeyType ' This is our key for the single byte.
Dim oB As Integer ' This is our universal interger to count with.
Dim oB2 As Integer ' This is another one of our universal intergers to count with.
Dim tmpBitLoc As Integer ' Temporary bit holder.
Dim bitFound As Boolean ' Specify when we use a data bit.
Dim CarryKey As Integer ' This is to specify how many times we will use the same keyset. More times saves space in the key.
Dim CariedKey As Integer ' Our CarryKey place holder.

CarryKey = 51 - Complexity ' 41 possible carriers.
CariedKey = 1 ' Set the carrier to be reset on first use.

' Be sure we're working with seperate file names.
If File2Encrypt = SaveDataFileAs Then MsgBox "Source File, Destination File, and Key File must all have different names.", vbExclamation, "Cannot Encrypt.": Exit Function
If File2Encrypt = SaveKeyFileAs Then MsgBox "Source File, Destination File, and Key File must all have different names.", vbExclamation, "Cannot Encrypt.": Exit Function
If SaveDataFileAs = SaveKeyFileAs Then MsgBox "Source File, Destination File, and Key File must all have different names.", vbExclamation, "Cannot Encrypt.": Exit Function


' Open our files for reading and writing.
EncFile = OpenIBitFile(File2Encrypt) ' Source File
SaveFile = OpenOBitFile(SaveDataFileAs) ' Destination File
KeyFile = OpenOBitFile(SaveKeyFileAs) ' Key File


'Here we add the file name to the key so that we will know what to call the extracted file at decryption time.
For oB = 1 To Len(CrashID & FileTitle & CrashID) ' I add a crashID to each end of the file name to make sure it is completely visible by our system.
        oB2 = Asc(Mid(CrashID & FileTitle & CrashID, oB, 1))
        OutputBits KeyFile, oB2, 8 ' Write these bytes 1 by 1 so that the file will only need to be openned once.
Next oB


Dim EncFileLen As Long ' We need to see how big in bytes this file is.
EncFileLen = LOF(EncFile.FileNum) ' Get the length.


Progbar.Max = EncFileLen ' Setup the progress bar (obviously.)
Progbar.value = 0 ' And make sure you reset it.
StatusLabel = "Encrypting Data: 0 Bytes" ' First label use.

'We create our CarryKey value.
OutputBits KeyFile, CarryKey, 8 ' Write to the key file what our carrier number is.


Dim X As Long ' Create our X for our for x run.
For X = 1 To EncFileLen ' 1 to the end of the file.
        
        
        
        'Get our byte to work with.
        For oB = 0 To 7
            oBits(oB) = InputBit(EncFile) ' Put the bits in our temporary array.
        Next oB
        
        
        
        CariedKey = CariedKey - 1 ' Take 1 token off our carrier to keep tally.
        If CariedKey = 0 Then ' If it's time to create a new key.
                'Create our random bits and our random string length.
                Randomize
                ByteKey.bitsLen = Int(Rnd * (Complexity - 10)) + 10 ' Random number 10 - 50
                
                For oB = 0 To 7 ' Clear our bit locations
                    ByteKey.bitloc(oB) = 60
                Next oB
                
                For oB = 0 To 7 ' Create our byte string. '-1' for 0 to 49.
getnewbitloc:         ' This is our return to try again...
                    tmpBitLoc = Int(Rnd * ByteKey.bitsLen)  ' get a temporary bit location in our string.
                    For oB2 = 0 To oB  ' Check to see if this location is taken.
                        If ByteKey.bitloc(oB2) = tmpBitLoc Then GoTo getnewbitloc  ' if it is, go back and try again.
                    Next oB2
                    ByteKey.bitloc(oB) = tmpBitLoc ' Empty location, save it.
                Next oB
            
                CariedKey = CarryKey ' Reset our carrier.
        End If
        
        
        
        'Create our full string to be written.
        For oB = 0 To ByteKey.bitsLen
            bitFound = False ' reset our trigger
            For oB2 = 0 To 7 ' check to see if we're going to use a data bit.
                If ByteKey.bitloc(oB2) = oB Then
                    nBits(oB) = oBits(oB2) ' We insert a data bit.
                    bitFound = True ' Set our trigger
                    Exit For
                End If
            Next oB2
            
            If bitFound = False Then
                    nBits(oB) = Int(Rnd * 2) + 1 ' if the bit wasn't triggered, insert a fake bit.
                    If nBits(oB) = 2 Then nBits(oB) = 1 Else nBits(oB) = 0
            End If
        Next oB
        
        If CariedKey = CarryKey Then ' If the carrier was reset, write the key to the file.
            'Write the key to the file.
            'Write To File bits 0-7 & strlen
            For oB = 0 To 7
                OutputBits KeyFile, ByteKey.bitloc(oB), 8
            Next oB
            
            OutputBits KeyFile, ByteKey.bitsLen, 8
        End If
        
        'Write our new data to the data file.
        For oB = 0 To ByteKey.bitsLen - 1
            OutputBit SaveFile, Val(nBits(oB))
        Next oB
        
        Progbar.value = X
        

        
        If X Mod 100 = 0 Then StatusLabel = "Encrypting Data:" & Str(Loc(EncFile.FileNum)) & " Bytes": DoEvents
        
Next X

StatusLabel = "Closing Files..."

'Close our open files so they can be used while this application is open.
CloseOBitFile SaveFile
CloseOBitFile KeyFile
CloseIBitFile EncFile
        
End Function

Public Function decryptFile(EncFileName As String, KeyFileName As String, saveFileName As String, Optional Progbar As ProgressBar, Optional StatusLabel As Label) As String
Dim EncFile As BitFile ' File we're getting data from (i.e. picture, document, zip, etc.)
Dim KeyFile As BitFile ' File we're saving the key to.
Dim SaveFile As BitFile ' File we're saving the encrypted data to.
Dim oBits(7) As Integer ' This is going to be 0 - 8 of our bits in our byte.
Dim oB As Integer ' This is our universal interger to count with.
Dim oB2 As Integer ' This is another one of our universal intergers to count with.
Dim KeyInfo(8) As Integer ' This will store our temp data from our key.
Dim TempBits(50) As Integer ' This is our temporary string
Dim FileTitle As String ' This is the original title of the encrypted file.
Dim CarryKey As String ' This is the first 5 bits in the key that tells us how many times to use each key.
Dim CariedKey As Integer ' Our CarryKey place holder.


' Open our files for reading and writing.
EncFile = OpenIBitFile(EncFileName) ' Our Data File
KeyFile = OpenIBitFile(KeyFileName)  ' Key File
SaveFile = OpenOBitFile(saveFileName)  ' Destination File
'Open KeyFileName For Input As #15


Progbar.value = 0 ' Reset the progress bar.
Progbar.Max = LOF(KeyFile.FileNum) * 1.01 ' Setup the progress bar.
' I added the ".01" because there is a small amount of bit-wise overhead that I didn't want to get introuble over.


'Get the original filename.
Do
    FileTitle = FileTitle & Chr(InputBits(KeyFile, 8)) ' We add our bytes (8 bits at a time) to our string.
    If InStr(Len(CrashID) + 1, FileTitle, CrashID, vbTextCompare) Then Exit Do ' if we have 2 CrashIDs then we have our file name.
Loop
FileTitle = Replace(FileTitle, CrashID, "") ' Get rid of the crashIDs from the name.

StatusLabel = "Decrypting: 0 Bytes"

'Decryption time. This is the most simple part of the show.
On Error GoTo pof ' Sometimes if a file is too small, the progress bar will cause an error. I just ignore it.

CarryKey = InputBits(KeyFile, 8) ' Get our key use number.
CariedKey = 1
Do

    'Get our key string for our first bit.
    ' 9x8 bits. 8 bits for each of the real bit locations and 8 bits at the end to tell us the string length.
    CariedKey = CariedKey - 1
    If CariedKey = 0 Then
    For oB = 0 To 8
        KeyInfo(oB) = InputBits(KeyFile, 8)  ' KeyInfo(8) is the string length.
        Progbar.value = Loc(KeyFile.FileNum)
    Next oB
    CariedKey = CarryKey
    End If
    
    If EOF(KeyFile.FileNum) Then Exit Do ' If we're empty, there's nothing more for us to do. Exit do.
    
    'Grab our string of bits that we will extract the original real bits from.
    For oB = 0 To KeyInfo(8) - 1
        TempBits(oB) = InputBit(EncFile)
    Next oB
    
    
    'Reorder the bits in the byte from our trusty map we created.
    For oB = 0 To 7
        oBits(oB) = TempBits(KeyInfo(oB))
    Next oB
    
    
    'Write the byte to the file now that we have it configured just how we want.
    For oB = 0 To 7
        OutputBit SaveFile, Val(oBits(oB))
    Next oB
    

    If Loc(KeyFile.FileNum) Mod 100 = 0 Then StatusLabel = "Decrypting:" & Str(Loc(KeyFile.FileNum)) & " Bytes": DoEvents

    
Loop

'Close our open files so they can be used while this application is open.
CloseOBitFile SaveFile
CloseIBitFile EncFile
CloseIBitFile KeyFile

decryptFile = FileTitle
Exit Function
pof:
If Err = 380 Then Resume Next
MsgBox Err.Description, vbCritical, "Error: " & Err
Resume Next
End Function
