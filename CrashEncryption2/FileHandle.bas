Attribute VB_Name = "BitReadWrite"
'**************************************
' Name: Bit IO
' Description:This module allows you to
'     view a file as a collection of bits rath
'     er than as a collection of bytes. It all
'     ows you to read/write a single bit at a
'     time or read/write up to 32 bits at once
'     .
' By: Derek Haas
'
' Inputs:It's all explained in the code
'
' Returns:Same as above
'
' Side Effects:Don't try writing to a fi
'     le opened for reading, and don't try rea
'     ding from a file opened for writing - th
'     ere is no error checking for that and th
'     e results are unpredictable.

'Don 't try To read or write more than 32 bits at a time With the InputBits and OutputBits functions.
'    If you try To write a value With less bits than that value requires, the correct value will Not be written. For example, don't try to write the value 32 into a file using only 4 bits.
'    After every call To inputbits and inputbit, you should check For eof on the input file using this code:
    'inputbit/inputbits call here

    'If EOF(BitFile.FileNum) = True Then 'replace bitfile With the name of the variable
        'put code to exit loop or leave function
        '     here
    'End If

'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=2903&lngWId=1'for details.'**************************************


Type BitFile
    FileNum As Integer 'File handle
    holder As Byte 'holds a byte from file
    mask As Byte 'used To read bits
    End Type

Public Function OpenOBitFile(FileName As String) As BitFile

    'Parameters - Filename
    'Returns - Bitfile
    'What it does - Opens a file for output
    '     a single bit at a time
    'Example -dim OutputFile as bitfile
    'OutputFile = OpenOBitFile("C:\test.bit"
    '     )
    
    Dim bitfilename As BitFile
    FileNum = FreeFile 'get lowest available file handle
    Open FileName For Binary As FileNum 'open it
    bitfilename.FileNum = FileNum 'assign file number To structure
    bitfilename.holder = 0 'bit holder = 0
    bitfilename.mask = 128 'used To read individual bits
    OpenOBitFile = bitfilename
End Function


Public Function OpenIBitFile(FileName As String) As BitFile

    'Parameters - Filename
    'Returns - Bitfile
    'What it does - Opens a file for input a
    '     single bit at a time
    'Example -dim InputFile as bitfile
    'InputFile = OpenIBitFile("C:\command.co
    '     m")
    Dim bitfilename As BitFile
    FileNum = FreeFile 'get lowest available file handle
    Open FileName For Binary As FileNum 'open it
    bitfilename.FileNum = FileNum 'assign file number To structure
    bitfilename.holder = 0 'bit holder = 0
    bitfilename.mask = 128 'used To read individual bits
    OpenIBitFile = bitfilename
End Function


Public Sub CloseIBitFile(bitfilename As BitFile)

    'Parameters - bitfile
    'Returns - Nothing
    'What it does - Closes the file associat
    '     ed with a bitfile
    'Example - CloseIBitFile(InputFile)
    Close bitfilename.FileNum 'Close the file associated With the bitfile
End Sub


Public Sub CloseOBitFile(bitfilename As BitFile)

    'Parameters - bitfile
    'Returns - Nothing
    'What it does - Closes the file associat
    '     ed with a bitfile
    'Example - CloseOBitFile(OutputFile)

    If bitfilename.mask <> 128 Then 'If there is unwritten data...
        Put bitfilename.FileNum, , bitfilename.holder 'Write it now
    End If

    Close bitfilename.FileNum 'Close the file
End Sub


Public Sub OutputBit(ByRef bitfilename As BitFile, bit As Byte)

    'Parameters - bitfile, bit to write
    'Returns - nothing
    'What it does - Writes the specified bit
    '     to the file
    'Example - OutputBit(OutputFile, 1)

    If bit <> 0 Then
        bitfilename.holder = bitfilename.holder Or bitfilename.mask
        'the holder stores up written bits until
        '     there are 8
        'At that point vb's normal file handling
        '     facilities can write it
    End If

    bitfilename.mask = bitfilename.mask \ 2 'decrease mask by power of 2


    If bitfilename.mask = 0 Then 'if mask is empty
        Put bitfilename.FileNum, , bitfilename.holder 'write the Byte
        bitfilename.holder = 0 'reset holder and mask
        bitfilename.mask = 128
    End If

    
End Sub


Public Sub OutputBits(ByRef bitfilename As BitFile, ByVal code As Long, ByVal count As Integer)

    'Parameters - bitfile, data to write, nu
    '     mber of bits to use
    'Returns - nothing
    'What it does - Writes the specified inf
    '     o using the specified number of bits
    'Example - OutputBits(OutputFile, 28, 7)
    '
    Dim mask As Long
    mask = 2 ^ (count - 1)


    Do While mask <> 0


        If (mask And code) <> 0 Then 'if the bits match up...
            bitfilename.holder = bitfilename.holder Or bitfilename.mask 'put the bit In the holder
        End If

        bitfilename.mask = bitfilename.mask \ 2
        mask = mask \ 2


        If bitfilename.mask = 0 Then 'when there are 8 bits, write the holder To the file
            Put bitfilename.FileNum, , bitfilename.holder
            bitfilename.holder = 0 'and reset the holder and mask
            bitfilename.mask = 128
        End If

    Loop

End Sub


Public Function InputBit(ByRef bitfilename As BitFile) As Byte

    'Parameters - bitfile
    'returns - the next bit from the file
    'Example: bit = InputBit(InputBitFile)
    Dim value As Byte

    If bitfilename.mask = 128 Then 'if at End of previous Byte
        
        Get bitfilename.FileNum, , bitfilename.holder 'get a new Byte from file
    End If

    value = bitfilename.holder And bitfilename.mask 'get the bit
    bitfilename.mask = bitfilename.mask \ 2 'move the mask bit down one


    If bitfilename.mask = 0 Then
        bitfilename.mask = 128
    End If


    If value <> 0 Then 'return 0 or 1 depending on value
        InputBit = 1
    Else
        InputBit = 0
    End If

End Function


Public Function InputBits(ByRef bitfilename As BitFile, count As Integer) As Long

    'Parameters - bitfile, number of bits to
    '     read
    'returns - the value of the next count b
    '     its in the bitfile
    'Example: byte = InputBits(InputBitFile,
    '     8)
    'This function works just like inputbit
    '     except that it loops through and reads t
    '     he specified
    'number of bits and puts them into a tem
    '     porary holder
    Dim holder As Long
    Dim longmask As Long
    longmask = 2 ^ (count - 1)


    Do While (longmask <> 0)


        If bitfilename.mask = 128 Then
            
            Get bitfilename.FileNum, , bitfilename.holder
        End If


        If (bitfilename.holder And bitfilename.mask) <> 0 Then
            holder = holder Or longmask
        End If

        bitfilename.mask = bitfilename.mask \ 2
        longmask = longmask \ 2


        If bitfilename.mask = 0 Then
            bitfilename.mask = 128
        End If

    Loop

    
    InputBits = holder
End Function
