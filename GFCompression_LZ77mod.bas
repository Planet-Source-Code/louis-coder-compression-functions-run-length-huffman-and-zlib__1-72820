Attribute VB_Name = "GFCompression_LZ77mod"
Option Explicit
'(c)2001, 2002 by Louis. This module contains a variety of the LZ77
'compression algorithm that finds and 'shortens' repeating strings.
'
'NOTE: this code has not been finished yet (does not work).
'
'general use
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'LZ77RepeatStringStruct
Private Type LZ77RepeatStringStruct
    ByteStringLength As Long
    ByteString() As Byte
End Type
Dim LZ77RepeatStringStructNumber As Integer
Dim LZ77RepeatStringStructArray() As LZ77RepeatStringStruct
Dim IndexTableStringLength As Long
Dim IndexTableString() As Byte
'
'NOTE: a string compressed with the LZ77 compression has the following format:
'abcdefgXlmnopX
'X: reserved char, jump to index table and get index of repeat string that
'is to be inserted at the position of the reserved char
'

'***********************************LZ77 COMPRESSION***********************************

Public Function LZ77_CompressString(ByRef ByteStringLength As Long, ByRef ByteString() As Byte) As Boolean
    'on error resume next 'returns True for success or False for error
    Dim ReservedChar As Byte
    Dim ReservedCharCount As Long
    Dim ByteStringLengthUnchanged As Long
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    'preset
    ByteStringLengthUnchanged = ByteStringLength
    'begin
    Call LZ77_RepeatStringBuffer_Reset(LZ77RepeatStringStructNumber, LZ77RepeatStringStructArray())
    Call LZ77_RepeatStringBuffer_Create(ByteStringLength, ByteString()) 'creates a table of repeating strings
    '
    'DEBUG
'    Dim StructLoop As Integer
'    Debug.Print "REPEATING STRINGS:"
'    For StructLoop = 1 To LZ77RepeatStringStructNumber
'        Call DISPLAYBYTESTRING(LZ77RepeatStringStructArray(StructLoop).ByteString())
'    Next StructLoop
    'END OF DEBUG
    '
    Call LZ77_GetReservedChar(ByteStringLength, ByteString(), ReservedChar, ReservedCharCount)
    Call LZ77_CompressStringSub(ByteStringLength, ByteString(), ReservedChar, ReservedCharCount, LZ77RepeatStringStructNumber, LZ77RepeatStringStructArray())
    'add GFCompressionHeader
    '
    If GFCompressionHeader_Preset(GFCompressionHeaderStructVar) = False Then GoTo Error:
    If GFCompressionHeader_Write(ByteStringLength, ByteString(), ByteStringLength, ByteStringLengthUnchanged) = False Then GoTo Error:
    '
    LZ77_CompressString = True 'ok
    Exit Function
Error:
    LZ77_CompressString = False 'error
    Exit Function
End Function

Private Sub LZ77_CompressStringSub(ByRef ByteStringLength As Long, ByRef ByteString() As Byte, ByVal ReservedChar As Byte, ByVal ReservedCharCount As Long, ByVal LZ77RepeatStringStructNumber As Integer, ByRef LZ77RepeatStringStructArray() As LZ77RepeatStringStruct)
    'on error resume next
    Dim InputByteStringLength As Long
    Dim OutputByteStringLength As Long
    Dim OutputByteString() As Byte
    Dim Temp1 As Long
    Dim Temp2 As Long
    Dim Temp3 As Long
    Dim StructLoop As Integer
    'preset
    Call IndexTableString_Reset
    OutputByteStringLength = 0 'reset
    ReDim OutputByteString(1 To ByteStringLength) As Byte 'cannot exceed this size
    'begin
    For Temp1 = 1& To ByteStringLength
        For StructLoop = 1 To LZ77RepeatStringStructNumber
            For Temp3 = 1& To LZ77RepeatStringStructArray(StructLoop).ByteStringLength
                '
                If (Temp1 + Temp3 - 1&) > ByteStringLength Then GoTo Jump1: 'verify
                '
                If Not (ByteString(Temp1 + Temp3 - 1&) = LZ77RepeatStringStructArray(StructLoop).ByteString(Temp3)) Then
                    GoTo Jump1:
                End If
            Next Temp3
            'a repeat string was found
            OutputByteStringLength = OutputByteStringLength + 1&
            OutputByteString(OutputByteStringLength) = ReservedChar
            Call IndexTableString_AddIndex(StructLoop)
            Temp1 = Temp1 + LZ77RepeatStringStructArray(StructLoop).ByteStringLength - 1& 'step over the chars in the input string
            GoTo Jump2:
Jump1: 'no repeat string found
        Next StructLoop
        'no repeat string was found
        OutputByteStringLength = OutputByteStringLength + 1&
        OutputByteString(OutputByteStringLength) = ByteString(Temp1)
        If ByteString(Temp1) = ReservedChar Then _
            Call IndexTableString_AddIndex(0) '0 for reserved char appeared in original uncompressed string
Jump2: 'a repeat string was found
    Next Temp1
    'DEBUG
    '
    'NOTE: now we have the compressed string and the IndexTableString
    'as well as the repeat strings, let's combine them and create the final output string.
    'The output string has the following format (without line breaking):
    'xxxx: length of uncompressed string
    'r: reserved char
    'yyyy: length if compressed string (without table index string etc.)
    'o: output string with varuious length
    'zzzz: index table string length
    'i: index table string with various length
    'rsl: repeat string length
    'rs: repeat string
    '
    'The last two items may appear more than once in the compressed string.
    '
    InputByteStringLength = ByteStringLength 'original size of passed uncompressed string
    ByteStringLength = 4& + 1& + 4& + OutputByteStringLength + 4& + IndexTableStringLength
    ReDim ByteString(1 To ByteStringLength) As Byte
    Call CopyMemory(ByteString(1), InputByteStringLength, 4&)
    Call CopyMemory(ByteString(5), ReservedChar, 1&)
    Call CopyMemory(ByteString(6), OutputByteStringLength, 4&)
    If (OutputByteStringLength) Then 'verify (important)
        Call CopyMemory(ByteString(10), OutputByteString(1), OutputByteStringLength)
    End If
    Call CopyMemory(ByteString(10& + OutputByteStringLength), IndexTableStringLength, 4&)
    If (IndexTableStringLength) Then 'verify (important)
        Call CopyMemory(ByteString(10& + OutputByteStringLength + 4&), IndexTableString(1), IndexTableStringLength)
    End If
    'NOTE: now add the repeat string table to the return string
    InputByteStringLength = ByteStringLength
    For StructLoop = 1 To LZ77RepeatStringStructNumber
        ByteStringLength = ByteStringLength + 4& + LZ77RepeatStringStructArray(StructLoop).ByteStringLength
    Next StructLoop
    ReDim Preserve ByteString(1 To ByteStringLength) As Byte
    ByteStringLength = InputByteStringLength + 1&
    For StructLoop = 1 To LZ77RepeatStringStructNumber
        Call CopyMemory( _
            ByteString(ByteStringLength), _
            LZ77RepeatStringStructArray(StructLoop).ByteStringLength, _
            4&)
        ByteStringLength = ByteStringLength + 4&
        Call CopyMemory( _
            ByteString(ByteStringLength), _
            LZ77RepeatStringStructArray(StructLoop).ByteString(1), _
            LZ77RepeatStringStructArray(StructLoop).ByteStringLength)
        ByteStringLength = ByteStringLength + LZ77RepeatStringStructArray(StructLoop).ByteStringLength
    Next StructLoop
    '
    'NOTE: ByteStringLength is the next write pos, substract one to get
    'original byte string length (tested).
    '
    ReDim Preserve ByteString(1 To (ByteStringLength - 1&)) As Byte 'shrink byte string if necessary
End Sub

'***REPEAT STRING BUFFER***
'NOTE: the repeat string buffer is an array of the type LZ77RepeatStructStruct.
'It stores the length of repeating strings and the strings itselves.

Private Sub LZ77_RepeatStringBuffer_Create(ByVal ByteStringLength As Long, ByRef ByteString() As Byte)
    'on eror resume next
    Dim Temp1 As Long
    Dim Temp2 As Long
    Dim Temp3 As Long
    Dim Temp4 As Long
    Dim TempByteStringLength As Long
    Dim TempByteString() As Byte
    Dim RepeatStringLengthMax As Long
    Dim u As Long
    'u = Val(Testfrm.Text3)
    u = 15 '***TEMP***
    'begin
    For Temp1 = 1& To (ByteStringLength - 1&)
        For Temp2 = (Temp1 + 1&) To ByteStringLength
            '
            If ByteString(Temp1) = ByteString(Temp2) Then
                'e.g. 'got got ': Temp1 = 1, Temp2 = 5
                '
                'NOTE: the current string has equal chars at the positions Temp1 and Temp2.
                'Loop forward to determinate the end of the repeated string.
                'The repeating string cannot be longer than the distance between
                'the first two equal chars and also not longer than the complete byte string.
                '
                RepeatStringLengthMax = MIN((Temp2 - Temp1), ByteStringLength - Temp2)
                '
                For Temp3 = 1& To RepeatStringLengthMax
                    '
                    If (Not (ByteString(Temp1 + Temp3) = ByteString(Temp2 + Temp3))) Or _
                        (Temp3 = RepeatStringLengthMax) Then 'already check next char
                        '
                        TempByteStringLength = Temp3  'don't add one as current char is not equal in both string-pieces
                        If TempByteStringLength > u Then 'verify
                            '
                            'NOTE: only strings with at least 3 chars are replaced through
                            'a shorter reserved string.
                            '
                            ReDim TempByteString(1 To TempByteStringLength) As Byte
                            Call CopyMemory(TempByteString(1), ByteString(Temp1), TempByteStringLength)
                            '
                            'Debug.Print "ADDED:"
                            'Call DISPLAYBYTESTRING(TempByteString())
                            'Tempstr$ = "got flott got flott got got got"
                            '
                            Call LZ77_RepeatStringBuffer_AddItem(TempByteStringLength, TempByteString())
                            Temp1 = Temp1 + Temp3
                            Temp2 = Temp1 'will be increased by one through 'Next Temp2'
                            Exit For
                        Else
                            'Temp1 = Temp1 + 1 'no!
                            'Temp2 = Temp1 + 1 'reset (important, don't use 0 as Temp2 must run 'behind' Temp1)
                            Exit For
                        End If
                    End If
                    '
                Next Temp3
            End If
        Next Temp2
    Next Temp1
End Sub

Private Sub LZ77_RepeatStringBuffer_Reset(ByRef RepeatStringStructNumber As Integer, ByRef RepeatStringStructArray() As LZ77RepeatStringStruct)
    'on error resume next
    RepeatStringStructNumber = 0 'reset
    ReDim RepeatStringStructArray(1 To 1) As LZ77RepeatStringStruct 'reset
End Sub

Private Sub LZ77_RepeatStringBuffer_AddItem(ByVal ByteStringLength As Long, ByRef ByteString() As Byte)
    'on error resume next
    Dim StructLoop As Integer
    Dim Temp As Long
    'verify
    If ByteStringLength = 0 Then Exit Sub 'verify
    'begin
    'Call DISPLAYBYTESTRING(ByteString())
    For StructLoop = 1 To LZ77RepeatStringStructNumber
        If ByteStringLength = LZ77RepeatStringStructArray(StructLoop).ByteStringLength Then
            For Temp = 1 To ByteStringLength
                If Not (LZ77RepeatStringStructArray(StructLoop).ByteString(Temp) = ByteString(Temp)) Then _
                    GoTo Jump:
            Next Temp
            Exit Sub 'passed string already existing
Jump:
        End If
    Next StructLoop
    'add current byte string to buffer
    If Not (LZ77RepeatStringStructNumber = 32766) Then 'verify
        LZ77RepeatStringStructNumber = LZ77RepeatStringStructNumber + 1
        ReDim Preserve LZ77RepeatStringStructArray(1 To LZ77RepeatStringStructNumber) As LZ77RepeatStringStruct
        LZ77RepeatStringStructArray(LZ77RepeatStringStructNumber).ByteStringLength = ByteStringLength 'cannot be 0
        ReDim LZ77RepeatStringStructArray(LZ77RepeatStringStructNumber).ByteString(1 To ByteStringLength) As Byte
        Call CopyMemory(LZ77RepeatStringStructArray(LZ77RepeatStringStructNumber).ByteString(1), ByteString(1), ByteStringLength)
    Else
        Exit Sub 'error
    End If
End Sub

'***END OF REPEAT STRING BUFFER***
'***INDEXTABLESTRING***
'NOTE: the IndexTableString is a byte array that stores references to LZ77RepeatStringStructArray()
'elements (in the form of intger values). If an index is 0 then the reserved char must be added to the string
'to decompress, if the index is not 0 then LZ77RepeatStructStructArray([index]).ByteString() must be added
'to the string to decomrpess.

Private Sub IndexTableString_Reset()
    'on error resume next
    IndexTableStringLength = 0 'reset
    Dim IndexTableString(1 To 1) As Byte
End Sub

Private Sub IndexTableString_AddIndex(ByVal LZ77RepeatStringStructIndex As Integer)
    'on erro resume next
    '
    'NOTE: if the added index is 0 then the reserved char appeared in the
    'uncompressed string (no data of the repeat string buffer must be accessed).
    '
    IndexTableStringLength = IndexTableStringLength + 2&
    ReDim Preserve IndexTableString(1 To IndexTableStringLength) As Byte
    Call CopyMemory(IndexTableString(IndexTableStringLength - 1&), LZ77RepeatStringStructIndex, 2&) 'copy a 16 bit Integer value
End Sub

'***END OF INDEX TABLE STRING***
'***OTHER***

Private Sub LZ77_GetReservedChar(ByVal ByteStringLength As Long, ByRef ByteString() As Byte, ByRef ReservedChar As Byte, ByRef ReservedCharCount As Long)
    'on error resume next 'returns char that appears with the lowest frequency in passed byte string
    Dim CharCountArray(0 To 255) As Long
    Dim CharCountMin As Long
    Dim Temp As Long
    'begin
    For Temp = 1& To ByteStringLength
        CharCountArray(ByteString(Temp)) = CharCountArray(ByteString(Temp)) + 1&
    Next Temp
    CharCountMin = 256& ^ 3& 'preset
    For Temp = 0& To 255&
        If CharCountArray(Temp) < CharCountMin Then _
            CharCountMin = CharCountArray(Temp)
    Next Temp
    For Temp = 255& To 0& Step (-1&) 'prefer an other char than 0 because of DISPLAYBYTESTRING() for debugging
        If CharCountArray(Temp) = CharCountMin Then
            ReservedChar = CByte(Temp)
            ReservedCharCount = CharCountArray(Temp) 'although not used
            Exit Sub 'ok
        End If
    Next Temp
    ReservedChar = CByte(0)
    ReservedCharCount = 0& 'although not used
    Exit Sub 'error (should not happen)
End Sub

'***END OF OTHER***
'*******************************END OF LZ77 COMPRESSION********************************
'**********************************LZ77 DECOMPRESSION**********************************

Public Function LZ77_DecompressString(ByRef ByteStringLength As Long, ByRef ByteString() As Byte, ByRef BlockLengthProcessed As Long) As Boolean
    'on error resume next 'returns True for success or False for error
    Dim InputByteStringLength As Long
    Dim IndexTableStringPointer As Integer '1, 2, ...
    Dim StructIndex As Integer
    Dim ReservedChar As Byte
    Dim Temp As Long
    Dim TempByteStringLength As Long
    Dim TempByteString() As Byte
    Dim OutputByteStringLength  As Long 'length of compressed string without index table string etc.
    Dim OutputByteString() As Byte
    Dim OutputByteStringWritePos As Long
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    'preset
    '
    If GFCompressionHeader_Preset(GFCompressionHeaderStructVar) = False Then GoTo Error:
    If GFCompressionHeader_Read(ByteStringLength, ByteString(), GFCompressionHeaderStructVar.BlockLengthCompressed, GFCompressionHeaderStructVar.BlockLengthOriginal) = False Then GoTo Error:
    If GFCompressionHeader_Remove(ByteStringLength, ByteString(), GFCompressionHeaderStructVar, BlockLengthProcessed) = False Then GoTo Error:
    '
    InputByteStringLength = ByteStringLength 'original length of passed byte string without header
    '
    'preset
    Call CopyMemory(ByteStringLength, ByteString(1), 4&)
    If (ByteStringLength) Then 'verify (important)
        ReDim OutputByteString(1 To ByteStringLength) As Byte
    End If
    Call CopyMemory(ReservedChar, ByteString(5), 1&)
    Call CopyMemory(OutputByteStringLength, ByteString(6), 4&)
    Call CopyMemory(IndexTableStringLength, ByteString(10& + OutputByteStringLength), 4&)
    If (IndexTableStringLength) Then 'verify (important)
        ReDim IndexTableString(1 To IndexTableStringLength) As Byte
        Call CopyMemory(IndexTableString(1), ByteString(10& + OutputByteStringLength + 4&), IndexTableStringLength)
    End If
    Call LZ77_RepeatStringBuffer_Reset(LZ77RepeatStringStructNumber, LZ77RepeatStringStructArray())
    Temp = 10& + OutputByteStringLength + 4& + IndexTableStringLength
    Do
        If (Temp + 4&) > InputByteStringLength Then Exit Do
        Call CopyMemory(TempByteStringLength, ByteString(Temp), 4&)
        Temp = Temp + 4&
        If (Temp + TempByteStringLength) > InputByteStringLength Then Exit Do
        ReDim TempByteString(1 To TempByteStringLength) As Byte
        Call CopyMemory(TempByteString(1), ByteString(Temp), TempByteStringLength)
        Temp = Temp + TempByteStringLength
        Call LZ77_RepeatStringBuffer_AddItem(TempByteStringLength, TempByteString())
    Loop
    '
    'DEBUG
'    For Temp = 1 To LZ77RepeatStringStructNumber
'        Debug.Print "REPEAT STRING #" + LTrim$(Str$(Temp))
'        Call DISPLAYBYTESTRING(LZ77RepeatStringStructArray(Temp).ByteString())
'    Next Temp
    'DEBUG
    '
    'begin
    For Temp = 10& To (10& + OutputByteStringLength - 1&) 'compressed string starts at position 10
        Select Case ByteString(Temp)
        Case ReservedChar
            IndexTableStringPointer = IndexTableStringPointer + 2
            Call CopyMemory(StructIndex, IndexTableString(IndexTableStringPointer - 1), 2&)
            If StructIndex = 0 Then
                OutputByteStringWritePos = OutputByteStringWritePos + 1&
                OutputByteString(OutputByteStringWritePos) = ByteString(Temp)
            Else
                OutputByteStringWritePos = OutputByteStringWritePos + 1&
                Call CopyMemory(OutputByteString(OutputByteStringWritePos), _
                    LZ77RepeatStringStructArray(StructIndex).ByteString(1), _
                    LZ77RepeatStringStructArray(StructIndex).ByteStringLength)
                OutputByteStringWritePos = OutputByteStringWritePos + _
                    LZ77RepeatStringStructArray(StructIndex).ByteStringLength - 1& 'one has already been added
            End If
        Case Else
            OutputByteStringWritePos = OutputByteStringWritePos + 1&
            OutputByteString(OutputByteStringWritePos) = ByteString(Temp)
        End Select
    Next Temp
    ByteStringLength = OutputByteStringWritePos
    ReDim ByteString(1 To ByteStringLength) As Byte
    Call CopyMemory(ByteString(1), OutputByteString(1), ByteStringLength)
    LZ77_DecompressString = True 'ok
    Exit Function
Error:
    LZ77_DecompressString = False 'error
    Exit Function
End Function

'******************************END OF LZ77 DECOMPRESSION*******************************
'*****************************************OTHER****************************************

Private Function MIN(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    'on error Resume Next 'use for i.e. CopyMemory(a(1), ByVal b, MIN(UBound(a()), Len(b))
    If Value1 < Value2 Then
        MIN = Value1
    Else
        MIN = Value2
    End If
End Function

