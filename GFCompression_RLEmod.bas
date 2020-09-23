Attribute VB_Name = "GFCompression_RLEmod"
Option Explicit
'(c)2001 by Louis.
'RLE_[Decomrpess/Compress]String_VC
Private Declare Function DLLRLE_CompressString Lib "cmprss10.dll" Alias "RLE_CompressString" (ByVal ByteStringLength As Long, ByRef ByteString As Any) As Long
Private Declare Function DLLRLE_DecompressString Lib "cmprss10.dll" Alias "RLE_DecompressString" (ByVal ByteStringLength As Long, ByRef ByteString As Any) As Long
Private Declare Function GetCompressionByteString Lib "cmprss10.dll" (ByVal ByteStringLength As Long, ByRef ByteString As Any) As Long
'general use
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'*****************************************RLE******************************************
'NOTE: it follows the code of a simple run length encoding.
'If a repetition longer than 3 chars is detected, a reserved char is written to the
'compressed string. The following char determins the length of the repetition
'(1 based):
'
'3-254: repetition is 3 to 254 chars
'255: repetition is greater than 254 (2 * 254, 3 * 254 etc.) chars
'256: reserved char was compressed
'
'If a repetition of exactly 3 chars is detected nothing special is done as using
'a second reserved char for triple repetition has hardly effect (0.2% in a LZ77_RepeatStringBuffer_Create).
'
'The compressed string has the following format:
'
'SSSSRX[*?]
'
'S: size of uncompressed string
'R: reserved char
'X: compressed string

Public Function RLE_CompressString(ByRef ByteStringLength As Long, ByRef ByteString() As Byte) As Boolean
    'on error resume next 'returns True for success or False for error
    Dim InputByteStringLength As Long
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    'begin
    If IsVCCompressionAvailable = True Then
        InputByteStringLength = ByteStringLength
        RLE_CompressString = RLE_CompressString_VC(ByteStringLength, ByteString())
    Else
        InputByteStringLength = ByteStringLength
        RLE_CompressString = RLE_CompressString_VB(ByteStringLength, ByteString())
    End If
    '
    If GFCompressionHeader_Preset(GFCompressionHeaderStructVar) = False Then GoTo Error:
    If GFCompressionHeader_Write(ByteStringLength, ByteString(), ByteStringLength, InputByteStringLength) = False Then GoTo Error:
    '
    Exit Function
Error:
    RLE_CompressString = False 'error
    Exit Function
End Function

Private Function RLE_CompressString_VC(ByRef ByteStringLength As Long, ByRef ByteString() As Byte) As Boolean
    'on error resume next 'returns True for success, False for error
    Dim ByteStringLengthNew As Long
    'begin
    '
    ByteStringLengthNew = DLLRLE_CompressString(ByteStringLength, ByteString(1))
    '
    If Not (ByteStringLengthNew = 0) Then 'verify
        '
        ReDim ByteString(1 To ByteStringLengthNew) As Byte
        Call GetCompressionByteString(ByteStringLengthNew, ByteString(1))
        '
    Else
        ReDim ByteString(1 To 1) As Byte 'reset
    End If
    '
    If (ByteStringLengthNew = 0) And (Not (ByteStringLength = 0)) Then
        RLE_CompressString_VC = False 'error
    Else
        RLE_CompressString_VC = True 'ok
    End If
    '
End Function

Private Function RLE_CompressString_VB(ByRef ByteStringLength As Long, ByRef ByteString() As Byte) As Boolean
    'on error resume next
    Dim ReservedChar As Byte
    Dim Char As Long 'repeating char, use the type Long to save conversion time
    Dim CharNumber As Long 'number of char repititions
    Dim CharFrequencyArray(0 To 255) As Long
    Dim CharFrequencyMin As Long
    Dim InputByteStringLength As Long
    Dim OutputByteStringLength As Long
    Dim OutputByteString() As Byte
    Dim Temp1 As Long
    Dim Temp2 As Long
    'preset
    '
    'NOTE: before the original compression begins a reserved char must be determinated.
    'Therefore a char in the passed string with the lowest frequency is seeked.
    'This char becomes the reserved char and only appears in the compressed string
    'to display that a special decompression action is to be performed.
    '
    For Temp1 = 1& To ByteStringLength
        CharFrequencyArray(ByteString(Temp1)) = CharFrequencyArray(ByteString(Temp1)) + 1
    Next Temp1
    '
    CharFrequencyMin = 256& ^ 3& 'preset
    '
    For Temp1 = 0& To 255&
        If CharFrequencyArray(Temp1) < CharFrequencyMin Then CharFrequencyMin = CharFrequencyArray(Temp1)
    Next Temp1
    For Temp1 = 0& To 255&
        If CharFrequencyArray(Temp1) = CharFrequencyMin Then
            ReservedChar = CInt(Temp1) 'no problem if there's only 1 char in the string to compress, algorithm works anyway (tested)
            Exit For
        End If
    Next Temp1
    '
    OutputByteStringLength = ByteStringLength + CharFrequencyMin 'in worst case the reserved char must be compressed that often
    ReDim OutputByteString(1 To OutputByteStringLength) As Byte
    '
    If Not (ByteStringLength = 0&) Then
        If ByteString(1) = 0 Then Char = 255 Else Char = 0 'preset
    End If
    '
    'begin
    '
    OutputByteStringLength = 0& 'reset
    For Temp1 = 1& To ByteStringLength
        If ByteString(Temp1) = Char Then
            CharNumber = CharNumber + 1&
        Else
            Select Case CharNumber
            Case 0&
                'no char repetition
            Case 1&
                OutputByteStringLength = OutputByteStringLength + 1&
                OutputByteString(OutputByteStringLength) = ByteString(Temp1 - 1&)
                CharNumber = 0& 'reset
            Case 2&
                OutputByteStringLength = OutputByteStringLength + 2&
                OutputByteString(OutputByteStringLength - 1&) = ByteString(Temp1 - 2&)
                OutputByteString(OutputByteStringLength) = ByteString(Temp1 - 1&)
                CharNumber = 0& 'reset
            Case Else
                OutputByteStringLength = OutputByteStringLength + 1&
                OutputByteString(OutputByteStringLength) = ReservedChar
                For Temp2 = 1& To (-Int(-CharNumber / 254&))
                    If Not (Temp2 = (-Int(-CharNumber / 254&))) Then
                        OutputByteStringLength = OutputByteStringLength + 1&
                        OutputByteString(OutputByteStringLength) = 254
                    Else
                        OutputByteStringLength = OutputByteStringLength + 1&
                        OutputByteString(OutputByteStringLength) = CByte(CharNumber - ((Temp2 - 1&) * 254&) - 1&) '1 to 0 based
                    End If
                Next Temp2
                CharNumber = 0& 'reset
            End Select
            If ByteString(Temp1) = ReservedChar Then 'repetition of reserved char is not supported
                OutputByteStringLength = OutputByteStringLength + 2&
                OutputByteString(OutputByteStringLength - 1&) = ReservedChar
                OutputByteString(OutputByteStringLength) = 255
                CharNumber = 0& 'reset
            Else
                OutputByteStringLength = OutputByteStringLength + 1&
                OutputByteString(OutputByteStringLength) = ByteString(Temp1)
                CharNumber = 0& 'reset
            End If
        End If
        Char = ByteString(Temp1)
        If Temp1 = (ByteStringLength - 1&) Then
            If ByteString(Temp1 + 1&) = 0 Then Char = 255 Else Char = 0 'verify next char is NOT buffered (important, tested)
        End If
        If Char = ReservedChar Then 'repetition of reserved char is not supported (reserved char is that with the lowest frequency)
            If Not (Temp1 = ByteStringLength) Then 'otherwise Char is any way not checked any more
                If ByteString(Temp1 + 1&) = 0 Then Char = 255 Else Char = 0 'verify next char is NOT buffered (important, tested)
            End If
        End If
    Next Temp1
    'create final return string
    InputByteStringLength = ByteStringLength
    ByteStringLength = OutputByteStringLength + 5&
    ReDim ByteString(1 To ByteStringLength) As Byte
    Call CopyMemory(ByteString(1), InputByteStringLength, 4)
    Call CopyMemory(ByteString(5), ReservedChar, 1)
    Call CopyMemory(ByteString(6), OutputByteString(1), OutputByteStringLength)
    RLE_CompressString_VB = True 'ok
    Exit Function
Error:
    RLE_CompressString_VB = False 'error
    Exit Function
End Function

Public Function RLE_DecompressString(ByRef ByteStringLength As Long, ByRef ByteString() As Byte, ByRef BlockLengthProcessed As Long) As Boolean
    'on error resume next
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    'NOTE: BlockLengthProcessed is NOT equal to ByteStringLength.
    'preset
    '
    If GFCompressionHeader_Preset(GFCompressionHeaderStructVar) = False Then GoTo Error:
    If GFCompressionHeader_Read(ByteStringLength, ByteString(), GFCompressionHeaderStructVar.BlockLengthCompressed, GFCompressionHeaderStructVar.BlockLengthOriginal) = False Then GoTo Error:
    If GFCompressionHeader_Remove(ByteStringLength, ByteString(), GFCompressionHeaderStructVar, BlockLengthProcessed) = False Then GoTo Error:
    '
    'begin
    If IsVCCompressionAvailable = True Then
        RLE_DecompressString = RLE_DecompressString_VC(ByteStringLength, ByteString())
    Else
        RLE_DecompressString = RLE_DecompressString_VB(ByteStringLength, ByteString())
    End If
    Exit Function
Error:
    RLE_DecompressString = False 'error
    Exit Function
End Function

Public Function RLE_DecompressString_VC(ByRef ByteStringLength As Long, ByRef ByteString() As Byte) As Boolean
    'on error resume next 'returns True for success, False for error
    Dim ByteStringLengthNew As Long
    'begin (code almost equal to that of compression)
    '
    ByteStringLengthNew = DLLRLE_DecompressString(ByteStringLength, ByteString(1))
    '
    If Not (ByteStringLengthNew = 0) Then 'verify
        '
        ReDim ByteString(1 To ByteStringLengthNew) As Byte
        Call GetCompressionByteString(ByteStringLengthNew, ByteString(1))
        '
    Else
        ReDim ByteString(1 To 1) As Byte 'reset
    End If
    '
    If (ByteStringLengthNew = 0) And (Not (ByteStringLength = 0)) Then
        RLE_DecompressString_VC = False 'error
    Else
        RLE_DecompressString_VC = True 'ok
    End If
    '
End Function

Public Function RLE_DecompressString_VB(ByRef ByteStringLength As Long, ByRef ByteString() As Byte) As Boolean
    'on error Resume Next 'returns True for success, False for error
    Dim OutputByteStringWritePos As Long
    Dim OutputByteStringLength As Long
    Dim OutputByteString() As Byte
    Dim ReservedChar As Byte
    Dim Temp1 As Long
    Dim Temp2 As Long
    Dim Temp3 As Long
    'preset
    Call CopyMemory(OutputByteStringLength, ByteString(1), 4)
    Call CopyMemory(ReservedChar, ByteString(5), 1)
    ReDim OutputByteString(1 To OutputByteStringLength) As Byte
    'begin
    For Temp1 = 6& To ByteStringLength
        If ByteString(Temp1) = ReservedChar Then
            Temp2 = Temp1 + 1& 'read pos in input string
            Do
                If ByteString(Temp2) = 254 Then
                    For Temp3 = 1& To 254&
                        OutputByteStringWritePos = OutputByteStringWritePos + 1&
                        OutputByteString(OutputByteStringWritePos) = ByteString(Temp1 - 1&)
                    Next Temp3
                    Temp2 = Temp2 + 1& 'read next char of input string
                Else
                    If ByteString(Temp2) = 255 Then
                        OutputByteStringWritePos = OutputByteStringWritePos + 1&
                        OutputByteString(OutputByteStringWritePos) = ReservedChar
                        Exit Do
                    Else
                        For Temp3 = 1 To CLng(ByteString(Temp2)) + 1& '0 to 1 based
                            OutputByteStringWritePos = OutputByteStringWritePos + 1&
                            OutputByteString(OutputByteStringWritePos) = ByteString(Temp1 - 1&)
                        Next Temp3
                        Exit Do
                    End If
                End If
            Loop
            Temp1 = Temp2
        Else
            OutputByteStringWritePos = OutputByteStringWritePos + 1&
            OutputByteString(OutputByteStringWritePos) = ByteString(Temp1)
        End If
    Next Temp1
    'create final, decompressed return string
    ByteStringLength = OutputByteStringLength
    ReDim ByteString(1 To ByteStringLength) As Byte
    Call CopyMemory(ByteString(1), OutputByteString(1), ByteStringLength)
    RLE_DecompressString_VB = True 'ok
    Exit Function
End Function

'*************************************END OF RLE***************************************

