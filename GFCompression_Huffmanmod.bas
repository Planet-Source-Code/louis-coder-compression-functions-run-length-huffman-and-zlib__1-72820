Attribute VB_Name = "GFCompression_Huffmanmod"
Option Explicit
'(c)2001 by Louis.
Private Declare Function DLLHuffman_CompressString Lib "cmprss10.dll" Alias "Huffman_CompressString" (ByRef HT_CodeStringStructArray As Any, ByVal ByteStringLength As Long, ByRef ByteString As Any, ByVal CompressedStringLength As Long, ByRef CompressedString As Any) As Long
Private Declare Function DLLHuffman_DecompressString Lib "cmprss10.dll" Alias "Huffman_DecompressString" (ByRef HuffmanDecompressStructArray As Any, ByVal ByteStringLength As Long, ByRef ByteString As Any, ByVal BitReadStartPos As Long, ByVal OutputByteStringLength As Long, ByRef OutputByteString As Any) As Long
'general use
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'****************************************HUFFMAN***************************************
'NOTE: the huffman compression assigns shorter bit codes to chars that appear
'with a high frequency in the string to compress, and longer bit codes to chars that
'appear with a low frequency.
'The HT_CodeStringStruct contains the chars and the related bit code length and the
'bit code itself.
'The HT_CodeStringStruct is temporary used to create the HT_CodeStringStruct.
'
'Every char of the char set (code 0 - 255) gets a bit code assigned.
'The compressed string has the following format:
'
'FFFFTTTTB[*256]C[*TTTT]X
'F: original string (file) length
'T: total length (bytes) of all bit codes (code 0 - 255)
'B: length (bits) of related bit code
'C: bit code
'X: compressed string
'
'T + B + C: HuffmanTreeCode[String/ByteString]
'

Public Function Huffman_CompressString(ByRef ByteStringLength As Long, ByRef ByteString() As Byte) As Boolean
    'on error resume next
    Dim HT_CodeStringStructNumber As Integer
    Dim HT_CodeStringStructArray(0 To 255) As HT_CodeStringStruct
    Dim HT_CharInfoStructVar As HT_CharInfoStruct
    Dim CompressedStringLength As Long
    Dim CompressedString() As Byte
    Dim CompressedStringIndex As Long
    Dim TreeByteStringLength As Long
    Dim TreeByteString() As Byte
    Dim InputByteStringLength As Long
    'end of compression
    Dim ByteStringLengthUnchanged As Long
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    Dim Temp As Long
    Dim Tempdbl#
    'preset
    ByteStringLengthUnchanged = ByteStringLength
    'begin
    'create huffman tree struct
    HT_CodeStringStructNumber = 256 'preset
    Call HTCS_HT_CodeStringStruct_Define(ByteStringLength, ByteString(), HT_CharInfoStructVar, HT_CodeStringStructNumber, HT_CodeStringStructArray())
    'calculate length of the compressed data
    For Temp = 0 To 255
        Tempdbl# = Tempdbl# + _
            CDbl(HT_CharInfoStructVar.CharFrequencyArray(Temp)) * _
            CDbl(HT_CodeStringStructArray(Temp).CodeLength)
    Next Temp
    CompressedStringLength = CLng(-Int(-(Tempdbl# / 8#)))
    If Not (CompressedStringLength = 0) Then
        ReDim CompressedString(1 To CompressedStringLength) As Byte
    Else
        GoTo Error:
    End If
    'compress input string
    If IsVCCompressionAvailable = True Then
        Call Huffman_CompressString_VC(HT_CodeStringStructArray(), ByteStringLength, ByteString(), CompressedStringLength, CompressedString())
    Else
        Call Huffman_CompressString_VB(HT_CodeStringStructArray(), ByteStringLength, ByteString(), CompressedStringLength, CompressedString())
    End If
    'create the huffman tree code string
    Call HTCS_TreeCodeByteString_Define(HT_CodeStringStructArray(), TreeByteStringLength, TreeByteString())
    'create the final compressed string
    InputByteStringLength = ByteStringLength
    ByteStringLength = 4 + CompressedStringLength + TreeByteStringLength
    ReDim ByteString(1 To ByteStringLength) As Byte
    Call CopyMemory(ByteString(1), InputByteStringLength, 4)
    Call CopyMemory(ByteString(5), TreeByteString(1), TreeByteStringLength)
    Call CopyMemory(ByteString(5 + TreeByteStringLength), CompressedString(1), CompressedStringLength)
    'add GFCompressionHeader
    '
    If GFCompressionHeader_Preset(GFCompressionHeaderStructVar) = False Then GoTo Error:
    If GFCompressionHeader_Write(ByteStringLength, ByteString(), ByteStringLength, ByteStringLengthUnchanged) = False Then GoTo Error:
    '
    Huffman_CompressString = True 'ok
    Exit Function
Error:
    Huffman_CompressString = False 'error
    Exit Function
End Function

Private Sub Huffman_CompressString_VC(ByRef HT_CodeStringStructArray() As HT_CodeStringStruct, ByVal ByteStringLength As Long, ByRef ByteString() As Byte, ByVal CompressedStringLength As Long, ByRef CompressedString() As Byte)
    'on error resume next
    Call DLLHuffman_CompressString(HT_CodeStringStructArray(0), ByteStringLength, ByteString(1), CompressedStringLength, CompressedString(1))
End Sub

Private Sub Huffman_CompressString_VB(ByRef HT_CodeStringStructArray() As HT_CodeStringStruct, ByVal ByteStringLength As Long, ByRef ByteString() As Byte, ByVal CompressedStringLength As Long, ByRef CompressedString() As Byte)
    'on error resume next
    Dim CompressedStringBitWritePos As Long
    Dim CompressedStringIndex As Long
    Dim Temp1 As Long
    Dim Temp2 As Long
    'begin
    For Temp1 = 1& To ByteStringLength
        '
        'NOTE: CompressedStringBitWritePos indendicates the bit in
        'CompressedString() where the next code string is to be 'added'.
        '
        For Temp2 = 1& To HT_CodeStringStructArray(ByteString(Temp1)).CodeLength
            '
            CompressedStringBitWritePos = CompressedStringBitWritePos + 1&
            '
            If (HT_CodeStringStructArray(ByteString(Temp1)).CodeArray(Temp2)) Then
                '
                'NOTE: earlier 'down-converting' is faster:
                'a = CLng((9# - 1#) / 8# + 1#) 'slower
                'a = CLng((9# - 1#) / 8#) + 1& 'faster
                '
                CompressedStringIndex = ((CompressedStringBitWritePos - 1&) \ 8&) + 1&
'                CompressedString(CompressedStringIndex) = _
'                    CompressedString(CompressedStringIndex) _
'                    Or _
'                    HT_CodeStringStructArray(ByteString(Temp1)).CodeArray(Temp2) * _
'                    (2& ^ (7& - ((CompressedStringBitWritePos + 7&) Mod 8&)))
                'NOTE: Mod is much faster than Select Case (tested).
                'NOTE: copying Byte vars to Long vars before using Or did NOT increase speed.
                'NOTE: as 2 ^ is much slower than a Select Case statement checking
                '8 values we use the Select Case statement:
                '
                Select Case (CompressedStringBitWritePos Mod 8&)
                Case 1& 'propability is the same for every value
                    CompressedString(CompressedStringIndex) = CompressedString(CompressedStringIndex) _
                        Or HT_CodeStringStructArray(ByteString(Temp1)).CodeArray(Temp2) * 128& '2 ^ 7 = 128
                Case 2&
                    CompressedString(CompressedStringIndex) = CompressedString(CompressedStringIndex) _
                        Or HT_CodeStringStructArray(ByteString(Temp1)).CodeArray(Temp2) * 64& '2 ^ 6 = 64
                Case 3&
                    CompressedString(CompressedStringIndex) = CompressedString(CompressedStringIndex) _
                        Or HT_CodeStringStructArray(ByteString(Temp1)).CodeArray(Temp2) * 32& '2 ^ 5 = 32
                Case 4&
                    CompressedString(CompressedStringIndex) = CompressedString(CompressedStringIndex) _
                        Or HT_CodeStringStructArray(ByteString(Temp1)).CodeArray(Temp2) * 16& '2 ^ 4 = 16
                Case 5&
                    CompressedString(CompressedStringIndex) = CompressedString(CompressedStringIndex) _
                        Or HT_CodeStringStructArray(ByteString(Temp1)).CodeArray(Temp2) * 8& '2 ^ 3 = 8
                Case 6&
                    CompressedString(CompressedStringIndex) = CompressedString(CompressedStringIndex) _
                        Or HT_CodeStringStructArray(ByteString(Temp1)).CodeArray(Temp2) * 4& '2 ^ 2 = 4
                Case 7&
                    CompressedString(CompressedStringIndex) = CompressedString(CompressedStringIndex) _
                        Or HT_CodeStringStructArray(ByteString(Temp1)).CodeArray(Temp2) * 2& '2 ^ 1 = 1
                Case 0& 'as we removed '+ 7&' (see original calculation that is commented-out) things get mysterious and we must use 0 and not 8 here
                    CompressedString(CompressedStringIndex) = CompressedString(CompressedStringIndex) _
                        Or HT_CodeStringStructArray(ByteString(Temp1)).CodeArray(Temp2) '* 1& '2 ^ 0 = 1
                End Select
            End If
        Next Temp2
    Next Temp1
End Sub

Public Function Huffman_DecompressString(ByRef ByteStringLength As Long, ByRef ByteString() As Byte, ByRef BlockLengthProcessed As Long) As Boolean
    'on error resume next 'returns True for success or False for error
    Dim HT_CodeStringStructNumber As Integer
    Dim HT_CodeStringStructArray(0 To 255) As HT_CodeStringStruct
    Dim HT_TreeStringStructVar As HT_TreeStringStruct
    Dim HTDC_CodeStringStructArray(0 To 255) As HTDC_CodeStringStruct
    Dim OutputStringStartPos As Long 'start pos of decompressed string in compressed string (after tree data)
    Dim OutputByteStringLength As Long
    Dim OutputByteString() As Byte
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    'preset
    '
    If GFCompressionHeader_Preset(GFCompressionHeaderStructVar) = False Then GoTo Error:
    If GFCompressionHeader_Read(ByteStringLength, ByteString(), GFCompressionHeaderStructVar.BlockLengthCompressed, GFCompressionHeaderStructVar.BlockLengthOriginal) = False Then GoTo Error:
    If GFCompressionHeader_Remove(ByteStringLength, ByteString(), GFCompressionHeaderStructVar, BlockLengthProcessed) = False Then GoTo Error:
    '
    'retain huffman tree containing char codes
    HT_CodeStringStructNumber = 256
    Call HTDC_HT_TreeStringStructVar_Define(HT_TreeStringStructVar, ByteStringLength, ByteString())
    Call HTDC_HT_CodeStringStruct_Define(HT_CodeStringStructArray(), HT_TreeStringStructVar)
    Call HTDC_CodeStringStruct_Define(HT_CodeStringStructArray(), HTDC_CodeStringStructArray())
    '
    'NOTE: HTDC_CodeStringStructArray(0) contains the shortest,
    'HTDC_CodeStringStructArray(255) the longest bit code.
    'Note that the shortest bit code will also appear the most requently so
    'that the decompression becomes fast.
    '
    Call CopyMemory(OutputByteStringLength, ByteString(1), 4) 'get decompressed string length
    ReDim OutputByteString(1 To OutputByteStringLength) As Byte
    OutputStringStartPos = 4 + 4 + HT_TreeStringStructVar.TreeByteStringLength + 1 'ok
    'decompress string
    If IsVCCompressionAvailable = True Then
        Call Huffman_DecompressString_VC(HTDC_CodeStringStructArray(), ByteStringLength, ByteString(), _
             (OutputStringStartPos * 8), OutputByteStringLength, OutputByteString())
    Else
        Call Huffman_DecompressString_VB(HTDC_CodeStringStructArray(), ByteStringLength, ByteString(), _
             (OutputStringStartPos * 8), OutputByteStringLength, OutputByteString())
    End If
    'create final, decompressed return string
    ByteStringLength = OutputByteStringLength
    ReDim ByteString(1 To OutputByteStringLength) As Byte
    Call CopyMemory(ByteString(1), OutputByteString(1), OutputByteStringLength) 'transfer decompressed string
    Huffman_DecompressString = True 'ok
    Exit Function
Error:
    Huffman_DecompressString = False 'error
    Exit Function
End Function

Private Sub Huffman_DecompressString_VC(ByRef HTDC_CodeStringStructArray() As HTDC_CodeStringStruct, ByVal ByteStringLength As Long, ByRef ByteString() As Byte, ByVal BitReadStartPos As Long, ByVal OutputByteStringLength As Long, ByRef OutputByteString() As Byte)
    'on error resume next
    Call DLLHuffman_DecompressString(HTDC_CodeStringStructArray(0), ByteStringLength, ByteString(1), BitReadStartPos, OutputByteStringLength, OutputByteString(1))
End Sub

Private Sub Huffman_DecompressString_VB(ByRef HTDC_CodeStringStructArray() As HTDC_CodeStringStruct, ByVal ByteStringLength As Long, ByRef ByteString() As Byte, ByVal BitReadStartPos As Long, ByVal OutputByteStringLength As Long, ByRef OutputByteString() As Byte)
    'on error resume next
    Dim CodeBufLength As Long
    Dim CodeBufArray(1 To 256) As Byte 'current code from input string
    Dim ByteStringIndex As Long
    Dim ByteStringLong As Long 'part of ByteString()
    Dim BitReadPos As Long
    Dim Temp1 As Long
    Dim Temp2 As Long
    Dim Temp3 As Long
    Dim Temp4 As Long
    'begin
    BitReadPos = BitReadStartPos
    For Temp1 = 1& To OutputByteStringLength
        'Temp1 = write pos in output string
        For Temp2 = 1& To 2048& 'read up to 256 chars (2048 bits) into CodeBufArray()
            'Temp2 = write pos in buffer array
            BitReadPos = BitReadPos + 1&
            ByteStringIndex = ((BitReadPos - 1&) \ 8&)
            '
            'If (ByteString(ByteStringIndex) And (2& ^ (7& - ((BitReadPos + 7&) Mod 8&)))) Then
            'NOTE: copying Byte vars to Long vars before using And did NOT increase speed.
            '
            Select Case BitReadPos Mod 8&
            Case 1&
                If (ByteString(ByteStringIndex) And 128&) Then
                    CodeBufArray(Temp2) = 1
                Else
                    CodeBufArray(Temp2) = 0
                End If
            Case 2&
                If (ByteString(ByteStringIndex) And 64&) Then
                    CodeBufArray(Temp2) = 1
                Else
                    CodeBufArray(Temp2) = 0
                End If
            Case 3&
                If (ByteString(ByteStringIndex) And 32&) Then
                    CodeBufArray(Temp2) = 1
                Else
                    CodeBufArray(Temp2) = 0
                End If
            Case 4&
                If (ByteString(ByteStringIndex) And 16&) Then
                    CodeBufArray(Temp2) = 1
                Else
                    CodeBufArray(Temp2) = 0
                End If
            Case 5&
                If (ByteString(ByteStringIndex) And 8&) Then
                    CodeBufArray(Temp2) = 1
                Else
                    CodeBufArray(Temp2) = 0
                End If
            Case 6&
                If (ByteString(ByteStringIndex) And 4&) Then
                    CodeBufArray(Temp2) = 1
                Else
                    CodeBufArray(Temp2) = 0
                End If
            Case 7&
                If (ByteString(ByteStringIndex) And 2&) Then
                    CodeBufArray(Temp2) = 1
                Else
                    CodeBufArray(Temp2) = 0
                End If
            Case 0&
                If (ByteString(ByteStringIndex) And 1&) Then
                    CodeBufArray(Temp2) = 1
                Else
                    CodeBufArray(Temp2) = 0
                End If
            End Select
            '
            For Temp3 = HTDC_CodeStringStructArray(1).StartIndexArray(Temp2) To HTDC_CodeStringStructArray(1).EndIndexArray(Temp2) 'check out if string in CodeBufArray() is equal to any of the existing non-zero char strings
                'NOTE: HTDC_CodeStringStructArray(x).NonZeroLengthCharCodeNumber is constant for all x.
                If (Temp2 = HTDC_CodeStringStructArray(Temp3).CharCodeArrayLength) Then 'Temp2 is the current code string length
                    For Temp4 = 1& To HTDC_CodeStringStructArray(Temp3).CharCodeArrayLength 'check all chars of code
                        If Not (HTDC_CodeStringStructArray(Temp3).CharCodeArray(Temp4) = CodeBufArray(Temp4)) Then
                            GoTo Skip:
                        End If
                    Next Temp4
                    OutputByteString(Temp1) = HTDC_CodeStringStructArray(Temp3).Char
                    GoTo Jump:
Skip:
                End If
            Next Temp3
        Next Temp2
Jump:
    Next Temp1
End Sub

'************************************END OF HUFFMAN************************************

