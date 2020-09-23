Attribute VB_Name = "GFCompression_HuffmanTreemod"
Option Explicit
'(c)2001 by Louis. Code to create and handle a Huffman tree.
'
'NOTE: as the structure types of these module are participated in
'rather complicated operations detailed descriptions are given.
'HT is a short form for HuffmanTree.
'If the name of a structure type has the prefix CS then this type
'is only to be used for compressing a string, if its name has the prefix
'DC then this type is only to be used for decompressing a string.
'
'general use
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'HTCS_HT_CodeStringStruct_Define
'
'NOTE: this structure contains information about a char's frequency
'in the uncompressed string. Generally the frequency of a char
'is one of the most important information for creating a Huffman tree.
'
Public Type HT_CharInfoStruct
    CharArray(0 To 255) As Byte
    CharFrequencyArray(0 To 255) As Long
End Type
'HTCS_HT_CodeStringStruct_Define
'
'NOTE: this structure is temporary used to fill HT_CodeStringStruct.
'
Public Type HT_CodeStringCreationStruct
    ByteStringLength As Long
    ByteString(1 To 256) As Byte
    ByteStringFrequency As Long 'cannot exceed max. input byte string length
End Type
'HTCS_HT_CodeStringStruct_Define
'
'NOTE: the following structure contains the final code strings
'of the chars used in the input string.
'ByteString(x) can either be 0 or 1 for any x.
'
Public Type HT_CodeStringStruct
    CodeLength As Long
    CodeArray(1 To 256) As Byte 'code string length should not exceed (256 / 2) bytes
End Type
'Huffman_DecompressString
'
'NOTE: the following structure stores the Huffman tree code string,
'i.e. the byte string that was created out of all available code strings.
'
'
Public Type HT_TreeStringStruct
    TreeByteStringBitCount As Long
    TreeByteStringLength As Long
    TreeByteString() As Byte
End Type
'Huffman_DecompressString
Public Type HTDC_CodeStringStruct
    Char As Byte
    CharCodeArrayLength As Long
    CharCodeArray(1 To 256) As Byte
    StartIndexArray(0 To 255) As Long
    EndIndexArray(0 To 255) As Long
End Type

'*******************************HUFFMANTREE: COMPRESSION*******************************

Public Sub HTCS_HT_CodeStringStruct_Define(ByVal ByteStringLength As Long, ByRef ByteString() As Byte, ByRef HT_CharInfoStructVar As HT_CharInfoStruct, ByVal HT_CodeStringStructNumber As Integer, ByRef HT_CodeStringStructArray() As HT_CodeStringStruct)
    'on error resume next 'creates a Huffman tree, see annotations in code for further information; HT_CodeStringStructNumber is ignored (should be 256)
    Dim HuffmanCharStructArray(0 To 255) As HT_CodeStringCreationStruct
    Dim HuffmanCharStructVar As HT_CodeStringCreationStruct 'everything's 0, used for resetting
    Dim CharFrequencyMax As Double
    Dim LastUsedIndex As Long 'last used HuffmanCharStructArray() (where byte string frequency is not 0)
    Dim TempHuffmanCharStruct As HT_CodeStringCreationStruct
    Dim TempByte As Byte
    Dim Temp1 As Long
    Dim Temp2 As Long
    '
    'NOTE: this sub also initializes the passed structures, use their
    'information in the further compressing process.
    '
    'NOTE: view Huffman ([1-4]).htm for further information.
    'The information provided there is not that good, but sufficient.
    '
    'initialize structure
    For Temp1 = 0& To 255&
        HuffmanCharStructArray(Temp1).ByteStringLength = 1&
        HuffmanCharStructArray(Temp1).ByteString(1) = CByte(Temp1)
        HT_CharInfoStructVar.CharArray(Temp1) = CByte(Temp1)
    Next Temp1
    'preset char frequency
    For Temp1 = 1& To ByteStringLength
        HuffmanCharStructArray(ByteString(Temp1)).ByteStringFrequency = _
            HuffmanCharStructArray(ByteString(Temp1)).ByteStringFrequency + 1&
        HT_CharInfoStructVar.CharFrequencyArray(ByteString(Temp1)) = _
            HuffmanCharStructArray(ByteString(Temp1)).ByteStringFrequency
    Next Temp1
    'create Huffman tree
    Do
        'sort strings by their frequency, HuffnanCharStructArray(0) contains the char that appears the most frequent
        Temp2 = 0& 'reset
        LastUsedIndex = 255& 'preset
ReDo:
        'get highest frequency
        CharFrequencyMax = 0& 'reset
        For Temp1 = Temp2 To 255&
            If HuffmanCharStructArray(Temp1).ByteStringFrequency > CharFrequencyMax Then
                CharFrequencyMax = HuffmanCharStructArray(Temp1).ByteStringFrequency
            End If
        Next Temp1
        'put all chars with current highest frequency 'at front' (Temp2)
        For Temp1 = Temp2 To 255&
            If HuffmanCharStructArray(Temp1).ByteStringFrequency = CharFrequencyMax Then
                If Not (Temp1 = Temp2) Then 'verify exchanging is necessary
                    'exchange 'Temp1 through Temp2'
                    Call CopyMemory(TempHuffmanCharStruct, HuffmanCharStructArray(Temp2), Len(TempHuffmanCharStruct))
                    Call CopyMemory(HuffmanCharStructArray(Temp2), HuffmanCharStructArray(Temp1), Len(HuffmanCharStructArray(Temp1)))
                    Call CopyMemory(HuffmanCharStructArray(Temp1), TempHuffmanCharStruct, Len(TempHuffmanCharStruct))
                End If
                If CharFrequencyMax > 0& Then LastUsedIndex = Temp2
                Temp2 = Temp2 + 1&
            End If
        Next Temp1
        If Not (Temp2 = 256&) Then GoTo ReDo: 'will become 256 in any case (even if CharFrequencyMax is 0)
        If Not (LastUsedIndex > 0&) Then Exit Do 'finished if only one byte string with related code string is existing (note that index is 0 based)
        '
        'NOTE: the array was ordered so that the string that appears the most frequent comes first.
        'Note that you can generally say 'char' instead of 'string' for the first loop run.
        '
        'NOTE: the two strings with the lowest frequency are token, and their frequency
        'chars and code is added and the result is stored it in the structure of the first string.
        'The second structure is reset and may not be used any more.
        'Example:
        '1st structure content:
        'e
        '207
        '0
        '2nd structure content:
        's
        '205
        '1
        'new 1st structure content:
        'es
        '412
        '01
        '
        'When the two strings with the lowest frequency are combined
        'the structure is sorted again.
        '
        'NOTE: when the two code strings are combined.
        'The following rules are important:
        '
        'The code array related to 'LastUsedIndex' is always extended by a '0'
        'and the code array related to 'LastUsedIndex - 1' is extended by a '1'.
        '
        'The new code bits are appended to the existing code string,
        'note that the original direction is REVERSED, so the code strings
        'must be swapped at the end of this sub.
        '
        For Temp1 = 1& To HuffmanCharStructArray(LastUsedIndex).ByteStringLength
            'add one further code bit
            HT_CodeStringStructArray(HuffmanCharStructArray(LastUsedIndex).ByteString(Temp1)).CodeLength = _
                HT_CodeStringStructArray(HuffmanCharStructArray(LastUsedIndex).ByteString(Temp1)).CodeLength + 1&
            HT_CodeStringStructArray(HuffmanCharStructArray(LastUsedIndex).ByteString(Temp1)).CodeArray( _
                HT_CodeStringStructArray(HuffmanCharStructArray(LastUsedIndex).ByteString(Temp1)).CodeLength) = 0&  'zero
        Next Temp1
        For Temp1 = 1& To HuffmanCharStructArray(LastUsedIndex - 1&).ByteStringLength
            'add one further code bit
            HT_CodeStringStructArray(HuffmanCharStructArray(LastUsedIndex - 1).ByteString(Temp1)).CodeLength = _
                HT_CodeStringStructArray(HuffmanCharStructArray(LastUsedIndex - 1).ByteString(Temp1)).CodeLength + 1&
            HT_CodeStringStructArray(HuffmanCharStructArray(LastUsedIndex - 1).ByteString(Temp1)).CodeArray( _
                HT_CodeStringStructArray(HuffmanCharStructArray(LastUsedIndex - 1).ByteString(Temp1)).CodeLength) = 1&  'one
        Next Temp1
        'NOTE: first copy string, then change length information.
        Call CopyMemory(HuffmanCharStructArray(LastUsedIndex - 1&).ByteString(HuffmanCharStructArray(LastUsedIndex - 1&).ByteStringLength + 1&), _
            HuffmanCharStructArray(LastUsedIndex).ByteString(1), HuffmanCharStructArray(LastUsedIndex).ByteStringLength)
        HuffmanCharStructArray(LastUsedIndex - 1&).ByteStringLength = _
            HuffmanCharStructArray(LastUsedIndex - 1&).ByteStringLength + HuffmanCharStructArray(LastUsedIndex).ByteStringLength
        HuffmanCharStructArray(LastUsedIndex - 1&).ByteStringFrequency = _
            HuffmanCharStructArray(LastUsedIndex - 1&).ByteStringFrequency + HuffmanCharStructArray(LastUsedIndex).ByteStringFrequency
        'reset second structure content
        Call CopyMemory(HuffmanCharStructArray(LastUsedIndex), HuffmanCharStructVar, Len(HuffmanCharStructVar))
    Loop
    '
    'NOTE: to check if the created tree is corrent you take a piece of paper,
    'mark a starting point and go to the left for 0 or to the right for 1.
    'If a tree with no missing or surplus branches appears, everything's alright.
    '
    'swap code strings
    For Temp1 = 0& To 255&
        For Temp2 = 1& To (HT_CodeStringStructArray(Temp1).CodeLength \ 2&) '\, not /
            TempByte = HT_CodeStringStructArray(Temp1).CodeArray(Temp2)
            HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = HT_CodeStringStructArray(Temp1).CodeArray(HT_CodeStringStructArray(Temp1).CodeLength - Temp2 + 1&)
            HT_CodeStringStructArray(Temp1).CodeArray(HT_CodeStringStructArray(Temp1).CodeLength - Temp2 + 1&) = TempByte
        Next Temp2
    Next Temp1
    'end of creating huffman tree struct
End Sub

Public Function HTCS_TreeCodeByteString_Define(ByRef HT_CodeStringStructArray() As HT_CodeStringStruct, ByRef TreeCodeByteStringLength As Long, ByRef TreeCodeByteString() As Byte) As String
    'on error resume next 'returns a byte string containing the data of the created huffman tree, and additional length information
    Dim TreeCodeLengthTotal As Long 'length (in bytes) of all bit codes
    Dim ByteStringBitWritePos As Long
    Dim ByteStringIndex As Long
    Dim Temp1 As Long
    Dim Temp2 As Long
    Dim TempByte As Byte
    'calculate length of final return string
    For Temp1 = 0& To 255&
        TreeCodeLengthTotal = TreeCodeLengthTotal + _
            HT_CodeStringStructArray(Temp1).CodeLength
    Next Temp1
    TreeCodeByteStringLength = 4& + 256& + (-Int(-TreeCodeLengthTotal / 8&))
    ReDim TreeCodeByteString(1 To TreeCodeByteStringLength) As Byte
    'add code string length information
    Call CopyMemory(TreeCodeByteString(1), TreeCodeLengthTotal, 4)
    'add code length information
    For Temp1 = 0& To 255&
        TempByte = CByte(HT_CodeStringStructArray(Temp1).CodeLength)
        Call CopyMemory(ByVal VarPtr(TreeCodeByteString(5 + Temp1)), TempByte, 1)
    Next Temp1
    'add codes itself
    ByteStringBitWritePos = (4& * 8&) + (256& * 8&)
    For Temp1 = 0& To 255&
        '
        'NOTE: ByteStringBitWritePos indendicates the bit in
        'TreeCodeByteString() where the next code string is to be 'added'.
        '
        For Temp2 = 1& To HT_CodeStringStructArray(Temp1).CodeLength
            '
            ByteStringBitWritePos = ByteStringBitWritePos + 1&
            '
            If Not (HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 0) Then
                '
                'NOTE: I must admit that this is the first time I regret that I don't use C++.
                '
                ByteStringIndex = ((ByteStringBitWritePos - 1&) \ 8&) + 1& '\, not /
                'TreeCodeByteString( _
                '    ByteStringIndex) = _
                'TreeCodeByteString( _
                '    ByteStringIndex) _
                '    Or _
                '    HT_CodeStringStructArray(Temp1).CodeArray(Temp2) * _
                '    (2& ^ (7& - CLng(((ByteStringBitWritePos + 7#) Mod 8#))))
                Select Case (ByteStringBitWritePos Mod 8&)
                Case 1& 'propability is the same for every value
                    TreeCodeByteString(ByteStringIndex) = TreeCodeByteString(ByteStringIndex) _
                        Or HT_CodeStringStructArray(Temp1).CodeArray(Temp2) * 128& '2 ^ 7 = 128
                Case 2&
                    TreeCodeByteString(ByteStringIndex) = TreeCodeByteString(ByteStringIndex) _
                        Or HT_CodeStringStructArray(Temp1).CodeArray(Temp2) * 64& '2 ^ 6 = 64
                Case 3&
                    TreeCodeByteString(ByteStringIndex) = TreeCodeByteString(ByteStringIndex) _
                        Or HT_CodeStringStructArray(Temp1).CodeArray(Temp2) * 32& '2 ^ 5 = 32
                Case 4&
                    TreeCodeByteString(ByteStringIndex) = TreeCodeByteString(ByteStringIndex) _
                        Or HT_CodeStringStructArray(Temp1).CodeArray(Temp2) * 16& '2 ^ 4 = 16
                Case 5&
                    TreeCodeByteString(ByteStringIndex) = TreeCodeByteString(ByteStringIndex) _
                        Or HT_CodeStringStructArray(Temp1).CodeArray(Temp2) * 8& '2 ^ 3 = 8
                Case 6&
                    TreeCodeByteString(ByteStringIndex) = TreeCodeByteString(ByteStringIndex) _
                        Or HT_CodeStringStructArray(Temp1).CodeArray(Temp2) * 4& '2 ^ 2 = 4
                Case 7&
                    TreeCodeByteString(ByteStringIndex) = TreeCodeByteString(ByteStringIndex) _
                        Or HT_CodeStringStructArray(Temp1).CodeArray(Temp2) * 2& '2 ^ 1 = 1
                Case 0& 'as we removed '+ 7&' (see original calculation that is commented-out) things get mysterious and we must use 0 and not 8 here
                    TreeCodeByteString(ByteStringIndex) = TreeCodeByteString(ByteStringIndex) _
                        Or HT_CodeStringStructArray(Temp1).CodeArray(Temp2) '* 1& '2 ^ 0 = 1
                End Select
                '
            End If
        Next Temp2
    Next Temp1
    'end of creating code string
End Function

'***************************END OF HUFFMANTREE: COMPRESSION****************************
'******************************HUFFMANTREE: DECOMPRESSION******************************

Public Sub HTDC_HT_TreeStringStructVar_Define(ByRef HT_TreeStringStructVar As HT_TreeStringStruct, ByVal ByteStringLength As Long, ByRef ByteString() As Byte)
    'on error resume next
    If Not (ByteStringLength < 4) Then
        Call CopyMemory(HT_TreeStringStructVar.TreeByteStringBitCount, ByteString(5), 4)
        HT_TreeStringStructVar.TreeByteStringLength = 256 + -Int(-HT_TreeStringStructVar.TreeByteStringBitCount / 8)
        ReDim HT_TreeStringStructVar.TreeByteString(1 To HT_TreeStringStructVar.TreeByteStringLength) As Byte
        Call CopyMemory(HT_TreeStringStructVar.TreeByteString(1), ByteString(9), HT_TreeStringStructVar.TreeByteStringLength)
    Else
        HT_TreeStringStructVar.TreeByteStringLength = 0 'reset (error)
        HT_TreeStringStructVar.TreeByteStringBitCount = 0 'reset (error)
        ReDim HT_TreeStringStructVar.TreeByteString(1 To 1) As Byte 'reset (error)
    End If
End Sub

Public Sub HTDC_HT_CodeStringStruct_Define(ByRef HT_CodeStringStructArray() As HT_CodeStringStruct, ByRef HT_TreeStringStructVar As HT_TreeStringStruct)
    'on error resume next
    Dim ByteStringIndex As Long
    Dim BitReadPos As Long
    Dim Temp1 As Long
    Dim Temp2 As Long
    'begin
    For Temp1 = 0& To 255&
        HT_CodeStringStructArray(Temp1).CodeLength = CLng(HT_TreeStringStructVar.TreeByteString(Temp1 + 1&))
    Next Temp1
    BitReadPos = (256& * 8&)
    With HT_TreeStringStructVar
        For Temp1 = 0& To 255&
            For Temp2 = 1& To HT_CodeStringStructArray(Temp1).CodeLength
                BitReadPos = BitReadPos + 1&
                ByteStringIndex = ((BitReadPos - 1&) \ 8&) + 1& '\, not /
                'If (.TreeByteString(ByteStringIndex) And (2 ^ (7 - ((BitReadPos + 7) Mod 8)))) Then
                '    HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 1
                'Else
                '    HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 0
                'End If
                Select Case BitReadPos Mod 8&
                Case 1&
                    If (.TreeByteString(ByteStringIndex) And 128&) Then
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 1
                    Else
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 0
                    End If
                Case 2&
                    If (.TreeByteString(ByteStringIndex) And 64&) Then
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 1
                    Else
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 0
                    End If
                Case 3&
                    If (.TreeByteString(ByteStringIndex) And 32&) Then
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 1
                    Else
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 0
                    End If
                Case 4&
                    If (.TreeByteString(ByteStringIndex) And 16&) Then
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 1
                    Else
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 0
                    End If
                Case 5&
                    If (.TreeByteString(ByteStringIndex) And 8&) Then
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 1
                    Else
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 0
                    End If
                Case 6&
                    If (.TreeByteString(ByteStringIndex) And 4&) Then
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 1
                    Else
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 0
                    End If
                Case 7&
                    If (.TreeByteString(ByteStringIndex) And 2&) Then
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 1&
                    Else
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 0&
                    End If
                Case 0&
                    If (.TreeByteString(ByteStringIndex) And 1&) Then
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 1
                    Else
                        HT_CodeStringStructArray(Temp1).CodeArray(Temp2) = 0
                    End If
                End Select
            Next Temp2
        Next Temp1
    End With
End Sub

Public Sub HTDC_CodeStringStruct_Define(ByRef HT_CodeStringStructArray() As HT_CodeStringStruct, ByRef HTDC_CodeStringStructArray() As HTDC_CodeStringStruct)
    'on error resume next
    Dim CodeLengthMin As Long
    Dim CodeLengthMayBeZeroFlag As Boolean
    Dim NonZeroLengthCharCodeNumber As Integer
    Dim TempHTDC_CodeStringStruct As HTDC_CodeStringStruct
    Dim Temp1 As Long
    Dim Temp2 As Long
    '
    'NOTE: this sub sorts the items if the tree code struct
    'and transfers them to the decompress struct (tree code struct will stay unchanged).
    '
    For Temp1 = 0& To 255&
        HTDC_CodeStringStructArray(Temp1).Char = CByte(Temp1)
        HTDC_CodeStringStructArray(Temp1).CharCodeArrayLength = HT_CodeStringStructArray(Temp1).CodeLength
        Call CopyMemory(HTDC_CodeStringStructArray(Temp1).CharCodeArray(1), HT_CodeStringStructArray(Temp1).CodeArray(1), 256)
    Next Temp1
    'NOTE: now sort the codes so that the longest is located at 'the beginning'.
    Temp2 = 0& 'preset
ReDo:
    CodeLengthMin = 256& ^ 3& 'reset
    For Temp1 = Temp2 To 255&
        If CodeLengthMayBeZeroFlag = False Then
            If (HTDC_CodeStringStructArray(Temp1).CharCodeArrayLength < CodeLengthMin) And _
                (HTDC_CodeStringStructArray(Temp1).CharCodeArrayLength > 0&) Then
                CodeLengthMin = HTDC_CodeStringStructArray(Temp1).CharCodeArrayLength
            End If
        Else
            If HTDC_CodeStringStructArray(Temp1).CharCodeArrayLength < CodeLengthMin Then
                CodeLengthMin = HTDC_CodeStringStructArray(Temp1).CharCodeArrayLength
            End If
        End If
    Next Temp1
    If CodeLengthMin = 256& ^ 3& Then
        CodeLengthMayBeZeroFlag = True
        GoTo ReDo: 'place all zero-length code strings at end of structure array now
    End If
    If CodeLengthMayBeZeroFlag = False Then 'still searching for non-zero length code strings
        NonZeroLengthCharCodeNumber = NonZeroLengthCharCodeNumber + 1
    End If
    For Temp1 = Temp2 To 255&
        '
        'NOTE: the code strings of the structure will be compared with the char code
        'buffer when decompressing. As the shortest string appears the most frequent
        'it should be located at the beginning of the structure array to avoid senseless
        'looping as far as possible. Code strings with the length 0 must all (!) be located
        'at the end of the structure array.
        '
        If HTDC_CodeStringStructArray(Temp1).CharCodeArrayLength = CodeLengthMin Then
            If Not (Temp1 = Temp2) Then
                TempHTDC_CodeStringStruct = HTDC_CodeStringStructArray(Temp1)
                HTDC_CodeStringStructArray(Temp1) = HTDC_CodeStringStructArray(Temp2)
                HTDC_CodeStringStructArray(Temp2) = TempHTDC_CodeStringStruct
            End If
            Temp2 = Temp2 + 1&
            If Not (Temp2 = 256&) Then
                GoTo ReDo:
            End If
        End If
    Next Temp1
    For Temp1 = 0& To 255&
        HTDC_CodeStringStructArray(1).StartIndexArray(Temp1) = 0&
        HTDC_CodeStringStructArray(1).EndIndexArray(Temp1) = -1&
    Next Temp1
    Temp2 = 0& 'reset
    For Temp1 = 0& To 255&
        If Not (HTDC_CodeStringStructArray(Temp1).CharCodeArrayLength = Temp2) Then
            HTDC_CodeStringStructArray(1).EndIndexArray(Temp2) = (Temp1 - 1&)
            Temp2 = HTDC_CodeStringStructArray(Temp1).CharCodeArrayLength
            HTDC_CodeStringStructArray(1).StartIndexArray(Temp2) = Temp1
        Else
            If Temp1 = 255& Then
                HTDC_CodeStringStructArray(1).EndIndexArray(Temp2) = 255&
            End If
        End If
    Next Temp1
    HTDC_CodeStringStructArray(1).StartIndexArray(0) = 0&
    HTDC_CodeStringStructArray(1).EndIndexArray(0) = -1&
End Sub

'**************************END OF HUFFMANTREE: DECOMPRESSION***************************
