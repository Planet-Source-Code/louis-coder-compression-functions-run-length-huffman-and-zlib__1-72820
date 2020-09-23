Attribute VB_Name = "GFCompressionmod"
Option Explicit
'(c)2001, 2002 by Louis. Self-made compression functions.
'
'NOTE: the CallBackForm must contain the following sub:
'Public Sub GFCompression_CallBackSub(ByVal ProcedureName As String, _
'    ByVal FileNumberCurrent As Integer, ByVal FileNumberTotal As Integer, _
'    ByVal FileName As String, ByVal BytesProcessed As Long, ByVal BytesTotal As Long, _
'    ByRef CancelFlag As Boolean)
'    'on error resume next
'    '
'    'NOTE: if ProcedureName is GFCompression_CompressFile() or
'    'GFCompression_DecompressFile() then BytesProcessed and BytesTotal are
'    'valid. The call back sub is called after the whole file or the maximal possible
'    'block length has been processed.
'    'If ProcedureName is GFCompression_CompressionPack_Create() or
'    'GFCompression_CompressionPack_Unpack() then FileNumberCurrent and
'    'FileNumberTotal are valid. The call back sub is called before a file is packed
'    'or unpacked. FileName is always valid, it contains a full path to the file that is
'    'currently compressed or decompressed, packed or unpacked. The call back sub
'    'is also called if a file has been processed (ProcedureName + ": ok") or if there
'    'has been an error processing the current file (ProcedureName + ": error").
'    '
'End Sub
'
'NOTE: about this project:
'The GFCompression project consists of several modules,
'every module contains code for one type of compression only.
'Every type has one function called [c. method]_CompressString() and
'one called [c. method]_DecompressString().
'These functions may check if the fast VC compression is avaiable and
'call [c. method]_[De]compressString_VC() or, if the VC dll is not available
'call [c. method]_[De]comrpessString_VB().
'
'Annotations should explain how the compression algorithm works
'and for which data it works the best (highest compresion ratio).
'
'NOTE: speed optimize (important, partially tested):
'-within a time intensive sub/function always tell the compiler what type a
' 'fixed' constant should have (1&, 2!, 3# etc.)
'-try to use coherent variable types within calculations
'-when manipulating bits don't use 2 ^ but a Select Case statement (much faster)
'-when manipulating bits try to use Long values only (CopyMemory() needs hardly time)
'-Int() is as fast as making VB do any rounding operation (use Int())
'-first use Int() on a double var, then CLng() as Int() returns a double var, too
'
'NOTE: 'Fast' procedures:
'When a procedure name has the suffix 'Fast' (e.g. ZLib_CompressStringFast()) then
'the function is optimized for speed, that means ALL unnecessary CopyMemory()
'operations are avoided and no larger VB strings are used.
'The Fast-function must handle ByteStringStartPos parameters, which store the
'start pos of a byte string in the ByteString()-array. If e.g. the header of a compressed
'string is to be removed then the start pos may be greater than 1. All other Fast-
'functions must then pass ByteString(ByteStringStartPos) to API functions.
'The Fast functions are to be created out of copies of the original functions and
'are located in the same code domain like the original functions.
'The Fast functions do have one source- and one target byte string, there's no
'unnecessary back-copying from source to source.
'The 'Fast' functions create a 'ByVal or ByRef-mess', so change passing method if
'necessary (if you found a bug).
'
'NOTE: string operations:
'CopyMemory(ByVal s1, ByVal (StrPtr(s2) + pos), length)
'fails, it must be pos * 2, but this also doesn't work, there are alternating
'the chars of s2 and Chr$(0) in s1 :(
'CopyMemory(ByVal (StrPtr(s1)), ByVal (StrPtr(s2) + pos), length) works :))
'But then pos must still be doubled.
'Seems as if the usage of StrPtr leads to errors, ByteStrings can't be copied
'into a string like this, the string would be filled with '?' (at least when being
'displayed in VB through QuickInfo).
'Don't use StrPtr to copy strings into byte arrays and vice versa. Copy
'the to-copied part of the string to a temporary string and then do it the old
'way without StrPtr.
'
'NOTE: file numbers:
'File numbers are created by FreeFile(), pay attention that all
'file numbers are generated right before using Open, do not store
'several file numbers at the beginning of a sub/function.
'
'NOTE: external files necessary:
'-cmprss10.dll should be located in %winsysdir% to speed up
' RLE and Huffman compression (cmprss10.dll not finished yet)
'-cmprzlib.dll is necessary for ZLib compression
'
'IMPORTANT: the compression help dlls must have names that
'are not used by other programs!
'
'IsVCCompressionAvailable
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'general use
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
'other
'
Private Const GFCompressionWindowLength As Long = 1024000 'how many bytes are compressed at the same time
'
'NOTE: do not use a window larger than 1024000 bytes as the compression was not tested therefore.
'In a LZ77_RepeatStringBuffer_Create with 2 Mb window length VB crashed after calling CopyMemory()
'and then Redim Preserve, the reason therefore is completely unknown (VB sucks).
'Futhermore the length of a string passed to a Huffman compression function must not
'exceed (2147483647 / 8) minus some more bytes.
'
Private Const GFCompressionWindowLengthExtension As Long = 10240
'
'NOTE: if a file cannot really be compressed (e.g. an mp3 file) then the final compressed
'string can exceed the length of the read uncompressed string (as headers are added).
'That's why the string read out of the compressed file must be larger than the one read
'out of the uncompressed file (when decompressing then the read string has the length
'GFCompressionWindowLength + GFCompressionWindowLengthExtension).
'GFCompressionHeader_Remove() will shorten the read compressed string using
'the data stored in the compression header, so it doesn't matter when the read
'compressed string is longer than required.
'
'Version
Const Version = "v1.0"
'GFCompressionControlStruct - general information
Private Type GFCompressionControlStruct
    CallBackFormEnabledFlag As Boolean
    CallBackForm As Object
End Type
Dim GFCompressionControlStructVar As GFCompressionControlStruct
'GFCompressionHeaderStruct - stores the compressed and the compressed block size
Public Type GFCompressionHeaderStruct
    GFCompressionHeaderString As String * 20
    GFCompressionHeaderStructLength As Long
    BlockLengthCompressed As Long
    BlockLengthOriginal As Long
End Type

'********************************COMPRESSION INTERFACE*********************************
'NOTE: the following functions can be used by any project to compress a file.
'When a file is decompressed, the method is determined automatically as it has been
'written to the first 20 chars of the file to decompress by the compression function.

Public Sub GFCompression_CallBackForm_Enable(ByRef CallBackForm As Object)
    'on error resume next
    GFCompressionControlStructVar.CallBackFormEnabledFlag = True
    Set GFCompressionControlStructVar.CallBackForm = CallBackForm
End Sub

Public Sub GFCompression_CallBackForm_GetInfo(ByRef CallBackFormEnabledFlag As Boolean, ByRef CallBackForm As Object)
    'on error resume next
    CallBackFormEnabledFlag = GFCompressionControlStructVar.CallBackFormEnabledFlag
    Set CallBackForm = GFCompressionControlStructVar.CallBackForm
End Sub

Public Sub GFCompression_CallBackForm_Disable()
    'on error resume next
    GFCompressionControlStructVar.CallBackFormEnabledFlag = False
    Set GFCompressionControlStructVar.CallBackForm = Nothing
End Sub

Public Function GFCompression_CompressFile(ByVal CompressionName As String, ByVal CompressionMethodName As String, ByVal TempFileReturnEnabledFlag As Boolean, ByRef TempFileReturned As String) As Boolean
    On Error GoTo Error: 'important (if memory low); returns True for success, False for error
    Dim CompressionNameFileNumber As Integer
    Dim ByteStringLength As Long
    Dim ByteString() As Byte
    Dim BlockReadPos As Long
    Dim BlockLength As Long
    Dim BlockLengthMax As Long 'length of file to compress
    Dim BlockString As String
    Dim BlockLoop As Integer
    Dim TempFile As String
    Dim TempFileNumber As Integer
    '
    'NOTE: set TempFileReturnEnabledFlag to True to avoid that the file to compress is changed in any way.
    'The target project can then receive the name of a temp file that contains the compressed data of the input file.
    '
    'verify
    If (Dir$(CompressionName, vbNormal Or vbHidden Or vbSystem Or vbArchive) = "") Or (Right$(CompressionName, 1) = "\") Or (CompressionName = "") Then 'verify (some target project require to compress also hidden files)
        MsgBox "internal error in GFCompression_CompressFile(): file " + CompressionName + " not found !", vbOKOnly + vbExclamation
        GoTo Error:
    End If
    'preset
    TempFile = GenerateTempFileName(GetDirectoryName(CompressionName))
    TempFileNumber = FreeFile(0)
    Open TempFile For Output As #TempFileNumber 'create file and print header string
        Print #TempFileNumber, GetFileHeaderString(FileLen(CompressionName), CompressionName, CompressionMethodName);
    Close #TempFileNumber
    BlockReadPos = 1 'preset
    'begin
    CompressionNameFileNumber = FreeFile(0)
    Open CompressionName For Binary As #CompressionNameFileNumber
        BlockLengthMax = LOF(CompressionNameFileNumber)
        Do
            'read ByteString()
            BlockLength = GFCompressionWindowLength 'preset
            If (BlockLength + BlockReadPos - 1) > BlockLengthMax Then
                BlockLength = (BlockLengthMax - BlockReadPos + 1)
            End If
            If BlockLength = 0 Then Exit Do 'verify
            BlockString = String(BlockLength, Chr$(0))
            Get #CompressionNameFileNumber, BlockReadPos, BlockString
            ByteStringLength = BlockLength
            ReDim ByteString(1 To ByteStringLength) As Byte
            Call CopyMemory(ByteString(1), ByVal BlockString, BlockLength)
            'compress ByteString()
            Select Case LCase$(CompressionMethodName)
            Case "huffman"
                If Huffman_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
            Case "rle"
                If RLE_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
            Case "rle huffman", "huffman rle"
                If RLE_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
                If Huffman_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
            Case "lz77"
                If LZ77_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
            Case "lz77 rle huffman"
                If RLE_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
                If LZ77_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
                If Huffman_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
            Case "zlib"
                If ZLib_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
            Case Else
                Close #CompressionNameFileNumber 'important
                If Not ((Dir$(TempFile) = "") Or (Right$(TempFile, 1) = "\") Or (TempFile = "")) Then Kill TempFile 'make sure file is deleted
                GoTo Jump: 'leave file unchanged
            End Select
            'write ByteString()
            BlockString = String$(ByteStringLength, Chr$(0))
            Call CopyMemory(ByVal BlockString, ByteString(1), ByteStringLength)
            TempFileNumber = FreeFile(0)
            Open TempFile For Append As #TempFileNumber
                Print #TempFileNumber, BlockString;
            Close #TempFileNumber
            BlockReadPos = BlockReadPos + BlockLength
            BlockLoop = BlockLoop + 1
            'call the call-back sub
            If GFCompressionControlStructVar.CallBackFormEnabledFlag = True Then
                Call GFCompressionControlStructVar.CallBackForm.GFCompression_CallBackSub( _
                    "GFCompression_CompressFile()", 1, 1, CompressionName, BlockReadPos, BlockLengthMax, 0)
            End If
        Loop Until (BlockLoop = 32767) 'avoid endless loop
    Close #CompressionNameFileNumber
    If TempFileReturnEnabledFlag = False Then
        If CopyFile(TempFile, CompressionName, 0) = 0 Then
            MsgBox "internal error in GFCompression_CompressFile(): FileCopy() failed !", vbOKOnly + vbExclamation
            'continue
        End If
        If Not ((Dir$(TempFile) = "") Or (Right$(TempFile, 1) = "\") Or (TempFile = "")) Then Kill TempFile 'make sure temp file is deleted
    Else
        TempFileReturned = TempFile
    End If
Jump: 'jump here if passed compression method unknown
    GFCompression_CompressFile = True 'ok
    Exit Function
Error:
    Close #CompressionNameFileNumber 'make sure files are closed
    Close #TempFileNumber 'make sure files are closed
    If Not ((Dir$(TempFile) = "") Or (Right$(TempFile, 1) = "\") Or (TempFile = "")) Then Kill TempFile 'make sure temp file is deleted
    GFCompression_CompressFile = False 'error
    Exit Function
End Function

Private Function GetFileHeaderString(ByVal InputNameSize As Long, ByVal InputName As String, ByVal CompressionMethodName As String) As String
    'on error resume next 'returns a string that contains data about a compressed file's content
    GetFileHeaderString = _
        CompressionMethodName + String$(20 - Len(CompressionMethodName), Chr$(0)) + String$(5, Chr$(0)) 'Chr$(0) marks the end of method name, 5 chars are reserved
End Function

'NOTE: a compressed file's string may be decompressed as compressed string
'and vice versa (but warning - not tested!) (does not work with Fast functions).

Public Function GFCompression_DecompressFile(ByVal DecompressionName As String, ByVal TempFileReturnEnabledFlag As Boolean, ByRef TempFileReturned As String) As Boolean
    On Error GoTo Error: 'important (if memory low); returns True for success, False for error
    Dim DecompressionNameFileNumber As Integer
    Dim CompressionMethodName As String
    Dim ByteStringLength As Long
    Dim ByteString() As Byte
    Dim BlockReadPos As Long
    Dim BlockLength As Long
    Dim BlockLengthMax As Long 'length of file to decompress
    Dim BlockString As String
    Dim BlockLoop As Integer
    Dim TempFile As String
    Dim TempFileNumber As Integer
    Dim BlockLengthProcessed1 As Long
    Dim BlockLengthProcessed2 As Long
    Dim BlockLengthProcessed3 As Long
    Dim Tempstr$
    '
    'NOTE: set TempFileReturnEnabledFlag to True to avoid that the file to decompress is changed in any way.
    'The target project can then receive the name of a temp file that contains the decompressed data of the input file.
    '
    'verify
    If (Dir$(DecompressionName) = "") Or (Right$(DecompressionName, 1) = "\") Or (DecompressionName = "") Then 'Verify
        MsgBox "internal error in GFCompression_DecompressFile(): file " + DecompressionName + " not found !", vbOKOnly + vbExclamation
        GoTo Error:
    End If
    'begin
    DecompressionNameFileNumber = FreeFile(0)
    Open DecompressionName For Binary As #DecompressionNameFileNumber
        Tempstr$ = String$(20, Chr$(0))
        Get #DecompressionNameFileNumber, 1, Tempstr$
    Close #DecompressionNameFileNumber
    If Not (InStr(1, Tempstr$, Chr$(0), vbBinaryCompare) = 0) Then 'verify
        CompressionMethodName = Left$(Tempstr$, InStr(1, Tempstr$, Chr$(0), vbBinaryCompare) - 1)
    Else
        CompressionMethodName = Tempstr$
    End If
    TempFile = GenerateTempFileName(GetDirectoryName(DecompressionName))
    TempFileNumber = FreeFile(0)
    Open TempFile For Output As #TempFileNumber 'just create file
    Close #TempFileNumber
    'begin
    BlockReadPos = 26 'preset
    DecompressionNameFileNumber = FreeFile(0)
    Open DecompressionName For Binary As #DecompressionNameFileNumber
        BlockLengthMax = LOF(DecompressionNameFileNumber)
        Do
            'read ByteString()
            BlockLength = MAX(GFCompressionWindowLength + GFCompressionWindowLengthExtension, 1024) 'preset (reserve a minimum of space for header data)
            If (BlockLength + BlockReadPos - 1) > BlockLengthMax Then
                BlockLength = (BlockLengthMax - BlockReadPos + 1)
            End If
            If BlockLength < 1 Then Exit Do
            BlockString = String$(BlockLength, Chr$(0))
            Get #DecompressionNameFileNumber, BlockReadPos, BlockString
            ByteStringLength = BlockLength
            ReDim ByteString(1 To ByteStringLength) As Byte
            Call CopyMemory(ByteString(1), ByVal BlockString, BlockLength)
            'decompress ByteString()
            Select Case LCase$(CompressionMethodName)
            Case "huffman"
                If Huffman_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
                BlockLength = BlockLengthProcessed1
            Case "rle"
                If RLE_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
                BlockLength = BlockLengthProcessed1
            Case "rle huffman", "huffman rle"
                If Huffman_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
                If RLE_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed2) = False Then GoTo Error:
                BlockLength = BlockLengthProcessed1 '+ BlockLengthProcessed2 'sum up all processed block lengths 'no!
                '
                'NOTE: tests showed that for some reason we must merely add the length
                'of the first processed block, this length is equal to ByteStringLength
                'after using both compressions huffman and rle on the original input string.
                '
            Case "lz77"
                If LZ77_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
                BlockLength = BlockLengthProcessed1
            Case "lz77 rle huffman"
                If Huffman_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
                If LZ77_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed3) = False Then GoTo Error:
                If RLE_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed2) = False Then GoTo Error:
                BlockLength = BlockLengthProcessed1 + BlockLengthProcessed2 + BlockLengthProcessed3 'sum up all processed block lengths
                '
                'NOTE: the line above is probably wrong, but it hasn't been tested
                'as the LZ77 compression doesn't work anyway.
                '
            Case "zlib"
                If ZLib_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
                BlockLength = BlockLengthProcessed1
            Case Else
                Close #DecompressionNameFileNumber 'important
                If Not ((Dir$(TempFile) = "") Or (Right$(TempFile, 1) = "\") Or (TempFile = "")) Then Kill TempFile 'make sure file is deleted
                GoTo Jump: 'leave file unchanged
            End Select
            'write ByteString()
            BlockString = String$(ByteStringLength, Chr$(0))
            Call CopyMemory(ByVal BlockString, ByteString(1), ByteStringLength)
            TempFileNumber = FreeFile(0)
            Open TempFile For Append As #TempFileNumber
                Print #TempFileNumber, BlockString;
            Close #TempFileNumber
            BlockReadPos = BlockReadPos + BlockLength 'BlockLength was 'indirectly' set in GFCompressionHeader_Remove()
            BlockLoop = BlockLoop + 1
            'call the call-back sub
            If GFCompressionControlStructVar.CallBackFormEnabledFlag = True Then
                Call GFCompressionControlStructVar.CallBackForm.GFCompression_CallBackSub( _
                    "GFCompression_DecompressFile()", 1, 1, DecompressionName, BlockReadPos, BlockLengthMax, 0)
            End If
        Loop Until (BlockLoop = 32767) 'avoid endless loop
    Close #DecompressionNameFileNumber
    If TempFileReturnEnabledFlag = False Then
        If CopyFile(TempFile, DecompressionName, 0) = 0 Then
            MsgBox "internal error in GFCompression_DecompressFile(): FileCopy() failed !", vbOKOnly + vbExclamation
            'continue
        End If
        If Not ((Dir$(TempFile) = "") Or (Right$(TempFile, 1) = "\") Or (TempFile = "")) Then Kill TempFile 'make sure file is deleted
    Else
        TempFileReturned = TempFile
    End If
Jump: 'jump here if compression method unknown
    GFCompression_DecompressFile = True 'ok
    Exit Function
Error:
    Close #DecompressionNameFileNumber 'make sure files are closed
    Close #TempFileNumber 'make sure files are closed
    If Not ((Dir$(TempFile) = "") Or (Right$(TempFile, 1) = "\") Or (TempFile = "")) Then Kill TempFile 'make sure temp file is deleted
    GFCompression_DecompressFile = False 'error
    Exit Function
End Function

Public Sub GFCompression_DeleteTempFile(ByVal TempFile As String)
    'on error resume next 'to be called by target project
    If Not ((Dir$(TempFile) = "") Or (Right$(TempFile, 1) = "\") Or (TempFile = "")) Then Kill TempFile 'make sure file is deleted
End Sub

'NOTE: the following two string-compression functions are not that good as
'they copy around much memory. Don't use frequently on large strings.

Public Function GFCompression_CompressString(ByRef CompressionString As String, ByVal CompressionMethodName As String) As Boolean
    On Error GoTo Error: 'compresses string, if larger than GFCompressionWindowLength bytes then the string will be split up into blocks
    Dim ByteString() As Byte
    Dim ByteStringLength As Long
    Dim CompressedString As String
    Dim CompressedStringLength As Long 'how many chars are in use
    Dim DecompressionStringLength As Long
    Dim BlockStartPos As Long
    Dim Tempstr$
    'preset
    'allocate string memory at once to avoid memory moving
    CompressedString = String$(CLng(CSng(Len(CompressionString)) * 1.1! + 10240!), Chr$(0)) 'add some space for headers, huffman tables etc. (although the compressed string should be smaller than the passed one)
    '
    'NOTE: the following size header is optional, but it avoids that
    'GFCompression_DecompressString() moves around memory as
    'strings must be joined.
    '
    DecompressionStringLength = Len(CompressionString)
    Tempstr$ = String$(4, Chr$(0))
    Call CopyMemory(ByVal Tempstr$, DecompressionStringLength, 4)
    Mid$(CompressedString, 1, 25) = "DECOMPRESSEDSTRINGLENGTH="
    Mid$(CompressedString, 26, 4) = Tempstr$
    Tempstr$ = GetFileHeaderString(Len(CompressionString), "COMPRESSEDSTRING", CompressionMethodName)
    Mid$(CompressedString, 30, Len(Tempstr$)) = Tempstr$
    CompressedStringLength = 29 + Len(Tempstr$) 'string is large enough, we can 'waste' some chars
    '
    'begin
    For BlockStartPos = 1 To Len(CompressionString) Step GFCompressionWindowLength 'create approx. 1MB blocks
        '
        ByteStringLength = GFCompressionWindowLength
        If (BlockStartPos + ByteStringLength) > Len(CompressionString) Then
            ByteStringLength = (Len(CompressionString) - BlockStartPos + 1&)
        End If
        If (ByteStringLength = 0&) Or (BlockStartPos > Len(CompressionString)) Then Exit For 'verify
        '
        ReDim ByteString(1 To ByteStringLength) As Byte
        Tempstr$ = Mid$(CompressionString, BlockStartPos, ByteStringLength)
        Call CopyMemory(ByteString(1), ByVal Tempstr$, Len(Tempstr$))
        '
        Select Case LCase$(CompressionMethodName)
        Case "huffman"
            If Huffman_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
        Case "rle"
            If RLE_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
        Case "rle huffman", "huffman rle"
            If RLE_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
            If Huffman_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
        Case "lz77"
            If LZ77_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
        Case "lz77 rle huffman"
            If RLE_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
            If LZ77_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
            If Huffman_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
        Case "zlib"
            If ZLib_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
        Case Else
            GoTo Jump: 'leave file unchanged
        End Select
        '
        If (CompressedStringLength + ByteStringLength) > Len(CompressedString) Then 'should not happen
            Debug.Print "warning in GFCompression_CompressString(): large string resized !"
            CompressedString = CompressedString + String$(ByteStringLength, Chr$(0))
        End If
        '
        Tempstr$ = String$(ByteStringLength, Chr$(0))
        Call CopyMemory(ByVal Tempstr$, ByteString(1), ByteStringLength)
        Mid$(CompressedString, CompressedStringLength + 1&, ByteStringLength) = Tempstr$
        '
        CompressedStringLength = CompressedStringLength + ByteStringLength
        '
    Next BlockStartPos
Jump: 'jump to here if compression method unknown
    CompressionString = String$(CompressedStringLength, Chr$(0)) 'change string size
    Call CopyMemory(ByVal CompressionString, ByVal CompressedString, CompressedStringLength)
    GFCompression_CompressString = True 'ok
    Exit Function
Error:
    GFCompression_CompressString = False 'error
    Exit Function
End Function

Public Function GFCompression_DecompressString(ByRef DecompressionString As String) As Boolean
    On Error GoTo Error: 'returns True for success or False for error
    Dim DecompressedString As String
    Dim DecompressedStringLength As Long
    Dim CompressionMethodName As String
    Dim ByteStringLength As Long
    Dim ByteString() As Byte
    Dim BlockReadPos As Long
    Dim BlockLength As Long
    Dim BlockLengthMax As Long 'length of file to decompress
    Dim BlockString As String
    Dim BlockLoop As Integer
    Dim BlockLengthProcessed1 As Long
    Dim BlockLengthProcessed2 As Long
    Dim BlockLengthProcessed3 As Long
    Dim Tempstr$
    'preset
    If Mid$(DecompressionString, 1, 25) = "DECOMPRESSEDSTRINGLENGTH=" Then
        Tempstr$ = Mid$(DecompressionString, 26, 4)
        Call CopyMemory(DecompressedStringLength, ByVal Tempstr$, 4)
        DecompressedString = String$(DecompressedStringLength, Chr$(0))
        BlockReadPos = 30& 'preset
    Else
        DecompressedStringLength = Len(DecompressionString) * 2&
        DecompressedString = String$(DecompressedStringLength, Chr$(0))
        BlockReadPos = 1& 'preset
    End If
    DecompressedStringLength = 0 'reset (further, different use, real string length)
    'get compression method name
    Tempstr$ = String$(25, Chr$(0))
    '
    Call CopyMemory(ByVal (StrPtr(Tempstr$)), ByVal (StrPtr(DecompressionString) + (BlockReadPos - 1&) * 2&), 25& * 2&)
    '
    If Not (InStr(1, Tempstr$, Chr$(0), vbBinaryCompare) = 0) Then 'verify
        CompressionMethodName = Left$(Tempstr$, InStr(1, Tempstr$, Chr$(0), vbBinaryCompare) - 1)
    Else
        CompressionMethodName = Tempstr$
    End If
    BlockReadPos = BlockReadPos + 25& 'jump over compression method name
    'begin
    BlockLengthMax = Len(DecompressionString) 'all stuff copied and transfered (file->string) from GFCompression_DecompressFile()
    Do
        'read ByteString()
        BlockLength = MAX(GFCompressionWindowLength + GFCompressionWindowLengthExtension, 1024) 'preset (reserve a minimum of space for header data)
        If (BlockLength + BlockReadPos - 1) > BlockLengthMax Then
            BlockLength = (BlockLengthMax - BlockReadPos + 1)
        End If
        If BlockLength < 1& Then Exit Do 'verify
        BlockString = String$(BlockLength, Chr$(0))
        '
        ByteStringLength = BlockLength
        ReDim ByteString(1 To ByteStringLength) As Byte
        '
        Tempstr$ = Mid$(DecompressionString, BlockReadPos, BlockLength)
        Call CopyMemory(ByteString(1), ByVal Tempstr$, BlockLength)
        '
        'decompress ByteString()
        Select Case LCase$(CompressionMethodName)
        Case "huffman"
            If Huffman_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
            BlockLength = BlockLengthProcessed1
        Case "rle"
            If RLE_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
            BlockLength = BlockLengthProcessed1
        Case "rle huffman", "huffman rle"
            If Huffman_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
            If RLE_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed2) = False Then GoTo Error:
            BlockLength = BlockLengthProcessed1 '+ BlockLengthProcessed2 'sum up all processed block lengths 'no!
            '
            'NOTE: tests showed that for some reason we must merely add the length
            'of the first processed block, this length is equal to ByteStringLength
            'after using both compressions huffman and rle on the original input string.
            '
        Case "lz77"
            If LZ77_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
            BlockLength = BlockLengthProcessed1
        Case "lz77 rle huffman"
            If Huffman_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
            If LZ77_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed3) = False Then GoTo Error:
            If RLE_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed2) = False Then GoTo Error:
            BlockLength = BlockLengthProcessed1 + BlockLengthProcessed2 + BlockLengthProcessed3 'sum up all processed block lengths
            '
            'NOTE: the line above is probably wrong, but it hasn't been tested
            'as the LZ77 compression doesn't work anyway.
            '
        Case "zlib"
            If ZLib_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
            BlockLength = BlockLengthProcessed1
        Case Else
            GoTo Jump: 'leave file unchanged
        End Select
        'write ByteString()
        DecompressedStringLength = (DecompressedStringLength + ByteStringLength)
        If DecompressedStringLength > Len(DecompressedString) Then 'could happen of not size header added (if original string compressed by GFCompression_CompressFile())
            Debug.Print "warning in GFCompression_DecompressString(): large string append !"
            DecompressedString = DecompressedString + String$(Len(DecompressedString) - DecompressedStringLength, Chr$(0))
        End If
        'EITHER
        Tempstr$ = String$(ByteStringLength, Chr$(0))
        Call CopyMemory(ByVal Tempstr$, ByteString(1), ByteStringLength)
        Mid$(DecompressedString, DecompressedStringLength - ByteStringLength + 1&, ByteStringLength) = Tempstr$
        '
        BlockReadPos = BlockReadPos + BlockLength 'BlockLength was 'indirectly' set in GFCompressionHeader_Remove()
        BlockLoop = BlockLoop + 1
        '
        'call the call-back sub
        'If GFCompressionControlStructVar.CallBackFormEnabledFlag = True Then
        '    Call GFCompressionControlStructVar.CallBackForm.GFCompression_CallBackSub( _
        '        "GFCompression_DecompressString()", 1, 1, DecompressionName, BlockReadPos, BlockLengthMax, 0)
        'End If
    Loop Until (BlockLoop = 32767) 'avoid endless loop
Jump:
    'cut end of string if no size header used
    If Not (Len(DecompressedString) = DecompressedStringLength) Then
        Debug.Print "warning in GFCompression_DecompressString(): large string truncated !"
        DecompressedString = Left$(DecompressedString, DecompressedStringLength)
    End If
    'end of cutting string
    DecompressionString = String$(DecompressedStringLength, Chr$(0))
    Call CopyMemory(ByVal DecompressionString, ByVal DecompressedString, DecompressedStringLength)
    GFCompression_DecompressString = True 'ok
    Exit Function
Error:
    GFCompression_DecompressString = False 'error
    Exit Function
End Function

'***FAST***
'NOTE: the following functions compress/decompress (parts of) BYTE strings.

Public Function GFCompression_CompressStringFast(ByVal ByteStringCompressedStartPos As Long, ByRef ByteStringCompressedLength As Long, ByRef ByteStringCompressed() As Byte, ByVal ByteStringStartPos As Long, ByVal ByteStringLength As Long, ByRef ByteString() As Byte, ByVal CompressionMethodName As String) As Boolean
    On Error GoTo Error: 'compresses string, if larger than GFCompressionWindowLength bytes then the string will be split up into blocks
    Dim ByteStringCompressedLengthAdded As Long 'how much compressed data was added to the compressed string (compressed string start pos also manipulated by lower-level compression functions)
    Dim DecompressionStringLength As Long
    Dim BlockStartPos As Long
    Dim BlockLength As Long
    Dim Tempstr$
    'verify
    If ByteStringLength < 1 Then
        GFCompression_CompressStringFast = True 'ok
        Exit Function
    End If
    'preset
    'allocate string memory at once to avoid memory moving
    ByteStringCompressedLength = CLng(CSng(ByteStringLength) * 1.1! + 10240!) '***TEMP*** (not absolutely save, if compressed string larger than uncompressed one then crash)
    ReDim ByteStringCompressed(1 To ByteStringCompressedLength) As Byte
    '
    'NOTE: the following size header is optional, but it avoids that
    'GFCompression_DecompressString() moves around memory as
    'strings must be joined.
    '
    Tempstr$ = "DECOMPRESSEDSTRINGLENGTH="
    Call CopyMemory(ByteStringCompressed(1), ByVal Tempstr$, 25)
    DecompressionStringLength = ByteStringLength
    Call CopyMemory(ByteStringCompressed(26), DecompressionStringLength, 4)
    Tempstr$ = GetFileHeaderString(ByteStringLength, "COMPRESSEDSTRING", CompressionMethodName)
    Call CopyMemory(ByteStringCompressed(30), ByVal Tempstr$, Len(Tempstr$)) 'length should be 25
    ByteStringCompressedStartPos = 30 + Len(Tempstr$) 'string is large enough, we can 'waste' some chars; here we have START POS, in non-Fast function we have string LENGTH (30/29)!
    '
    'begin
    For BlockStartPos = 1 To ByteStringLength Step GFCompressionWindowLength 'create approx. 1MB blocks
        '
        BlockLength = GFCompressionWindowLength
        If (BlockStartPos + BlockLength) > ByteStringLength Then
            BlockLength = (ByteStringLength - BlockStartPos + 1&)
        End If
        If (BlockLength = 0&) Or (BlockStartPos > ByteStringLength) Then Exit For 'verify
        '
        Select Case LCase$(CompressionMethodName)
'        Case "huffman"
'            If Huffman_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
'        Case "rle"
'            If RLE_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
'        Case "rle huffman", "huffman rle"
'            If RLE_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
'            If Huffman_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
'        Case "lz77"
'            If LZ77_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
'        Case "lz77 rle huffman"
'            If RLE_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
'            If LZ77_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
'            If Huffman_CompressString(ByteStringLength, ByteString()) = False Then GoTo Error:
        Case "zlib"
            ByteStringCompressedLengthAdded = ByteStringCompressedLength 'length information passed to ZLib function is used and then set to the size of the added compressed data
            If ZLib_CompressStringFast(ByteStringCompressedStartPos, ByteStringCompressedLengthAdded, ByteStringCompressed(), (BlockStartPos + ByteStringStartPos - 1&), BlockLength, ByteString()) = False Then GoTo Error:
            '
            'NOTE: the ZLib_CompressStringFast() function has increased the start pos, it now points to the new compressed
            'data start, furthermore ByteStringCompressedLengthAdded bytes of compressed data have been added to the
            'compressed string.
            '
            ByteStringCompressedStartPos = ByteStringCompressedStartPos + ByteStringCompressedLengthAdded
        Case Else
            GoTo Jump: 'leave file unchanged
        End Select
        '
    Next BlockStartPos
    'shrinken compressed return-string
    ByteStringCompressedLength = (ByteStringCompressedStartPos - 1&)
    ReDim Preserve ByteStringCompressed(1 To ByteStringCompressedLength) As Byte 'SHRINKEN byte string (probably faster than enlarging)
    'finished, leave
Jump: 'jump to here if compression method unknown
    GFCompression_CompressStringFast = True 'ok
    Exit Function
Error:
    GFCompression_CompressStringFast = False 'error
    Exit Function
End Function

Public Function GFCompression_DecompressStringFast(ByVal ByteStringDecompressedStartPos As Long, ByRef ByteStringDecompressedLength As Long, ByRef ByteStringDecompressed() As Byte, ByVal ByteStringStartPos As Long, ByVal ByteStringLength As Long, ByRef ByteString() As Byte) As Boolean
    On Error GoTo Error: 'returns True for success or False for error
    Dim ByteStringDecompressedLengthAdded As Long
    Dim CompressionMethodName As String
    Dim BlockReadPos As Long
    Dim BlockLength As Long
    Dim BlockLengthMax As Long 'length of file to decompress
    Dim BlockLoop As Integer
    Dim BlockLengthProcessed1 As Long
    Dim BlockLengthProcessed2 As Long
    Dim BlockLengthProcessed3 As Long
    Dim Tempstr$
    'preset
    BlockReadPos = ByteStringStartPos 'mostly 1, or higher value
    Tempstr$ = String$(25, Chr$(0))
    Call CopyMemory(ByVal Tempstr$, ByteString(ByteStringStartPos), 25)
    If Tempstr$ = "DECOMPRESSEDSTRINGLENGTH=" Then
        Call CopyMemory(ByteStringDecompressedLength, ByteString(ByteStringStartPos + 25), 4)
        ReDim ByteStringDecompressed(1 To (ByteStringDecompressedStartPos + ByteStringDecompressedLength - 1&)) As Byte 'the callig sub/function may want to add additional data in front of decompressed string (leave space free there)
        BlockReadPos = BlockReadPos + 29& 'preset (read next char here)
    Else 'too dangerous or/and slow, we won't have a string without size information anyway
        MsgBox "internal error in GFCompression_DecompressStringFast(): uncompressed string size information missing !", vbOKOnly + vbExclamation
        GoTo Error:
    End If
    'get compression method name
    Tempstr$ = String$(25, Chr$(0))
    Call CopyMemory(ByVal Tempstr$, ByteString(BlockReadPos), 25)
    If Not (InStr(1, Tempstr$, Chr$(0), vbBinaryCompare) = 0) Then 'verify
        CompressionMethodName = Left$(Tempstr$, InStr(1, Tempstr$, Chr$(0), vbBinaryCompare) - 1)
    Else
        CompressionMethodName = Tempstr$
    End If
    BlockReadPos = BlockReadPos + 25& 'jump over compression method name
    'begin
    BlockLengthMax = ByteStringLength 'all stuff copied and transfered (file->string) from GFCompression_DecompressFile()
    Do
        '
        'read ByteString()
        BlockLength = MAX(GFCompressionWindowLength + GFCompressionWindowLengthExtension, 1024) 'preset (reserve a minimum of space for header data)
        If (BlockLength + BlockReadPos - 1) > BlockLengthMax Then
            BlockLength = (BlockLengthMax - BlockReadPos + 1)
        End If
        If BlockLength < 1& Then Exit Do 'verify
        '
        'decompress ByteString()
        Select Case LCase$(CompressionMethodName)
'        Case "huffman"
'            If Huffman_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
'            BlockLength = BlockLengthProcessed1
'        Case "rle"
'            If RLE_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
'            BlockLength = BlockLengthProcessed1
'        Case "rle huffman", "huffman rle"
'            If Huffman_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
'            If RLE_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed2) = False Then GoTo Error:
'            BlockLength = BlockLengthProcessed1 '+ BlockLengthProcessed2 'sum up all processed block lengths 'no!
'            '
'            'NOTE: tests showed that for some reason we must merely add the length
'            'of the first processed block, this length is equal to ByteStringLength
'            'after using both compressions huffman and rle on the original input string.
'            '
'        Case "lz77"
'            If LZ77_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
'            BlockLength = BlockLengthProcessed1
'        Case "lz77 rle huffman"
'            If Huffman_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
'            If LZ77_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed3) = False Then GoTo Error:
'            If RLE_DecompressString(ByteStringLength, ByteString(), BlockLengthProcessed2) = False Then GoTo Error:
'            BlockLength = BlockLengthProcessed1 + BlockLengthProcessed2 + BlockLengthProcessed3 'sum up all processed block lengths
'            '
'            'NOTE: the line above is probably wrong, but it hasn't been tested
'            'as the LZ77 compression doesn't work anyway.
'            '
        Case "zlib"
            ByteStringDecompressedLengthAdded = ByteStringDecompressedLength 'length information passed to ZLib dll
            If ZLib_DecompressStringFast(ByteStringDecompressedStartPos, ByteStringDecompressedLengthAdded, ByteStringDecompressed(), BlockReadPos, BlockLength, ByteString(), BlockLengthProcessed1) = False Then GoTo Error:
            ByteStringDecompressedStartPos = ByteStringDecompressedStartPos + ByteStringDecompressedLengthAdded
            BlockLength = BlockLengthProcessed1
        Case Else
            GoTo Jump: 'leave file unchanged
        End Select
        'write ByteString()
        '
        'ByteStringDecompressedLength was already set (size information in header)
        BlockReadPos = BlockReadPos + BlockLength 'BlockLength was 'indirectly' set in GFCompressionHeader_Remove()
        BlockLoop = BlockLoop + 1
        '
        'call the call-back sub
        'If GFCompressionControlStructVar.CallBackFormEnabledFlag = True Then
        '    Call GFCompressionControlStructVar.CallBackForm.GFCompression_CallBackSub( _
        '        "GFCompression_DecompressString()", 1, 1, DecompressionName, BlockReadPos, BlockLengthMax, 0)
        'End If
    Loop Until (BlockLoop = 32767) 'avoid endless loop
Jump:
    GFCompression_DecompressStringFast = True 'ok
    Exit Function
Error:
    GFCompression_DecompressStringFast = False 'error
    Exit Function
End Function

'***END OF FAST***
'****************************END OF COMPRESSION INTERFACE******************************
'***********************************COMPRESSIONPACK************************************
'NOTE: a compression pack is a file that contains several compressed files.
'The compression pack is created by the code of this module, and also
'unpacking is done by the GFCompression code.
'The string in the CompressionPackFile has the following format:
'
'GFCOMPRESSIONPACKFILE1.0ffrrlllls[*llll]nnnnC:\Skins\BaseSkin\Skin.datdddd[...]nnn...
'ff: number of files in packet
'rr: number of strings
'llll: string length
's: string data
'nnnn: file name length
'dddd: file data length
'
'The file name is stored as complete path to the included file.
'The identification string at the begin of the file must be exactly 25 chars long.

Public Function GFCompression_CompressionPack_Create(ByVal CompressionPackFile As String, ByVal FileNumber As Integer, ByRef FileArray() As String, ByVal CompressionMethodName As String, ByVal StringNumber As Integer, ByRef StringArray() As String) As Boolean
    On Error GoTo Error: 'important (if a file is locked); returns True for success, False for error
    Dim CompressionPackFileNumber As Integer
    Dim BlockReadPos As Long
    Dim BlockLength As Long
    Dim BlockString As String
    Dim TempFile As String
    Dim TempFileNumber As Integer
    Dim FileLoop As Integer
    Dim Temp As Long
    Dim Tempstr$
    '
    'NOTE: call this function to create a compression pack file out of the files in the passed file array.
    'If CompressionMethodName is "", no compression will be used.
    'It is possible to save additional strings in the CompressionPackFile that can be requested
    'seperately to allow the target project to e.g. determinate the output directory before unpacking.
    '
    'begin
    CompressionPackFileNumber = FreeFile(0)
    Open CompressionPackFile For Output As #CompressionPackFileNumber 'create file and write header
        'write identification string
        Print #CompressionPackFileNumber, "GFCOMPRESSIONPACKFILE" + Left$(Version, 4);
        'write total file number
        Tempstr$ = String$(2, Chr$(0))
        Call CopyMemory(ByVal Tempstr$, FileNumber, 2)
        Print #CompressionPackFileNumber, Tempstr$; 'print total number of files
        'write additional strings
        Tempstr$ = String$(2, Chr$(0))
        Call CopyMemory(ByVal Tempstr$, StringNumber, 2)
        Print #CompressionPackFileNumber, Tempstr$; 'print number of strings
        For FileLoop = 1 To StringNumber
            Tempstr$ = String$(4, Chr$(0))
            Temp = Len(StringArray(FileLoop))
            Call CopyMemory(ByVal Tempstr$, Temp, 4)
            Print #CompressionPackFileNumber, Tempstr$; 'print number of strings
            Print #CompressionPackFileNumber, StringArray(FileLoop);
        Next FileLoop
        'write file data
        For FileLoop = 1 To FileNumber
            'call the call-back sub
            If GFCompressionControlStructVar.CallBackFormEnabledFlag = True Then
                Call GFCompressionControlStructVar.CallBackForm.GFCompression_CallBackSub( _
                    "GFCompression_CompressionPack_Create()", FileLoop, FileNumber, FileArray(FileLoop), 0, 0, 0)
            End If
            'verify file
            If (Dir$(FileArray(FileLoop)) = "") Or (Right$(FileArray(FileLoop), 1) = "\") Or (FileArray(FileLoop) = "") Then 'verify
                MsgBox "internal error in GFCompression_CompressionPack_Create(): file '" + FileArray(FileLoop) + "' not found !", vbOKOnly + vbExclamation
                GoTo Jump:
            End If
            'compress file
            If GFCompression_CompressFile(FileArray(FileLoop), CompressionMethodName, True, TempFile) = False Then GoTo Error:
            'write file information
            Tempstr$ = String$(4, Chr$(0))
            Temp = Len(FileArray(FileLoop))
            Call CopyMemory(ByVal Tempstr$, Temp, 4) 'copy file name length
            Print #CompressionPackFileNumber, Tempstr$;
            Print #CompressionPackFileNumber, FileArray(FileLoop);
            'write file
            BlockReadPos = 1 'preset
            TempFileNumber = FreeFile(0)
            Open TempFile For Binary As #TempFileNumber
                Tempstr$ = String$(4, Chr$(0))
                Temp = LOF(TempFileNumber)
                Call CopyMemory(ByVal Tempstr$, Temp, 4) 'copy file data length
                Print #CompressionPackFileNumber, Tempstr$;
                Do 'read file in blocks to save memory
                    BlockLength = 512000 'preset
                    If (BlockReadPos + BlockLength - 1) > LOF(TempFileNumber) Then 'verify
                        BlockLength = LOF(TempFileNumber) - BlockReadPos + 1
                    End If
                    If BlockLength = 0 Then Exit Do 'verify
                    BlockString = String(BlockLength, Chr$(0))
                    Get #TempFileNumber, BlockReadPos, BlockString
                    Print #CompressionPackFileNumber, BlockString;
                    BlockReadPos = BlockReadPos + BlockLength
                Loop
            Close #TempFileNumber
            '
            'NOTE: the temp file was created by GFCompression_CompressFile(),
            'its content was transferred to CompressionPackFile.
            '
            If Not ((Dir$(TempFile) = "") Or (Right$(TempFile, 1) = "\") Or (TempFile = "")) Then Kill TempFile 'make sure temp file is deleted
Jump: 'jump here if a file to include was not found
        Next FileLoop
    Close #CompressionPackFileNumber
    If GFCompressionControlStructVar.CallBackFormEnabledFlag = True Then
        Call GFCompressionControlStructVar.CallBackForm.GFCompression_CallBackSub( _
            "GFCompression_CompressionPack_Create(): ok", 0, 0, "", 0, 0, 0)
    End If
    GFCompression_CompressionPack_Create = True 'ok
    Exit Function
Error:
    Close #CompressionPackFileNumber 'make sure file is closed
    Close #TempFileNumber 'make sure file is closed
    If Not ((Dir$(TempFile) = "") Or (Right$(TempFile, 1) = "\") Or (TempFile = "")) Then Kill TempFile 'make sure temp file is deleted
    If GFCompressionControlStructVar.CallBackFormEnabledFlag = True Then
        Call GFCompressionControlStructVar.CallBackForm.GFCompression_CallBackSub( _
            "GFCompression_CompressionPack_Create(): error", 0, 0, "", 0, 0, 0)
    End If
    GFCompression_CompressionPack_Create = False 'error
    Exit Function
End Function

Public Function GFCompression_CompressionPack_GetStringArray(ByVal CompressionPackFile As String, ByRef StringNumber As Integer, ByRef StringArray() As String) As Boolean
    'on error resume next 'returns True for success, False for error
    Dim CompressionPackFileNumber As Integer
    Dim StringLength As Long
    Dim StringLoop As Integer
    Dim Tempstr$
    '
    'NOTE: this function initializes the passed string array with the strings that
    'are to be included within a CompressionPackFile.
    'Note that StringNumber may be 0, but if the function returns True
    'everything's alright though.
    '
    'verify
    If (Dir$(CompressionPackFile) = "") Or (Right$(CompressionPackFile, 1) = "\") Or (CompressionPackFile = "") Then 'verify
        MsgBox "internal error in GFCompression_CompressionPack_GetStrignArray(): file '" + CompressionPackFile + "' not found !", vbOKOnly + vbExclamation
        GoTo Error:
    End If
    'begin
    CompressionPackFileNumber = FreeFile(0)
    Open CompressionPackFile For Binary As #CompressionPackFileNumber
        'read identification string
        Tempstr$ = String$(25, Chr$(0)) 'identification string must have the length 25
        Get #CompressionPackFileNumber, 1, Tempstr$
        If Not (Left$(Tempstr$, Len("GFCOMPRESSIONPACKFILE")) = "GFCOMPRESSIONPACKFILE") Then
            MsgBox "internal error in GFCompression_CompressionPack_GetStringArray(): file '" + CompressionPackFile + "' has an invalid format !", vbOKOnly + vbExclamation
            GoTo Error:
        End If
        'skip over total file number
        Tempstr$ = String$(2, Chr$(0))
        Get #CompressionPackFileNumber, , Tempstr$
        'read strings
        Tempstr$ = String$(2, Chr$(0))
        Get #CompressionPackFileNumber, , Tempstr$
        Call CopyMemory(StringNumber, ByVal Tempstr$, 2)
        If Not (StringNumber = 0) Then
            ReDim StringArray(1 To StringNumber) As String
        Else
            ReDim StringArray(1 To 1) As String 'reset
        End If
        For StringLoop = 1 To StringNumber
            Tempstr$ = String$(4, Chr$(0))
            Get #CompressionPackFileNumber, , Tempstr$
            Call CopyMemory(StringLength, ByVal Tempstr$, 4)
            StringArray(StringLoop) = String$(StringLength, Chr$(0))
            Get #CompressionPackFileNumber, , StringArray(StringLoop)
        Next StringLoop
    Close #CompressionPackFileNumber
    GFCompression_CompressionPack_GetStringArray = True 'ok
    Exit Function
Error:
    Close #CompressionPackFileNumber 'make sure file is closed
    GFCompression_CompressionPack_GetStringArray = False 'error
    Exit Function
End Function

Public Function GFCompression_CompressionPack_Unpack(ByVal CompressionPackFile As String, ByVal OutputDirectory As String, Optional ByVal UnpackName As String = "") As Boolean
    On Error GoTo Error: 'important (if a file is locked); returns True for success, False for error
    Dim CompressionPackFileNumber As Integer
    Dim OutputNameLength As Long
    Dim OutputName As String
    Dim OutputNameNumber As Integer 'running index of file to unpack
    Dim OutputNameNumberTotal As Integer 'total number of files in packet
    Dim OutputNameFileNumber As Integer
    Dim BlockLengthTotal As Long
    Dim BlockLength As Long
    Dim BlockString As String
    Dim StringLength As Long
    Dim StringNumber As Integer 'will be skipped
    Dim StringArray() As String 'will be skipped
    Dim StringLoop As Integer
    Dim Tempstr$
    '
    'NOTE: the OutputDirectory must already exist, this function will not create it
    '(you can use GFCreateDirectory() to do so).
    '
    'NOTE: if UnpackName is not "", this function will unpack this file ONLY.
    'If UnpackName is not found, the function will not return error but the calling sub
    'should check if UnpackName has been created.
    '
    'verify
    If (Dir$(CompressionPackFile) = "") Or (Right$(CompressionPackFile, 1) = "\") Or (CompressionPackFile = "") Then 'verify
        MsgBox "internal error in GFCompression_CompressionPack_Unpack(): file '" + CompressionPackFile + "' not found !", vbOKOnly + vbExclamation
        GoTo Error:
    End If
    If Len(OutputDirectory) = 0 Then GoTo Error:
    If Not (Right$(OutputDirectory, 1) = "\") Then OutputDirectory = OutputDirectory + "\" 'verify
    'begin
    CompressionPackFileNumber = FreeFile(0)
    Open CompressionPackFile For Binary As #CompressionPackFileNumber
        'read identification string
        Tempstr$ = String$(25, Chr$(0)) 'identification string must have the length 25
        Get #CompressionPackFileNumber, 1, Tempstr$
        If Not (Left$(Tempstr$, Len("GFCOMPRESSIONPACKFILE")) = "GFCOMPRESSIONPACKFILE") Then
            MsgBox "internal error in GFCompression_CompressionPack_Unpack(): file '" + CompressionPackFile + "' has an invalid format !", vbOKOnly + vbExclamation
            GoTo Error:
        End If
        'read number of files in packet
        Tempstr$ = String$(2, Chr$(0))
        Get #CompressionPackFileNumber, , Tempstr$
        Call CopyMemory(OutputNameNumberTotal, ByVal Tempstr$, 2)
        'skip strings
        Tempstr$ = String$(2, Chr$(0))
        Get #CompressionPackFileNumber, , Tempstr$
        Call CopyMemory(StringNumber, ByVal Tempstr$, 2)
        If Not (StringNumber = 0) Then
            ReDim StringArray(1 To StringNumber) As String
        Else
            ReDim StringArray(1 To 1) As String 'reset
        End If
        For StringLoop = 1 To StringNumber
            Tempstr$ = String$(4, Chr$(0))
            Get #CompressionPackFileNumber, , Tempstr$
            Call CopyMemory(StringLength, ByVal Tempstr$, 4)
            StringArray(StringLoop) = String$(StringLength, Chr$(0))
            Get #CompressionPackFileNumber, , StringArray(StringLoop)
        Next StringLoop
        'begin unpacking
        Do
            OutputNameNumber = OutputNameNumber + 1 'index of current file
            'read OutputName
            If (EOF(CompressionPackFileNumber) Or (Seek(CompressionPackFileNumber) > LOF(CompressionPackFileNumber))) Then Exit Do
            Tempstr$ = String$(4, Chr$(0))
            Get #CompressionPackFileNumber, , Tempstr$
            Call CopyMemory(OutputNameLength, ByVal Tempstr$, 4)
            OutputName = String$(OutputNameLength, Chr$(0))
            Get #CompressionPackFileNumber, , OutputName
            'NOTE: create final OutputName, note that the name (path) saved in file is that of the original file
            OutputName = OutputDirectory + GetFileName(OutputName)
            'call the call-back sub
            If GFCompressionControlStructVar.CallBackFormEnabledFlag = True Then
                Call GFCompressionControlStructVar.CallBackForm.GFCompression_CallBackSub( _
                    "GFCompression_CompressionPack_Unpack()", OutputNameNumber, OutputNameNumberTotal, OutputName, 0, 0, 0)
            End If
            'read block length
            Tempstr$ = String$(4, Chr$(0))
            Get #CompressionPackFileNumber, , Tempstr$
            Call CopyMemory(BlockLengthTotal, ByVal Tempstr$, 4)
            '
            'NOTE: BlockLengthTotal is decreased by the amount of bytes that are
            'read with each block, if BlockLengthTotal is 0 the reading process is finished.
            '
            If ((UCase$(GetFileName(OutputName)) = UCase$(UnpackName)) Or (UnpackName = "")) And _
                (Not ((Right$(OutputName, 1) = "\") Or (OutputName = ""))) Then 'verify, too (important)
                'transfer file data from compression pack file to OutputName
                OutputNameFileNumber = FreeFile(0)
                Open OutputName For Output As OutputNameFileNumber
                    Do 'copy file in blocks to save memory
                        BlockLength = 512000 'preset
                        If (BlockLengthTotal - BlockLength) < 1 Then 'verify
                            BlockLength = BlockLengthTotal
                        End If
                        If BlockLength = 0 Then Exit Do 'verify
                        BlockString = String$(BlockLength, Chr$(0))
                        Get #CompressionPackFileNumber, , BlockString
                        Print #OutputNameFileNumber, BlockString;
                        BlockLengthTotal = BlockLengthTotal - BlockLength
                    Loop
                Close OutputNameFileNumber
                'decompress OutputName
                If GFCompression_DecompressFile(OutputName, False, Tempstr$) = False Then GoTo Error:
            Else
                'skip current file
                Seek (CompressionPackFileNumber), (Seek(CompressionPackFileNumber) + BlockLengthTotal + 1) 'add one (important, if BlockLengthTotal is 0)
            End If
        Loop
    Close #CompressionPackFileNumber
    If GFCompressionControlStructVar.CallBackFormEnabledFlag = True Then
        Call GFCompressionControlStructVar.CallBackForm.GFCompression_CallBackSub( _
            "GFCompression_CompressionPack_Unpack(): ok", 0, 0, "", 0, 0, 0)
    End If
    GFCompression_CompressionPack_Unpack = True 'ok
    Exit Function
Error:
    Close #CompressionPackFileNumber 'make sure file is closed
    Close #OutputNameFileNumber 'make sure file is closed
    If GFCompressionControlStructVar.CallBackFormEnabledFlag = True Then
        Call GFCompressionControlStructVar.CallBackForm.GFCompression_CallBackSub( _
            "GFCompression_CompressionPack_Unpack(): error", 0, 0, "", 0, 0, 0)
    End If
    GFCompression_CompressionPack_Unpack = False 'error
    Exit Function
End Function

'*******************************END OF COMPRESSIONPACK*********************************
'*********************************GFCOMPRESSIONHEADER**********************************
'NOTE: always write the GFCompressionHeader after compressing a string and remove
'it before decompressing a string.
'The GFComrpessionHeader is necessary to determinate where a window in a
'compressed (!) string ends (NOT after GFCompressionWindowLength bytes).
'
'NOTE: the values of BlockLength[Compressed/Original] do NOT include the
'GFCompressionHeader size.

Public Function GFCompressionHeader_Preset(ByRef GFCompressionHeaderStructVar As GFCompressionHeaderStruct) As Boolean
    'on error resume next 'returns always True
    GFCompressionHeaderStructVar.GFCompressionHeaderStructLength = Len(GFCompressionHeaderStructVar)
    GFCompressionHeader_Preset = True 'ok
    Exit Function
End Function

Public Function GFCompressionHeader_Write(ByRef ByteStringLength As Long, ByRef ByteString() As Byte, ByVal BlockLengthCompressed As Long, ByVal BlockLengthOriginal As Long) As Boolean
    'on error resume next 'returns True for success or False for error; see annotations at top of GFCompressionHeader code
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    'preset
    '
    GFCompressionHeaderStructVar.GFCompressionHeaderStructLength = Len(GFCompressionHeaderStructVar)
    GFCompressionHeaderStructVar.GFCompressionHeaderString = "GFCOMPRESSIONHEADER "
    GFCompressionHeaderStructVar.BlockLengthCompressed = BlockLengthCompressed
    GFCompressionHeaderStructVar.BlockLengthOriginal = BlockLengthOriginal
    '
    ByteStringLength = GFCompressionHeaderStructVar.GFCompressionHeaderStructLength + BlockLengthCompressed
    ReDim Preserve ByteString(1 To ByteStringLength) As Byte
    Call CopyMemory(ByteString(1 + GFCompressionHeaderStructVar.GFCompressionHeaderStructLength), ByteString(1), BlockLengthCompressed) 'move already existing content 'rightwards'
    '
    'begin
    If Not (ByteStringLength < GFCompressionHeaderStructVar.GFCompressionHeaderStructLength) Then
        Call CopyMemory(ByteString(1), GFCompressionHeaderStructVar, GFCompressionHeaderStructVar.GFCompressionHeaderStructLength)
        GFCompressionHeader_Write = True 'ok
    Else
        GFCompressionHeader_Write = False 'error
    End If
End Function

Public Function GFCompressionHeader_Read(ByVal ByteStringLength As Long, ByRef ByteString() As Byte, ByRef BlockLengthCompressed As Long, ByRef BlockLengthOriginal As Long) As Boolean
    'on error resume next 'returns True for success or False for error; see annotations at top of GFCompressionHeader code
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    'preset
    GFCompressionHeaderStructVar.GFCompressionHeaderStructLength = Len(GFCompressionHeaderStructVar)
    'begin
    If Not (ByteStringLength < GFCompressionHeaderStructVar.GFCompressionHeaderStructLength) Then
        Call CopyMemory(GFCompressionHeaderStructVar, ByteString(1), GFCompressionHeaderStructVar.GFCompressionHeaderStructLength)
        If GFCompressionHeaderStructVar.GFCompressionHeaderString = "GFCOMPRESSIONHEADER " Then 'verify
            BlockLengthCompressed = GFCompressionHeaderStructVar.BlockLengthCompressed
            BlockLengthOriginal = GFCompressionHeaderStructVar.BlockLengthOriginal
            GFCompressionHeader_Read = True 'ok
        Else
            MsgBox "internal error in GFCompressionHeader_Read(): wrong compression header string:" + Chr$(10) + "'" + GFCompressionHeaderStructVar.GFCompressionHeaderString + "' !", vbOKOnly + vbExclamation
            GoTo Error:
        End If
    Else
        GoTo Error:
    End If
    Exit Function
Error:
    BlockLengthCompressed = 0 'reset
    BlockLengthOriginal = 0 'reset
    GFCompressionHeader_Read = False 'error
    Exit Function
End Function

Public Function GFCompressionHeader_Remove(ByRef ByteStringLength As Long, ByRef ByteString() As Byte, ByRef GFCompressionHeaderStructVar As GFCompressionHeaderStruct, ByRef BlockLengthProcessed As Long) As Boolean
    'on error resume next 'returns always True
    '
    BlockLengthProcessed = (GFCompressionHeaderStructVar.BlockLengthCompressed + GFCompressionHeaderStructVar.GFCompressionHeaderStructLength)
    Call CopyMemory(ByteString(1), ByteString(1 + GFCompressionHeaderStructVar.GFCompressionHeaderStructLength), GFCompressionHeaderStructVar.BlockLengthCompressed) 'we only need the original block
    ByteStringLength = GFCompressionHeaderStructVar.BlockLengthCompressed
    ReDim Preserve ByteString(1 To ByteStringLength) As Byte
    '
    GFCompressionHeader_Remove = True 'ok
    Exit Function
End Function

'***FAST***

Public Function GFCompressionHeader_PresetFast(ByRef GFCompressionHeaderStructVar As GFCompressionHeaderStruct) As Boolean
    'on error resume next
    GFCompressionHeader_PresetFast = GFCompressionHeader_Preset(GFCompressionHeaderStructVar)
End Function

Public Function GFCompressionHeader_WriteFast(ByVal ByteStringStartPos As Long, ByRef ByteStringLength As Long, ByRef ByteString() As Byte, ByVal BlockLengthCompressed As Long, ByVal BlockLengthOriginal As Long) As Boolean
    'on error resume next 'returns True for success or False for error; see annotations at top of GFCompressionHeader code
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    'preset
    '
    GFCompressionHeaderStructVar.GFCompressionHeaderStructLength = Len(GFCompressionHeaderStructVar)
    GFCompressionHeaderStructVar.GFCompressionHeaderString = "GFCOMPRESSIONHEADER "
    GFCompressionHeaderStructVar.BlockLengthCompressed = BlockLengthCompressed
    GFCompressionHeaderStructVar.BlockLengthOriginal = BlockLengthOriginal
    '
    'begin
    If ByteStringStartPos < GFCompressionHeaderStructVar.GFCompressionHeaderStructLength Then 'there must be space at the beginning of the byte string
        GFCompressionHeader_WriteFast = False 'error
    Else
        ByteStringStartPos = ByteStringStartPos - GFCompressionHeaderStructVar.GFCompressionHeaderStructLength 'do NOT return changed start pos, function-internal use only
        'ByteStringLength = ByteStringLength + GFCompressionHeaderStructVar.GFCompressionHeaderStructLength 'no! callig sub/function must already have added header space
        Call CopyMemory(ByteString(ByteStringStartPos), GFCompressionHeaderStructVar, GFCompressionHeaderStructVar.GFCompressionHeaderStructLength)
        GFCompressionHeader_WriteFast = True 'ok
    End If
End Function

Public Function GFCompressionHeader_ReadFast(ByVal ByteStringStartPos As Long, ByVal ByteStringLength As Long, ByRef ByteString() As Byte, ByRef BlockLengthCompressed As Long, ByRef BlockLengthOriginal As Long) As Boolean
    'on error resume next 'returns True for success or False for error; see annotations at top of GFCompressionHeader code
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    'preset
    GFCompressionHeaderStructVar.GFCompressionHeaderStructLength = Len(GFCompressionHeaderStructVar)
    'begin
    If Not (ByteStringLength < GFCompressionHeaderStructVar.GFCompressionHeaderStructLength) Then
        Call CopyMemory(GFCompressionHeaderStructVar, ByteString(ByteStringStartPos), GFCompressionHeaderStructVar.GFCompressionHeaderStructLength)
        If GFCompressionHeaderStructVar.GFCompressionHeaderString = "GFCOMPRESSIONHEADER " Then 'verify
            BlockLengthCompressed = GFCompressionHeaderStructVar.BlockLengthCompressed
            BlockLengthOriginal = GFCompressionHeaderStructVar.BlockLengthOriginal
            GFCompressionHeader_ReadFast = True 'ok
        Else
            MsgBox "internal error in GFCompressionHeader_ReadFast(): wrong compression header string:" + Chr$(10) + "'" + GFCompressionHeaderStructVar.GFCompressionHeaderString + "' !", vbOKOnly + vbExclamation
            GoTo Error:
        End If
    Else
        GoTo Error:
    End If
    Exit Function
Error:
    BlockLengthCompressed = 0 'reset
    BlockLengthOriginal = 0 'reset
    GFCompressionHeader_ReadFast = False 'error
    Exit Function
End Function

Public Function GFCompressionHeader_RemoveFast(ByRef ByteStringStartPos As Long, ByRef ByteStringLength As Long, ByRef ByteString() As Byte, ByRef GFCompressionHeaderStructVar As GFCompressionHeaderStruct, ByRef BlockLengthProcessed As Long) As Boolean
    'on error resume next 'returns always True
    '
    BlockLengthProcessed = (GFCompressionHeaderStructVar.BlockLengthCompressed + GFCompressionHeaderStructVar.GFCompressionHeaderStructLength)
    ByteStringStartPos = ByteStringStartPos + GFCompressionHeaderStructVar.GFCompressionHeaderStructLength
    'ByteStringLength = ByteStringLength - GFCompressionHeaderStructVar.GFCompressionHeaderStructLength 'no! calling sub/function must pay attention to removed header's size
    '
    GFCompressionHeader_RemoveFast = True 'ok
    Exit Function
End Function

'***END OF FAST***
'*****************************END OF GFCOMPRESSIONHEADER*******************************
'****************************************OTHER*****************************************

Public Function IsVCCompressionAvailable() As Boolean
    'on error Resume Next 'returns True if the fast VC compression can be used, False if not
    Dim WinSysDir As String
    WinSysDir = String$(260, Chr(0)) 'MAX_PATH
    Call GetSystemDirectory(WinSysDir, 260)
    If Not (InStr(1, WinSysDir, Chr(0), vbBinaryCompare)) = 0 Then 'verify
        WinSysDir = Left$(WinSysDir, InStr(1, WinSysDir, Chr(0), vbBinaryCompare) - 1)
    End If
    If Not (Right$(WinSysDir, 1) = "\") Then WinSysDir = WinSysDir + "\" 'verify
    If Not (Dir$(WinSysDir + "cmprss10.dll") = "") Then
        IsVCCompressionAvailable = True
    Else
        IsVCCompressionAvailable = False
    End If
End Function

Public Sub DEBUG_DISPLAYBYTESTRING(ByRef ByteString() As Byte, ByVal DisplayStartPos As Long, ByRef DisplayEndPos As Long)
    'on error resume next 'very useful, use it!
    Dim Tempstr$
    'begin
    Tempstr$ = String$(DisplayEndPos - DisplayStartPos + 1, Chr$(0))
    Call CopyMemory(ByVal Tempstr$, ByteString(1), MIN(Len(Tempstr$), UBound(ByteString())))
    Debug.Print Tempstr$
End Sub

'************************************END OF OTHER**************************************
'**********************************GENERAL FUNCTIONS***********************************

Private Function GetFileName(ByVal GetFileNameName As String) As String
    'on error Resume Next 'returns chars after last backslash or nothing
    Dim GetFileNameLoop As Integer
    GetFileName = "" 'reset
    For GetFileNameLoop = Len(GetFileNameName) To 1 Step (-1)
        If Mid$(GetFileNameName, GetFileNameLoop, 1) = "\" Then
            GetFileName = Right$(GetFileNameName, Len(GetFileNameName) - GetFileNameLoop)
            Exit For
        End If
    Next GetFileNameLoop
End Function

Private Function GetDirectoryName(ByVal GetDirectoryNameName As String) As String
    'on error Resume Next 'returns chars from string begin to (including) last backslash or nothing
    Dim GetDirectoryNameLoop As Integer
    GetDirectoryName = "" 'reset
    For GetDirectoryNameLoop = Len(GetDirectoryNameName) To 1 Step (-1)
        If Mid$(GetDirectoryNameName, GetDirectoryNameLoop, 1) = "\" Then
            GetDirectoryName = Left$(GetDirectoryNameName, GetDirectoryNameLoop)
            Exit For
        End If
    Next GetDirectoryNameLoop
End Function

Private Function GenerateTempFileName(ByVal TempFilePath As String) As String 'copied from NN99 (06-13-2001)
    'on error Resume Next 'returns name of a non-existing file in TempFilePath, file name has following format: ########.tmp
    Dim GenerateTempFileLoop As Integer
    If (Not (Right$(TempFilePath, 1) = "\")) And (Not (TempFilePath = "")) Then 'verify
        TempFilePath = TempFilePath + "\"
    End If
    Do
        GenerateTempFileName = TempFilePath + Format$((Rnd(1) * 1E+08!), "00000000") + ".tmp"
        GenerateTempFileLoop = GenerateTempFileLoop + 1 'save is save
    Loop Until (Dir$(GenerateTempFileName) = "") Or (GenerateTempFileLoop = 32767)
End Function

Private Function MIN(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    'on error Resume Next 'use for i.e. CopyMemory(a(1), ByVal b, MIN(UBound(a()), Len(b))
    If Value1 < Value2 Then
        MIN = Value1
    Else
        MIN = Value2
    End If
End Function

Private Function MAX(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    'on error Resume Next 'use for e.g. CopyMemory(a(1), ByVal b, MIN(UBound(a()), Len(b))
    If Value1 > Value2 Then
        MAX = Value1
    Else
        MAX = Value2
    End If
End Function

'***END OF MODULE***

