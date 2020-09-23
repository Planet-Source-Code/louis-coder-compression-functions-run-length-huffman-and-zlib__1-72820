Attribute VB_Name = "GFCompression_ZLibmod"
Option Explicit
'(c)2001 by Louis.
'
'NOTE: this module provides two functions for compressing and decompressing
'a string using the ZLib compression ((c) by Jean-loup Gailly & Mark Adler).
'It is the best compression method, but copyrighted and thus not usable in
'commercial programs.
'
'ZLIB_[Compress/Decompress]String
Private Declare Function ZLIBDLL_CompressString Lib "cmprzlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function ZLIBDLL_DecompressString Lib "cmprzlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
'general use
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'*****************************************ZLIB*****************************************

Public Function ZLib_CompressString(ByRef ByteStringLength As Long, ByRef ByteString() As Byte) As Boolean
    'on error resume next 'returns True for success, False for error
    Dim CompressionByteStringLength As Long
    Dim CompressionByteString() As Byte
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    '
    'NOTE: the ZLib compression functions return an error value, but this value is
    'not tested as I'm not sure if that value is constant in all ZLib.dll versions.
    'NOTE: the passed byte string length should be limited to GFCompressionWindowLength.
    '
    'preset
    CompressionByteStringLength = CLng(CSng(ByteStringLength) * 1.1! + 10240!)
    ReDim CompressionByteString(1 To CompressionByteStringLength) As Byte
    'begin
    '
    Call ZLIBDLL_CompressString(CompressionByteString(1), CompressionByteStringLength, ByteString(1), ByteStringLength)
    '
    If Not ((CompressionByteStringLength = 0) And (Not (ByteStringLength = 0))) Then 'verify
        'NOTE: success, copy compressed string to passed string.
        '
        If GFCompressionHeader_Preset(GFCompressionHeaderStructVar) = False Then GoTo Error:
        If GFCompressionHeader_Write(ByteStringLength, ByteString(), CompressionByteStringLength, ByteStringLength) = False Then GoTo Error:
        '
        Call CopyMemory(ByteString(1 + GFCompressionHeaderStructVar.GFCompressionHeaderStructLength), CompressionByteString(1), CompressionByteStringLength)
        ZLib_CompressString = True 'ok
    Else
        GoTo Error:
    End If
    Exit Function
Error:
    'NOTE: error, leave passed byte string unchanged.
    ZLib_CompressString = False 'error
    Exit Function
End Function

Public Function ZLib_DecompressString(ByRef ByteStringLength As Long, ByRef ByteString() As Byte, ByRef BlockLengthProcessed As Long) As Boolean
    'on error resume next 'returns True for success or False for error
    Dim DecompressionByteStringLength As Long
    Dim DecompressionByteString() As Byte
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    'preset
    '
    If GFCompressionHeader_Preset(GFCompressionHeaderStructVar) = False Then GoTo Error:
    If GFCompressionHeader_Read(ByteStringLength, ByteString(), GFCompressionHeaderStructVar.BlockLengthCompressed, GFCompressionHeaderStructVar.BlockLengthOriginal) = False Then GoTo Error:
    If GFCompressionHeader_Remove(ByteStringLength, ByteString(), GFCompressionHeaderStructVar, BlockLengthProcessed) = False Then GoTo Error:
    '
    DecompressionByteStringLength = GFCompressionHeaderStructVar.BlockLengthOriginal
    ReDim DecompressionByteString(1 To DecompressionByteStringLength) As Byte
    '
    'begin
    '
    Call ZLIBDLL_DecompressString(DecompressionByteString(1), DecompressionByteStringLength, ByteString(1), GFCompressionHeaderStructVar.BlockLengthCompressed)
    '
    If Not ((DecompressionByteStringLength = 0) And (Not (ByteStringLength = 0))) Then 'verify
        '
        'NOTE: success, copy decompressed string back to passed string.
        '
        ByteStringLength = DecompressionByteStringLength
        ReDim ByteString(1 To ByteStringLength) As Byte
        '
        Call CopyMemory(ByteString(1), DecompressionByteString(1), DecompressionByteStringLength)
        ZLib_DecompressString = True 'ok
    Else
        GoTo Error:
    End If
    Exit Function
Error:
    '
    'NOTE: error, leave passed byte string unchanged.
    '
    ByteStringLength = 0 'processed bytes
    ZLib_DecompressString = False 'error
    Exit Function
End Function

'***FAST***
'NOTE: the calling sub/function MUST size the target array passed to the compression functions.

Public Function ZLib_CompressStringFast(ByRef CompressedByteStringStartPos As Long, ByRef CompressedByteStringLength As Long, ByRef CompressedByteString() As Byte, ByRef ByteStringStartPos As Long, ByRef ByteStringLength As Long, ByRef ByteString() As Byte) As Boolean
    'on error resume next 'returns True for success, False for error
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    '
    'NOTE: the ZLib compression functions return an error value, but this value is
    'not tested as I'm not sure if that value is constant in all ZLib.dll versions.
    'NOTE: the passed byte string length should be limited to GFCompressionWindowLength.
    'NOTE: CompressedByteStringStartPos can be set by the calling function.
    'NOTE: CompressedByteStringLength must be set to the total length of CompressedByteString(),
    'when this function is left then CompressedByteStringLength is set to the size of the added
    'compressed data.
    '
    'preset
    CompressedByteStringStartPos = CompressedByteStringStartPos + Len(GFCompressionHeaderStructVar) 'create space for later header adding
    'begin
    '
    Call ZLIBDLL_CompressString(CompressedByteString(CompressedByteStringStartPos), CompressedByteStringLength, ByteString(ByteStringStartPos), ByteStringLength)
    '
    If Not ((CompressedByteStringLength = 0&) And (Not (ByteStringLength = 0&))) Then 'verify
        'NOTE: success, copy compressed string to passed string.
        '
        If GFCompressionHeader_PresetFast(GFCompressionHeaderStructVar) = False Then GoTo Error:
        'NOTE: there must be some space in front of the original byte string data to make this work:
        If GFCompressionHeader_WriteFast(CompressedByteStringStartPos, CompressedByteStringLength, CompressedByteString(), CompressedByteStringLength, ByteStringLength) = False Then GoTo Error:
        '
        ZLib_CompressStringFast = True 'ok
    Else
        GoTo Error:
    End If
    Exit Function
Error:
    'NOTE: error, leave passed byte string unchanged.
    ZLib_CompressStringFast = False 'error
    Exit Function
End Function

Public Function ZLib_DecompressStringFast(ByRef DecompressedByteStringStartPos As Long, ByRef DecompressedByteStringLength As Long, ByRef DecompressedByteString() As Byte, ByVal ByteStringStartPos As Long, ByRef ByteStringLength As Long, ByRef ByteString() As Byte, ByRef BlockLengthProcessed As Long) As Boolean
    'on error resume next 'returns True for success or False for error
    Dim GFCompressionHeaderStructVar As GFCompressionHeaderStruct
    '
    'NOTE: ByteStringStartPos is increased when removing header, but length of processed
    'block includes header so don't return the increased value (pass ByVal).
    '
    'preset
    '
    If GFCompressionHeader_PresetFast(GFCompressionHeaderStructVar) = False Then GoTo Error:
    If GFCompressionHeader_ReadFast(ByteStringStartPos, ByteStringLength, ByteString(), GFCompressionHeaderStructVar.BlockLengthCompressed, GFCompressionHeaderStructVar.BlockLengthOriginal) = False Then GoTo Error:
    If GFCompressionHeader_RemoveFast(ByteStringStartPos, ByteStringLength, ByteString(), GFCompressionHeaderStructVar, BlockLengthProcessed) = False Then GoTo Error:
    '
    'begin
    '
    Call ZLIBDLL_DecompressString(DecompressedByteString(DecompressedByteStringStartPos), DecompressedByteStringLength, ByteString(ByteStringStartPos), GFCompressionHeaderStructVar.BlockLengthCompressed)
    '
    If Not ((DecompressedByteStringLength = 0&) And (Not (ByteStringLength = 0&))) Then 'verify
        ZLib_DecompressStringFast = True 'ok
    Else
        GoTo Error:
    End If
    Exit Function
Error:
    '
    'NOTE: error, leave passed byte string unchanged.
    '
    ByteStringLength = 0& 'processed bytes
    ZLib_DecompressStringFast = False 'error
    Exit Function
End Function

'***END OF FAST***
'*************************************END OF ZLIB**************************************

