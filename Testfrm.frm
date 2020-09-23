VERSION 5.00
Begin VB.Form Testfrm 
   Caption         =   "GFCompression Test"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame2 
      Caption         =   "Test on file selected in list"
      Height          =   3255
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "Compress and Decompress selected file (Test)"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   4455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "create compression pack"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2340
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "unpack compression pack"
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   2340
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Compress and Decompress 'till no doubt that there's no error"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Text            =   "C:\"
         ToolTipText     =   "select directory containing files on which to test compression"
         Top             =   300
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Compress with RLE"
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   660
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Compress with Huffman"
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Compress with LZ77"
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   1500
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Decompress"
         Height          =   375
         Left            =   4680
         TabIndex        =   7
         Top             =   2340
         Width           =   2175
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Compress with ZLib"
         Height          =   375
         Left            =   4680
         TabIndex        =   6
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Compress with RLE Huffman"
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   2760
         Width           =   2175
      End
      Begin VB.FileListBox File1 
         Height          =   1065
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "select a file to test compression on (WARNING: file may get screwed up if compression fails)"
         Top             =   660
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Length"
      Height          =   1275
      Left            =   4740
      TabIndex        =   15
      Top             =   3360
      Width           =   2295
      Begin VB.Label Label1 
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LZ77 Test"
      Height          =   375
      Left            =   2460
      TabIndex        =   13
      Top             =   4740
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   12
      ToolTipText     =   "testing results"
      Top             =   3420
      Width           =   4515
   End
End
Attribute VB_Name = "Testfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'(c)2001, 2004 by Louis. Test form for GFCompressionmod.
'
'Downloaded from www.louis-coder.com.
'These are compression functions made by Louis Coder. The ZLib compression
'was created by Jean-loup Gailly & Mark Adler (freely downloadable in the Internet).
'
'To compress a file, call GFCompression_CompressFile( _
'    CompressionName, CompressionMethodName, _
'    TempFileReturnEnabledFlag, TempFileReturned)
'
'CompressionName is the name of the file to compress. The file will be compressed block-wise,
'the blocks are first written into a temporary file, and then the temp file is named as the original
'file (original file will be overwritten) if TempFileReturnEnabledFlag is False.
'If TempFileReturnEnabledFlag is True, the original file will not be overwritten but you get the
'name of the temp file in TempFileReturned.
'
'CompressionName is the compression method. The following methods are available:
'-"huffman"
'   -the file data will be compressed using the Huffman compression
'-"rle"
'   -the file data will be compressed using a run length encoding
'-"rle huffman" or "huffman rle"
'   -RLE and then Huffman compression will be used
'-"zlib"
'   -the file data will be compressed using the ZLib compression by
'    Jean-loup Gailly & Mark Adler. The interface to the dll (cmprzlib.dll, must be located on the
'    target machine) was written by Louis Coder.
'
'The LZ77 compression does not work, don't try to use it.
'
'Please note that the file is compressed (and decompressed) in blocks, so even extremely
'large files can be compressed. The whole header stuff required for the decompression is added
'automatically. That's why GFCompression_DecompressFile() will automatically know the
'compression method of a file having been compressed with the GFCompression functions.
'
'You can create a 'Compression Pack' containing several files and extra data (strings)
'using the CompressionPack functions. Just pass an array of files, the compression method
'and an array of strings (e.g. copyright- or password info for the files). If you don't need to save
'string data (that's optional) the just pass StringNumber = 0.
'When using GFCompression_CompressionPack_Unpack() on a compression pack, all contained
'files will be unpacked to OutputDirectory, no matter where (in which sub directories) they were
'originally located (merely the file name is retained).
'
'IMPORTANT: when having debugged your application (that uses GFCompression functions)
'then compile it optimized for speed and DISABLE ALL ERROR CHECKING OPTIONS in the
'extended compiling options window (mainly the array boundaries check must be disabled).
'
'The compression has been tested successfully and
'was used in Toricxs (www.toricxs.com).
'If you have questions, mail louis@louis-coder.com.
'
'DEBUG
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Sub Text1_KeyPress(KeyAscii As Integer)
    'on error resume next
    Select Case KeyAscii
    Case 13
        If Not (Dir(Text1.Text, vbDirectory) = "") Then
            File1.Path = Text1.Text
            File1.Refresh
        End If
    End Select
End Sub

'*********************************DEBUG COMMAND CLICKS*********************************

Private Sub Command1_Click()
    'on error resume next
    Dim InputName As String
    'preset
    InputName = File1.Path
    If Not (Right$(InputName, 1) = "\") Then InputName = InputName + "\"
    InputName = InputName + File1.FileName
    'begin
    Call DEBUG_CompressionCheck(InputName)
End Sub

Private Sub Command2_Click()
    'on error resume next
    Dim FileNumber As Integer
    Dim FileArray(1 To 4) As String
    Dim NULLARRAYSTRING() As String
    'preset
    FileNumber = 4
    FileArray(1) = "C:\Command.com"
    FileArray(2) = "C:\Claw.bmp"
    FileArray(3) = "C:\Windows\Calc.exe"
    FileArray(4) = "C:\Windows\DrvSpace.exe"
    'begin
    Debug.Print GFCompression_CompressionPack_Create("C:\CPack.dat", FileNumber, FileArray(), "rle huffman", 0, NULLARRAYSTRING())
End Sub

Private Sub Command3_Click()
    'on error resume next
    Debug.Print GFCompression_CompressionPack_Unpack("C:\CPack.dat", "C:\Unzipped\")
End Sub

Private Sub Command4_Click()
    'on error resume next
    Dim InputName As String
    Dim FileLoop As Integer
    'begin
    For FileLoop = 1 To File1.ListCount
        InputName = File1.Path
        If Not (Right$(InputName, 1) = "\") Then InputName = InputName + "\"
        InputName = InputName + File1.List(FileLoop - 1)
        If FileLen(InputName) < 1000000 Then 'around 1MB
            Call DEBUG_CompressionCheck(InputName)
        End If
    Next FileLoop
End Sub

Private Sub Command5_Click()
    'on error resume next
    Dim TempByteStringLength As Long
    Dim TempByteString() As Byte
    Dim Tempstr$
    'begin
    'Open "c:\command.com" For Binary As #1
    '    Tempstr$ = String$(LOF(1), Chr$(0))
    '    Get #1, 1, Tempstr$
    'Close #1
    Tempstr$ = "This is a very important test because a test is very important"
    TempByteStringLength = Len(Tempstr$)
    Call GETBYTESTRINGFROMSTRING(TempByteStringLength, TempByteString, Tempstr$)
    Call LZ77_CompressString(TempByteStringLength, TempByteString())
    Call LZ77_DecompressString(TempByteStringLength, TempByteString(), 0&)
    Call DEBUG_DISPLAYBYTESTRING(TempByteString(), 1, 128)
End Sub

Private Sub Command6_Click() 'compress with lre
    'on error resume next
    Dim CompressionName As String
    Dim Tempstr$
    'preset
    CompressionName = File1.Path
    If Not (Right$(CompressionName, 1) = "\") Then CompressionName = CompressionName + "\"
    CompressionName = CompressionName + File1.List(File1.ListIndex)
    'begin
    Call GFCompression_CompressFile(CompressionName, "rle", False, Tempstr$)
    Label1.Caption = LTrim$(Str$(FileLen(CompressionName)))
End Sub

Private Sub Command7_Click() 'compress with huffman
    'on error resume next
    Dim CompressionName As String
    Dim Tempstr$
    'preset
    CompressionName = File1.Path
    If Not (Right$(CompressionName, 1) = "\") Then CompressionName = CompressionName + "\"
    CompressionName = CompressionName + File1.List(File1.ListIndex)
    'begin
    Dim t As Single
    t = Timer
    Call GFCompression_CompressFile(CompressionName, "huffman", False, Tempstr$)
    Label1.Caption = LTrim$(Str$(FileLen(CompressionName)))
    MsgBox Timer - t
End Sub

Private Sub Command8_Click() 'compress with lz77
    'on error resume next
    Dim CompressionName As String
    Dim Tempstr$
    'preset
    CompressionName = File1.Path
    If Not (Right$(CompressionName, 1) = "\") Then CompressionName = CompressionName + "\"
    CompressionName = CompressionName + File1.List(File1.ListIndex)
    'begin
    Call GFCompression_CompressFile(CompressionName, "lz77", False, Tempstr$)
    Label1.Caption = LTrim$(Str$(FileLen(CompressionName)))
End Sub

Private Sub Command9_Click() 'compress with zlib
    'on error resume next
    Dim CompressionName As String
    Dim Tempstr$
    'preset
    CompressionName = File1.Path
    If Not (Right$(CompressionName, 1) = "\") Then CompressionName = CompressionName + "\"
    CompressionName = CompressionName + File1.List(File1.ListIndex)
    'begin
    Call GFCompression_CompressFile(CompressionName, "zlib", False, Tempstr$)
    Label1.Caption = LTrim$(Str$(FileLen(CompressionName)))
End Sub

Private Sub Command10_Click() 'decompress
    Dim DecompressionName As String
    Dim Tempstr$
    'preset
    DecompressionName = File1.Path
    If Not (Right$(DecompressionName, 1) = "\") Then DecompressionName = DecompressionName + "\"
    DecompressionName = DecompressionName + File1.List(File1.ListIndex)
    'begin
    If GFCompression_DecompressFile(DecompressionName, False, Tempstr$) = False Then
        MsgBox "error decompressing file !", vbOKOnly + vbExclamation
    End If
    Label1.Caption = LTrim$(Str$(FileLen(DecompressionName)))
End Sub

Private Sub Command11_Click()
    'on error resume next
    Dim CompressionName As String
    Dim Tempstr$
    'preset
    CompressionName = File1.Path
    If Not (Right$(CompressionName, 1) = "\") Then CompressionName = CompressionName + "\"
    CompressionName = CompressionName + File1.List(File1.ListIndex)
    'begin
    Dim t As Single
    t = Timer
    Call GFCompression_CompressFile(CompressionName, "rle huffman", False, Tempstr$)
    Label1.Caption = LTrim$(Str$(FileLen(CompressionName)))
    MsgBox Timer - t
End Sub

'*****************************END OF DEBUG COMMAND CLICKS******************************
'****************************************OTHER*****************************************

Private Sub DEBUG_CompressionCheck(ByVal InputName As String)
    'on error resume next
    Dim InputNameString As String
    Dim TempFile As String
    Dim TempFileString As String
    Dim Tempstr$
    'verify
    If (Dir(InputName) = "") Or (Right$(InputName, 1) = "\") Or (Len(InputName) = 0) Then
        MsgBox "internal error in DEBUG_CompressionCheck(): file '" + InputName + "' not found !", vbOKOnly + vbExclamation
        Exit Sub
    End If
    'begin
    Tempstr$ = Text2.Text + "TESTING: " + InputName + Chr$(13) + Chr$(10)
    Text2.Text = Tempstr$
    Text2.SelStart = Len(Text2.Text)
    Text2.SelLength = 0
    Text2.Refresh 'important
    Dim t As Single
    t = Timer
    Call GFCompression_CompressFile(InputName, "rle huffman", True, TempFile)
    If Not (FileLen(InputName) = 0) Then 'verify (avoid division through 0)
        Tempstr$ = Text2.Text + LTrim$(Str$(CSng(FileLen(TempFile)) / CSng(FileLen(InputName)))) + Chr$(13) + Chr$(10)
        Text2.Text = Tempstr$
    Else
        Tempstr$ = Text2.Text + "0 byte file" + Chr$(13) + Chr$(10)
        Text2.Text = Tempstr$
    End If
    'MsgBox "TempFile: " + TempFile
    Call GFCompression_DecompressFile(TempFile, False, Tempstr$)
    Debug.Print Timer - t
    If IsFileEqual(InputName, TempFile) = True Then
        Tempstr$ = Text2.Text + "YEEHA!" + Chr$(13) + Chr$(10)
        Text2.Text = Tempstr$
        Text2.SelStart = Len(Text2.Text)
        Text2.SelLength = 0
        Text2.Refresh 'important
    Else
        Dim String1 As String
        Dim String2 As String
        Open InputName For Binary As #1
        Open TempFile For Binary As #2
            String1 = String$(LOF(1), Chr$(0))
            String2 = String$(LOF(2), Chr$(0))
            Get #1, 1, String1
            Get #2, 1, String2
            Dim Temp As Long
            For Temp = 1 To Len(String1)
                If Not (Mid$(String1, Temp, 1) = Mid$(String2, Temp, 1)) Then
                    Tempstr$ = Text2.Text + "ERROR AT FILE POSITION: " + LTrim$(Str$(Temp)) + Chr$(13) + Chr$(10)
                    Text2.Text = Tempstr$
                    Tempstr$ = Text2.Text + "FILE LENGTH: " + LTrim$(Str$(FileLen(TempFile))) + Chr$(13) + Chr$(10)
                    Text2.Text = Tempstr$
                    Text2.SelStart = Len(Text2.Text)
                    Text2.SelLength = 0
                    Text2.Refresh
                    MsgBox "ERROR"
                    Exit For
                End If
            Next Temp
        Close #2
        Close #1
    End If
    If Not ((Dir(TempFile) = "") Or (Right$(TempFile, 1) = "\") Or (Len(TempFile) = 0)) Then Kill TempFile
    Exit Sub
End Sub

Private Function IsFileEqual(ByVal File1 As String, ByVal File2 As String) As Boolean 'can be used as a general function
    'on error resume next 'returns True if both files are equal, False if not, then FirstChangePos points to the first char that differs
    Dim BlockStartPos1 As Long
    Dim BlockLength1 As Long
    Dim BlockString1 As String
    Dim BlockStartPos2 As Long
    Dim BlockLength2 As Long
    Dim BlockString2 As String
    Dim File1FileNumber As Integer
    Dim File2FileNumber As Integer
    'verify
    If (Dir(File1) = "") Or (Right$(File1, 1) = "\") Or (Len(File1) = 0) Then GoTo Error: 'verify
    If (Dir(File2) = "") Or (Right$(File2, 1) = "\") Or (Len(File2) = 0) Then GoTo Error: 'verify
    'preset
    BlockStartPos1 = 1 'preset
    BlockStartPos2 = 1 'preset
    'begin
    File1FileNumber = FreeFile(0)
    Open File1 For Binary As #File1FileNumber
    File2FileNumber = FreeFile(0)
    Open File2 For Binary As #File2FileNumber
        Do
            BlockLength1 = 512000
            BlockLength2 = 512000
            If (BlockStartPos1 + BlockLength1 - 1) > LOF(File1FileNumber) Then
                BlockLength1 = LOF(File1FileNumber) - BlockStartPos1 + 1
            End If
            If (BlockStartPos2 + BlockLength2 - 1) > LOF(File2FileNumber) Then
                BlockLength2 = LOF(File2FileNumber) - BlockStartPos2 + 1
            End If
            '
            If Not (BlockLength1 = BlockLength2) Then GoTo Error:
            If BlockLength1 < 1 Then Exit Do 'ok
            '
            BlockString1 = String$(BlockLength1, Chr$(0))
            BlockString2 = String$(BlockLength2, Chr$(0))
            Get #File1FileNumber, BlockStartPos1, BlockString1
            Get #File2FileNumber, BlockStartPos2, BlockString2
            If Len(BlockString1) = Len(BlockString2) Then
                If Not (InStr(1, BlockString1, BlockString2, vbBinaryCompare) = 1) Then
                    GoTo Error:
                End If
            End If
            BlockStartPos1 = BlockStartPos1 + BlockLength1
            BlockStartPos2 = BlockStartPos2 + BlockLength2
        Loop
    Close #File2FileNumber
    Close #File1FileNumber
    IsFileEqual = True 'ok
    Exit Function
Error:
    Close #File1FileNumber
    Close #File2FileNumber
    IsFileEqual = False 'error
    Exit Function
End Function

'************************************END OF OTHER**************************************
