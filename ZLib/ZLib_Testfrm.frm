VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'(c)2001 by Louis, ZLib.dll (c) by Jean-loup Gailly & Mark Adler. Test project for ZLib.dll.
'
'NOTE: remove directory 'F:\GF\GFCompression\ZLib\' when copying declarations to target project.
'
Private Declare Function compress Lib "F:\GF\GFCompression\ZLib\zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "F:\GF\GFCompression\ZLib\zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal level As Long) As Long
Private Declare Function uncompress Lib "F:\GF\GFCompression\ZLib\zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Private Sub Form_Load()
    'on error resume next
    Dim SourceByteStringLength As Long
    Dim SourceByteString() As Byte
    Dim TargetByteStringLength As Long
    Dim TargetByteString() As Byte
    'begin
    Open "C:\Command.com" For Binary As #1
        SourceByteStringLength = LOF(1)
        ReDim SourceByteString(1 To SourceByteStringLength) As Byte
        Get #1, 1, SourceByteString()
    Close #1
    TargetByteStringLength = CLng(CSng(SourceByteStringLength) * 1.05!) 'add five percent for safety
    ReDim TargetByteString(1 To TargetByteStringLength) As Byte
    'compress
    Debug.Print compress(TargetByteString(1), TargetByteStringLength, SourceByteString(1), SourceByteStringLength)
    Debug.Print TargetByteStringLength
    'decompress
    Call uncompress(SourceByteString(1), SourceByteStringLength, TargetByteString(1), TargetByteStringLength)
    Debug.Print SourceByteStringLength
End Sub
