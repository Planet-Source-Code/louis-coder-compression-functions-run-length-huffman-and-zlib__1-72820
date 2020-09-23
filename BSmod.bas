Attribute VB_Name = "BSmod"
Option Explicit
'(c)2001-2003 by Louis.
'
'NOTE: do not enable 'on error Resume Next to increase speed
'(speed is the main reason to implement BSmod).
'NOTE: Byte vars seem internally to be converted to long when checking its
'value, thus If ByteString(1) = 0& is faster than If ByteString(1) = 0 (tested).
'But attention: allocation is faster with Integer values, thus
'Let ByteString(1) = 0 is faster than Let ByteString(1) = 0&.
'Also Select Case is NOT faster when comparing Byte- with Long values.
'Select Case ByteString(1): Case 1%: End Select is faster than
'Select Case ByteString(1): Case 1&: End Select.
'
#Const TargetProjectDefinesByteUCaseTableFlag = True
#Const TargetProjectDefinesByteLCaseTableFlag = True
'
'general use
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Sub BYTESTRING_COPYMEMORY Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'GETBYTESTRINGLENGTH
Dim ByteStringLengthMax As Long 'use global var to increase speed
'GETBYTESTRINGLENGTHFIXED
Dim ByteStringLengthFixed As Long
'BYTEUCASE
Dim ByteUCaseTableDefinedFlag As Boolean
Dim ByteUCaseTable(0 To 255) As Byte
'BYTELCASE
Dim ByteLCaseTableDefinedFlag As Boolean
Dim ByteLCaseTable(0 To 255) As Byte
'other
Dim ByteString1LengthGlobal As Long 'used to save time of memory allocation
Dim ByteString2LengthGlobal As Long 'used to save time of memory allocation
Dim TempGlobal As Long 'used to save time of memory allocation

'***BYTE STRING FUNCTIONS***
'NOTE: the following subs/functions are used to manipulate 'byte strings',
'i.e. byte arrays that replaced normal strings to increase speed.
'NOTE: generally you must differ between 1) GETBYTESTRINGLENGTH() and 2) UBound().
'1) returns 'how much is used' of the maximal possible byte string length, and
'2) returns the maximal possible byte string length (important to avoid crashing program with CopyMemory()).

Public Sub DefineByteUCaseTable()
    'on error Resume Next
    Dim Temp As Long
    ByteUCaseTableDefinedFlag = True
    For Temp = 0 To 255
        ByteUCaseTable(Temp) = Asc(UCase$(Chr$(Temp)))
    Next Temp
End Sub

Public Sub DefineByteLCaseTable()
    'on error Resume Next
    Dim Temp As Long
    ByteLCaseTableDefinedFlag = True
    For Temp = 0 To 255
        ByteLCaseTable(Temp) = Asc(LCase$(Chr$(Temp)))
    Next Temp
End Sub

'************************************COPY FUNCTIONS************************************

Public Sub BYTESTRINGCOPY(ByRef TargetByteString() As Byte, ByRef SourceByteString() As Byte)
    'on error resume next
    Dim SourceByteStringLength As Long
    'preset
    SourceByteStringLength = UBound(SourceByteString())
    'begin
    If (SourceByteStringLength > 0&) Then 'verify
        ReDim TargetByteString(1 To SourceByteStringLength) As Byte
        Call CopyMemory(TargetByteString(1), SourceByteString(1), SourceByteStringLength)
    Else
        ReDim TargetByteString(1 To 1) As Byte
    End If
End Sub

Public Sub BYTESTRINGCOPYEX(ByRef TargetByteString() As Byte, ByRef SourceByteString() As Byte, ByVal SourceByteStringLength As Long)
    'on error resume next
    If (SourceByteStringLength > 0&) Then 'verify
        ReDim TargetByteString(1 To SourceByteStringLength) As Byte
        Call CopyMemory(TargetByteString(1), SourceByteString(1), SourceByteStringLength)
    Else
        ReDim TargetByteString(1 To 1) As Byte
    End If
End Sub

Public Sub BYTESTRINGCOPYFIXED(ByRef TargetByteString() As Byte, ByRef SourceByteString() As Byte)
    'on error resume next 'does not resize target byte string, does not clear target byte string
    Call CopyMemory(TargetByteString(1), SourceByteString(1), BS_MIN(UBound(SourceByteString()), UBound(TargetByteString())))
End Sub

'********************************END OF COPY FUNCTIONS*********************************
'*********************************CONVERSION FUNCTIONS*********************************
'NOTE: the following subs/functions are used to create byte strings out of string and vice versa.

Public Sub GETBYTESTRINGFROMSTRING(ByVal ByteStringLengthTotal As Long, ByRef ByteString() As Byte, ByVal NormalString As String)
    'on error Resume Next 'redims the passed ByteString() array
    ReDim ByteString(1 To ByteStringLengthTotal) As Byte
    Call CopyMemory(ByteString(1), ByVal NormalString, BS_MIN(ByteStringLengthTotal, Len(NormalString)))
End Sub

Public Sub GETFIXEDBYTESTRINGFROMSTRING(ByVal ByteStringLengthTotal As Long, ByRef ByteString() As Byte, ByVal NormalString As String)
    'on error Resume Next 'does NOT redim the passed ByteString() array
    Call CopyMemory(ByteString(1), ByVal NormalString, BS_MIN(ByteStringLengthTotal, Len(NormalString)))
End Sub

Public Sub GETSTRINGFROMBYTESTRING(ByRef ByteString() As Byte, ByRef NormalString As String)
    'on error Resume Next
    Dim ByteStringLength As Long
    ByteStringLength = GETBYTESTRINGLENGTH(ByteString())
    NormalString = String$(ByteStringLength, Chr$(0))
    Call CopyMemory(ByVal NormalString, ByteString(1), ByteStringLength)
End Sub

Public Function GETRETURNSTRINGFROMBYTESTRING(ByRef ByteString() As Byte) As String
    'on error Resume Next
    Dim ByteStringLength As Long
    Dim Tempstr$
    ByteStringLength = GETBYTESTRINGLENGTH(ByteString())
    Tempstr$ = String$(ByteStringLength, Chr$(0))
    Call CopyMemory(ByVal Tempstr$, ByteString(1), ByteStringLength)
    GETRETURNSTRINGFROMBYTESTRING = Tempstr$
End Function

'*****************************END OF CONVERSION FUNCTIONS******************************
'*********************************INFORMATION FUNCTIONS********************************
'NOTE: the following subs/functions ar used to gain misc information about byte strings.
'The information functions should be optimized for speed.

Public Function GETBYTESTRINGLENGTH(ByRef ByteString() As Byte) As Long
    'on error Resume Next 'returns appearence of first Chr$(0) in ByteString() or UBound(ByteString()); a special algorithm is used that increases the checking speed compared to a simple loop
    ByteStringLengthMax = UBound(ByteString())
    For TempGlobal = 1& To ByteStringLengthMax Step 2&
        If ByteString(TempGlobal) = 0& Then
            If Not (TempGlobal = 1&) Then
                If ByteString(TempGlobal - 1) = 0& Then
                    GETBYTESTRINGLENGTH = (TempGlobal - 2&)
                    Exit Function
                Else
                    GETBYTESTRINGLENGTH = (TempGlobal - 1&)
                    Exit Function
                End If
            Else
                GETBYTESTRINGLENGTH = (TempGlobal - 1&)
                Exit Function
            End If
        End If
    Next TempGlobal
    If ByteString(ByteStringLengthMax) = 0& Then
        GETBYTESTRINGLENGTH = ByteStringLengthMax - 1&
    Else
        GETBYTESTRINGLENGTH = ByteStringLengthMax
    End If
    Exit Function
End Function

Public Function GETBYTESTRINGLENGTHMAX(ByRef ByteString() As Byte) As Long
    'on error resume next 'professional-looking UBound()
    GETBYTESTRINGLENGTHMAX = UBound(ByteString()) 'ByteString() must have been Dim-ed already
End Function

Public Function BYTESTRINGISEQUAL(ByRef ByteString1() As Byte, ByRef ByteString2() As Byte, ByVal IgnoreCapitalizationFlag As Boolean) As Boolean
    'on error Resume Next 'returns True if data (!) in the two passed byte strings is equal
    Dim ByteString1LengthGlobal As Long
    Dim ByteString2LengthGlobal As Long
    Dim TempGlobal As Long
    '
    'NOTE: when using large loops to check byte strings for equalness,
    'first check if (ByteString1(1) = ByteString2(1)) = True before calling
    'this function.
    '
    ByteString1LengthGlobal = GETBYTESTRINGLENGTH(ByteString1())
    ByteString2LengthGlobal = GETBYTESTRINGLENGTH(ByteString2())
    If Not (ByteString1LengthGlobal = ByteString2LengthGlobal) Then
        BYTESTRINGISEQUAL = False
        Exit Function
    Else
        If IgnoreCapitalizationFlag = False Then
            For TempGlobal = 1& To ByteString1LengthGlobal 'or ByteString2LengthGlobal
                If Not (ByteString1(TempGlobal) = ByteString2(TempGlobal)) Then
                    BYTESTRINGISEQUAL = False
                    Exit Function
                End If
            Next TempGlobal
        Else
            For TempGlobal = 1& To ByteString1LengthGlobal 'or ByteString2LengthGlobal
                If Not (BYTEUCASE(ByteString1(TempGlobal)) = BYTEUCASE(ByteString2(TempGlobal))) Then
                    BYTESTRINGISEQUAL = False
                    Exit Function
                End If
            Next TempGlobal
        End If
    End If
    BYTESTRINGISEQUAL = True
    Exit Function
End Function

Public Function BYTESTRINGISEQUAL3(ByRef ByteString1() As Byte, ByRef ByteString2() As Byte, ByVal IgnoreCapitalizationFlag As Boolean) As Boolean
    'on error Resume Next 'returns True if data (!) in the two passed byte strings is equal
    '
    'NOTE: use this function instead of BYTESTRINGISEQUAL3() if both passed byte
    'strings have an UBound >= 3. For this case this faster function here can be used.
    '
    If IgnoreCapitalizationFlag = False Then
        If ByteString1(1) = ByteString2(1) Then
            If ByteString1(2) = ByteString2(2) Then
                If ByteString1(3) = ByteString2(3) Then
                    'continue below
                Else
                    GoTo NotEqual:
                End If
            Else
                GoTo NotEqual:
            End If
        Else
            GoTo NotEqual:
        End If
    Else
        If BYTEUCASE(ByteString1(1)) = BYTEUCASE(ByteString2(1)) Then
            If BYTEUCASE(ByteString1(2)) = BYTEUCASE(ByteString2(2)) Then
                If BYTEUCASE(ByteString1(3)) = BYTEUCASE(ByteString2(3)) Then
                    'continue below
                Else
                    GoTo NotEqual:
                End If
            Else
                GoTo NotEqual:
            End If
        Else
            GoTo NotEqual:
        End If
    End If
    '
    ByteString1LengthGlobal = GETBYTESTRINGLENGTH(ByteString1())
    ByteString2LengthGlobal = GETBYTESTRINGLENGTH(ByteString2())
    If Not (ByteString1LengthGlobal = ByteString2LengthGlobal) Then
        BYTESTRINGISEQUAL3 = False
        Exit Function
    Else
        If IgnoreCapitalizationFlag = False Then
            For TempGlobal = 1& To ByteString1LengthGlobal 'or ByteString2LengthGlobal
                If Not (ByteString1(TempGlobal) = ByteString2(TempGlobal)) Then
                    BYTESTRINGISEQUAL3 = False
                    Exit Function
                End If
            Next TempGlobal
        Else
            For TempGlobal = 1& To ByteString1LengthGlobal 'or ByteString2LengthGlobal
                If Not (BYTEUCASE(ByteString1(TempGlobal)) = BYTEUCASE(ByteString2(TempGlobal))) Then
                    BYTESTRINGISEQUAL3 = False
                    Exit Function
                End If
            Next TempGlobal
        End If
    End If
    BYTESTRINGISEQUAL3 = True
    Exit Function
NotEqual:
    BYTESTRINGISEQUAL3 = False
    Exit Function
End Function

Public Sub BYTESTRINGISEQUALFIXED_SETLENGTH(ByVal ByteStringLengthFixedPassed As Long)
    'on error resume next
    ByteStringLengthFixed = ByteStringLengthFixedPassed
End Sub

Public Function BYTESTRINGISEQUALFIXED(ByRef ByteString1() As Byte, ByRef ByteString2() As Byte, ByVal IgnoreCapitalizationFlag As Boolean) As Boolean
    'on error Resume Next 'returns True if data (!) in the two passed byte strings is equal (both byte strings msut have a fixed length that was previously passed to BYTESTRINGISEQUALFIXED_SETLENGTH())
    Dim TempGlobal As Long
    '
    'NOTE: when using large loops to check byte strings for equalness,
    'first check if (ByteString1(1) = ByteString2(1)) = True before calling
    'this function.
    '
    If IgnoreCapitalizationFlag = False Then
        For TempGlobal = 1& To ByteStringLengthFixed
            If Not (ByteString1(TempGlobal) = ByteString2(TempGlobal)) Then
                BYTESTRINGISEQUALFIXED = False
                Exit Function
            End If
        Next TempGlobal
    Else
        For TempGlobal = 1& To ByteStringLengthFixed
            If Not (BYTEUCASE(ByteString1(TempGlobal)) = BYTEUCASE(ByteString2(TempGlobal))) Then
                BYTESTRINGISEQUALFIXED = False
                Exit Function
            End If
        Next TempGlobal
    End If
    BYTESTRINGISEQUALFIXED = True
    Exit Function
End Function

'NOTE: the following function was not really faster than the original BYTESTRINGISEQUAL() function.
'Public Function BYTESTRINGISEQUAL2(ByRef ByteString1() As Byte, ByRef ByteString2() As Byte, ByVal IgnoreCapitalizationFlag As Boolean, ByVal ByteString1LengthMax As Long, ByVal ByteString2LengthMax As Long) As Boolean
'    'on error resume next
'    Dim Temp As Long
'    '
'    'NOTE: this function is faster than BYTESTRINGISEQUAL().
'    'Use this function if the maximal byte string length is known
'    '(this function saves GETBYTESTRINGLENGTH() calls).
'    '
'    'preset
'    'begin
'    If IgnoreCapitalizationFlag = False Then
'        For Temp = 1 To BS_MIN(ByteString1LengthMax, ByteString2LengthMax)
'            If ByteString1(Temp) = ByteString2(Temp) Then
'                'ok
'            Else
'                BYTESTRINGISEQUAL2 = False
'                Exit Function
'            End If
'        Next Temp
'        If ByteString1LengthMax > ByteString2LengthMax Then
'            For Temp = (ByteString2LengthMax + 1&) To ByteString1LengthMax
'                If ByteString1(Temp) = 0& Then
'                    Exit For 'end of string
'                Else
'                    BYTESTRINGISEQUAL2 = False
'                    Exit Function
'                End If
'            Next Temp
'        Else
'            For Temp = (ByteString1LengthMax + 1&) To ByteString2LengthMax
'                If ByteString2(Temp) = 0& Then
'                    Exit For 'end of string
'                Else
'                    BYTESTRINGISEQUAL2 = False
'                    Exit Function
'                End If
'            Next Temp
'        End If
'    Else
'        For Temp = 1 To BS_MIN(ByteString1LengthMax, ByteString2LengthMax)
'            If BYTEUCASE(ByteString1(Temp)) = BYTEUCASE(ByteString2(Temp)) Then
'                'ok
'            Else
'                BYTESTRINGISEQUAL2 = False
'                Exit Function
'            End If
'        Next Temp
'        If ByteString1LengthMax > ByteString2LengthMax Then
'            For Temp = (ByteString2LengthMax + 1&) To ByteString1LengthMax
'                If ByteString1(Temp) = 0& Then
'                    Exit For 'end of string
'                Else
'                    BYTESTRINGISEQUAL2 = False
'                    Exit Function
'                End If
'            Next Temp
'        Else
'            For Temp = (ByteString1LengthMax + 1&) To ByteString2LengthMax
'                If ByteString2(Temp) = 0& Then
'                    Exit For 'end of string
'                Else
'                    BYTESTRINGISEQUAL2 = False
'                    Exit Function
'                End If
'            Next Temp
'        End If
'    End If
'    BYTESTRINGISEQUAL2 = True 'if not quit before
'    Exit Function
'End Function

Public Function InStrByte(ByVal StartPos As Long, ByRef ByteString1() As Byte, ByRef ByteString2() As Byte, ByVal CompareMethod As Integer) As Long
    'on error Resume Next
    Dim ByteString2Pos As Long
    Dim ByteString2Length As Long
    Dim Temp As Long
    'verify
    If StartPos < 1& Then
        InStrByte = 0&
        Exit Function
    End If
    'preset
    ByteString2Pos = 1&
    If ByteString2Pos > UBound(ByteString2()) Then
        InStrByte = 0& 'error
        Exit Function 'verify
    End If
    ByteString2Length = GETBYTESTRINGLENGTH(ByteString2())
    'begin
    Select Case CompareMethod
    Case vbBinaryCompare
        For Temp = StartPos To GETBYTESTRINGLENGTH(ByteString1())
            If ByteString1(Temp) = ByteString2(ByteString2Pos) Then
                ByteString2Pos = ByteString2Pos + 1&
                If ByteString2Pos > ByteString2Length Then
                    InStrByte = Temp - ByteString2Pos + 2& 'ok
                    Exit Function
                End If
            Else
                ByteString2Pos = 1& 'reset (important)
            End If
        Next Temp
    Case Else 'i.e. vbTextCompare
        For Temp = StartPos To GETBYTESTRINGLENGTH(ByteString1())
            If BYTEUCASE(ByteString1(Temp)) = BYTEUCASE(ByteString2(ByteString2Pos)) Then
                ByteString2Pos = ByteString2Pos + 1&
                If ByteString2Pos > ByteString2Length Then
                    InStrByte = Temp - ByteString2Pos + 2& 'ok
                    Exit Function
                End If
            Else
                ByteString2Pos = 1& 'reset (important)
            End If
        Next Temp
    End Select
    InStrByte = 0& 'error
    Exit Function
End Function

'NOTE: the following function was not faster than the original InStrByte() function.
'Public Function InStrByte2(ByVal StartPos As Long, ByRef ByteString1() As Byte, ByRef ByteString2() As Byte, ByVal CompareMethod As Integer, ByVal ByteString1LengthMax As Long, ByVal ByteString2LengthMax As Long) As Long
'    'on error resume next
'    Dim ByteString2Pos As Long
'    Dim ByteString2LengthMaxMinusOne As Long
'    Dim Temp As Long
'    'verify
'    If StartPos < 1& Then
'        InStrByte2 = 0&
'        Exit Function
'    End If
'    If (ByteString2(1) = 0&) Or (ByteString2LengthMax = 0) Then
'        InStrByte2 = 0&
'        Exit Function
'    End If
'    'preset
'    ByteString2LengthMaxMinusOne = ByteString2LengthMax - 1&
'    'begin
'    Select Case CompareMethod
'    Case vbBinaryCompare
'        For Temp = StartPos To ByteString1LengthMax
'            If (ByteString2Pos) > ByteString2LengthMaxMinusOne Then
'                InStrByte2 = (Temp - ByteString2Pos)
'                Exit Function
'            End If
'            If ByteString2(1& + ByteString2Pos) = 0& Then
'                InStrByte2 = (Temp - ByteString2Pos)
'                Exit Function
'            End If
'            If ByteString1(Temp) = ByteString2(1& + ByteString2Pos) Then
'                ByteString2Pos = ByteString2Pos + 1&
'            Else
'                ByteString2Pos = 0& 'reset
'            End If
'        Next Temp
'        If (ByteString2Pos) Then
'            InStrByte2 = (Temp - ByteString2Pos)
'        Else
'            InStrByte2 = 0&
'        End If
'    Case vbTextCompare
'        For Temp = StartPos To ByteString1LengthMax
'            If (ByteString2Pos) > ByteString2LengthMaxMinusOne Then
'                InStrByte2 = (Temp - ByteString2Pos)
'                Exit Function
'            End If
'            If ByteString2(1& + ByteString2Pos) = 0& Then
'                InStrByte2 = (Temp - ByteString2Pos)
'                Exit Function
'            End If
'            If BYTEUCASE(ByteString1(Temp)) = BYTEUCASE(ByteString2(1& + ByteString2Pos)) Then
'                ByteString2Pos = ByteString2Pos + 1&
'            Else
'                ByteString2Pos = 0& 'reset
'            End If
'        Next Temp
'        If (ByteString2Pos) Then
'            InStrByte2 = (Temp - ByteString2Pos)
'        Else
'            InStrByte2 = 0&
'        End If
'    End Select
'    Exit Function
'End Function

Public Sub DISPLAYBYTESTRING(ByRef ByteString() As Byte)
    'on error Resume Next 'use for debugging
    Dim Tempstr$
    Call GETSTRINGFROMBYTESTRING(ByteString(), Tempstr$)
    Debug.Print Tempstr$
End Sub

Public Sub DISPLAYBYTESTRINGDEC(ByRef ByteString() As Byte, Optional ByVal ByteStringLength As Long = -1&)
    'on error resume next 'displays byte string data like in a hex editor, but in decimal
    Dim CharFor As Long
    'preset
    If ByteStringLength = -1& Then ByteStringLength = UBound(ByteString())
    'begin
    For CharFor = 1 To ByteStringLength
        If ByteString(CharFor) < 100 Then
            If ByteString(CharFor) < 10 Then
                Debug.Print "00" + CStr(ByteString(CharFor)) + " ";
            Else
                Debug.Print "0" + CStr(ByteString(CharFor)) + " ";
            End If
        Else
            Debug.Print CStr(ByteString(CharFor)) + " ";
        End If
    Next CharFor
    Debug.Print ""
End Sub

'*****************************END OF INFORMATION FUNCTIONS*****************************
'********************************MANIPULATION FUNCTIONS********************************
'NOTE: the following subs/functions are used to manipulate byte strings.
'The manipulation function should be optimized for speed.

Public Sub BYTESTRINGSIZE(ByRef ByteString() As Byte, ByVal ByteStringLengthMaxNew As Long)
    'on error resume next 'use instead of ReDim x (looks more professional)
    ReDim ByteString(1 To BS_MAX(1, ByteStringLengthMaxNew)) As Byte
End Sub

Public Sub BYTESTRINGRESIZE(ByRef ByteString() As Byte, ByVal ByteStringLengthMaxNew As Long)
    'on error resume next 'use instead of ReDim Preserve x (looks more professional)
    ReDim Preserve ByteString(1 To BS_MAX(1, ByteStringLengthMaxNew)) As Byte
End Sub

Public Sub BYTESTRINGCLEAR(ByRef ByteString() As Byte)
    'on error resume next
    Call BYTESTRINGLEFT(ByteString(), 0&)
End Sub

Public Sub BYTESTRINGLEFT(ByRef ByteString() As Byte, ByVal RetainCharNumber As Long)
    'on error Resume Next 'retains left part of ByteString() and sets the rest to 0
    Dim ByteStringLBound As Long
    Dim Temp As Long
    'preset
    ByteStringLBound = LBound(ByteString())
    'begin
    For Temp = (RetainCharNumber + 1&) To UBound(ByteString())
        If Not (Temp < ByteStringLBound) Then 'verify (important)
            ByteString(Temp) = 0 'reset
        End If
    Next Temp
End Sub
    
Public Sub BYTESTRINGMID(ByRef TargetByteString() As Byte, ByRef SourceByteString() As Byte, ByVal CopyStartPos As Long, ByVal CopyLength As Long)
    'on error Resume Next 'like '[...] = Mid$([...])'
    Dim SourceByteStringLength As Long
    'preset
    SourceByteStringLength = GETBYTESTRINGLENGTH(SourceByteString())
    If (CopyStartPos < 1&) Or (CopyStartPos > SourceByteStringLength) Then GoTo Error: 'verify
    If (CopyStartPos + CopyLength - 1&) > SourceByteStringLength Then CopyLength = SourceByteStringLength - CopyStartPos + 1&
    'begin
    If Not (CopyLength < 1&) Then 'verify
        ReDim TargetByteString(1 To CopyLength) As Byte
        Call CopyMemory(TargetByteString(1), SourceByteString(CopyStartPos), CopyLength)
    Else
        GoTo Error:
    End If
    Exit Sub
Error:
    ReDim TargetByteString(1 To 1) As Byte 'verify UBound() etc. will not fail in further actions
    TargetByteString(1) = 0 'reset
    Exit Sub
End Sub
    
Public Sub BYTESTRINGRIGHT(ByRef ByteString() As Byte, ByVal RetainCharNumber As Long)
    'on error Resume Next 'retains left part of ByteString() and sets the rest to 0
    Dim ByteStringLBound As Long
    Dim Temp As Long
    'preset
    ByteStringLBound = LBound(ByteString())
    'begin
    For Temp = 1& To (GETBYTESTRINGLENGTH(ByteString()) - RetainCharNumber)
        If Not (Temp < ByteStringLBound) Then 'verify (important)
            ByteString(Temp) = 0 'reset
        End If
    Next Temp
End Sub

Public Sub BYTESTRINGTRIM(ByRef ByteString() As Byte)
    'on error Resume Next
    Dim ByteStringLength As Long
    Dim Temp As Long
    'preset
    ByteStringLength = GETBYTESTRINGLENGTH(ByteString())
    'cut left spaces
    For Temp = 1& To ByteStringLength
        If Not (ByteString(Temp) = 32&) Then
            If (Temp > 1&) Then
                Call BYTESTRINGCUT(ByteString(), 1, (Temp - 1), 0)
                ByteStringLength = GETBYTESTRINGLENGTH(ByteString()) 'refresh (important)
            End If
            Exit For
        End If
    Next Temp
    'cut right spaces
    For Temp = ByteStringLength To 1& Step (-1&)
        If Not (ByteString(Temp) = 32&) Then
            If (Temp < ByteStringLength) Then
                Call BYTESTRINGCUT(ByteString(), (Temp + 1), (ByteStringLength - Temp), 0)
            End If
            Exit For
        End If
    Next Temp
End Sub

Public Sub BYTESTRINGREMOVESPACE(ByRef ByteString() As Byte)
    'on error Resume Next 'removes all (!) spaces in ByteString()
    Dim ByteStringWritePos As Long
    Dim ByteStringReadPos As Long
    Dim ByteStringLength As Long
    'preset
    ByteStringLength = GETBYTESTRINGLENGTH(ByteString())
    'begin
    For ByteStringReadPos = 1& To ByteStringLength
        If Not (ByteString(ByteStringReadPos) = 32&) Then
            ByteStringWritePos = ByteStringWritePos + 1&
            ByteString(ByteStringWritePos) = ByteString(ByteStringReadPos) 'is moved at left only (using one byte string is possible)
        End If
    Next ByteStringReadPos
    For ByteStringReadPos = (ByteStringWritePos + 1&) To ByteStringLength
        ByteString(ByteStringReadPos) = 0 'reset
    Next ByteStringReadPos
End Sub

Public Function BYTEUCASE(ByRef ByteChar As Byte) As Byte
    'on error Resume Next
    #If TargetProjectDefinesByteUCaseTableFlag = False Then
        If ByteUCaseTableDefinedFlag = False Then Call DefineByteUCaseTable
    #End If
    BYTEUCASE = ByteUCaseTable(ByteChar)
End Function

Public Sub BYTESTRINGUCASE(ByRef ByteString() As Byte, ByVal ByteStringLengthMax As Long)
    'on error resume next
    Dim Temp As Long
    'preset
    #If TargetProjectDefinesByteUCaseTableFlag = False Then
        If ByteUCaseTableDefinedFlag = False Then Call DefineByteUCaseTable
    #End If
    'begin
    For Temp = 1& To ByteStringLengthMax
        ByteString(Temp) = ByteUCaseTable(ByteString(Temp))
    Next Temp
End Sub

Public Function BYTELCASE(ByRef ByteChar As Byte) As Byte
    'on error Resume Next
    #If TargetProjectDefinesByteLCaseTableFlag = False Then
        If ByteLCaseTableDefinedFlag = False Then Call DefineByteLCaseTable
    #End If
    BYTELCASE = ByteLCaseTable(ByteChar)
End Function

Public Sub BYTESTRINGLCASE(ByRef ByteString() As Byte, ByVal ByteStringLengthMax As Long)
    'on error resume next
    Dim Temp As Long
    'preset
    #If TargetProjectDefinesByteLCaseTableFlag = False Then
        If ByteLCaseTableDefinedFlag = False Then Call DefineByteLCaseTable
    #End If
    'begin
    For Temp = 1& To ByteStringLengthMax
        ByteString(Temp) = ByteLCaseTable(ByteString(Temp))
    Next Temp
End Sub

Public Function BYTESTRINGVAL(ByRef ByteString() As Byte) As Double
    'on error Resume Next 'does not check for Double overflow
    Dim ByteStringLengthMax As Long
    Dim NumberEndIndex As Long 'where number before comma ends
    Dim ReturnValue As Double
    Dim Temp As Long
    'preset
    ByteStringLengthMax = UBound(ByteString())
    NumberEndIndex = ByteStringLengthMax
    'begin
    For Temp = 1 To ByteStringLengthMax
        Select Case ByteString(Temp)
        Case 48 '0
        Case 49 '1
        Case 50 '2
        Case 51 '3
        Case 52 '4
        Case 53 '5
        Case 54 '6
        Case 55 '7
        Case 56 '8
        Case 57 '9
        Case 46 '.
            NumberEndIndex = (Temp - 1)
            Exit For
        Case Else
            NumberEndIndex = (Temp - 1)
            Exit For
        End Select
    Next Temp
    For Temp = 1 To ByteStringLengthMax
        Select Case ByteString(Temp)
        Case 48 '0
            ReturnValue = ReturnValue + (ByteString(Temp) - 48) * (10& ^ (NumberEndIndex - Temp))
        Case 49 '1
            ReturnValue = ReturnValue + (ByteString(Temp) - 48) * (10& ^ (NumberEndIndex - Temp))
        Case 50 '2
            ReturnValue = ReturnValue + (ByteString(Temp) - 48) * (10& ^ (NumberEndIndex - Temp))
        Case 51 '3
            ReturnValue = ReturnValue + (ByteString(Temp) - 48) * (10& ^ (NumberEndIndex - Temp))
        Case 52 '4
            ReturnValue = ReturnValue + (ByteString(Temp) - 48) * (10& ^ (NumberEndIndex - Temp))
        Case 53 '5
            ReturnValue = ReturnValue + (ByteString(Temp) - 48) * (10& ^ (NumberEndIndex - Temp))
        Case 54 '6
            ReturnValue = ReturnValue + (ByteString(Temp) - 48) * (10& ^ (NumberEndIndex - Temp))
        Case 55 '7
            ReturnValue = ReturnValue + (ByteString(Temp) - 48) * (10& ^ (NumberEndIndex - Temp))
        Case 56 '8
            ReturnValue = ReturnValue + (ByteString(Temp) - 48) * (10& ^ (NumberEndIndex - Temp))
        Case 57 '9
            ReturnValue = ReturnValue + (ByteString(Temp) - 48) * (10& ^ (NumberEndIndex - Temp))
        Case 46 '.
            NumberEndIndex = NumberEndIndex + 1 '10 ^ (-1) will be used for next number
        Case Else
            Exit For
        End Select
    Next Temp
    BYTESTRINGVAL = ReturnValue
End Function

Public Function BYTESTRINGVALLONG(ByRef ByteString() As Byte) As Long
    'on error Resume Next 'save return value in a var of the type Long
    Dim ByteStringLength As Long
    Dim UseCharNumber As Long 'how many chars of ByteString() are used to get number
    Dim Temp As Long
    'preset
    ByteStringLength = GETBYTESTRINGLENGTH(ByteString())
    UseCharNumber = ByteStringLength 'preset
    'begin
    For Temp = 1& To ByteStringLength
        If (ByteString(Temp) < 48&) Or (ByteString(Temp) > 57&) Then
            UseCharNumber = (Temp - 1&)
            Exit For
        End If
    Next Temp
    For Temp = 1& To UseCharNumber
        BYTESTRINGVALLONG = BYTESTRINGVALLONG + 10& ^ (UseCharNumber - Temp) * (ByteString(Temp) - 48)
    Next Temp
End Function

Public Sub BYTESTRINGCUT(ByRef ByteString() As Byte, ByVal CutStartPos As Long, ByVal CutLength As Long, Optional ByVal RightEndReplaceAsc As Byte = 0, Optional ByVal ByteStringLengthFixedFlag As Boolean = True)
    'on error Resume Next 'removes data in byte string leftwards and deltes 'right end' of moved data through overwriting it with passed asc
    Dim ByteStringLength As Long
    Dim MoveBlockStartPos As Long
    Dim MoveBlockLength As Long
    Dim Temp As Long
    'NOTE: (ByteStringLength - (CutStartPos + CutLength - 1) + 1) is the length of block to move leftwards.
    'preset
    If ByteStringLengthFixedFlag = True Then
        ByteStringLength = GETBYTESTRINGLENGTH(ByteString())
    Else
        'NOTE: when wanting to cut data at the end of the byte string
        'this function would originally return an error, now it doesn't any more.
        ByteStringLength = UBound(ByteString())
    End If
    If (CutStartPos < 1&) Or (CutStartPos > ByteStringLength) Then 'verify
        MsgBox "internal error in BYTESTRINGCUT(): passed value invalid !", vbOKOnly + vbExclamation
        Exit Sub
    End If
    If (CutStartPos + CutLength - 1&) > ByteStringLength Then
        CutLength = ByteStringLength - CutStartPos + 1&
    End If
    MoveBlockStartPos = CutStartPos + CutLength
    MoveBlockLength = ByteStringLength - MoveBlockStartPos + 1&
    'begin
    If Not (MoveBlockStartPos > ByteStringLength) Then
        If (MoveBlockStartPos + MoveBlockLength - 1&) > UBound(ByteString()) Then 'verify
            MoveBlockLength = UBound(ByteString()) - MoveBlockStartPos + 1&
        End If
        Call CopyMemory(ByteString(CutStartPos), ByteString(MoveBlockStartPos), MoveBlockLength)
    Else
        'nothing to move
    End If
    For Temp = (ByteStringLength - CutLength + 1&) To ByteStringLength
        ByteString(Temp) = RightEndReplaceAsc 'reset
    Next Temp
End Sub

Public Sub BYTESTRINGINSERT(ByRef ByteString() As Byte, ByVal InsertStartPos As Long, ByVal InsertString As String, Optional ByVal ByteStringLengthFixedFlag As Boolean = True)
    'on error Resume Next 'inserts InsertString at given position (without overwriting original chars)
    Dim ByteStringLength As Long
    Dim MoveBlockStartPos As Long
    Dim MoveBlockTargetPos As Long 'new start pos
    Dim MoveBlockLength As Long
    'preset
    If ByteStringLengthFixedFlag = True Then
        ByteStringLength = GETBYTESTRINGLENGTH(ByteString())
    Else
        'NOTE: when wanting to append data to the end of the byte string
        'this function would originally return an error, now it doesn't any more.
        ByteStringLength = UBound(ByteString())
    End If
    If (InsertStartPos < 1) Or (InsertStartPos > ByteStringLength) Then 'verify
        MsgBox "internal error in BYTESTRINGINSERT(): passed value invalid !", vbOKOnly + vbExclamation
        Exit Sub
    End If
    MoveBlockStartPos = InsertStartPos
    MoveBlockTargetPos = (MoveBlockStartPos + Len(InsertString))
    MoveBlockLength = ByteStringLength - MoveBlockStartPos + 1&
    'begin
    If Not (MoveBlockTargetPos > ByteStringLength) Then
        If (MoveBlockTargetPos + MoveBlockLength - 1&) > UBound(ByteString()) Then 'verify
            MoveBlockLength = UBound(ByteString()) - MoveBlockTargetPos + 1&
        End If
        'NOTE: now the block is moved to the right side (in BYTESTRINGCUT() it is moved to the left side).
        Call CopyMemory(ByteString(MoveBlockTargetPos), ByteString(MoveBlockStartPos), MoveBlockLength)
    Else
        'nothing to move
    End If
    Call CopyMemory(ByteString(InsertStartPos), ByVal InsertString, BS_MIN(Len(InsertString), (ByteStringLength - InsertStartPos + 1&)))
End Sub

Public Sub BYTESTRINGINSERTByte(ByRef ByteString() As Byte, ByVal InsertStartPos As Long, ByRef InsertByteString() As Byte, Optional ByVal ByteStringLengthFixedFlag As Boolean = True)
    'on error Resume Next 'inserts InsertByteString() at given position (without overwriting original chars)
    Dim ByteStringLength As Long
    Dim InsertByteStringLength As Long
    Dim MoveBlockStartPos As Long
    Dim MoveBlockTargetPos As Long 'new start pos
    Dim MoveBlockLength As Long
    'preset
    If ByteStringLengthFixedFlag = True Then
        ByteStringLength = GETBYTESTRINGLENGTH(ByteString())
    Else
        'NOTE: when wanting to append data to the end of the byte string
        'this function would originally return an error, now it doesn't any more.
        ByteStringLength = UBound(ByteString())
    End If
    InsertByteStringLength = GETBYTESTRINGLENGTH(InsertByteString())
    If (InsertStartPos < 1&) Or (InsertStartPos > ByteStringLength) Then 'verify
        MsgBox "internal error in BYTESTRINGINSERT(): passed value invalid !", vbOKOnly + vbExclamation
        Exit Sub
    End If
    MoveBlockStartPos = InsertStartPos
    MoveBlockTargetPos = (MoveBlockStartPos + InsertByteStringLength)
    MoveBlockLength = ByteStringLength - MoveBlockStartPos + 1&
    'begin
    If Not (MoveBlockTargetPos > ByteStringLength) Then
        If (MoveBlockTargetPos + MoveBlockLength - 1&) > UBound(ByteString()) Then 'verify
            MoveBlockLength = UBound(ByteString()) - MoveBlockTargetPos + 1&
        End If
        'NOTE: now the block is moved to the right side (in BYTESTRINGCUT() it is moved to the left side).
        Call CopyMemory(ByteString(MoveBlockTargetPos), ByteString(MoveBlockStartPos), MoveBlockLength)
    Else
        'nothing to move
    End If
    Call CopyMemory(ByteString(InsertStartPos), InsertByteString(1), BS_MIN(InsertByteStringLength, (ByteStringLength - InsertStartPos + 1&)))
End Sub

'****************************END OF MANIPULATION FUNCTIONS*****************************
'************************************FILE FUNCTIONS************************************

Public Function GetDirectoryNameByte(ByRef PathByteString() As Byte, ByRef DirectoryByteString() As Byte) As Boolean
    'on error Resume Next 'does not size passed arrays; returns True for directory name transferred, False for error
    Dim Temp As Long
    For Temp = GETBYTESTRINGLENGTH(PathByteString()) To 1 Step (-1)
        If PathByteString(Temp) = 92 Then 'back slash will be included
            Call CopyMemory(DirectoryByteString(1), PathByteString(1), BS_MIN(Temp, UBound(DirectoryByteString())))
            GetDirectoryNameByte = True 'ok
            Exit Function
        End If
    Next Temp
    GetDirectoryNameByte = False 'error
    Exit Function
End Function

'********************************END OF FILE FUNCTIONS*********************************
'************************************BORDER STRINGS************************************

Public Function GetBorderedStringByte(ByRef MainString() As Byte, ByRef BorderedStringStartStringPassed() As Byte, ByRef BorderedStringEndStringPassed() As Byte, ByRef BorderedString() As Byte, ByVal ByteStringLength As Long) As Boolean
    'on error Resume Next 'BorderStrings are excluded; pass border string including Chr$(1) for border string located at string start and Chr$(255) for located at string end
    Dim BorderedStringStartPos As Long 'position in MainString
    Dim BorderedStringEndPos As Long 'position in MainString
    Dim BorderStringLengthMax As Long
    Dim Temp As Long
    '
    'NOTE: this string function has been manipulated to be used together with byte strings.
    'Function returns True if BorderedString() has been initialized, False if not.
    'NOTE: in the byte string version of this function Chr$(1) represents Chr$(0)
    'as there was the problem that the string end could not be determined any more.
    'NOTE: BorderedString[Start/End]String() must have the same maximal length.
    '
    'preset
    BorderStringLengthMax = UBound(BorderedStringStartStringPassed()) 'see notes above
    ReDim BorderedStringStartString(1 To BorderStringLengthMax) As Byte 'byte string will be changed
    ReDim BorderedStringEndString(1 To BorderStringLengthMax) As Byte 'byte string will be changed
    Call CopyMemory(BorderedStringStartString(1), BorderedStringStartStringPassed(1), BorderStringLengthMax)
    Call CopyMemory(BorderedStringEndString(1), BorderedStringEndStringPassed(1), BorderStringLengthMax)
    '
    For Temp = 1 To BorderStringLengthMax 'there may be only one 'special char' in every string
        If BorderedStringStartString(Temp) = 1 Then
            Call CopyMemory(BorderedStringStartString(Temp), BorderedStringStartString(Temp + 1), BorderStringLengthMax - Temp)
            BorderedStringStartString(BorderStringLengthMax) = 0 'reset
            If BorderedStringStartString(1) = 0 Then 'check first to increase speed
                BorderedStringStartPos = 1
            Else
                If BorderedStringStartString(2) = 0 Then
                    BorderedStringStartPos = 2
                Else
                    BorderedStringStartPos = 1 + GETBYTESTRINGLENGTH(BorderedStringStartString())
                End If
            End If
        End If
        If BorderedStringStartString(Temp) = 255 Then
            Call CopyMemory(BorderedStringStartString(Temp), BorderedStringStartString(Temp + 1), BorderStringLengthMax - Temp)
            BorderedStringStartString(BorderStringLengthMax) = 0 'reset
            If BorderedStringStartString(1) = 0 Then 'check first to increase speed
                BorderedStringStartPos = GETBYTESTRINGLENGTH(MainString())
            Else
                If BorderedStringStartString(2) = 0 Then
                    BorderedStringStartPos = GETBYTESTRINGLENGTH(MainString())
                Else
                    BorderedStringStartPos = GETBYTESTRINGLENGTH(MainString()) - GETBYTESTRINGLENGTH(BorderedStringStartString())
                End If
            End If
        End If
        If BorderedStringEndString(Temp) = 1 Then
            Call CopyMemory(BorderedStringEndString(Temp), BorderedStringEndString(Temp + 1), BorderStringLengthMax - Temp)
            BorderedStringEndString(BorderStringLengthMax) = 0 'reset
            If BorderedStringEndString(1) = 0 Then 'check first to increase speed
                BorderedStringEndPos = 1
            Else
                If BorderedStringEndString(2) = 0 Then
                    BorderedStringEndPos = 2
                Else
                    BorderedStringEndPos = 1 + GETBYTESTRINGLENGTH(BorderedStringEndString())
                End If
            End If
        End If
        If BorderedStringEndString(Temp) = 255 Then
            Call CopyMemory(BorderedStringEndString(Temp), BorderedStringEndString(Temp + 1), BorderStringLengthMax - Temp)
            BorderedStringEndString(BorderStringLengthMax) = 0 'reset
            If BorderedStringEndString(1) = 0 Then 'check first to increase speed
                BorderedStringEndPos = GETBYTESTRINGLENGTH(MainString())
            Else
                If BorderedStringEndString(2) = 0 Then
                    BorderedStringEndPos = GETBYTESTRINGLENGTH(MainString()) - 1
                Else
                    BorderedStringEndPos = GETBYTESTRINGLENGTH(MainString()) - GETBYTESTRINGLENGTH(BorderedStringEndString())
                End If
            End If
        End If
    Next Temp
    If BorderedStringStartPos = 0 Then 'if not set yet
        BorderedStringStartPos = InStrByte(1, MainString(), BorderedStringStartString(), vbBinaryCompare)
        If Not (BorderedStringStartPos = 0) Then 'verify
            BorderedStringStartPos = BorderedStringStartPos + 1
        Else
            Call BYTESTRINGLEFT(BorderedString(), 0) 'reset (error)
            GetBorderedStringByte = False 'error
            Exit Function
        End If
    End If
    If BorderedStringEndPos = 0 Then 'if not set yet
        BorderedStringEndPos = InStrByte(1, MainString(), BorderedStringEndString(), vbBinaryCompare)
        If Not (BorderedStringEndPos = 0) Then 'verify
            BorderedStringEndPos = BorderedStringEndPos - 1
        Else
            Call BYTESTRINGLEFT(BorderedString(), 0) 'reset (error)
            GetBorderedStringByte = False 'error
            Exit Function
        End If
    End If
    'begin
    If Not ((BorderedStringStartPos < 1) Or (BorderedStringEndPos < 1) Or _
        (BorderedStringStartPos > GETBYTESTRINGLENGTH(MainString())) Or (BorderedStringEndPos > GETBYTESTRINGLENGTH(MainString())) Or _
        (BorderedStringStartPos > BorderedStringEndPos)) Then 'verify
        Call CopyMemory(BorderedString(1), MainString(BorderedStringStartPos), BS_MIN((BorderedStringEndPos - BorderedStringStartPos + 1), UBound(BorderedString()))) 'ok
        Call BYTESTRINGLEFT(BorderedString(), (BorderedStringEndPos - BorderedStringStartPos + 1))
        GetBorderedStringByte = True 'ok
    Else
        Call BYTESTRINGLEFT(BorderedString(), 0) 'reset (error)
        GetBorderedStringByte = False 'error
    End If
End Function
    
Public Function GetBorderedStringByteEx(ByRef MainString() As Byte, ByRef BorderedStringStartStringPassed() As Byte, ByRef BorderedStringEndStringPassed() As Byte, ByRef BorderedString() As Byte, ByVal ByteStringLength As Long) As Boolean
    'on error Resume Next 'BorderStrings are excluded; pass border string including Chr$(1) for border string located at string start and Chr$(255) for located at string end
    Dim BorderedStringStartPos As Long 'position in MainString
    Dim BorderedStringStartStringIndex As Long 'how many start strings must have appeared before
    Dim BorderedStringEndPos As Long 'position in MainString
    Dim BorderedStringEndStringIndex As Long  'how many end strings must have appeared before
    Dim AsterixByteString(1 To 1) As Byte
    Dim StartPos As Long 'first '*'
    Dim EndPos As Long 'second '*'
    Dim MainStringLength As Long
    Dim BorderedStringLengthMax As Long
    Dim BorderStringLengthMax As Long
    Dim Temp As Long
    Dim TempByteString() As Byte
    '
    'NOTE: the start/end string may contain '*#*', i.e.
    'start string = '-*2*' end string = '-*3*' filters 'world' out of 'hello - you - world'.
    '
    'NOTE: this string function has been manipulated to be used together with byte strings.
    'Function returns True if BorderedString() has been initialized, False if not.
    'NOTE: if the bordered string is longer than the target byte string, the function
    'will add points to the target string to display it is not complete,
    'e.g. Always look on the Bride....
    'The length of the target string should not be smaller than 3.
    'NOTE: in the byte string version of this function Chr$(1) represents Chr$(0)
    'as there was the problem that the string end could not be determined any more.
    '
    'preset
    AsterixByteString(1) = 42
    BorderStringLengthMax = UBound(BorderedStringStartStringPassed()) 'see notes above
    ReDim BorderedStringStartString(1 To BorderStringLengthMax) As Byte 'byte string will be changed
    ReDim BorderedStringEndString(1 To BorderStringLengthMax) As Byte 'byte string will be changed
    Call CopyMemory(BorderedStringStartString(1), BorderedStringStartStringPassed(1), BorderStringLengthMax)
    Call CopyMemory(BorderedStringEndString(1), BorderedStringEndStringPassed(1), BorderStringLengthMax)
    '
    StartPos = InStrByte(1, BorderedStringStartString(), AsterixByteString(), vbBinaryCompare)
    If Not (StartPos = 0) Then
        EndPos = InStr(StartPos + 1, BorderedStringStartString(), AsterixByteString(), vbBinaryCompare)
        If Not (EndPos = 0) Then
            Call BYTESTRINGMID(TempByteString(), BorderedStringStartString(), StartPos + 1, EndPos - 1)
            BorderedStringStartStringIndex = BYTESTRINGVALLONG(TempByteString())
            If BorderedStringStartStringIndex < 1 Then BorderedStringStartStringIndex = 1 'verify
            Call BYTESTRINGCUT(BorderedStringStartString(), StartPos, EndPos - StartPos + 2, 0) 'cut out i.e. *2*
        Else
            BorderedStringStartStringIndex = 1
        End If
    Else
        BorderedStringStartStringIndex = 1
    End If
    StartPos = InStrByte(1, BorderedStringEndString(), AsterixByteString(), vbBinaryCompare)
    If Not (StartPos = 0) Then
        EndPos = InStr(StartPos + 1, BorderedStringEndString(), AsterixByteString(), vbBinaryCompare)
        If Not (EndPos = 0) Then
            Call BYTESTRINGMID(TempByteString(), BorderedStringEndString(), StartPos + 1, EndPos - 1)
            BorderedStringEndStringIndex = BYTESTRINGVALLONG(TempByteString())
            If BorderedStringEndStringIndex < 1 Then BorderedStringEndStringIndex = 1 'verify
            Call BYTESTRINGCUT(BorderedStringEndString(), StartPos, EndPos - StartPos + 2, 0) 'cut out i.e. *2*
        Else
            BorderedStringEndStringIndex = 1
        End If
    Else
        BorderedStringEndStringIndex = 1
    End If
    'preset (2)
    For Temp = 1 To BorderStringLengthMax 'there may be only one 'special char' in every string
        If BorderedStringStartString(Temp) = 1 Then
            Call CopyMemory(BorderedStringStartString(Temp), BorderedStringStartString(Temp + 1), BorderStringLengthMax - Temp)
            BorderedStringStartString(BorderStringLengthMax) = 0 'reset
            If BorderedStringStartString(1) = 0 Then 'check first to increase speed
                BorderedStringStartPos = 1
            Else
                If BorderedStringStartString(2) = 0 Then
                    BorderedStringStartPos = 2
                Else
                    BorderedStringStartPos = 1 + GETBYTESTRINGLENGTH(BorderedStringStartString())
                End If
            End If
        End If
        If BorderedStringStartString(Temp) = 255 Then
            Call CopyMemory(BorderedStringStartString(Temp), BorderedStringStartString(Temp + 1), BorderStringLengthMax - Temp)
            BorderedStringStartString(BorderStringLengthMax) = 0 'reset
            If BorderedStringStartString(1) = 0 Then 'check first to increase speed
                BorderedStringStartPos = GETBYTESTRINGLENGTH(MainString())
            Else
                If BorderedStringStartString(2) = 0 Then
                    BorderedStringStartPos = GETBYTESTRINGLENGTH(MainString())
                Else
                    BorderedStringStartPos = GETBYTESTRINGLENGTH(MainString()) - GETBYTESTRINGLENGTH(BorderedStringStartString())
                End If
            End If
        End If
        If BorderedStringEndString(Temp) = 1 Then
            Call CopyMemory(BorderedStringEndString(Temp), BorderedStringEndString(Temp + 1), BorderStringLengthMax - Temp)
            BorderedStringEndString(BorderStringLengthMax) = 0 'reset
            If BorderedStringEndString(1) = 0 Then 'check first to increase speed
                BorderedStringEndPos = 1
            Else
                If BorderedStringEndString(2) = 0 Then
                    BorderedStringEndPos = 2
                Else
                    BorderedStringEndPos = 1 + GETBYTESTRINGLENGTH(BorderedStringEndString())
                End If
            End If
        End If
        If BorderedStringEndString(Temp) = 255 Then
            Call CopyMemory(BorderedStringEndString(Temp), BorderedStringEndString(Temp + 1), BorderStringLengthMax - Temp)
            BorderedStringEndString(BorderStringLengthMax) = 0 'reset
            If BorderedStringEndString(1) = 0 Then 'check first to increase speed
                BorderedStringEndPos = GETBYTESTRINGLENGTH(MainString())
            Else
                If BorderedStringEndString(2) = 0 Then
                    BorderedStringEndPos = GETBYTESTRINGLENGTH(MainString()) - 1
                Else
                    BorderedStringEndPos = GETBYTESTRINGLENGTH(MainString()) - GETBYTESTRINGLENGTH(BorderedStringEndString())
                End If
            End If
        End If
    Next Temp
    'begin
    If BorderedStringStartPos = 0 Then 'if not set yet
        For Temp = 1 To BorderedStringStartStringIndex
            BorderedStringStartPos = InStrByte(BorderedStringStartPos + 1, MainString(), BorderedStringStartString(), vbBinaryCompare)
        Next Temp
        If Not (BorderedStringStartPos = 0) Then 'verify
            BorderedStringStartPos = BorderedStringStartPos + GETBYTESTRINGLENGTH(BorderedStringStartString())
        Else
            Call BYTESTRINGLEFT(BorderedString(), 0) 'reset (error)
            GetBorderedStringByteEx = False 'error
            Exit Function
        End If
    End If
    If BorderedStringEndPos = 0 Then 'if not set yet
        For Temp = 1 To BorderedStringEndStringIndex
            BorderedStringEndPos = InStrByte(BorderedStringEndPos + 1, MainString(), BorderedStringEndString(), vbBinaryCompare)
        Next Temp
        If Not (BorderedStringEndPos = 0) Then 'verify
            BorderedStringEndPos = BorderedStringEndPos - 1
        Else
            Call BYTESTRINGLEFT(BorderedString(), 0) 'reset (error)
            GetBorderedStringByteEx = False 'error
            Exit Function
        End If
    End If
    'begin
    MainStringLength = GETBYTESTRINGLENGTH(MainString())
    If Not ((BorderedStringStartPos < 1) Or (BorderedStringEndPos < 1) Or _
        (BorderedStringStartPos > MainStringLength) Or (BorderedStringEndPos > MainStringLength) Or _
        (BorderedStringStartPos > BorderedStringEndPos)) Then 'verify
        If (BorderedStringEndPos - BorderedStringStartPos + 1) > UBound(BorderedString()) Then
            BorderedStringLengthMax = UBound(BorderedString()) 'do here to increase speed (not done that often)
            Call CopyMemory(BorderedString(1), MainString(BorderedStringStartPos), BorderedStringLengthMax - 3) 'ok
            BorderedString(BorderedStringLengthMax - 2) = 46 '.
            BorderedString(BorderedStringLengthMax - 1) = 46 '.
            BorderedString(BorderedStringLengthMax) = 46 '.
        Else
            Call CopyMemory(BorderedString(1), MainString(BorderedStringStartPos), BS_MIN((BorderedStringEndPos - BorderedStringStartPos + 1), UBound(BorderedString()))) 'ok
            Call BYTESTRINGLEFT(BorderedString(), (BorderedStringEndPos - BorderedStringStartPos + 1))
        End If
        GetBorderedStringByteEx = True 'ok
    Else
        Call BYTESTRINGLEFT(BorderedString(), 0) 'reset (error)
        GetBorderedStringByteEx = False 'error
    End If
End Function

'********************************END OF BORDER STRINGS*********************************
'****************************************OTHER*****************************************

'NOTE: to avoid name conflicts MIN() is called BS_MIN() (ByteString_MIN()) here.
'In general, all public functions that are to be used by multiple projects must have
'a definite name, add the 'sub system prefix' to create definite names.

Public Function BS_MIN(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    'on error Resume Next 'use for i.e. CopyMemory(a(1), ByVal b, BS_MIN(UBound(a()), Len(b))
    If Value1 < Value2 Then
        BS_MIN = Value1
    Else
        BS_MIN = Value2
    End If
End Function

Public Function BS_MAX(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    'on error Resume Next 'use in combination with ReDim()
    If Value1 > Value2 Then
        BS_MAX = Value1
    Else
        BS_MAX = Value2
    End If
End Function

