Attribute VB_Name = "Global3"
Option Explicit
'-------------------------------
' Constants for Token processing
'-------------------------------
Public Const TK_none = 0
Public Const TK_uint = 1
Public Const TK_str = 2
Public Const TK_dword = 3
Public Const TK_float = 4
Public Const TK_uint4float = 5
Public Const TK_2sint3float = 6
Public Const TK_sint = 7
Public Const TK_2uint2float = 8
Public Const TK_uintfloat = 9
Public Const TK_dworduint = 10
Public Const TK_tokuintfloat = 11
Public Const TK_uintnocr = 21
Public Const TK_literal = 0
Public Const TK_level = 1
Public Const TK_embedded = 2
Public Const TK_type = 3
Public Const TK_count = 4
Public Const TK_precis = 5
Public Const TK_embedded_yes = 1
Public Const TK_embedded_no = 0

Public fso As FileSystemObject
Public Param() As String
Public ExpectedFileMask As String
Public SourceFile As String

Public SourceOffset As Long      'source offset
Public DebugMode As Boolean   'debugmode
Public QuietMode As Boolean   'quietmode
Public TargetFileHandle As Long      'thdl
Public TabDepth As Long      'workcl
Public PreviousToken As Long      'Prev Token
Public LoggerFileHandle As Long      'lhdl
Public LoggingMode As Boolean   'logit
Public PreviousOffset As Long      'prev offset

Public Tokens() As String
Public Home As String
Public OtherHome As String

Public TokenTarget() As Byte

'----------------------------------------------
' Remap 2 bytes into integer
'----------------------------------------------
Public Function Peek2(buf() As Byte, start As Long) As Integer
Dim I As Integer
    I = 0
    CopyMemory ByVal VarPtr(I), buf(start), 2
    Peek2 = I
End Function

'----------------------------------------------
' ReMap four bytes into Long
'----------------------------------------------
Public Function Peek4(buf() As Byte, start As Long) As Long
Dim L As Long
    L = 0
    CopyMemory ByVal VarPtr(L), buf(start), 4
    Peek4 = L
End Function

'--------------------------------------------
' ReMap 4 bytes into Single Float
'--------------------------------------------
Public Function Peek4FL(buf() As Byte, start As Long) As Single
Dim F As Single
    F = 0#
    CopyMemory ByVal VarPtr(F), buf(start), 4
    Peek4FL = F
End Function

'---------------------------------------------
' Recursive routine to expand consecutive tokens
'---------------------------------------------
Public Function DoSomeTags(sourcesize As Long) As Boolean
Dim ret As Boolean
Dim Token As Long
Dim preamble As String
Dim outline As String
Dim TokenName As String
Dim nexttokenoffset As Long
Dim tempnexttoken As Long
Dim nexttoken As Long
Dim valuetype As Long
Dim valuecount As Long
Dim dplaces As Long
On Error GoTo errtrap
    Do While SourceOffset < sourcesize
        Token = Peek2(TokenTarget, SourceOffset)    ' current token numeric value
        TokenName = Tokens(TK_literal, Token)       ' current token name
        If TokenName = "" Then
            If Not QuietMode Then
                MsgBox "Unknown token found at:" & vbCrLf & "Offset=" & CStr(SourceOffset) & vbCrLf & "Token=" & CStr(Token) & _
                        vbCrLf & "Previous Token at:" & vbCrLf & "Offset=" & CStr(PreviousOffset) & vbCrLf & "Token=" & CStr(PreviousToken), _
                    vbOKOnly, _
                    "Error"
            End If
            DoSomeTags = False  ' Pull out of all recursions
            Exit Function
        End If
        PreviousToken = Token   ' save incase next token is bad
        PreviousOffset = SourceOffset   ' save incase next token is bad
        
        If TabDepth = 1 Then    ' do for primary tokens only
            If DebugMode Then If LoggingMode Then Logger ("Starting on token " & TokenName)
        End If
        
        preamble = String(TabDepth, vbTab)
        outline = preamble & TokenName & " ( "  ' all tokens start this way
        nexttokenoffset = Peek4(TokenTarget, SourceOffset + 4) - 1 ' use its length to see where next token is at this level
        SourceOffset = SourceOffset + 9   'skip  header to token parms etc
        nexttoken = nexttokenoffset + SourceOffset  ' add token offset to file offset to get next real token location
        
        If Tokens(TK_type, Token) <> 0 Then     ' do only for registered types cased below
            valuetype = Tokens(TK_type, Token)
            valuecount = Tokens(TK_count, Token)    'mostly for number of 4 byte parms
            dplaces = Tokens(TK_precis, Token)  ' mostly for float decimal places (ignored in practice)
            
            Select Case valuetype
            
                Case TK_uint            ' unsigned integer values
                    If dplaces = 1 Then
                        If nexttokenoffset = 0 Then
                            outline = outline & ")"
                            FileWrite TargetFileHandle, outline, True
                            GoTo Ecase1
                        End If
                    End If
                    outline = Do_UINT(Token, outline)
                    
                Case TK_uintnocr        ' unsigned integers and NO cr at end of line
                    outline = Do_UINTNOCR(Token, outline)
                    
                Case TK_sint            ' signed integer values
                    outline = Do_SINT(Token, outline)
                    
                Case TK_uint4float      ' one UINt and several float values
                    outline = Do_UINT4FLOAT(Token, outline)
                    
                Case TK_2uint2float     ' couple of uint and several float values
                    outline = Do_2UINT2FLOAT(Token, outline)
                    
                Case TK_dworduint       ' a dword (bit mask) and several uint
                    outline = Do_DWORDUINT(Token, outline)
                    
                Case TK_uintfloat       ' a uint and a single float
                    outline = Do_UINTFLOAT(Token, outline, True)
                    
                Case TK_tokuintfloat    ' an embedded token(s) and a uint and single float
                    FileWrite TargetFileHandle, outline, True
                    tempnexttoken = Peek4(TokenTarget, SourceOffset + 4) + 8 + SourceOffset
                    TabDepth = TabDepth + 1
                    ret = DoSomeTags(tempnexttoken)
                    If Not ret Then
                        DoSomeTags = False
                        Exit Function
                    End If
                    TabDepth = TabDepth - 1
                    outline = ""
                    outline = Do_UINTFLOAT(Token, outline, False)
                    outline = outline & vbCrLf & preamble & ")"
                    FileWrite TargetFileHandle, outline, True
                    
                Case TK_2sint3float     ' some sint and several float
                    outline = Do_2SINT3FLOAT(Token, outline)
                    
                Case TK_str             ' a variable length unicode string
                    outline = Do_STR(Token, outline)
                    
                Case TK_dword           ' a dword bit mask
                    outline = Do_DWORD(Token, outline)
                    
                Case TK_float           ' one or more float values
                    outline = Do_FLOAT(Token, outline)
                    
            End Select
Ecase1:
        End If  ' end of registered token process
        
        If Tokens(TK_embedded, Token) = TK_embedded_yes Then    ' did it have any embedded tokens?
            FileWrite TargetFileHandle, outline, True           ' output the collected line so far
            outline = ""
            TabDepth = TabDepth + 1     ' increase the token depth
            ret = DoSomeTags(nexttoken) ' recurse to process embeddid tokens
            If Not ret Then             ' hopefully all were registered
                DoSomeTags = False
                Exit Function
            End If
            TabDepth = TabDepth - 1     ' drop back to current level
            FileWrite TargetFileHandle, preamble & ")", True    ' write out the end of this levels token
        End If
        
        If TabDepth = 1 Then
            If DebugMode Then If LoggingMode Then Logger ("Finished token " & TokenName)
        End If
        
    Loop    ' keep processing all embedded tokens in this length of the file.
    DoSomeTags = True
    
    Exit Function
    
    
errtrap:
Call MsgBox("A fatal error occurred  while uncompressing a .w file  - Error # " & Err.Number _
            & vbCrLf & "Error description: " & Err.Description _
            , vbCritical, App.Title)

    
    
End Function

Public Function Do_UINT(Token As Long, outline As String) As String
Dim valuecount As Long
    valuecount = Tokens(TK_count, Token)    'pull number of uints expected
    outline = UINTparm(valuecount, outline) ' put them into the output string
    If Tokens(TK_embedded, Token) <> TK_embedded_yes Then ' if not embedded end the token (parens) and output the info
        outline = outline & ")"
        FileWrite TargetFileHandle, outline, True
        outline = ""
    End If
    Do_UINT = outline
End Function

Public Function UINTparm(valuecount As Long, outline As String) As String
Dim I As Long
Dim value As Long
    ' assemble all values into output line
    For I = 1 To valuecount
        value = Peek4(TokenTarget, SourceOffset)
        'consider values larger than 4294967296
        SourceOffset = SourceOffset + 4
        outline = outline & CStr(value) & " "
    Next
    UINTparm = outline
End Function

Public Function Do_UINTNOCR(Token As Long, outline As String) As String
Dim valuecount As Long
    valuecount = Tokens(TK_count, Token)
    outline = UINTparm(valuecount, outline)
    If Tokens(TK_embedded, Token) <> TK_embedded_yes Then
        outline = outline & ") "
        FileWriteNoCR TargetFileHandle, outline, True
        outline = ""
    End If
    Do_UINTNOCR = outline
End Function

Public Function Do_SINT(Token As Long, outline As String) As String
Dim valuecount As Long
    valuecount = Tokens(TK_count, Token)
    outline = SINTparm(valuecount, outline)
    If Tokens(TK_embedded, Token) <> TK_embedded_yes Then
        outline = outline & ")"
        FileWrite TargetFileHandle, outline, True
        Do_SINT = ""
        Exit Function
    End If
    Do_SINT = outline
End Function

Public Function SINTparm(valuecount As Long, outline As String) As String
Dim I As Long
Dim value As Long
    For I = 1 To valuecount
        value = Peek4(TokenTarget, SourceOffset)
        SourceOffset = SourceOffset + 4
        outline = outline & value & " "
    Next
    SINTparm = outline
End Function

Public Function Do_UINT4FLOAT(Token As Long, outline As String) As String
Dim valuecount As Long
Dim dplaces As Long
Dim value As Long
Dim gcnt As Long
Dim j As Long

    valuecount = Tokens(TK_count, Token)
    dplaces = Tokens(TK_precis, Token)
    value = Peek4(TokenTarget, SourceOffset)
    gcnt = value
    'consider values larger than 4294967296
    SourceOffset = SourceOffset + 4
    outline = outline & CStr(value) & " "
    For j = 1 To gcnt
        outline = FLOATparm(dplaces, outline, 3)
    Next
    If Tokens(TK_embedded, Token) <> TK_embedded_yes Then
        outline = outline & ")"
        FileWrite TargetFileHandle, outline, True
        Do_UINT4FLOAT = ""
        Exit Function
    End If
    Do_UINT4FLOAT = outline
End Function

Public Function FLOATparm(valuecount As Long, outline As String, dplaces As Long) As String
Dim I As Long
Dim value As Single
    For I = 1 To valuecount
        value = Peek4FL(TokenTarget, SourceOffset)
        SourceOffset = SourceOffset + 4
        outline = outline & CStr(value) & " "
    Next
    FLOATparm = outline
End Function

Public Function Do_2UINT2FLOAT(Token As Long, outline As String) As String
Dim valuecount As Long
Dim dplaces As Long

    valuecount = Tokens(TK_count, Token)
    dplaces = Tokens(TK_precis, Token)
    outline = UINTparm(1, outline)
    FileWriteNoCR TargetFileHandle, outline, True
    outline = " "
    outline = UINTparm(valuecount - 1, outline)
    outline = FLOATparm(dplaces, outline, 3)
    FileWrite TargetFileHandle, outline, True
    Do_2UINT2FLOAT = ""
End Function

Public Function Do_DWORDUINT(Token As Long, outline As String) As String
Dim valuecount As Long

    valuecount = Tokens(TK_count, Token)
    outline = DWORDparm(valuecount, outline)
    valuecount = Tokens(TK_precis, Token)
    outline = UINTparm(valuecount, outline)
    If Tokens(TK_embedded, Token) <> TK_embedded_yes Then
        outline = outline & ")"
        FileWrite TargetFileHandle, outline, True
        outline = ""
    End If
    Do_DWORDUINT = outline
End Function

Public Function DWORDparm(valuecount As Long, outline As String) As String
Dim list As String
Dim I As Long
Dim j As Long
Dim fullbyte As Long
Dim byte1 As Long
Dim byte2 As Long

    list = "0123456789abcdef"
    For I = 1 To valuecount
        For j = 3 To 0 Step -1
            fullbyte = TokenTarget(SourceOffset + j)
            byte1 = Int(fullbyte / 16)
            byte2 = fullbyte - (byte1 * 16)
            outline = outline & Mid(list, byte1 + 1, 1) & Mid(list, byte2 + 1, 1)
        Next
        SourceOffset = SourceOffset + 4
        outline = outline & " "
    Next
    DWORDparm = outline
End Function

Public Function Do_UINTFLOAT(Token As Long, outline As String, cls As Boolean) As String
Dim valuecount As Long
Dim dplaces As Long

    valuecount = Tokens(TK_count, Token)
    dplaces = Tokens(TK_precis, Token)
    outline = UINTparm(valuecount, outline)
    outline = FLOATparm(dplaces, outline, 3)
    If Tokens(TK_embedded, Token) <> TK_embedded_yes Then
        If cls Then
            outline = outline & ")"
            FileWrite TargetFileHandle, outline, True
        Else
            FileWriteNoCR TargetFileHandle, outline, True
        End If
        outline = ""
    End If
    Do_UINTFLOAT = outline
End Function

Public Function Do_2SINT3FLOAT(Token As Long, outline As String) As String
Dim valuecount As Long
Dim dplaces As Long

    valuecount = Tokens(TK_count, Token)
    dplaces = Tokens(TK_precis, Token)
    outline = SINTparm(valuecount, outline)
    outline = FLOATparm(dplaces, outline, 3)
    If Tokens(TK_embedded, Token) <> TK_embedded_yes Then
        outline = outline & ")"
        FileWrite TargetFileHandle, outline, True
        outline = ""
    End If
    Do_2SINT3FLOAT = outline
End Function

Public Function Do_STR(Token As Long, outline As String) As String
Dim valuecount As Long

    valuecount = Tokens(TK_count, Token)
    outline = STRparm(valuecount, outline)
    If Tokens(TK_embedded, Token) <> TK_embedded_yes Then
        outline = outline & ")"
        FileWrite TargetFileHandle, outline, True
        outline = ""
    End If
    Do_STR = outline

End Function

Public Function STRparm(valuecount As Long, outline As String) As String
Dim I As Long
Dim size As Long
Dim str As String

    For I = 1 To valuecount
        size = Peek2(TokenTarget, SourceOffset)     ' get the string size in unicode characters
        SourceOffset = SourceOffset + 2     ' skip over to start of string
        str = String(size + 1, " ")         ' preallocated a target buffer
        CopyMemory ByVal StrPtr(str), TokenTarget(SourceOffset), size * 2   ' assemble into output string type
        outline = outline & str & " "       ' then drop into output line
        SourceOffset = SourceOffset + (size * 2)    ' account for size in bytes.
    Next
    STRparm = outline
End Function

Public Function Do_DWORD(Token As Long, outline As String) As String
Dim valuecount As Long

    valuecount = Tokens(TK_count, Token)
    outline = DWORDparm(valuecount, outline)
    If Tokens(TK_embedded, Token) <> TK_embedded_yes Then
        outline = outline & ")"
        FileWrite TargetFileHandle, outline, True
        outline = ""
    End If
    Do_DWORD = outline

End Function

Public Function Do_FLOAT(Token As Long, outline As String) As String
Dim valuecount As Long
Dim dplaces As Long
    
    valuecount = Tokens(TK_count, Token)
    dplaces = Tokens(TK_precis, Token)
    outline = FLOATparm(valuecount, outline, dplaces)
    If Tokens(TK_embedded, Token) <> TK_embedded_yes Then
        outline = outline & ")"
        FileWrite TargetFileHandle, outline, True
        outline = ""
    End If
    Do_FLOAT = outline

End Function

'--------------------------------------------------------------
' Uncompress source file using zlib dll
'--------------------------------------------------------------
Public Function InflateTile(src As String, tok As String) As Long
Dim F As Long
Dim prefix As String
Dim targetsize As Long
Dim sourcesize As Long
Dim source() As Byte
Dim rslt As Long


    prefix = String(8, " ")
    F = FreeFile(1)
    Open src For Binary Access Read As #F
    Get #F, 1, prefix
    
    If prefix = "SIMISA@@" Then
        Close #F
        InflateTile = -98
        Exit Function
    End If
    
    If prefix <> "SIMISA@F" Then
        Close #F
        InflateTile = -99
        Exit Function
    End If
    
    Get #F, 9, targetsize
    sourcesize = LOF(F) - 16    ' ignore file prefix
    ReDim source(1 To sourcesize) As Byte
    Get #F, 17, source
    Close #F
    
    ReDim TokenTarget(1 To targetsize) As Byte
    rslt = Uncompress(TokenTarget(1), targetsize, source(1), sourcesize)
    InflateTile = rslt
End Function

Public Sub InitValues()
    Set fso = CreateObject("Scripting.FileSystemObject")
    DebugMode = False
    QuietMode = True
    TargetFileHandle = 0
    TabDepth = 0
    PreviousToken = 0
    LoggerFileHandle = 0
    LoggingMode = True
    ExpectedFileMask = ""
    Home = CurDir
    OtherHome = App.path
    SourceFile = ""
    
End Sub

'-------------------------------------------------
' seperate command line into parameters and
' store them into Param array
'-------------------------------------------------
Public Sub PullParms(tmp As String)
Dim pcnt As Long
Dim pstart As Long
Dim L As Long
Dim I As Long
Dim MatchQuote As Boolean
    
    tmp = Trim(tmp) & " "   ' clean up parameters
    Do While InStr(1, tmp, "  ") <> 0
        tmp = Replace(tmp, "  ", " ")
    Loop
    
    L = Len(tmp)
    MatchQuote = False  ' allow for quoted file names
    ReDim Preserve Param(0 To 0)
    pstart = 1
    pcnt = 0
    For I = 1 To L
        If MatchQuote Then
            If Mid(tmp, I, 1) = Chr(34) Then
                MatchQuote = False
            End If
        ElseIf Mid(tmp, I, 1) = " " Then
            If pstart <> I Then
                pcnt = pcnt + 1
                ReDim Preserve Param(0 To pcnt)
                Param(pcnt) = Mid(tmp, pstart, I - pstart)
            End If
            pstart = I + 1
        ElseIf Mid(tmp, I, 1) = Chr(34) Then
            MatchQuote = True
        End If
    Next
    If pstart < L Then
        pcnt = pcnt + 1
        ReDim Preserve Param(0 To pcnt)
        If MatchQuote Then
            Param(pcnt) = Mid(tmp, pstart) & Chr(34)
        Else
            Param(pcnt) = Mid(tmp, pstart)
        End If
    End If
    Param(0) = CStr(pcnt)
End Sub

Public Sub ParseParms()
Dim SourceParmFound As Boolean
Dim TargetParmFound As Boolean

    SourceParmFound = False
    TargetParmFound = False
    QuietMode = False
    DebugMode = False
    
    If CLng(Param(0)) > 0 Then  ' at least one parameter
        If Mid(Param(1), 1, 1) = "/" Then   ' is it a switch
            If LCase(Mid(Param(1), 2, 1)) = "d" Then    ' is it the debug switch
                DebugMode = True
            End If
            If LCase(Mid(Param(1), 2, 1)) = "?" Then    ' is it the help switch
          '      Call ShowHelp          ' display a help panel (future)
                End                     ' and exit program
            End If
        Else        ' not a switch
            DebugMode = False
            SourceFile = Param(1)   ' assume it is the source file, but make sure it exists
            If Not fso.FileExists(SourceFile) Then
                End
            Else
                SourceParmFound = True
            End If
        End If
    End If
    
    If CLng(Param(0)) > 1 Then      ' at least two parameters
        If SourceParmFound Then
            ExpectedFileMask = Param(2)
            If FilePath(ExpectedFileMask) = "" Then ExpectedFileMask = FilePath(SourceFile) & ExpectedFileMask
            If FileRoot(ExpectedFileMask) = "" Then ExpectedFileMask = FilePath(ExpectedFileMask) & "*." & FileExtension(ExpectedFileMask)
            TargetParmFound = True
        Else
            SourceFile = Param(2)   ' assume it is the source file, but make sure it exists
            If Not fso.FileExists(SourceFile) Then
                End
            Else
                SourceParmFound = True
            End If
        End If
    End If
    
    If SourceParmFound And TargetParmFound Then
        QuietMode = True
    End If
    
    If DebugMode And Not QuietMode Then
        LoggingMode = True
    End If
    
    If CLng(Param(0)) > 2 Then  ' at least 3 parms
        If TargetParmFound Then
            If Not QuietMode Then MsgBox "Unidentified parameter " & Param(3), vbOKOnly, "Error"
            End
        Else
            ExpectedFileMask = Param(3)
            If FilePath(ExpectedFileMask) = "" Then ExpectedFileMask = FilePath(SourceFile) & ExpectedFileMask
            If FileRoot(ExpectedFileMask) = "" Then ExpectedFileMask = FilePath(ExpectedFileMask) & "*." & FileExtension(ExpectedFileMask)
            TargetParmFound = True
        End If
    End If
    
    If SourceParmFound And TargetParmFound Then
        QuietMode = True
    End If
    
    If DebugMode And Not QuietMode Then
        LoggingMode = True
    End If

    If CLng(Param(0)) > 3 Then
        If Not QuietMode Then MsgBox "Unidentified parameter " & Param(4), vbOKOnly, "Error"
        End
    End If
    

End Sub

Public Function FileOpen(filename As String, inout As String) As Long
Dim hdl As Long

    hdl = FreeFile(0)
    If StrComp("Write", inout, vbTextCompare) = 0 Then
        Open filename For Output As #hdl
    Else
        Open filename For Input As #hdl
    End If
    FileOpen = hdl
End Function

Public Function FilePath(str As String) As String
'-----------------------------------
'Extract the file path portion including trailing \
'-----------------------------------
Dim L As Long
Dim ptr As Long

    L = Len(str)
    If L < 1 Then
        FilePath = ""
        Exit Function
    End If
    ptr = InStrRev(str, "\", -1)
    If ptr = 0 Then
        FilePath = ""
        Exit Function
    End If
    FilePath = Mid(str, 1, ptr)
End Function

Public Function FileTitle(str As String) As String
'--------------------------------------------
'Extract the full filename without any path
'--------------------------------------------
Dim L As Long
Dim ptr As Long

    L = Len(str)
    If L < 1 Then
        FileTitle = ""
        Exit Function
    End If
    ptr = InStrRev(str, "\", -1)
    If ptr = 0 Then
        FileTitle = str
        Exit Function
    End If
    FileTitle = Mid(str, ptr + 1)
End Function

Public Function FileRoot(str As String) As String
'-------------------------------------------
'Extract the filename without the extension
'-------------------------------------------
Dim tmp As String
Dim ptr As Long

    tmp = FileTitle(str)
    ptr = InStr(1, tmp, ".")
    If ptr = 0 Then
        FileRoot = tmp
    Else
        FileRoot = Mid(tmp, 1, ptr - 1)
    End If
End Function

Public Function FileExtension(str As String) As String
'-----------------------------------------
'Extract just the file extension value
'-----------------------------------------
Dim tmp As String
Dim ptr As Long

    tmp = FileTitle(str)
    ptr = InStr(1, tmp, ".")
    If ptr = 0 Then
        FileExtension = ""
    Else
        FileExtension = Mid(tmp, ptr + 1)
    End If
End Function



Public Sub Logger(str As String)
    
    If Not LoggingMode Then Exit Sub
    FileWrite LoggerFileHandle, Now & vbTab & str, False
End Sub

Public Sub FileWrite(hdl As Long, line As String, uni As Boolean)

If Mid(line, Len(line), 1) = " " Then line = Mid(line, 1, Len(line) - 1)

    If uni Then
        Print #hdl, StrConv(line & vbCrLf, vbUnicode);
    Else
        Print #hdl, line
    End If
End Sub

Public Sub FileWriteNoCR(hdl As Long, line As String, uni As Boolean)
    If uni Then
        Print #hdl, StrConv(line, vbUnicode);
    Else
        Print #hdl, line;
    End If
End Sub

Public Function FileRead(hdl As Long) As String
Dim line As String

    Line Input #hdl, line
    If EOF(hdl) Then
        FileRead = "*EOF*"
    Else
        FileRead = line
    End If
End Function

Public Sub FileClose(hdl As Long)
    Close #hdl
End Sub

Public Function ItemExtract(position As Long, list As String, sep As String) As String
'---------------------------------------------
' Used to make a copy of a section of string data
' that is delimited by given seperators
'---------------------------------------------
Dim I As Long
Dim L As Long
Dim S As Long
Dim ptr As Long
Dim eptr As Long

    If position < 1 Then
        ItemExtract = ""
        Exit Function
    End If
    S = Len(sep)
    If S < 1 Then
        ItemExtract = ""
        Exit Function
    End If
    L = Len(list)
    If L < 1 Then
        ItemExtract = ""
        Exit Function
    End If
    eptr = 0
    For I = 1 To position
        ptr = eptr + Len(sep)
        eptr = InStr(ptr, list, sep, vbTextCompare)
        If eptr = 0 Then
            If I = position Then
                ItemExtract = Mid(list, ptr)
                Exit Function
            End If
            ItemExtract = ""
            Exit Function
        End If
    Next
    ItemExtract = Mid(list, ptr, eptr - ptr)
End Function

Public Function ItemCount(list As String, sep As String) As Long
'--------------------------------------------
' See how many delimited strings exist
' by definition this is always one more
' than the number of seperators found
'--------------------------------------------
Dim I As Long
Dim ptr As Long
    
    I = 1
    ptr = 1
    Do
        ptr = InStr(ptr, list, sep, vbTextCompare)
        If ptr = 0 Then
            ItemCount = I
            Exit Function
        End If
        I = I + 1
        ptr = ptr + Len(sep)
    Loop
End Function

