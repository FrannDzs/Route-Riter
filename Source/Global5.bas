Attribute VB_Name = "Global5"
Option Explicit
Option Compare Text


Public dbgitemcount As Long
Public dbgnextitem As Long
Public dbgfindfull As Long
Public dbgdonext As Long
Public cSourceFile() As Byte
Public SourceFileHandle As Long
Public SourceIdx As Long
Public FileType As String
Public ParensDepth As Long
Public TargetOffset As Long

Public Function CNAnum(parmidx As Long) As Single
Dim str As String
        str = NextItem(1, parmidx, " ")
        str = Replace(str, ",", ".")
        str = Replace(str, "'", "")
        CNAnum = CSng(Val(str))
End Function
Public Sub CompressMe(tFile As String, tfile2 As String, booDidNotComp As Boolean)
Dim size As Long
Dim ret As Boolean
Dim CompBuf() As Byte
Dim Sourcehdr As String
Dim targetsize As Long
Dim i As Long
Dim rslt As Long
On Error GoTo Errtrap

Call InitValues2

Source = tFile
Target = tfile2


If Right$(tFile, 1) = "s" Then
kind = S_Type
ElseIf Right$(tFile, 1) = "t" Then
kind = T_Type
ElseIf Right$(tFile, 1) = "w" Then
kind = W_Type
Else
strReport = strReport & tFile & " is an invalid compressed file type." & vbCrLf
booDidNotComp = True
Exit Sub
End If
  
    'access the input file - ASSUME it is unicode text for now
    SourceFileHandle = FileOpen2(Source, "Read")
    'need to know its exact size in bytes (not characters)
    size = LOF(SourceFileHandle)
    ReDim cSourceFile(1 To size)
    'read whole file in and then close file
    Get #SourceFileHandle, 1, cSourceFile
    Close #SourceFileHandle
    'cheater trick to move contents into string intact
    SourceFile = cSourceFile
   
    'Lets see if this is a source file
    If ((cSourceFile(1) <> 255) Or (cSourceFile(2) <> 254)) Then
    strReport = strReport & tFile & " Source file is NOT a unicode MSTS file." & vbCrLf
    booDidNotComp = True
      Exit Sub
    End If
    If Mid$(SourceFile, 18, 4) <> "JINX" Then
    strReport = strReport & tFile & " Source file is NOT an un-compressed MSTS file." & vbCrLf
            booDidNotComp = True
            Exit Sub
    End If
    If Mid$(SourceFile, 25, 1) <> "t" Then
    strReport = strReport & tFile & " Source file is NOT an un-compressed unicode MSTS file." & vbCrLf
        booDidNotComp = True
        Exit Sub
    End If
    
    If kind = T_Type Then

        If Mid$(SourceFile, 24, 1) <> "x" Then
        Call MsgBox("Terrain file was NOT created by the DeCompiler.", vbExclamation, App.Title)
            booDidNotComp = True
            Exit Sub
        End If
    End If
    
'    SaveSetting App.ExeName, "Settings", "Source", Source
'    SaveSetting App.ExeName, "Settings", "Target", Target
'    SaveSetting App.ExeName, "Settings", "Kind", CStr(kind)
'
    'Freeup the memory
    ReDim cSourceFile(1 To 1)
    'Save off the original msts header
    Sourcehdr = Mid$(SourceFile, 1, 32)
    'then strip it from source file string
    SourceFile = Mid$(SourceFile, InStr(1, SourceFile, vbCrLf) + 2)
    'Make string into one very long line of text with only single spaces as seperators
    SourceFile = Replace(SourceFile, vbCrLf, " ")
    SourceFile = Replace(SourceFile, vbTab, " ")
    SourceFile = Trim$(SourceFile) & " "
    Do While (InStr(1, SourceFile, "  ") <> 0)
        SourceFile = Replace(SourceFile, "  ", " ")
    Loop
    
    ReDim cSourceFile(1 To Len(SourceFile))
    For i = 1 To Len(SourceFile)
        cSourceFile(i) = Asc(Mid$(SourceFile, i, 1))
    Next
   ' Stop
    'Move original header over from unicode to ansi buffer
    For i = 1 To Len(Sourcehdr) - 1
        TokenTarget(i) = Asc(Mid$(Sourcehdr, i + 1, 1))  '********************************
    Next
    'Tag it as a binary version
    TokenTarget(24) = Asc("b")
    'Replace the missing linefeed
    TokenTarget(32) = 10
    
    'go compile the input string to the target buffer
    ret = DoNextToken(1, Len(SourceFile))
    If ret = False Then
    Call MsgBox(Lang(393) & tFile _
                & vbCrLf & "failed to comress properly." _
                , vbCritical, App.Title)
    booDidNotComp = True
    Exit Sub
    End If
    'If the output is just a compiled token file,then the extension is .Token
    'If the output is also compressed, the extension is .Compressed
    
  
    TargetFileHandle = FileOpen2(Target, "Write")
    'Make output buffer exact size
    ReDim Preserve TokenTarget(1 To TargetOffset - 1)
    
    If kind = T_Type Then
        TokenTarget(23) = Asc("6")
    End If
    If ((kind = W_Type) Or (kind = S_Type)) Then
        'set initial compress buffer plenty big enough
        targetsize = 12 + CLng(TargetOffset * 1.1)
        ReDim CompBuf(1 To targetsize)
        'Dont include first 16 bytes in compression size
        TargetOffset = TargetOffset - 17
        'start compression after first 16 bytes
        rslt = compress(CompBuf(17), VarPtr(targetsize), TokenTarget(17), TargetOffset)
        'replicate the original header
        For i = 1 To 16
            CompBuf(i) = TokenTarget(i)
        Next
        'Tag it as a compressed format and save the source buffer size
        CompBuf(8) = Asc("F")
        Poke4 CompBuf, 9, TargetOffset
        'make the compressed buffer exact size and dump it to file
        ReDim Preserve CompBuf(1 To targetsize + 16)
        Put #TargetFileHandle, 1, CompBuf
        'frmComp.lblStatus.Caption = "Completed compile - " & CStr(targetsize + 16) & " Bytes."
    Else
        Put #TargetFileHandle, 1, TokenTarget
        'frmComp.lblStatus.Caption = "Completed compile - " & CStr(TargetOffset) & " Bytes."
    End If
    Close #TargetFileHandle
    
    Exit Sub
Errtrap:
   
End Sub


Public Sub InitValues2()
    Home = CurDir
    OtherHome = App.path
    SourceFile = vbNullString
    TargetFile = vbNullString
    Source = vbNullString
    Target = vbNullString
    ParensDepth = 0
    SourceIdx = 1
    TargetOffset = 33
    ReDim TokenTarget(1 To 32)
    
End Sub

'----------------------------------------------
' Remap integer into 2 bytes
'----------------------------------------------
Public Sub Poke2(buf() As Byte, Start As Long, num As Integer)
    On Error GoTo ReGrow
    If Start > UBound(buf) - 10 Then GoTo ReGrow
    CopyMemory buf(Start), ByVal VarPtr(num), 2
    Exit Sub
ReGrow:
    ReDim Preserve buf(LBound(buf) To UBound(buf) + 10000)
    On Error GoTo 0
    CopyMemory buf(Start), ByVal VarPtr(num), 2
    
End Sub

'----------------------------------------------
' ReMap Long into 4 bytes
'----------------------------------------------
Public Sub Poke4(buf() As Byte, Start As Long, num As Long)
    On Error GoTo ReGrow
    If Start > UBound(buf) - 10 Then GoTo ReGrow
    CopyMemory buf(Start), ByVal VarPtr(num), 4
    Exit Sub
ReGrow:
    ReDim Preserve buf(LBound(buf) To UBound(buf) + 10000)
    On Error GoTo 0
    CopyMemory buf(Start), ByVal VarPtr(num), 4
End Sub

'----------------------------------------------
' ReMap Currency into 4 bytes
'----------------------------------------------
Public Sub Poke4U(buf() As Byte, Start As Long, num As Currency)
Dim l As Long
    l = CLng(num)
    On Error GoTo ReGrow
    If Start > UBound(buf) - 10 Then GoTo ReGrow
    CopyMemory buf(Start), ByVal VarPtr(l), 4
    Exit Sub
ReGrow:
    ReDim Preserve buf(LBound(buf) To UBound(buf) + 10000)
    On Error GoTo 0
    CopyMemory buf(Start), ByVal VarPtr(l), 4
End Sub

'--------------------------------------------
' ReMap Single Float into 4 bytes
'--------------------------------------------
Public Sub Poke4FL(buf() As Byte, Start As Long, num As Single)
    On Error GoTo ReGrow
    If Start > UBound(buf) - 10 Then GoTo ReGrow
    CopyMemory buf(Start), ByVal VarPtr(num), 4
    Exit Sub
ReGrow:
    ReDim Preserve buf(LBound(buf) To UBound(buf) + 10000)
    On Error GoTo 0
    CopyMemory buf(Start), ByVal VarPtr(num), 4
End Sub

'--------------------------------------------
' ReMap Unicode string into consecutive bytes
'--------------------------------------------
Public Sub PokeStr(buf() As Byte, Start As Long, str As String)
Dim usize As Long
    usize = LenB(str)
    On Error GoTo ReGrow
     If Start > UBound(buf) - usize - 10 Then GoTo ReGrow


    CopyMemory buf(Start), ByVal StrPtr(str), usize
    Exit Sub
ReGrow:
    ReDim Preserve buf(LBound(buf) To UBound(buf) + 10000)
    On Error GoTo 0
    CopyMemory buf(Start), ByVal StrPtr(str), usize
End Sub

Public Function FileOpen2(FileName As String, inout As String) As Long
Dim hdl As Long

    hdl = FreeFile(0)
    If StrComp("Write", inout, vbTextCompare) = 0 Then
        Open FileName For Binary Access Write As #hdl
    Else
        Open FileName For Binary Access Read As #hdl
    End If
    FileOpen2 = hdl
End Function

'Public Function FilePath(str As String) As String
''-----------------------------------
''Extract the file path portion including trailing \
''-----------------------------------
'Dim l As Long
'Dim ptr As Long
'
'    l = Len(str)
'    If l < 1 Then
'        FilePath = vbNullString
'        Exit Function
'    End If
'    ptr = InStrRev(str, "\", -1)
'    If ptr = 0 Then
'        FilePath = vbNullString
'        Exit Function
'    End If
'    FilePath = Mid$(str, 1, ptr)
'End Function
'
'Public Function FileTitle(str As String) As String
''--------------------------------------------
''Extract the full filename without any path
''--------------------------------------------
'Dim l As Long
'Dim ptr As Long
'
'    l = Len(str)
'    If l < 1 Then
'        FileTitle = vbNullString
'        Exit Function
'    End If
'    ptr = InStrRev(str, "\", -1)
'    If ptr = 0 Then
'        FileTitle = str
'        Exit Function
'    End If
'    FileTitle = Mid$(str, ptr + 1)
'End Function
'
'Public Function FileRoot(str As String) As String
''-------------------------------------------
''Extract the filename without the extension
''-------------------------------------------
'Dim tmp As String
'Dim ptr As Long
'
'    tmp = FileTitle(str)
'    ptr = InStr(1, tmp, ".")
'    If ptr = 0 Then
'        FileRoot = tmp
'    Else
'        FileRoot = Mid$(tmp, 1, ptr - 1)
'    End If
'End Function

'Public Function FileExtension(str As String) As String
''-----------------------------------------
''Extract just the file extension value
''-----------------------------------------
'Dim tmp As String
'Dim ptr As Long
'
'    tmp = FileTitle(str)
'    ptr = InStr(1, tmp, ".")
'    If ptr = 0 Then
'        FileExtension = vbNullString
'    Else
'        FileExtension = Mid$(tmp, ptr + 1)
'    End If
'End Function


'Public Sub FileWrite(hdl As Long, line As String, uni As Boolean)
'    If Mid$(line, Len(line), 1) = " " Then line = Mid$(line, 1, Len(line) - 1)
'    If uni Then
'        Print #hdl, StrConv(line & vbCrLf, vbUnicode);
'    Else
'        Print #hdl, line
'    End If
'End Sub

'Public Sub FileWriteNoCR(hdl As Long, line As String, uni As Boolean)
'    If uni Then
'        Print #hdl, StrConv(line, vbUnicode);
'    Else
'        Print #hdl, line;
'    End If
'End Sub

'Public Function FileRead(hdl As Long) As String
'Dim line As String
'
'    Line Input #hdl, line
'    If EOF(hdl) Then
'        FileRead = "*EOF*"
'    Else
'        FileRead = line
'    End If
'End Function

'Public Sub FileClose(hdl As Long)
'    Close #hdl
'End Sub

Public Function NextItem(position As Long, offset As Long, Sep As String) As String

dbgnextitem = dbgnextitem + 1
'---------------------------------------------
' Used to make a copy of a section of string data
' that is delimited by given seperators
'---------------------------------------------
Dim i As Long
Dim s As Long
Dim ptr As Long
Dim eptr As Long
On Error GoTo Errtrap
' Stop
    If position = 1 Then
        eptr = SearchStr(offset, Sep, Len(SourceFile))
        NextItem = Mid$(SourceFile, offset, eptr - offset)
        Exit Function
    End If
    If position < 1 Then
        NextItem = vbNullString
        Exit Function
    End If
    eptr = offset - 1
    For i = 1 To position
        ptr = eptr + 1
        eptr = SearchStr(ptr, Sep, Len(SourceFile))
        If eptr = 0 Then
            If i = position Then
                NextItem = Mid$(SourceFile, ptr)
                Exit Function
            End If
            NextItem = vbNullString
            Exit Function
        End If
    Next
    NextItem = Mid$(SourceFile, ptr, eptr - ptr)
    Exit Function
Errtrap:
  
End Function

Public Function SearchStr(Start As Long, Sep As String, Max As Long) As Long
Dim eptr As Long
Dim QuoteFound As Boolean
' If kind = T_Type Then
' For eptr = start To max
' If cSourceFile(eptr) = Asc(Sep) Then Exit For
' Next
' Else
    QuoteFound = False
    For eptr = Start To Max
        If QuoteFound Then
            If cSourceFile(eptr) = 34 Then QuoteFound = False
        Else
            If cSourceFile(eptr) = 34 Then QuoteFound = True
        End If
        If Not QuoteFound Then
            If cSourceFile(eptr) = Asc(Sep) Then Exit For
        End If
    Next
    'End If
    SearchStr = eptr
End Function


Public Function ItemCount(list As String, Sep As String) As Long
dbgitemcount = dbgitemcount + 1
'--------------------------------------------
' See how many delimited strings exist
' by definition this is always one more
' than the number of seperators found
'--------------------------------------------
Dim i As Long
Dim ptr As Long
    
    i = 1
    ptr = 1
    Do
        ptr = InStr(ptr, list, Sep, vbTextCompare)
        If ptr = 0 Then
            ItemCount = i
            Exit Function
        End If
        i = i + 1
        ptr = ptr + Len(Sep)
    Loop
End Function

Public Function DoNextToken(ByRef srcidx As Long, srcend As Long) As Boolean
dbgdonext = dbgdonext + 1
Dim line As String
Dim CurrentOffset As Long
Dim lineidx As Long
Dim num As Long
Dim cnum As Currency
Dim fnum As Single
Dim Token As Long
Dim TokenFound As Boolean
Dim ret As Boolean
Dim i As Long
Dim j As Long
Dim k As Long
Dim gcnt As Long
Dim rcnt As Long
Dim parmidx As Long
Dim parmend As Long
Dim possiblelabel As String
Dim hexlist As String
Dim bytl As Byte
Dim byth As Byte
Dim bytf As Byte
Dim str As String
Dim tempend As Long
Dim TokenName As String
Dim TokenType As Long
Dim TokenCount As Long
Dim TokenPrecis As Long
Dim TokenEmbed As Boolean, strTemp As String, numTemp As Single

On Error GoTo Errtrap


If srcidx = 1 Then
        srcend = FindFullParens(1, SearchStr(srcidx, "(", srcend), srcend)
    End If

If srcidx > srcend Then
        DoNextToken = False
        Exit Function
    End If



    parmidx = srcidx
    parmend = srcend
    hexlist = "0123456789abcdef"
    ParensDepth = ParensDepth + 1
    
    Do While parmidx < srcend
    
        If TargetOffset >= (UBound(TokenTarget) - 100000) Then
            ReDim Preserve TokenTarget(1 To TargetOffset + 500000)
        End If
        
        TokenName = LCase(NextItem(1, parmidx, " "))
 '       Debug.Print TokenName
'        frmComp.lblStatus.Caption = "Processing Token " & TokenName & " at offset " & CStr(parmidx)
'        frmComp.lblStatus.Refresh
        
        ' look for label strings here !!
        possiblelabel = NextItem(2, parmidx, " ")
        If possiblelabel = "(" Then possiblelabel = vbNullString
        
        parmidx = SearchStr(parmidx, "(", Len(SourceFile))
        
        parmend = FindFullParens(ParensDepth, parmidx, srcend)
        parmidx = parmidx + 2
        
        Token = Tokenz.Item(TokenName).TokNumber
        TokenType = Tokenz.Item(TokenName).TokType
        TokenCount = Tokenz.Item(TokenName).TokCount
        TokenEmbed = Tokenz.Item(TokenName).TokEmbed
        TokenPrecis = Tokenz.Item(TokenName).TokPrecis
                
        CurrentOffset = TargetOffset
        
        Poke2 TokenTarget, CurrentOffset, CInt(Token)
        'Poke2 TokenTarget, CurrentOffset + 2, 0
        If kind = W_Type Then
            Poke2 TokenTarget, CurrentOffset + 2, 4
        Else
            Poke2 TokenTarget, CurrentOffset + 2, 0
        End If


        CurrentOffset = CurrentOffset + 4
        Poke4 TokenTarget, CurrentOffset, 1
        TokenTarget(CurrentOffset + 4) = 0
        If possiblelabel <> vbNullString Then
            TokenTarget(CurrentOffset + 4) = Len(possiblelabel)
            PokeStr TokenTarget, CurrentOffset + 5, possiblelabel
        End If
        
        
        
        Select Case TokenType
        
            Case TK_none            'All
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                
            Case TK_uint            'All
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                If TokenPrecis <> 0 Then
                    If Mid$(SourceFile, parmidx, 1) = ")" Then TokenCount = 0
                End If

                For i = 1 To TokenCount
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
'                    parmidx = InStr(parmidx + 1, SourceFile, " ") + 1
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                
            Case TK_str             'all
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    str = NextItem(1, parmidx, " ")
                    str = Replace(str, Chr$(34), "")
                    Poke2 TokenTarget, TargetOffset, CInt(Len(str))
                    PokeStr TokenTarget, TargetOffset + 2, str
                    TargetOffset = TargetOffset + 2 + LenB(str)
                    parmidx = SearchStr(parmidx, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                    If Not ret Then
                        DoNextToken = False
                        Exit Function
                    End If
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset


            
            Case TK_dword           ' S and W files
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    str = NextItem(1, parmidx, " ")
                    If str = ")" Then Exit For  '********
                    For j = 7 To 1 Step -2
                        bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                        byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                        bytf = (byth * 16) + bytl
                        TokenTarget(TargetOffset) = bytf
                        TargetOffset = TargetOffset + 1
                    Next
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset

            Case TK_float               'All
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                fnum = CNAnum(parmidx)
                    'fnum = CSng(NextItem(1, parmidx, " "))
                    Poke4FL TokenTarget, TargetOffset, fnum
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
            
            Case TK_uint4float          ' W files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                cnum = CCur(NextItem(1, parmidx, " "))
                If cnum > 2147483647 Then
                    num = CLng(cnum - 4294967296#)
                Else
                    num = CLng(cnum)
                End If
                Poke4 TokenTarget, TargetOffset, num
                TargetOffset = TargetOffset + 4
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                gcnt = CLng(cnum)
                For j = 1 To gcnt
                    For i = 1 To TokenPrecis
'                    strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                    fnum = CNAnum(parmidx)
                        'fnum = CSng(NextItem(1, parmidx, " "))
                        Poke4FL TokenTarget, TargetOffset, fnum
                        TargetOffset = TargetOffset + 4
                        parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                    Next
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
            
            Case TK_2sint3float             ' W files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    num = CLng(NextItem(1, parmidx, " "))
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                For i = 1 To TokenPrecis
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
fnum = CNAnum(parmidx)
                    'fnum = CSng(NextItem(1, parmidx, " "))
                    Poke4FL TokenTarget, TargetOffset, fnum
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
            
            Case TK_sint                    ' W files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    num = CLng(NextItem(1, parmidx, " "))
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                            
            Case TK_2uint2float             ' None - superceded
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                For i = 1 To TokenPrecis
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                    fnum = CNAnum(parmidx)
                    'fnum = CSng(NextItem(1, parmidx, " "))
                    Poke4FL TokenTarget, TargetOffset, fnum
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                            
            Case TK_uintfloat               ' T and W files
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                For i = 1 To TokenPrecis
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                fnum = CNAnum(parmidx)
                    'fnum = CSng(NextItem(1, parmidx, " "))
                    Poke4FL TokenTarget, TargetOffset, fnum
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                
            Case TK_dworduint               ' W files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    str = NextItem(1, parmidx, " ")
                    For j = 7 To 1 Step -2
                        bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                        byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                        bytf = (byth * 16) + bytl
                        TokenTarget(TargetOffset) = bytf
                        TargetOffset = TargetOffset + 1
                    Next
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                For i = 1 To TokenPrecis
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
            
            Case TK_tokuintfloat            ' W files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                tempend = FindFullParens(ParensDepth, parmidx, srcend)
                ret = DoNextToken(parmidx, tempend)
                If Not ret Then
                    DoNextToken = False
                    Exit Function
                End If
                For i = 1 To TokenCount
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                For i = 1 To TokenPrecis
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
fnum = CNAnum(parmidx)
                   ' fnum = CSng(NextItem(1, parmidx, " "))
                    Poke4FL TokenTarget, TargetOffset, fnum
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
            
            Case TK_uintuint            ' S files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                cnum = CCur(NextItem(1, parmidx, " "))
                If cnum > 2147483647 Then
                    num = CLng(cnum - 4294967296#)
                Else
                    num = CLng(cnum)
                End If
                Poke4 TokenTarget, TargetOffset, num
                TargetOffset = TargetOffset + 4
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                gcnt = CLng(cnum)
                
                Select Case TokenPrecis
                
                    Case TK_uint
                        For j = 1 To gcnt
                            For i = 1 To TokenCount
                                cnum = CCur(NextItem(1, parmidx, " "))
                                If cnum > 2147483647 Then
                                    num = CLng(cnum - 4294967296#)
                                Else
                                    num = CLng(cnum)
                                End If
                                Poke4 TokenTarget, TargetOffset, num
                                TargetOffset = TargetOffset + 4
                                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                            Next
                        Next
                        
                    Case TK_sint
                        For j = 1 To gcnt
                            For i = 1 To TokenCount
                                num = CLng(NextItem(1, parmidx, " "))
                                Poke4 TokenTarget, TargetOffset, num
                                TargetOffset = TargetOffset + 4
                                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                            Next
                        Next
                    
                    Case TK_float
                        For j = 1 To gcnt
                            For i = 1 To TokenCount
'                            strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                              fnum = CNAnum(parmidx)
                                'fnum = CSng(NextItem(1, parmidx, " "))
                                Poke4FL TokenTarget, TargetOffset, fnum
                                TargetOffset = TargetOffset + 4
                                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                            Next
                        Next
                        
                    Case TK_dword
                        For k = 1 To gcnt
                            For i = 1 To TokenCount
                                str = NextItem(1, parmidx, " ")
                                For j = 7 To 1 Step -2
                                    bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                                    byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                                    bytf = (byth * 16) + bytl
                                    TokenTarget(TargetOffset) = bytf
                                    TargetOffset = TargetOffset + 1
                                Next
                                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                            Next
                        Next
                    
                End Select
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
            
            Case TK_uintfloatdword          ' S files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                For i = 1 To TokenPrecis
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                fnum = CNAnum(parmidx)
                    'fnum = CSng(NextItem(1, parmidx, " "))
                    Poke4FL TokenTarget, TargetOffset, fnum
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                str = NextItem(1, parmidx, " ")
                For j = 7 To 1 Step -2
                    bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                    byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                    bytf = (byth * 16) + bytl
                    TokenTarget(TargetOffset) = bytf
                    TargetOffset = TargetOffset + 1
                Next
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                
            Case TK_dworduintfloat          ' S files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    str = NextItem(1, parmidx, " ")
                    For j = 7 To 1 Step -2
                        bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                        byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                        bytf = (byth * 16) + bytl
                        TokenTarget(TargetOffset) = bytf
                        TargetOffset = TargetOffset + 1
                    Next
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                For i = 1 To TokenPrecis
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                  fnum = CNAnum(parmidx)
                   ' fnum = CSng(NextItem(1, parmidx, " "))
                    Poke4FL TokenTarget, TargetOffset, fnum
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                cnum = CCur(NextItem(1, parmidx, " "))
                If cnum > 2147483647 Then
                    num = CLng(cnum - 4294967296#)
                Else
                    num = CLng(cnum)
                End If
                Poke4 TokenTarget, TargetOffset, num
                TargetOffset = TargetOffset + 4
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                
            Case TK_uintfloat6              ' S files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                For i = 1 To TokenPrecis
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                 fnum = CNAnum(parmidx)
                    'fnum = CSng(NextItem(1, parmidx, " "))
                    Poke4FL TokenTarget, TargetOffset, fnum
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                                        
            Case TK_mixed1                  ' S files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                str = NextItem(1, parmidx, " ")
                For j = 7 To 1 Step -2
                    bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                    byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                    bytf = (byth * 16) + bytl
                    TokenTarget(TargetOffset) = bytf
                    TargetOffset = TargetOffset + 1
                Next
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                str = NextItem(1, parmidx, " ")
                If str = ")" Then Exit Do
                cnum = CCur(str)
                If cnum > 2147483647 Then
                    num = CLng(cnum - 4294967296#)
                Else
                    num = CLng(cnum)
                End If
                Poke4 TokenTarget, TargetOffset, num
                TargetOffset = TargetOffset + 4
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                num = CLng(NextItem(1, parmidx, " "))
                Poke4 TokenTarget, TargetOffset, num
                TargetOffset = TargetOffset + 4
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                str = NextItem(1, parmidx, " ")
                If str = ")" Then Exit Do
                cnum = CCur(str)
                If cnum > 2147483647 Then
                    num = CLng(cnum - 4294967296#)
                Else
                    num = CLng(cnum)
                End If
                Poke4 TokenTarget, TargetOffset, num
                TargetOffset = TargetOffset + 4
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Do While NextItem(1, parmidx, " ") <> ")"
                    num = CLng(NextItem(1, parmidx, " "))
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Loop
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
            
            Case TK_uintplus                ' S files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                Do
                    str = NextItem(1, parmidx, " ")
                    If str = ")" Then Exit Do
                    cnum = CCur(str)
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Loop
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                
            Case TK_uintplusfloat               ' S files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                Do While NextItem(1, parmidx, " ") <> ")"
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                fnum = CNAnum(parmidx)
                    'fnum = CSng(NextItem(1, parmidx, " "))
                    Poke4FL TokenTarget, TargetOffset, fnum
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Loop
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
            
            Case TK_mixed3                  ' S files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                str = NextItem(1, parmidx, " ")
                For j = 7 To 1 Step -2
                    bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                    byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                    bytf = (byth * 16) + bytl
                    TokenTarget(TargetOffset) = bytf
                    TargetOffset = TargetOffset + 1
                Next
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                For i = 1 To 2
                    num = CLng(NextItem(1, parmidx, " "))
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                For i = 1 To 2
                    str = NextItem(1, parmidx, " ")
                    For j = 7 To 1 Step -2
                        bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                        byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                        bytf = (byth * 16) + bytl
                        TokenTarget(TargetOffset) = bytf
                        TargetOffset = TargetOffset + 1
                    Next
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                tempend = FindFullParens(ParensDepth, parmidx, srcend)
                ret = DoNextToken(parmidx, tempend)
                If Not ret Then
                    DoNextToken = False
                    Exit Function
                End If
                Do While NextItem(1, parmidx, " ") <> ")"
                    If Not IsNumeric(NextItem(1, parmidx, " ")) Then
                        tempend = FindFullParens(ParensDepth, parmidx, srcend)
                        ret = DoNextToken(parmidx, tempend)
                        If Not ret Then
                            DoNextToken = False
                            Exit Function
                        End If
                    Else
                        Do While NextItem(1, parmidx, " ") <> ")"
                            cnum = CCur(NextItem(1, parmidx, " "))
                            If cnum > 2147483647 Then
                                num = CLng(cnum - 4294967296#)
                            Else
                                num = CLng(cnum)
                            End If
                            Poke4 TokenTarget, TargetOffset, num
                            TargetOffset = TargetOffset + 4
                            parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                        Loop
                    End If
                Loop
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                

            Case TK_mixed4                  ' S files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                str = NextItem(1, parmidx, " ")
                For j = 7 To 1 Step -2
                    bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                    byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                    bytf = (byth * 16) + bytl
                    TokenTarget(TargetOffset) = bytf
                    TargetOffset = TargetOffset + 1
                Next
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                For i = 1 To 2
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                For i = 1 To 2
                    str = NextItem(1, parmidx, " ")
                    For j = 7 To 1 Step -2
                        bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                        byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                        bytf = (byth * 16) + bytl
                        TokenTarget(TargetOffset) = bytf
                        TargetOffset = TargetOffset + 1
                    Next
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                tempend = FindFullParens(ParensDepth, parmidx, srcend)
                ret = DoNextToken(parmidx, tempend)
                If Not ret Then
                    DoNextToken = False
                    Exit Function
                End If
                Do While NextItem(1, parmidx, " ") <> ")"
                    If Not IsNumeric(NextItem(1, parmidx, " ")) Then
                        tempend = FindFullParens(ParensDepth, parmidx, srcend)
                        ret = DoNextToken(parmidx, tempend)
                        If Not ret Then
                            DoNextToken = False
                            Exit Function
                        End If
                    Else
                        Do While NextItem(1, parmidx, " ") <> ")"
'                        strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                          fnum = CNAnum(parmidx)
                           'fnum = CSng(NextItem(1, parmidx, " "))
                            Poke4FL TokenTarget, TargetOffset, fnum
                            TargetOffset = TargetOffset + 4
                            parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                        Loop
                    End If
                Loop
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
            
            Case TK_uintnocr                ' W files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                            
            Case TK_mixed2                  ' S files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                str = NextItem(1, parmidx, " ")
                For j = 7 To 1 Step -2
                    bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                    byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                    bytf = (byth * 16) + bytl
                    TokenTarget(TargetOffset) = bytf
                    TargetOffset = TargetOffset + 1
                Next
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                cnum = CCur(NextItem(1, parmidx, " "))
                If cnum > 2147483647 Then
                    num = CLng(cnum - 4294967296#)
                Else
                    num = CLng(cnum)
                End If
                Poke4 TokenTarget, TargetOffset, num
                TargetOffset = TargetOffset + 4
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                tempend = FindFullParens(ParensDepth, parmidx, srcend)
                ret = DoNextToken(parmidx, tempend)
                If Not ret Then
                    DoNextToken = False
                    Exit Function
                End If
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
               fnum = CNAnum(parmidx)
                'fnum = CSng(NextItem(1, parmidx, " "))
                Poke4FL TokenTarget, TargetOffset, fnum
                TargetOffset = TargetOffset + 4
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                num = CLng(NextItem(1, parmidx, " "))
                Poke4 TokenTarget, TargetOffset, num
                TargetOffset = TargetOffset + 4
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Do While NextItem(1, parmidx, " ") <> ")"
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Loop
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
            
            Case TK_tokfloat                ' S files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                tempend = FindFullParens(ParensDepth, parmidx, srcend)
                ret = DoNextToken(parmidx, tempend)
                If Not ret Then
                    DoNextToken = False
                    Exit Function
                End If
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
              fnum = CNAnum(parmidx)
               ' fnum = CSng(NextItem(1, parmidx, " "))
                Poke4FL TokenTarget, TargetOffset, fnum
                TargetOffset = TargetOffset + 4
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
            
            Case TK_struint                 ' T files only
                
             
               TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                For i = 1 To TokenCount
                    str = NextItem(1, parmidx, " ")
                    str = Replace(str, Chr$(34), "")
                    Poke2 TokenTarget, TargetOffset, CInt(Len(str))
                    PokeStr TokenTarget, TargetOffset + 2, str
                    TargetOffset = TargetOffset + 2 + LenB(str)
                    parmidx = SearchStr(parmidx, " ", Len(SourceFile)) + 1
                Next
                For i = 1 To TokenPrecis
              
                    cnum = CCur(NextItem(1, parmidx, " "))
                    If cnum > 2147483647 Then
                        num = CLng(cnum - 4294967296#)
                    Else
                        num = CLng(cnum)
                    End If
                    Poke4 TokenTarget, TargetOffset, num
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Next
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                
            Case TK_dwordfloat              ' T files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                str = NextItem(1, parmidx, " ")
                For j = 7 To 1 Step -2
                    bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                    byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                    bytf = (byth * 16) + bytl
                    TokenTarget(TargetOffset) = bytf
                    TargetOffset = TargetOffset + 1
                Next
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Do While NextItem(1, parmidx, " ") <> ")"
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                    fnum = CNAnum(parmidx)
                    'fnum = CSng(NextItem(1, parmidx, " "))
                    Poke4FL TokenTarget, TargetOffset, fnum
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Loop
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset

            Case TK_buffer                  ' T files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                num = CLng(NextItem(1, parmidx, " "))
                parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                For i = 1 To num
                    bytf = CByte(NextItem(1, parmidx, " "))
                    TokenTarget(TargetOffset) = bytf
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                    TargetOffset = TargetOffset + 1
                Next
                Do While parmidx < parmend
                    str = NextItem(1, parmidx, " ")
                    For j = 7 To 1 Step -2
                        bytl = InStr(1, hexlist, Mid$(str, j + 1, 1), vbTextCompare) - 1
                        byth = InStr(1, hexlist, Mid$(str, j, 1), vbTextCompare) - 1
                        bytf = (byth * 16) + bytl
                        TokenTarget(TargetOffset) = bytf
                        TargetOffset = TargetOffset + 1
                    Next
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Loop
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset
                
            Case TK_somefloat               ' T files only
                TargetOffset = CurrentOffset + 5 + (2 * Len(possiblelabel))
                Do While NextItem(1, parmidx, " ") <> ")"
'                strTemp = NextItem(1, parmidx, " ")
'                numTemp = Val(strTemp)
'                fnum = CSng(numTemp)
                   fnum = CNAnum(parmidx)
                   ' fnum = CSng(NextItem(1, parmidx, " "))
                    Poke4FL TokenTarget, TargetOffset, fnum
                    TargetOffset = TargetOffset + 4
                    parmidx = SearchStr(parmidx + 1, " ", Len(SourceFile)) + 1
                Loop
                If TokenEmbed Then
                    ret = DoNextToken(parmidx, parmend)
                End If
                Poke4 TokenTarget, CurrentOffset, (TargetOffset - CurrentOffset - 4)
                CurrentOffset = TargetOffset

        End Select
    parmidx = SearchStr(parmidx, " ", Len(SourceFile)) + 1
    Loop
    
    ParensDepth = ParensDepth - 1
    srcidx = parmidx
    DoNextToken = True
    Exit Function
Errtrap:
   
End Function

Public Function FindFullParens(StartDepth As Long, StartIdx As Long, StartEnd As Long) As Long
dbgfindfull = dbgfindfull + 1
Dim x As Long
Dim X1 As Long
Dim X2 As Long
Dim Y As Long

    x = StartIdx
    Y = StartDepth
'    X = InStr(X, SourceFile, "(") + 1
    x = SearchStr(x, "(", Len(SourceFile)) + 1
    Do
        X1 = SearchStr(x, ")", Len(SourceFile))
        X2 = SearchStr(x, "(", Len(SourceFile))
        If X1 = 0 Then X1 = StartEnd
        If X2 = 0 Then X2 = StartEnd
        x = X1
        If X2 < X1 Then x = X2
        If x > StartEnd Then
            FindFullParens = StartEnd
            Exit Function
        End If
        If Mid$(SourceFile, x, 1) = ")" Then Y = Y - 1
        If Mid$(SourceFile, x, 1) = "(" Then Y = Y + 1
        If Y < StartDepth Then
            FindFullParens = x
            Exit Function
        End If
        x = x + 1
    Loop
End Function

