VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmComp 
   Caption         =   "Comp"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   1920
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim tfile As String
Dim size As Long
Dim srcbuf() As Byte
Dim ret As Boolean
Dim CompBuf() As Byte
Dim Sourcehdr As String
Dim targetsize As Long

    
    Call InitValues2
    
    Call Init_Tokenz
    
CDL1.Filter = "MSTS Files (*.s;*.t;*.w)|*.s;*.t;*.w"
CDL1.DialogTitle = "Open MSTS File"
CDL1.FilterIndex = 1
CDL1.Action = 1
tfile = CDL1.filename
If Right$(tfile, 1) = "s" Then
kind = S_Type
ElseIf Right$(tfile, 1) = "t" Then
kind = T_Type
ElseIf Right$(tfile, 1) = "w" Then
kind = W_Type
Else
Call MsgBox("An invalid compressed file type has been selected.", vbCritical, App.Title)
Exit Sub
End If
  
    'The following few lines must be expanded to allow user to choose file and file type
    
'    tFile = "C:\Program Files\Microsoft Games\Train Simulator\ROUTES\USA2\Tiles\-01a192d4-2nd.Expanded"
'    kind = T_Type

'    tFile = "C:\Program Files\Microsoft Games\Train Simulator\ROUTES\USA2\Shapes\barfence10m.Expanded"
'    kind = S_Type

'    tFile = "C:\Program Files\Microsoft Games\Train Simulator\ROUTES\USA2\World\w-012583+014758.Expanded"
'    kind = W_Type
'

    'access the input file - ASSUME it is unicode text for now
    SourceFileHandle = FileOpen2(tfile, "Read")
    'need to know its exact size in bytes (not characters)
    size = LOF(SourceFileHandle)
    ReDim srcbuf(1 To size)
    'read whole file in and then close file
    Get #SourceFileHandle, 1, srcbuf
    Close #SourceFileHandle
    'cheater trick to move contents into string intact
    SourceFile = srcbuf
    'Freeup the memory
    ReDim srcbuf(1 To 1)
    'Save off the original msts header
    Sourcehdr = Mid(SourceFile, 1, 32)
    'then strip it from source file string
    SourceFile = Mid(SourceFile, InStr(1, SourceFile, vbCrLf) + 2)
    'Make string into one very long line of text with only single spaces as seperators
    SourceFile = Replace(SourceFile, vbCrLf, " ")
    SourceFile = Replace(SourceFile, vbTab, " ")
    SourceFile = Trim(SourceFile) & " "
    Do While (InStr(1, SourceFile, "  ") <> 0)
        SourceFile = Replace(SourceFile, "  ", " ")
    Loop
    
    'Move original header over from unicode to ansi buffer
    For i = 1 To Len(Sourcehdr) - 1
        TokenTarget(i) = Asc(Mid(Sourcehdr, i + 1, 1))
    Next
    'Tag it as a binary version
    TokenTarget(24) = Asc("b")
    'Replace the missing linefeed
    TokenTarget(32) = 10
    
    'go compile the input string to the target buffer
    ret = DoNextToken(1, Len(SourceFile))
    
    'If the output is just a compiled token file,then the extension is .Token
    'If the output is also compressed, the extension is .Compressed
    
    tfile = Replace(tfile, ".Expanded", ".Token")
    If ((kind = W_Type) Or (kind = S_Type)) Then
        tfile = Replace(tfile, ".Token", ".Compressed")
    End If
    
    TargetFileHandle = FileOpen2(tfile, "Write")
    'Make output buffer exact size
    ReDim Preserve TokenTarget(1 To TargetOffset - 1)
    
    If ((kind = W_Type) Or (kind = S_Type)) Then
        'set initial compress buffer plenty big enough
        targetsize = 12 + CLng(TargetOffset * 1.1)
        ReDim CompBuf(1 To targetsize)
        'Dont include first 16 bytes in compression size
        TargetOffset = TargetOffset - 17
        'start compression after first 16 bytes
        rstl = Compress(CompBuf(17), VarPtr(targetsize), TokenTarget(17), TargetOffset)
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
    Else
'        Put #TargetFileHandle, 1, "SIMISA@@@@@@@@@@JINX0t6b______" & vbCrLf
        Put #TargetFileHandle, 1, TokenTarget
    End If
    Close #TargetFileHandle
        
'    MsgBox "DoNextToken=" & CStr(dbgdonext) & vbCrLf & "ItemExtract=" & CStr(dbgitemextract) & vbCrLf & "ItemCount=" & CStr(dbgitemcount) & vbCrLf & "FindFull=" & CStr(dbgfindfull), vbOKOnly, "Counters"
        
    End
End Sub
