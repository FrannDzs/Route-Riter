VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ComponentOne ActiveX Converter: Version 7 to 8"
   ClientHeight    =   4170
   ClientLeft      =   3930
   ClientTop       =   2310
   ClientWidth     =   6405
   Icon            =   "Convert7to8.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6405
   Begin VB.CommandButton cmd 
      Caption         =   "Select project to convert"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   3885
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   2640
      TabIndex        =   6
      Top             =   5280
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton cmdButton2 
      Caption         =   "Convert all projects in the selected folder"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   4800
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   240
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lst 
      Height          =   1665
      IntegralHeight  =   0   'False
      ItemData        =   "Convert7to8.frx":0442
      Left            =   120
      List            =   "Convert7to8.frx":0444
      TabIndex        =   0
      Top             =   1800
      Width           =   6015
   End
   Begin VB.Label Label4 
      Caption         =   "Batch Conversion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Double-click a folder to select it."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   120
      X2              =   6240
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Files:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: The source code for this utility is included in the ComponentOne ActiveX Studio distribution package."
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   6135
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"Convert7to8.frx":0446
      Height          =   390
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6150
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' conversion tables
Dim g_VbpRef() As String
Dim g_FrmRef() As String
Dim g_CtlName() As String
Dim g_GUID() As String

' table dimensions
Dim g_NVbpRef As Integer
Dim g_NFrmRef As Integer
Dim g_NCtlName As Integer
Dim g_NGUID As Integer

' constants
Const OLD_VERS = 0
Const NEW_VERS = 1

' convert control declaration and properties
Sub ConvertControl(s$)
    Dim i%
    
    ' declarations
    For i = 0 To g_NCtlName - 1
        StringReplace s, g_CtlName(OLD_VERS, i), g_CtlName(NEW_VERS, i)
    Next i
    Print #2, s
    
    ' loop converting the control properties
    While Not EOF(1)
        Line Input #1, s
        
        ' no blobs in version 8 (vspdf7 used them, just skip the line)
        If InStr(s, "OleObjectBlob") > 0 Or InStr(s, "_ConvInfo") > 0 Then
            Line Input #1, s
        End If
        
        ' handle properties that are sub-objects
        If Left(Trim(s), 14) = "BeginProperty " Then
            For i = 0 To g_NGUID - 1
                StringReplace s, g_GUID(OLD_VERS, i), g_GUID(NEW_VERS, i)
            Next
        End If
        
        ' handle nested controls
        If Left(Trim(s), 6) = "Begin " Then
            ConvertControl s
        
        ' at the end of the control, bail out
        ElseIf Trim(s) = "End" Then
            Print #2, s
            Exit Sub
            
        ' other stuff passes on directly
        Else
            Print #2, s
        End If
    Wend
    
    ' we should never get here
    Debug.Print "unexpected EOF"

End Sub

' convert control types
Sub ConvertCode(ByRef s$)
    Dim i%, j%, s1$, s2$
    For i = 0 To g_NCtlName - 1
        
        ' get names
        s1 = g_CtlName(OLD_VERS, i)
        s2 = g_CtlName(NEW_VERS, i)
        
        ' convert full names
        StringReplace s, s1, s2
        
        ' convert library names
        j = InStr(s1, ".")
        If j > 0 Then s1 = Left(s1, j)
        j = InStr(s2, ".")
        If j > 0 Then s2 = Left(s2, j)
        StringReplace s, s1, s2
        
    Next i
End Sub

' convert form
Sub ConvertForm(ByVal pth$, fn$)
    Dim s$
    Dim bConv As Boolean
    Dim i As Integer
    
    ' if the file has a path spec, don't prepend the project path
    If InStr(fn, "\") = 0 Then pth = ""
    
    ' make a backup copy of the file
    ' we just append '.bak' to the file name
    FileCopy fn, fn & ".bak"
    
    ' open files to read (backup) and write (converted)
    Open pth & fn & ".bak" For Input As #1
    Open pth & fn For Output As #2
    
    ' scan the file
    While Not EOF(1)
        Line Input #1, s
        
        ' convert object references
        bConv = False
        For i = 0 To g_NFrmRef - 1
            If StringReplace(s, g_FrmRef(OLD_VERS, i), g_FrmRef(NEW_VERS, i)) Then
                bConv = True
            End If
        Next i
        
        If bConv = True Then
           Print #2, s
        
        ' convert controls
        ElseIf Left(Trim(s), 6) = "Begin " Then
            ConvertControl s
            
        ' convert code and event declarations,
        ' then copy whatever is left to the new file
        Else
            ConvertCode s
            Print #2, s
        End If
    Wend
    
    ' done with this form, close both files
    Close #1
    Close #2
End Sub

' replace strings, return value to indicate whether chages were made
Function StringReplace(s$, find$, repl$) As Boolean
    Dim result$
    result = Replace(s, find, repl)
    If result <> s Then
        s = result
        StringReplace = True
    End If
End Function

' prompt for project and convert it
Private Sub cmd_Click()
    Dim prj$
    
    ' open project
    dlg.CancelError = False
    dlg.FileName = ""
    dlg.DefaultExt = "vbp"
    dlg.Filter = "Visual Basic Projecs (*.vbp)|*.vbp"
    dlg.Flags = cdlOFNFileMustExist
    dlg.ShowOpen
    prj = dlg.FileName
    If prj = "" Then Exit Sub
    
    ' convert project
    ConvertProject prj, True

End Sub

' convert project
Private Sub ConvertProject(prj As String, confirm As Boolean)
    Dim pth$, s$, i%, j%
    Dim bConv As Boolean
    
    ' clear list
    lst.Clear
    
    ' parse project path and file name
    i = 0
    While InStr(i + 1, prj, "\") > 0
        i = InStr(i + 1, prj, "\")
    Wend
    pth = Left(prj, i)
    prj = Mid(prj, i + 1)
    
    ' read the whole project file
    Open pth & prj For Binary As #1
    s = Input(LOF(1), 1)
    Close #1
    
    ' make sure we have something to convert
    bConv = False
    For i = 0 To g_NVbpRef - 1
        If InStr(s, g_VbpRef(OLD_VERS, i)) > 0 Then
            bConv = True
        End If
    Next i
        
    If bConv = False Then
        MsgBox "Couldn't find any convertible controls in '" & prj & "'." & vbCrLf & vbCrLf & "Conversion cancelled.", vbInformation
        Exit Sub
    End If
    
    ' confirm conversion
    If confirm Then
        i = MsgBox("Ready to convert '" & prj & "' from Version 7 to 8." & vbCrLf & vbCrLf & _
                   "The original files will be backed up in files with an additional '.bak' extension." & vbCrLf & vbCrLf & _
                   "Press OK to proceed or Cancel if you changed your mind.", _
                   vbOKCancel Or vbInformation)
        If i <> vbOK Then Exit Sub
    End If
    
    ' replace library references
    For i = 0 To g_NVbpRef - 1
        StringReplace s, g_VbpRef(OLD_VERS, i), g_VbpRef(NEW_VERS, i)
    Next i
    
    ' don't disturb me while I work
    Me.Enabled = False
    
    ' make a backup copy of the project file
    ' we just append '.bak' to the file name
    FileCopy pth & prj, pth & prj & ".bak"
    
    ' write out updated project file
    Open prj For Output As #1
    Print #1, s
    Close #1
    lst.AddItem prj & "... OK"
    
    ' find all forms in the project
    i = 1
    Do
        i = InStr(i, s, "Form=")
        If i = 0 Then Exit Do
        If Mid(s, i + 5, 1) <> """" Then
            j = InStr(i, s, ".frm")
            If j <> 0 Then lst.AddItem Mid(s, i + 5, j - i - 5 + 4)
        End If
        i = i + 1
    Loop
    
    ' find all UserControls in the project
    i = 1
    Do
        i = InStr(i, s, "UserControl=")
        If i = 0 Then Exit Do
        If Mid(s, i + 12, 1) <> """" Then
            j = InStr(i, s, ".ctl")
            If j <> 0 Then lst.AddItem Mid(s, i + 12, j - i - 12 + 4)
        End If
        i = i + 1
    Loop
    
    ' process forms
    For i = 1 To lst.ListCount - 1
        ConvertForm pth, lst.List(i)
        lst.ListIndex = i
        lst.List(i) = lst.List(i) & "... OK"
        Refresh
        DoEvents
    Next
    
    ' show confirmation and warnings
    If confirm Then
        s = "Project '" & prj & "' has been converted from Version 7 to 8." & vbCrLf & vbCrLf & _
            "Copies of the original files were saved with a 'bak' extension in case you want to restore the original project." & vbCrLf & vbCrLf & _
            "You should now open the new project in VB and compile it to make sure there are no syntax errors." & vbCrLf & vbCrLf
        MsgBox s, vbInformation Or vbOKOnly
    End If
    
    ' all done, re-activate form
    Me.Enabled = True

End Sub

' ----------------------------------------------------------------------------
' For batch conversion only
'

Private Sub cmdButton2_Click()
    Dim i%, j%, cnt%
    Dim f$
    
    ' loop through all directories under the selected one
    For i = 0 To Dir1.ListCount - 1
        File1.Path = Dir1.List(i)
         
        ' loop through all files in this directory
        For j = 0 To File1.ListCount - 1
            f = File1.List(j)
            
            ' if a VB project, add the full path to the file name and
            ' convert project
            If Right(f, 4) = ".vbp" Then
                f = File1.Path + "\" + f
                Debug.Print "Converting " & f
                ChDir File1.Path
                ConvertProject f, False
                cnt = cnt + 1
            End If
        
        Next j  ' next file
    Next i ' next directory
    
    MsgBox "Done converting " & cnt & " projects.", vbInformation, "Done"
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
    
    ' populate conversion tables
    PopConvTables
    
End Sub

' populate conversion tables
Private Sub PopConvTables()

    ' set table dimensions
    g_NVbpRef = 27
    g_NFrmRef = 27
    g_NCtlName = 29
    g_NGUID = 6

    ' redim tables
    ReDim g_VbpRef(2, g_NVbpRef)
    ReDim g_FrmRef(2, g_NFrmRef)
    ReDim g_CtlName(2, g_NCtlName)
    ReDim g_GUID(2, g_NGUID)
    
    ' ---------------------------------------------------
    ' list of controls handled.  Keep an eye on this - as the
    '  list grows the form design may need adjustment.
    
    Label1(0) = "This utility converts VB projects that use " + _
      "VSFlex7, VSView7, VSReport7, SizerOne, VSSpell6, " + _
      "C1Query, TrueData6, TrueDataLite7, C1Chart7, TrueDBList7, TrueDBGrid7, and XArrayDB7" + _
      " controls to the new Version 8 controls."
    
    ' ---------------------------------------------------
    ' load tables: Version 7

    ' old references stored in vbp files
    g_VbpRef(OLD_VERS, 0) = "Object={A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0; VSPrint7.ocx"             ' VSPrinter
    g_VbpRef(OLD_VERS, 1) = "Object={6871D5CE-1A9F-11D4-9A1F-F7280EC6F828}#1.0#0; VSDraw7.ocx"              ' VSDraw
    g_VbpRef(OLD_VERS, 2) = "Object={8FF4C5A0-1A0C-11D4-9A1F-EF40A7BBFB28}#1.0#0; VSVPort7.ocx"             ' VSViewPort
    g_VbpRef(OLD_VERS, 3) = "Object={1B04A20A-C295-476C-BA28-DC6D9110E7A3}#1.0#0; VSPDF.ocx"                ' VSPdf
    g_VbpRef(OLD_VERS, 4) = "Object={49C98174-5A47-443B-9ADD-5F4880F7096D}#1.0#0; AwkOne.ocx"               ' AwkOne
    g_VbpRef(OLD_VERS, 5) = "Object={9E883861-2808-4487-913D-EA332634AC0D}#1.0#0; SizerOne.ocx"             ' SizerOne, TabOne
    g_VbpRef(OLD_VERS, 6) = "Object={08769121-33BD-11D3-BD95-B44CFE3A3C4B}#1.0#0; VSSpell6.ocx"             ' VSSpell
    g_VbpRef(OLD_VERS, 7) = "Object={83DCDF03-433E-11D3-BD95-C5F237C8B472}#1.0#0; VSThes6.ocx"              ' VSThesaurus
    g_VbpRef(OLD_VERS, 8) = "Object={429F6260-B945-11D3-9A1F-9E6707138531}#1.0#0; vsFlex7N.ocx"             ' VSFlexGrid (Unicode Light)
    g_VbpRef(OLD_VERS, 9) = "Object={C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0; vsFlex7l.ocx"             ' VSFlexGrid (Light)
    g_VbpRef(OLD_VERS, 10) = "Object={D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0; vsFlex7u.ocx"            ' VSFlexGrid (Unicode ADO)
    g_VbpRef(OLD_VERS, 11) = "Object={D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0; vsFlex7d.ocx"            ' VSFlexGrid (DAO)
    g_VbpRef(OLD_VERS, 12) = "Object={D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0; vsFlex7.ocx"             ' VSFlexGrid (ADO)
    g_VbpRef(OLD_VERS, 13) = "Object={DCC46394-4B19-11D3-BD95-D426EF2C7949}#1.0#0; VSStr7.ocx"              ' VSFlexString
    g_VbpRef(OLD_VERS, 14) = "Object={3A6F7F80-45E5-11D4-AC3E-ADBCE8B30410}#1.0#0; VSRpt7.ocx"              ' VSReport
    g_VbpRef(OLD_VERS, 15) = "Object={DEF7CB00-83C0-11D0-A0F1-00A024703500}#1.0#0; C1Q1.OCX"                ' C1Query
    g_VbpRef(OLD_VERS, 16) = "Object={0D6235BA-DBA2-11D1-B5DF-0060976089D0}#1.0#0; truedc6.ocx"             ' TData
    g_VbpRef(OLD_VERS, 17) = "Object={0D623681-DBA2-11D1-B5DF-0060976089D0}#1.0#0; tdcl7.ocx"               ' TDataLite
    g_VbpRef(OLD_VERS, 18) = "Object={C643EB3F-235C-4181-9B55-36A268833718}#7.0#0; Olch2x7.ocx"             ' C1Chart2D
    g_VbpRef(OLD_VERS, 19) = "Object={A4F5504C-4D7B-4827-87C7-7CA6D5794D06}#7.0#0; Olch3x7.ocx"             ' C1Chart3D
    g_VbpRef(OLD_VERS, 20) = "Object={9487F13A-8164-4CB5-97BD-CFA9A776D71F}#7.0#0; Olch2xu7.ocx"            ' C1Chart2D (UNICODE)
    g_VbpRef(OLD_VERS, 21) = "Object={7DA9DE68-6056-4010-8A8D-B76808352C30}#7.0#0; Olch3xu7.ocx"            ' C1Chart3D (UNICODE)
    g_VbpRef(OLD_VERS, 22) = "Object={0D6236CD-DBA2-11D1-B5DF-0060976089D0}#7.0#0; tdbl7.ocx"               ' tdblist7(Icursor)
    g_VbpRef(OLD_VERS, 23) = "Object={DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0; todl7.ocx"               ' tdblist7(OLEDB)
    g_VbpRef(OLD_VERS, 24) = "Object={0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0; tdbg7.ocx"               ' TDBGrid7 (ICursor)
    g_VbpRef(OLD_VERS, 25) = "Object={DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0; todg7.ocx"               ' TDBGrid7 (OLEDB)
    g_VbpRef(OLD_VERS, 26) = "Reference=*\G{0D6236A9-DBA2-11D1-B5DF-0060976089D0}#7.0#0#"               ' XArrayDB7

    ' old references stored in frm/ctl files
    g_FrmRef(OLD_VERS, 0) = "Object = ""{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0""; ""VSPrint7.ocx"""   ' VSPrinter
    g_FrmRef(OLD_VERS, 1) = "Object = ""{6871D5CE-1A9F-11D4-9A1F-F7280EC6F828}#1.0#0""; ""VSDraw7.ocx"""    ' VSDraw
    g_FrmRef(OLD_VERS, 2) = "Object = ""{8FF4C5A0-1A0C-11D4-9A1F-EF40A7BBFB28}#1.0#0""; ""VSVPort7.ocx"""   ' VSViewPort
    g_FrmRef(OLD_VERS, 3) = "Object = ""{1B04A20A-C295-476C-BA28-DC6D9110E7A3}#1.0#0""; ""VSPDF.ocx"""      ' VSPDF
    g_FrmRef(OLD_VERS, 4) = "Object = ""{49C98174-5A47-443B-9ADD-5F4880F7096D}#1.0#0""; ""AwkOne.ocx"""     ' AwkOne
    g_FrmRef(OLD_VERS, 5) = "Object = ""{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0""; ""SizerOne.ocx"""   ' SizerOne, TabOne
    g_FrmRef(OLD_VERS, 6) = "Object = ""{08769121-33BD-11D3-BD95-B44CFE3A3C4B}#1.0#0""; ""VSSpell6.ocx"""   ' VSSpell
    g_FrmRef(OLD_VERS, 7) = "Object = ""{83DCDF03-433E-11D3-BD95-C5F237C8B472}#1.0#0""; ""VSThes6.ocx"""    ' VSThesaurus
    g_FrmRef(OLD_VERS, 8) = "Object = ""{429F6260-B945-11D3-9A1F-9E6707138531}#1.0#0""; ""vsFlex7N.ocx"""   ' VSFlexGrid (Unicode Light)
    g_FrmRef(OLD_VERS, 9) = "Object = ""{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0""; ""vsFlex7l.ocx"""   ' VSFlexGrid (Light)
    g_FrmRef(OLD_VERS, 10) = "Object = ""{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0""; ""vsFlex7u.ocx"""  ' VSFlexGrid (Unicode)
    g_FrmRef(OLD_VERS, 11) = "Object = ""{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0""; ""vsFlex7d.ocx"""  ' VSFlexGrid (DAO)
    g_FrmRef(OLD_VERS, 12) = "Object = ""{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0""; ""vsFlex7.ocx"""   ' VSFlexGrid (ADO)
    g_FrmRef(OLD_VERS, 13) = "Object = ""{DCC46394-4B19-11D3-BD95-D426EF2C7949}#1.0#0""; ""VSStr7.ocx"""    ' VSFlexString
    g_FrmRef(OLD_VERS, 14) = "Object = ""{3A6F7F80-45E5-11D4-AC3E-ADBCE8B30410}#1.0#0""; ""VSRpt7.ocx"""    ' VSReport
    g_FrmRef(OLD_VERS, 15) = "Object = ""{DEF7CB00-83C0-11D0-A0F1-00A024703500}#1.0#0""; ""C1Q1.OCX"""      ' C1Query
    g_FrmRef(OLD_VERS, 16) = "Object = ""{0D6235BA-DBA2-11D1-B5DF-0060976089D0}#1.0#0""; ""truedc6.ocx"""   ' TData
    g_FrmRef(OLD_VERS, 17) = "Object = ""{0D623681-DBA2-11D1-B5DF-0060976089D0}#1.0#0""; ""tdcl7.ocx"""     ' TDataLite
    g_FrmRef(OLD_VERS, 18) = "Object = ""{C643EB3F-235C-4181-9B55-36A268833718}#7.0#0""; ""Olch2x7.ocx"""   ' C1Chart2D
    g_FrmRef(OLD_VERS, 19) = "Object = ""{A4F5504C-4D7B-4827-87C7-7CA6D5794D06}#7.0#0""; ""Olch3x7.ocx"""   ' C1Chart3D
    g_FrmRef(OLD_VERS, 20) = "Object = ""{9487F13A-8164-4CB5-97BD-CFA9A776D71F}#7.0#0""; ""Olch2xu7.ocx"""  ' C1Chart2D (UNICODE)
    g_FrmRef(OLD_VERS, 21) = "Object = ""{7DA9DE68-6056-4010-8A8D-B76808352C30}#7.0#0""; ""Olch3xu7.ocx"""  ' C1Chart3D (UNICODE)
    g_FrmRef(OLD_VERS, 22) = "Object = ""{0D6236CD-DBA2-11D1-B5DF-0060976089D0}#7.0#0""; ""tdbl7.ocx"""     ' TDBList(ICursor)
    g_FrmRef(OLD_VERS, 23) = "Object = ""{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0""; ""todl7.ocx"""     ' TDBList (OLEDB)
    g_FrmRef(OLD_VERS, 24) = "Object = ""{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0""; ""tdbg7.ocx"""     ' TDBGrid7 (ICursor)
    g_FrmRef(OLD_VERS, 25) = "Object = ""{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0""; ""todg7.ocx"""     ' TDBGrid7 (OLEDB)
    g_FrmRef(OLD_VERS, 26) = "Object = ""{0D6236A9-DBA2-11D1-B5DF-0060976089D0}#7.0#0""; ""xadb7.ocx"""     ' XArrayDB7

    ' old control class names
    g_CtlName(OLD_VERS, 0) = "VSPrinter7LibCtl.VSPrinter"               ' VSPrinter
    g_CtlName(OLD_VERS, 1) = "VSDraw7LibCtl.VSDraw"                     ' VSDraw
    g_CtlName(OLD_VERS, 2) = "VSViewPort7LibCtl.VSViewPort"             ' VSViewPort
    g_CtlName(OLD_VERS, 3) = "VSPDFLibCtl.VSPDF"                        ' VSPDF
    g_CtlName(OLD_VERS, 4) = "AwkOneLibCtl.AwkOne"                      ' AwkOne
    g_CtlName(OLD_VERS, 5) = "SizerOneLibCtl.TabOne"                    ' TabOne
    g_CtlName(OLD_VERS, 6) = "SizerOneLibCtl.ElasticOne"                ' ElasticOne
    g_CtlName(OLD_VERS, 7) = "VSSPELL6LibCtl.VSSpell"                   ' VSSpell
    g_CtlName(OLD_VERS, 8) = "VSTHES6LibCtl.VSThesaurus"                ' VSTHesaurus
    g_CtlName(OLD_VERS, 9) = "VSFlex7NCtl.VSFlexGrid"                   ' VSFlexGrid (Light Unicode)
    g_CtlName(OLD_VERS, 10) = "VSFlex7LCtl.VSFlexGrid"                  ' VSFlexGrid (Light)
    g_CtlName(OLD_VERS, 11) = "VSFlex7UCtl.VSFlexGrid"                  ' VSFlexGrid (Unicode)
    g_CtlName(OLD_VERS, 12) = "VSFlex7DAOCtl.VSFlexGrid"                ' VSFlexGrid (DAO)
    g_CtlName(OLD_VERS, 13) = "VSFlex7Ctl.VSFlexGrid"                   ' VSFlexGrid (ADO)
    g_CtlName(OLD_VERS, 14) = "VSSTR7LibCtl.VSFlexString"               ' VSFlexString
    g_CtlName(OLD_VERS, 15) = "VSREPORTLibCtl.VSReport"                 ' VSReport
    g_CtlName(OLD_VERS, 16) = "C1Query10Ctl.C1Query"                    ' C1Query
    g_CtlName(OLD_VERS, 17) = "C1Query10Ctl.C1QueryFrame"               ' C1Query
    g_CtlName(OLD_VERS, 18) = "TrueData60Ctl.TData"                     ' TData
    g_CtlName(OLD_VERS, 19) = "TrueDataLite70Ctl.TDataLite"             ' TDataLite
    g_CtlName(OLD_VERS, 20) = "C1Chart2D7.Chart2D"                      ' C1Chart2D
    g_CtlName(OLD_VERS, 21) = "C1Chart3D7.Chart3D"                      ' C1Chart3D
    g_CtlName(OLD_VERS, 22) = "C1Chart2D7U.Chart2D"                     ' C1Chart2D Unicode
    g_CtlName(OLD_VERS, 23) = "C1Chart3D7U.Chart3D"                     ' C1Chart3D Unicode
    g_CtlName(OLD_VERS, 24) = "TrueDBList70"                            ' TDBList
    g_CtlName(OLD_VERS, 25) = "TrueOleDBList70"                         ' TDBList OLEDB
    g_CtlName(OLD_VERS, 26) = "TrueDBGrid70"                            ' TDBGrid7 (ICursor)
    g_CtlName(OLD_VERS, 27) = "TrueOleDBGrid70"                         ' TDBGrid7 (OLEDB)
    g_CtlName(OLD_VERS, 28) = "XArrayDBObject"                          ' XArrayDB7
    

    ' guids used to instantiate old objects that are sub-properties
    g_GUID(OLD_VERS, 0) = "{8F5A70A3-B6D3-11D3-9A1F-800A5BACB530}"      ' VSReport Layout class
    g_GUID(OLD_VERS, 1) = "{8F5A70A1-B6D3-11D3-9A1F-800A5BACB530}"      ' VSReport DataSource class
    g_GUID(OLD_VERS, 2) = "{8AFA5902-B17C-11D3-9A1F-C87A5EC37F33}"      ' VSReport Group class
    g_GUID(OLD_VERS, 3) = "{E5849A61-ADD9-11D3-BDEB-000000000000}"      ' VSReport Section class
    g_GUID(OLD_VERS, 4) = "{E5849A63-ADD9-11D3-BDEB-000000000000}"      ' VSReport Field class
    g_GUID(OLD_VERS, 5) = "{3A6F7F8D-45E5-11D4-AC3E-ADBCE8B30410}"      ' VSReport Report class
    
    ' ---------------------------------------------------
    ' load tables: version 8

    ' new references stored in vbp files
    g_VbpRef(NEW_VERS, 0) = "Object={54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0; VSPrint8.ocx"             ' VSPrinter
    g_VbpRef(NEW_VERS, 1) = "Object={D3F92121-EFAA-4B5C-B91B-3D6A8FFD1477}#1.0#0; VSDraw8.ocx"              ' VSDraw
    g_VbpRef(NEW_VERS, 2) = "Object={96548BD2-D0BF-46B1-B519-8F2268D49306}#1.0#0; VSVPort8.ocx"             ' VSViewPort
    g_VbpRef(NEW_VERS, 3) = "Object={1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0; VSPDF8.ocx"               ' VSPdf
    g_VbpRef(NEW_VERS, 4) = "Object={E7F57B23-8EC4-4B47-A3F2-06800D054C07}#1.0#0; C1Awk.ocx"                ' AwkOne
    g_VbpRef(NEW_VERS, 5) = "Object={0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0; C1Sizer.ocx"              ' SizerOne, TabOne
    g_VbpRef(NEW_VERS, 6) = "Object={CDF1175F-8D7D-431B-8B61-069CB3A80DD6}#1.0#0; Spell8.ocx"               ' VSSpell
    g_VbpRef(NEW_VERS, 7) = "Object={05BD37E5-B82F-49E6-9A0A-97BE4815460C}#1.0#0; Thes8.ocx"                ' VSThesaurus
    g_VbpRef(NEW_VERS, 8) = "Object={9DB12C0E-8736-49E6-9CA1-896A1E67D6F3}#1.0#0; vsFlex8N.ocx"             ' VSFlexGrid (Unicode Light)
    g_VbpRef(NEW_VERS, 9) = "Object={1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0; vsFlex8l.ocx"             ' VSFlexGrid (Light)
    g_VbpRef(NEW_VERS, 10) = "Object={C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0; vsFlex8u.ocx"            ' VSFlexGrid (Unicode ADO)
    g_VbpRef(NEW_VERS, 11) = "Object={FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0; vsFlex8d.ocx"            ' VSFlexGrid (DAO)
    g_VbpRef(NEW_VERS, 12) = "Object={BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0; vsFlex8.ocx"             ' VSFlexGrid (ADO)
    g_VbpRef(NEW_VERS, 13) = "Object={1C81E0B1-BFC2-42C4-A910-97E0FA9F83C9}#1.0#0; VSStr8.ocx"              ' VSFlexString
    g_VbpRef(NEW_VERS, 14) = "Object={C8CF160E-7278-4354-8071-850013B36892}#1.0#0; VSRpt8.ocx"              ' VSReport
    g_VbpRef(NEW_VERS, 15) = "Object={605925BE-4799-4093-A2E7-39323147E70E}#1.0#0; C1Query8.OCX"            ' C1Query
    g_VbpRef(NEW_VERS, 16) = "Object={7FEC7313-D161-427C-A141-48E17931414B}#1.0#0; truedc8.ocx"             ' TData
    g_VbpRef(NEW_VERS, 17) = "Object={E8E54757-DFC9-473C-98B3-15D72AB7870B}#1.0#0; tdcl8.ocx"               ' TDataLite
    g_VbpRef(NEW_VERS, 18) = "Object={0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0; Olch2x8.ocx"             ' C1Chart2D
    g_VbpRef(NEW_VERS, 19) = "Object={AC82DA6D-CD3F-43BF-AF2E-56591B5585D8}#8.0#0; Olch3x8.ocx"             ' C1Chart3D
    g_VbpRef(NEW_VERS, 20) = "Object={75634CE7-D088-4C44-8F7C-3C117CE5857B}#8.0#0; Olch2xu8.ocx"            ' C1Chart2D (UNICODE)
    g_VbpRef(NEW_VERS, 21) = "Object={5C9704A4-FE02-45EC-A1E3-7773F3CB0D5A}#8.0#0; Olch3xu8.ocx"            ' C1Chart3D (UNICODE)
    g_VbpRef(NEW_VERS, 22) = "Object={CC225F54-31E2-494D-83EA-7C88B58F46B0}#8.0#0; tdbl8.ocx"               ' TDBL (ICursor)
    g_VbpRef(NEW_VERS, 23) = "Object={60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0; todl8.ocx"               ' TDBL (OLEDB)
    g_VbpRef(NEW_VERS, 24) = "Object={77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0; tdbg8.ocx"               ' TDBGrid8 (ICursor)
    g_VbpRef(NEW_VERS, 25) = "Object={562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0; todg8.ocx"               ' TDBGrid8 (OLEDB)
    g_VbpRef(NEW_VERS, 26) = "Reference=*\G{C10EF3FA-DEBF-4189-8859-C35CA400BBA8}#8.0#0#"               ' XArrayDB8
    
    ' new references stored in frm/ctl files
    g_FrmRef(NEW_VERS, 0) = "Object = ""{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0""; ""VSPrint8.ocx"""   ' VSPrinter
    g_FrmRef(NEW_VERS, 1) = "Object = ""{D3F92121-EFAA-4B5C-B91B-3D6A8FFD1477}#1.0#0""; ""VSDraw8.ocx"""    ' VSDraw
    g_FrmRef(NEW_VERS, 2) = "Object = ""{96548BD2-D0BF-46B1-B519-8F2268D49306}#1.0#0""; ""VSVPort8.ocx"""   ' VSViewPort
    g_FrmRef(NEW_VERS, 3) = "Object = ""{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0""; ""VSPDF8.ocx"""     ' VSPdf
    g_FrmRef(NEW_VERS, 4) = "Object = ""{E7F57B23-8EC4-4B47-A3F2-06800D054C07}#1.0#0""; ""C1Awk.ocx"""      ' AwkOne
    g_FrmRef(NEW_VERS, 5) = "Object = ""{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0""; ""C1Sizer.ocx"""    ' SizerOne, TabOne
    g_FrmRef(NEW_VERS, 6) = "Object = ""{CDF1175F-8D7D-431B-8B61-069CB3A80DD6}#1.0#0""; ""Spell8.ocx"""     ' VSSpell
    g_FrmRef(NEW_VERS, 7) = "Object = ""{05BD37E5-B82F-49E6-9A0A-97BE4815460C}#1.0#0""; ""Thes8.ocx"""      ' VSThesaurus
    g_FrmRef(NEW_VERS, 8) = "Object = ""{9DB12C0E-8736-49E6-9CA1-896A1E67D6F3}#1.0#0""; ""vsFlex8N.ocx"""   ' VSFlexGrid (Unicode Light)
    g_FrmRef(NEW_VERS, 9) = "Object = ""{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0""; ""vsFlex8l.ocx"""   ' VSFlexGrid (Light)
    g_FrmRef(NEW_VERS, 10) = "Object = ""{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0""; ""vsFlex8u.ocx"""  ' VSFlexGrid (Unicode ADO)
    g_FrmRef(NEW_VERS, 11) = "Object = ""{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0""; ""vsFlex8d.ocx"""  ' VSFlexGrid (DAO)
    g_FrmRef(NEW_VERS, 12) = "Object = ""{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0""; ""vsFlex8.ocx"""   ' VSFlexGrid (ADO)
    g_FrmRef(NEW_VERS, 13) = "Object = ""{1C81E0B1-BFC2-42C4-A910-97E0FA9F83C9}#1.0#0""; ""VSStr8.ocx"""    ' VSFlexString
    g_FrmRef(NEW_VERS, 14) = "Object = ""{C8CF160E-7278-4354-8071-850013B36892}#1.0#0""; ""VSRpt8.ocx"""    ' VSReport
    g_FrmRef(NEW_VERS, 15) = "Object = ""{605925BE-4799-4093-A2E7-39323147E70E}#1.0#0""; ""C1Query8.OCX"""  ' C1Query
    g_FrmRef(NEW_VERS, 16) = "Object = ""{7FEC7313-D161-427C-A141-48E17931414B}#1.0#0""; ""truedc8.ocx"""   ' TData
    g_FrmRef(NEW_VERS, 17) = "Object = ""{E8E54757-DFC9-473C-98B3-15D72AB7870B}#1.0#0""; ""tdcl8.ocx"""     ' TDataLite
    g_FrmRef(NEW_VERS, 18) = "Object = ""{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0""; ""Olch2x8.ocx"""   ' C1Chart2D
    g_FrmRef(NEW_VERS, 19) = "Object = ""{AC82DA6D-CD3F-43BF-AF2E-56591B5585D8}#8.0#0""; ""Olch3x8.ocx"""   ' C1Chart3D
    g_FrmRef(NEW_VERS, 20) = "Object = ""{75634CE7-D088-4C44-8F7C-3C117CE5857B}#8.0#0""; ""Olch2xu8.ocx"""  ' C1Chart2D (UNICODE)
    g_FrmRef(NEW_VERS, 21) = "Object = ""{5C9704A4-FE02-45EC-A1E3-7773F3CB0D5A}#8.0#0""; ""Olch3xu8.ocx"""  ' C1Chart3D (UNICODE)
    g_FrmRef(NEW_VERS, 22) = "Object = ""{CC225F54-31E2-494D-83EA-7C88B58F46B0}#8.0#0""; ""tdbl8.ocx"""     ' TDBL
    g_FrmRef(NEW_VERS, 23) = "Object = ""{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0""; ""todl8.ocx"""     ' TDBL OLEDB
    g_FrmRef(NEW_VERS, 24) = "Object = ""{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0""; ""tdbg8.ocx"""     ' TDBGrid8 (ICursor)
    g_FrmRef(NEW_VERS, 25) = "Object = ""{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0""; ""todg8.ocx"""     ' TDBGrid8 (OLEDB)
    g_FrmRef(NEW_VERS, 26) = "Object = ""{C10EF3FA-DEBF-4189-8859-C35CA400BBA8}#8.0#0""; ""xadb8.ocx"""     ' XArrayDB8
    
    ' new control class names
    g_CtlName(NEW_VERS, 0) = "VSPrinter8LibCtl.VSPrinter"               ' VSPrinter
    g_CtlName(NEW_VERS, 1) = "VSDraw8LibCtl.VSDraw"                     ' VSDraw
    g_CtlName(NEW_VERS, 2) = "VSViewPort8LibCtl.VSViewPort"             ' VSViewPort
    g_CtlName(NEW_VERS, 3) = "VSPDF8LibCtl.VSPDF8"                      ' VSPDF
    g_CtlName(NEW_VERS, 4) = "C1AwkLibCtl.C1Awk"                        ' AwkOne
    g_CtlName(NEW_VERS, 5) = "C1SizerLibCtl.C1Tab"                      ' TabOne
    g_CtlName(NEW_VERS, 6) = "C1SizerLibCtl.C1Elastic"                  ' ElasticOne
    g_CtlName(NEW_VERS, 7) = "VSSPELL8LibCtl.VSSpell"                   ' VSSpell
    g_CtlName(NEW_VERS, 8) = "VSTHES8LibCtl.VSThesaurus"                ' VSTHesaurus
    g_CtlName(NEW_VERS, 9) = "VSFlex8NCtl.VSFlexGrid"                   ' VSFlexGrid (Light Unicode)
    g_CtlName(NEW_VERS, 10) = "VSFlex8LCtl.VSFlexGrid"                  ' VSFlexGrid (Light)
    g_CtlName(NEW_VERS, 11) = "VSFlex8UCtl.VSFlexGrid"                  ' VSFlexGrid (Unicode)
    g_CtlName(NEW_VERS, 12) = "VSFlex8DAOCtl.VSFlexGrid"                ' VSFlexGrid (DAO)
    g_CtlName(NEW_VERS, 13) = "VSFlex8Ctl.VSFlexGrid"                   ' VSFlexGrid (ADO)
    g_CtlName(NEW_VERS, 14) = "VSStr8LibCtl.VSFlexString"               ' VSFlexString
    g_CtlName(NEW_VERS, 15) = "VSReport8LibCtl.VSReport"                ' VSReport
    g_CtlName(NEW_VERS, 16) = "C1Query80Ctl.C1Query"                    ' C1Query
    g_CtlName(NEW_VERS, 17) = "C1Query80Ctl.C1QueryFrame"               ' C1Query
    g_CtlName(NEW_VERS, 18) = "TrueData80Ctl.TData"                     ' TData
    g_CtlName(NEW_VERS, 19) = "TrueDataLite80Ctl.TDataLite"             ' TDataLite
    g_CtlName(NEW_VERS, 20) = "C1Chart2D8.Chart2D"                      ' C1Chart2D
    g_CtlName(NEW_VERS, 21) = "C1Chart3D8.Chart3D"                      ' C1Chart3D
    g_CtlName(NEW_VERS, 22) = "C1Chart2D8U.Chart2D"                     ' C1Chart2D Unicode
    g_CtlName(NEW_VERS, 23) = "C1Chart3D8U.Chart3D"                     ' C1Chart3D Unicode
    g_CtlName(NEW_VERS, 24) = "TrueDBList80"                            ' TDBList
    g_CtlName(NEW_VERS, 25) = "TrueOleDBList80"                         ' TDBListOLE
    g_CtlName(NEW_VERS, 26) = "TrueDBGrid80"                            ' TDBGrid8 (ICursor)
    g_CtlName(NEW_VERS, 27) = "TrueOleDBGrid80"                         ' TDBGrid8 (OLEDB)
    g_CtlName(NEW_VERS, 28) = "XArrayDBObject"                          ' XArrayDB7
    
    ' guids used to instantiate new objects that are sub-properties
    g_GUID(NEW_VERS, 0) = "{D853A4F1-D032-4508-909F-18F074BD547A}"      ' VSReport Layout class
    g_GUID(NEW_VERS, 1) = "{D1359088-0913-44ea-AE50-6A7CD77D4C50}"      ' VSReport DataSource class
    g_GUID(NEW_VERS, 2) = "{E862F8BF-E806-4b39-9A11-0BBED515338B}"      ' VSReport Group class
    g_GUID(NEW_VERS, 3) = "{673CB92F-28D3-421f-86CD-1099425A5037}"      ' VSReport Section class
    g_GUID(NEW_VERS, 4) = "{6AC1BBA5-107E-4f07-BCF0-DF757735D0A8}"      ' VSReport Field class
    g_GUID(NEW_VERS, 5) = "{C0489F4D-67C4-4069-9E3E-95245F084E3F}"      ' VSReport Report class
    
End Sub
