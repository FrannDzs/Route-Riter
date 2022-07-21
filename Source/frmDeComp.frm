VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDeComp 
   Caption         =   "DeComp"
   ClientHeight    =   4395
   ClientLeft      =   165
   ClientTop       =   915
   ClientWidth     =   8085
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDeComp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Proceed"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdTarget 
      Caption         =   "Target File"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "Source File"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.OptionButton optType 
      Caption         =   "Terrain Tile File"
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton optType 
      Caption         =   "Shape File"
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton optType 
      Caption         =   "World File"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   6720
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      UseMnemonic     =   0   'False
      Width           =   7575
   End
   Begin VB.Label Label1 
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblTarget 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   5895
   End
   Begin VB.Label lblSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   5895
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmDeComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdProceed_Click()
Dim line As String
Dim thdl As Long
Dim Token As Long
Dim i As Long
Dim tmp As String
Dim pos As Long
Dim rslt As Long
Dim Ename As String
Dim tgtdir As String
Dim Sourcebuf As String
Dim Oname As String
Dim str As String
Dim strb(1 To 32) As Byte

    If ((Source = "") Or (Target = "")) Then
        frmDeComp.lblStatus.Caption = "Both the Source and Target file must be specified above."
        Exit Sub
    End If
               
    rslt = InflateTile(Source)
    
    If rslt = -98 Then
        frmDeComp.lblStatus.Caption = Source & " already unicode."
        Exit Sub
    End If
    If rslt = -99 Then
        frmDeComp.lblStatus.Caption = Source & " is not a compressed W,S or T file."
        Exit Sub
    End If
    
'    If DebugMode Then
'        rslt = FreeFile(1)
'        Open Oname For Binary Access Write As #rslt
'        Put #rslt, 1, "SIMISA@@@@@@@@@@"
'        Put #rslt, 17, TokenTarget
'        Close #rslt
'    End If
    
    SaveSetting App.EXEName, "Settings", "Source", Source
    SaveSetting App.EXEName, "Settings", "Target", Target
    SaveSetting App.EXEName, "Settings", "Kind", CStr(Kind)
        
    str = String(16, " ")
    For i = 2 To 32 Step 2
        strb(i - 1) = TokenTarget(i / 2)
        strb(i) = 0
    Next
    CopyMemory ByVal StrPtr(str), strb(1), 32
    
    i = InStrRev(str, "b")
    Mid(str, i, 1) = "t"
    If Kind = T_Type Then
        Mid(str, i - 1, 1) = "x"
    End If
    SourceOffset = 17
    TargetFileHandle = FreeFile(1)
    Open Target For Output As TargetFileHandle
    Print #TargetFileHandle, Chr(255) & Chr(254);
    Print #TargetFileHandle, StrConv("SIMISA@@@@@@@@@@", vbUnicode);
    Print #TargetFileHandle, StrConv(str & vbCrLf, vbUnicode);
    TabDepth = 0
    PreviousToken = 0
    PreviousOffset = 0
    
'   Call the recursive tag expander
    Call DoSomeTags(UBound(TokenTarget), True)
    
    Close #TargetFileHandle
    
    frmDeComp.lblStatus.Caption = "Finished " & Target & " at " & CStr(FileLen(Target)) & " Bytes."
    frmDeComp.lblStatus.Refresh

End Sub

Private Sub cmdSource_Click()
Dim tgtdir As String
Dim rslt As Long

    If Kind = W_Type Then
        frmDeComp.cdlg.DefaultExt = "w"
        frmDeComp.cdlg.Filter = "World Files (*.w)|*.w|All Files (*.*)|*.*"
    ElseIf Kind = S_Type Then
        frmDeComp.cdlg.DefaultExt = "s"
        frmDeComp.cdlg.Filter = "Shape Files (*.s)|*.s|All Files (*.*)|*.*"
    Else
        frmDeComp.cdlg.DefaultExt = "t"
        frmDeComp.cdlg.Filter = "Terrain Files (*.t)|*.t|All Files (*.*)|*.*"
    End If
    frmDeComp.cdlg.DialogTitle = "Select the Source File"
    If Source <> "" Then
        frmDeComp.cdlg.filename = FileTitle(Source)
        frmDeComp.cdlg.InitDir = FilePath(Source)
    Else
        frmDeComp.cdlg.filename = ""
        frmDeComp.cdlg.InitDir = TSRoot
    End If
    frmDeComp.cdlg.FilterIndex = 0
    frmDeComp.cdlg.CancelError = True
    frmDeComp.cdlg.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer
    frmDeComp.lblSource.Caption = ""
    
    On Error GoTo SkipdlgChange
    frmDeComp.cdlg.ShowOpen
    On Error GoTo 0
    Source = frmDeComp.cdlg.filename
    frmDeComp.lblSource.Caption = Source
    frmDeComp.lblSource.Refresh
    
    
SkipdlgChange:
    On Error GoTo 0
        

End Sub

Private Sub cmdTarget_Click()
Dim rslt As Long
    
    rslt = vbNo
    rslt = MsgBox("Do you wish to overwrite the input file(s)?", vbYesNo, "DeCompiler")
    If rslt = vbYes Then
        Target = Source
    Else
        frmDeComp.cdlg.CancelError = True
        frmDeComp.cdlg.DefaultExt = "Expanded"
        frmDeComp.cdlg.DialogTitle = "Select the Target file location"
        If Target <> "" Then
            frmDeComp.cdlg.filename = FileTitle(Target)
            frmDeComp.cdlg.InitDir = FilePath(Target)
        Else
            frmDeComp.cdlg.filename = FileRoot(Source) & ".Expanded"
            frmDeComp.cdlg.InitDir = TSRoot
        End If
        frmDeComp.cdlg.Filter = "working Files (*.Expanded)|*.Expanded"
        frmDeComp.cdlg.FilterIndex = 0
        frmDeComp.cdlg.Flags = cdlOFNPathMustExist Or cdlOFNExplorer
        frmDeComp.lblTarget.Caption = ""
        
        On Error GoTo SkipdlgChange
        frmDeComp.cdlg.ShowOpen
        On Error GoTo 0
        Target = frmDeComp.cdlg.filename
    End If
    frmDeComp.lblTarget.Caption = Target
    frmDeComp.lblTarget.Refresh
    
SkipdlgChange:
    On Error GoTo 0
        
End Sub

Private Sub Form_Load()

    frmDeComp.Caption = App.ProductName
    
    Call InitValues

    TSRoot = GetTSRoot()

    Call Init_Tokens
    
    Source = GetSetting(App.EXEName, "Settings", "Source", "")
    Target = GetSetting(App.EXEName, "Settings", "Target", "")
    Kind = CLng(GetSetting(App.EXEName, "Settings", "Kind", "0"))
    
    frmDeComp.optType(0).value = False
    frmDeComp.optType(1).value = False
    frmDeComp.optType(2).value = False
    frmDeComp.optType(Kind).value = True
    
    frmDeComp.lblSource.Caption = Source
    frmDeComp.lblTarget.Caption = Target
    
    Me.Show
    
       
    
End Sub

Private Sub mnuAbout_Click()
    Load frmAbout
End Sub

Private Sub optType_Click(Index As Integer)
    Kind = Index
End Sub
