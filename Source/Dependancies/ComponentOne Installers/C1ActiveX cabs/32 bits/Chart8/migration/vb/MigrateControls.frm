VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   2070
   ClientTop       =   2955
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Top             =   6120
      Width           =   1215
   End
   Begin VB.ListBox lstConvertedModules 
      Height          =   1815
      Left            =   360
      TabIndex        =   9
      Top             =   3840
      Width           =   6615
   End
   Begin VB.CommandButton cmdStartOver 
      Caption         =   "Start &Over"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select Project..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   7080
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstTo 
      Height          =   840
      Left            =   3600
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
   End
   Begin VB.ListBox lstFrom 
      Height          =   840
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label lblFinish 
      Caption         =   "lblFinish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   13
      Top             =   6000
      Width           =   4455
   End
   Begin VB.Label lblThree 
      Alignment       =   2  'Center
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label lblTwo 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblOne 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblChooseControls 
      Caption         =   "lblChoseControls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label lblProjectName 
      Alignment       =   1  'Right Justify
      Caption         =   "lblProjectName"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label lblTo 
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblFrom 
      Caption         =   "From:"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Cntrls As New clsControlData

Private projectPath As String
Private projectName As String
Private vbpContents As String

Private Sub doStep1(onoff As Boolean)
  cmdSelect.Enabled = onoff
  lblOne.Enabled = onoff
  If onoff Then
    lblProjectName.Caption = ""
    lstFrom.Clear
    lstTo.Clear
    lstConvertedModules.Clear
  End If
End Sub

Private Sub doStep2(onoff As Boolean)
  lblTwo.Enabled = onoff
  lblFrom.Enabled = onoff
  lblTo.Enabled = onoff
  lstFrom.Enabled = onoff
  lstTo.Enabled = onoff
  lblChooseControls.Enabled = onoff
  cmdContinue.Enabled = onoff
  cmdStartOver.Enabled = onoff
End Sub

Private Sub doStep3(onoff As Boolean)
  lblThree.Enabled = onoff
  lblFinish.Enabled = onoff
  lblFinish.Visible = onoff
  cmdFinish.Enabled = onoff
End Sub

Private Sub cmdSelect_Click()
  Dim i As Integer, j As Integer
  
  '=======================================================
  ' open existing project
  
  ' get VB project name
  '
  dlg.CancelError = False
  dlg.FileName = ""
  dlg.DefaultExt = "vbp"
  dlg.Filter = "Visual Basic Projects (*.vbp)|*.vbp"
  dlg.Flags = cdlOFNFileMustExist
  dlg.ShowOpen
  projectName = dlg.FileName
  If projectName = "" Then Exit Sub
  
  lblProjectName.Caption = projectName

  ' parse project path and file name
  '
  i = 0
  While InStr(i + 1, projectName, "\") > 0
      i = InStr(i + 1, projectName, "\")
  Wend
  projectPath = Left(projectName, i)
  projectName = Mid(projectName, i + 1)
  
  ' read the whole project file
  '
  Open projectPath & projectName For Binary As #1
  vbpContents = Input(LOF(1), 1)
  Close #1

  '=======================================================
  ' make sure it has an appropriate Control project
  '
  If Not Cntrls.CheckForControlUsage(vbpContents) Then
      Call MsgBox("This project does not use " & Cntrls.ControlsTitle _
        & vbCrLf & vbCrLf & "Conversion cancelled.", vbInformation)
      Exit Sub
  End If
  
  '=======================================================
  ' allow control group selection
  '
  Call doStep1(False)
  Call doStep2(True)
  
  Call Cntrls.FillFromListbox(lstFrom)
  Call Cntrls.FillToListbox(lstTo)

End Sub

Private Sub pickModules(modIDStr As String, modExtension As String)
  Dim i As Integer, j As Integer
  Dim iLen As Integer, eLen As Integer
  iLen = Len(modIDStr)
  eLen = Len(modExtension)
  
  i = 1
  Do
    i = InStr(i, vbpContents, modIDStr)
    If i = 0 Then Exit Do
    If Mid(vbpContents, i - 1, 1) = vbLf Then
      If Mid(vbpContents, i + iLen, 1) <> """" Then
        j = InStr(i, vbpContents, modExtension)
        If j <> 0 Then
          Call lstConvertedModules.AddItem(Mid(vbpContents, i + iLen, j - i - iLen + eLen))
        End If
      End If
    End If
    i = i + 1
  Loop
End Sub

Private Sub pickNamedModules(modIDStr As String, modExtension As String)
  Dim i As Integer, j As Integer
  Dim iLen As Integer, eLen As Integer
  iLen = Len(modIDStr)
  eLen = Len(modExtension)
  
  i = 1
  Do
    i = InStr(i, vbpContents, modIDStr)
    If i = 0 Then Exit Do
    If Mid(vbpContents, i - 1, 1) = vbLf Then
      j = InStr(i, vbpContents, ";")
      If j <> 0 Then
        j = j + 1
        While Mid(vbpContents, j, 1) = " "
          j = j + 1
        Wend
        i = j
        j = InStr(i, vbpContents, modExtension)
        If j <> 0 Then
          Call lstConvertedModules.AddItem(Mid(vbpContents, i, j - i + eLen))
        End If
      End If
    End If
    i = i + 1
  Loop
End Sub

Private Sub ConvertModule(fn As String)
  Dim buf As String, pth As String
  Dim fnModified As Boolean

  ' if the file has a path spec, don't prepend the project path
  '
  If InStr(fn, "\") = 0 Then
    pth = ""
  Else
    pth = projectPath
  End If

  ' make a backup copy of the file
  ' we just append '.bak' to the file name
  ' the file will have two extensions, which is OK in Win32
  '
  FileCopy pth & fn, pth & fn & ".bak"
  
  ' open files to read (backup) and write (converted)
  '
  Open pth & fn & ".bak" For Input As #1
  Open pth & fn For Output As #2
  
  fnModified = False
  
  ' scan the file
  '
  While Not EOF(1)
    ' read a line
    '
    Line Input #1, buf
    
    ' make any convertions
    If Cntrls.ConvertControls(buf) Then
      fnModified = True
    End If
        
    Print #2, buf
  Wend
  ' done with this module, close both files
  '
  Close #1
  Close #2

  If Not fnModified Then
    Call Kill(pth & fn & ".bak")
  End If

End Sub

Private Sub doConversions()
  
  '=======================================================
  ' convert project
  Call Cntrls.ClearUnselectedControls(lstFrom)
  Call Cntrls.SetTargetControls(lstTo)

  ' replace library references.
  '
  Call Cntrls.ConvertVBPControls(vbpContents)
  
  ' make a backup copy of the project file
  ' we just append '.bak' to the file name
  ' the file will have two extensions, which is OK in Win32
  '
  FileCopy projectName, projectName & ".bak"
  
  ' write out updated project file
  '
  Open projectName For Output As #1
  Print #1, vbpContents
  Close #1
  
  lstConvertedModules.AddItem projectName & "... OK"
  
  ' find all modules in the project
  '
  Call pickModules("Form=", ".frm")
  Call pickModules("UserControl=", ".ctl")
  Call pickNamedModules("Module=", ".bas")
  Call pickNamedModules("Class=", ".cls")

  vbpContents = ""

  ' process files
  '
  Dim i As Integer
  For i = 1 To lstConvertedModules.ListCount - 1
    Call ConvertModule(lstConvertedModules.List(i))
    lstConvertedModules.ListIndex = i
    lstConvertedModules.List(i) = lstConvertedModules.List(i) & "... OK"
    Refresh
    DoEvents
  Next
  
  ' show confirmation and warnings
  '
  lblFinish.Caption = "Project '" & projectName & "' has been converted to " & lstTo.List(lstTo.ListIndex) & vbCrLf & vbCrLf & _
      "Copies of the original files were saved with a 'bak' extension in case you want to restore the original project." & vbCrLf & vbCrLf & _
      "You should now open the new project in VB and compile it to make sure there are no syntax errors." & vbCrLf & vbCrLf
End Sub

Private Sub cmdContinue_Click()
  If lstFrom.ListCount = 1 Then
    If lstFrom.List(0) = lstTo.List(lstTo.ListIndex) Then
      Call MsgBox("All " & Cntrls.ControlsTitle & " already match the target controls." _
        & vbCrLf & vbCrLf & "Conversion cancelled.", vbInformation)
      Exit Sub
    End If
  End If
  
  Call doStep2(False)
  Me.Enabled = False
  
  Call doConversions
  
  Me.Enabled = True
  Call doStep3(True)
End Sub

Private Sub cmdStartOver_Click()
  vbpContents = ""

  Call doStep2(False)
  Call doStep1(True)
End Sub

Private Sub cmdFinish_Click()
  Call doStep3(False)
  Call doStep1(True)
End Sub

Private Sub Form_Load()
  Call doStep1(True)
  Call doStep2(False)
  Call doStep3(False)
  
  Cntrls.LoadControlData
  
  Me.Caption = "Migration of " + Cntrls.ControlsTitle
  lblChooseControls.Caption = "Select the appropriate set of " + _
    Cntrls.ControlsTitle + " in each list, and press Continue to " + _
    "begin, or press Start Over to select a different project."
  
  Debug.Print "Height: "; Me.Height / Screen.TwipsPerPixelY
  Debug.Print "Width : "; Me.Width / Screen.TwipsPerPixelX
End Sub

