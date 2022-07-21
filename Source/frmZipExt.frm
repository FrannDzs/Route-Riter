VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form frmNewZip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MyZipp"
   ClientHeight    =   9075
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   12420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8LCtl.VSFlexGrid VsFg1 
      Height          =   6135
      Left            =   360
      TabIndex        =   21
      Top             =   1920
      Width           =   11535
      _cx             =   20346
      _cy             =   10821
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmZipExt.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   840
      TabIndex        =   17
      ToolTipText     =   "Adjust zip compression 0=None 9=Highest (but slow)"
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Max             =   9
      SelStart        =   6
      TickStyle       =   1
      Value           =   6
      TextPosition    =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Quit"
      Height          =   255
      Left            =   10560
      TabIndex        =   16
      Top             =   600
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   8040
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Route Compression Options"
      Height          =   1335
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton Command2 
         Caption         =   "Process"
         Height          =   255
         Left            =   6120
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Produce multiple Zip files?"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Approx. size of each .zip in Mbytes."
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label Label4 
         Caption         =   "Zip File name: "
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   6855
      End
      Begin VB.Label Label3 
         Caption         =   "Route Name: "
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   10560
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1335
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   7335
      Begin VB.CheckBox Check1 
         Caption         =   "Use Relative Paths"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Include Sub-Directories"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Fullpath"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Wildcards"
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog dlgAdd 
      Left            =   11760
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Add a file to the archive"
      MaxFileSize     =   32000
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8700
      Width           =   12420
      _ExtentX        =   21908
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Processing:"
            TextSave        =   "Processing:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   11760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "High"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   20
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Low"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   19
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Compression Level  =   6"
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   960
      Width           =   1815
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "Archives"
      Begin VB.Menu mnuNew 
         Caption         =   "New Archive"
      End
      Begin VB.Menu mnuMulti 
         Caption         =   "New Multi-Part Archive"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Archive"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close Archive"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuProcess 
      Caption         =   "Extract Files"
      Begin VB.Menu mnuExt 
         Caption         =   "Extract All"
      End
      Begin VB.Menu mnuExtSel 
         Caption         =   "Extract Selected"
      End
      Begin VB.Menu mnuExtCur 
         Caption         =   "Exract to Current Folder"
      End
   End
   Begin VB.Menu mnuAddFiles 
      Caption         =   "Add Files"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add File(s)"
      End
      Begin VB.Menu mnuAddFold 
         Caption         =   "Add Folder"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "Search"
      Begin VB.Menu mnuFind 
         Caption         =   "Find File"
      End
      Begin VB.Menu mnuFindFiles 
         Caption         =   "Find Files(Pattern Match)"
      End
   End
End
Attribute VB_Name = "frmNewZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text


Public WithEvents Zip As sawzipng.Archive
Attribute Zip.VB_VarHelpID = -1

Dim Lfile As sawzipng.FileInfo
Dim strFile() As String
Dim tempRow() As Long
Dim booFullPath As Boolean
Dim booSubDir As Boolean
Dim strWildcards As String
Dim CompLevel As Integer
Dim booRelative As Boolean
Dim strRelative As String
Dim ExtFile As Long
Dim strSavePath As String
Private Sub FillGrid()
Dim i As Long, strTemp As String
Dim indexes() As Variant
Dim names(0) As Variant
On Error GoTo ErrTrap

VsFg1.Rows = 1
For i = 0 To Zip.Count - 1
Set Lfile = Zip.GetFileInfo(i)
SB1.Panels(2).Text = Lfile.Filename
strTemp = Lfile.Filename & vbTab & Lfile.CompressionSize & vbTab & Lfile.UncompressedSize & vbTab
strTemp = strTemp & "     " & Lfile.ModificationDate
names(0) = Lfile.Filename
indexes = Zip.GetIndexes(names)
If indexes(0) > -1 Then
strTemp = Str(indexes(0)) & vbTab & strTemp
End If
Label1:
VsFg1.AddItem strTemp

Next i
SB1.Panels(3).Text = Str(Zip.DirCount)
SB1.Panels(4).Text = Str(Zip.FileCount)
MousePointer = 0

Exit Sub
ErrTrap:
GoTo Label1
End Sub








Private Sub SetLang()
Me.Caption = Lang(249)
Label6(0).Caption = Lang(250)
Label6(1).Caption = Lang(251)
Label2.Caption = Lang(252)
Frame2.Caption = Lang(253)
Label3.Caption = Lang(254)
Label4.Caption = Lang(255)
Label5.Caption = Lang(257)
Check2.Caption = Lang(256)
Command2.Caption = Lang(258)
mnuFiles.Caption = Lang(598)
mnuNew.Caption = Lang(599)
mnuMulti.Caption = Lang(600)
mnuOpen.Caption = Lang(601)
mnuClose.Caption = Lang(602)
mnuExit.Caption = Lang(38)
mnuProcess.Caption = Lang(603)
mnuExt.Caption = Lang(604)
mnuExtSel.Caption = Lang(605)
mnuExtCur.Caption = Lang(606)
mnuAddFiles.Caption = Lang(607)
mnuAdd.Caption = Lang(608)
mnuAddFold.Caption = Lang(609)
mnuSearch.Caption = Lang(610)
mnuFind.Caption = Lang(611)
mnuFindFiles.Caption = Lang(612)


End Sub

Private Sub Check1_Click(Index As Integer)
If Index = 0 Then
If Check1(0).Value = 0 Then
booFullPath = False
Else
booFullPath = True
Check1(2).Value = 0
strRelative = vbNullString
End If
End If
If Index = 1 Then
If Check1(1).Value = 0 Then
booSubDir = False
Else
booSubDir = True
End If
End If
If Index = 2 Then
If Check1(2).Value = 0 Then
booRelative = False
strRelative = vbNullString
ElseIf Check1(2).Value = 1 Then
booRelative = True
booFullPath = False
Check1(0).Value = 0


End If
End If
End Sub

Private Sub Command1_Click()
Call FillGrid
End Sub

Private Sub Command2_Click()

Dim PartLength As Long, strFolder As String

If Check2.Value = 1 Then
If Val(Text2) = 0 Then
Call MsgBox(Lang(541) & vbCrLf & Lang(542), vbExclamation, App.Title)

Exit Sub
End If
CDL1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt
CDL1.Filename = ZipName

    CDL1.ShowSave
    If CDL1.Filename = vbNullString Then
    Exit Sub
    End If
    If Len(CDL1.Filename) > 0 Then
    strZipName = CDL1.Filename
    Label4.Caption = "Zip File name:  " & strZipName & ".zip"

    If Right$(strZipName, 4) = ".zip" Then
    strZipName = Left$(strZipName, Len(strZipName) - 4)
    End If
    DoEvents
    
    PartLength = Val(Text2)
    PartLength = PartLength * 1000000
    MousePointer = 11
        If Zip Is Nothing Then
            Set Zip = New Archive
           
        Else
            Zip.Close
        End If
        Zip.Create CDL1.Filename, CM_CREATE_SPAN, PartLength
        booMulti = True
        End If
        
ElseIf Check2.Value = 0 Then
CDL1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt
CDL1.Filter = "Zip|*.zip"
CDL1.Filename = ZipName

    CDL1.ShowSave
    If CDL1.Filename = vbNullString Then
    Exit Sub
    End If
    If Len(CDL1.Filename) > 0 Then
    strZipName = CDL1.Filename
    Label4.Caption = "Zip File name:  " & strZipName

    If Right$(strZipName, 4) <> ".zip" Then
    strZipName = strZipName & ".zip"
    
    End If
    DoEvents
   
    MousePointer = 11
        If Zip Is Nothing Then
            Set Zip = New Archive
           
        Else
            Zip.Close
        End If
        Zip.Create CDL1.Filename, CM_CREATE
        booMulti = False
        End If
        

End If
strFolder = NewZipPath
DoEvents

x = InStrRev(strFolder, "\")
strRelative = Left$(strFolder, x - 1)
Zip.RootPath = strRelative

Zip.AddFolder strFolder, booSubDir, booFullPath, CompLevel, SM_SMART_SAFE, 65536
DoEvents

Call FillGrid

MousePointer = 0

End Sub

Private Sub Command3_Click()
Unload Me

End Sub


Private Sub Form_Load()
Call SetLang
booFullPath = False
booSubDir = True
booRelative = True
booMulti = False
strWildcards = Text1
CompLevel = Slider1.Value
If FromZip = 1 Then
Frame2.Visible = True
Label3.Caption = "Route Path:  " & NewZipPath
Label4.Caption = "Zip File name:  " & ZipName & ".zip"

End If
If FromZip = 2 Then
    Call mnuNew_Click
    Zip.RootPath = NewZipPath
    If Len(CDL1.Filename) = 0 Then Exit Sub
    Zip.AddFolder NewZipPath & "Routes", True, False
    Zip.AddFolder NewZipPath & "Trains", True, False
    Call FillGrid

      End If
     
FromZip = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrTrap
 
        If Not Zip.Closed Then
            Zip.Close
            
        End If
   
    Unload Me
Exit Sub
ErrTrap:
Exit Sub
    
End Sub


Private Sub mnuAdd_Click()
Dim x As Integer, strPath As String, a$, xx As Integer, Y As Integer
Dim xy As Integer, i As Integer, Z As Integer
On Error GoTo ErrTrap

PB1.Visible = False
If Zip Is Nothing Then
Call MsgBox(Lang(543) & vbCrLf & Lang(544), vbExclamation, "Error")

Exit Sub
End If
dlgAdd.Flags = cdlOFNLongNames + cdlOFNAllowMultiselect
dlgAdd.DialogTitle = "Choose File(s) to Add"
dlgAdd.Filename = "*.*"
dlgAdd.ShowOpen
If Len(dlgAdd.Filename) > 0 Then
a$ = dlgAdd.Filename
x = InStr(a$, " ")
If x = 0 Then
If booRelative = True Then
Z = InStrRev(a$, "\")
Z = InStrRev(a$, "\", Z - 1)
strRelative = Left$(a$, Z - 1)
Zip.RootPath = strRelative
End If


Zip.AddFile a$, booFullPath, CompLevel, SM_SMART_SAFE, 65536

Call FillGrid
Exit Sub
End If

strPath = Left$(a$, x - 1)

xx = x + 1
xy = xx
Do
x = InStr(xx, a$, " ")
Y = Y + 1
xx = x + 1
Loop While x > 0
ReDim strFile(0 To Y - 1)
xx = xy

For i = 0 To Y - 1
x = InStr(xx, a$, " ")
If x = 0 Then
strFile(i) = Mid$(a$, xx)
Exit For
End If
strFile(i) = Mid$(a$, xx, x - xx)
xx = x + 1
Next i

For i = 0 To Y - 1
If booRelative = True Then
Z = InStrRev(strPath, "\")
Z = InStrRev(a$, "\", Z - 1)
strRelative = Left$(strPath, Z - 1)
Zip.RootPath = strRelative
End If
Zip.AddFile strPath & strFile(i), booFullPath, CompLevel, SM_SMART_SAFE, 65536
Next i
Call FillGrid
End If


Exit Sub
ErrTrap:

Call MsgBox("You have selected more than the maximum number of files possible at one time. Reduce your" _
            & vbCrLf & "                                                      selection and try again." _
            , vbCritical, "An Error Occurred")


End Sub

Private Sub mnuAddFiles_Click()
Check1(1).Value = 1
Check1(2).Value = 1
End Sub

Private Sub mnuAddFold_Click()
Dim Dir As String, strFolder As String, strRelative As String, x As Integer

PB1.Visible = False
If Zip Is Nothing Then
Call MsgBox(Lang(543) & vbCrLf & Lang(544), vbExclamation, "Error")

Exit Sub
End If
Dir = SelectDir(Me.hwnd, "Select folder to add")
        If Len(Dir) > 0 Then
strFolder = Dir
        End If
MousePointer = 11
ZOrder

If booRelative = True Then
x = InStrRev(strFolder, "\")
strRelative = Left$(strFolder, x - 1)
Zip.RootPath = strRelative
End If



If Len(strFolder) > 0 And Text1 = vbNullString Then

Zip.AddFolder strFolder, booSubDir, booFullPath, CompLevel, SM_SMART_SAFE, 65536
Call FillGrid
ElseIf Len(strFolder) > 0 And Text1 <> vbNullString Then
Zip.AddFolderWithWildcard strFolder, strWildcards, booSubDir, booFullPath, CompLevel, SM_SMART_SAFE, 65536
Call FillGrid
End If
MousePointer = 0
'Zip.Close
'If booMulti = True Then
'Name strZipName As strZipName & ".zip"
'End If
'Zip.AddFolderWithWildcard "d:\msts-backups\portogden", "*.*", True, True, -1, SM_SMART_SAFE, 65536
End Sub


Private Sub mnuClose_Click()
 If Not Zip.Closed Then
            Zip.Close
            VsFg1.Rows = 1
        End If
End Sub

Private Sub mnuExit_Click()
Unload Me

End Sub


Private Sub mnuExt_Click()
Dim i As Long, Dir As String

On Error GoTo ErrTrap
ExtFile = 1
PB1.Visible = True
If Zip Is Nothing Then
Call MsgBox(Lang(543) & vbCrLf & Lang(545), vbExclamation, "Error")

Exit Sub
End If
Dir = SelectDir(Me.hwnd, "Select destination directory")
        If Len(Dir) > 0 Then
strSavePath = Dir
        End If


strSavePath = InputBox("Save to: ", "Confirm Save Path", strSavePath, 3360, 1080)
If strSavePath = vbNullString Then Exit Sub
MousePointer = 11
If Len(strSavePath) > 0 Then

If Zip.FileCount > 0 Then
   i = 0
   Do
   
       Zip.Extract i, strSavePath, booFullPath
       i = i + 1
   Loop While i < Zip.FileCount
End If

End If
MousePointer = 0
SB1.Panels(2).Text = "Extract Completed"
Exit Sub
ErrTrap:

End Sub

Private Sub mnuExtCur_Click()
Dim x As Integer, i As Long


ExtFile = 1
PB1.Visible = True
If Zip Is Nothing Then
Call MsgBox(Lang(543) & vbCrLf & Lang(545), vbExclamation, "Error")

Exit Sub
End If
MousePointer = 11
If Len(CDL1.Filename) > 0 Then
x = InStrRev(CDL1.Filename, "\")
strSavePath = Left$(CDL1.Filename, x - 1)
Else
Exit Sub
End If
If Zip.FileCount > 0 Then
   i = 0
   Do
       Zip.Extract i, strSavePath, False
       i = i + 1
   Loop While i < Zip.FileCount
End If
MousePointer = 0
SB1.Panels(2).Text = "Extract Completed"
End Sub

Private Sub mnuExtSel_Click()
Dim i As Long, Dir As String, tempTtl As Long
Dim tempIdx As Long

ExtFile = 1
PB1.Visible = True
If Zip Is Nothing Then
Call MsgBox(Lang(543) & vbCrLf & Lang(545), vbExclamation, "Error")

Exit Sub
End If

Dir = SelectDir(Me.hwnd, "Select destination directory")
        If Len(Dir) > 0 Then
strSavePath = Dir
        End If


strSavePath = InputBox("Save Selected Files to: ", "Confirm Save Path", strSavePath, 3360, 1080)
If strSavePath = vbNullString Then Exit Sub

MousePointer = 11
If Len(strSavePath) > 0 Then
ReDim tempRow(VsFg1.SelectedRows - 1)

For i = 0 To VsFg1.SelectedRows - 1
tempRow(i) = VsFg1.SelectedRow(i)
Next i
tempTtl = VsFg1.SelectedRows - 1

For i = 0 To tempTtl
VsFg1.Select tempRow(i), 0
tempIdx = Val(VsFg1.Cell(flexcpText))

       Zip.Extract tempIdx, strSavePath, booFullPath
       
 Next i

End If
MousePointer = 0
SB1.Panels(2).Text = Lang(547)
End Sub


Private Sub mnuFind_Click()
Dim FoundIndex As Long, strFile As String


If Zip Is Nothing Then
Call MsgBox(Lang(543) & vbCrLf & Lang(546), vbExclamation, "Error")

Exit Sub
End If

strFile = InputBox("Filename: ", "Search For", , 3360, 2080)
If strFile = vbNullString Then Exit Sub
FoundIndex = Zip.FindFile(strFile, FF_DEFAULT, Not booFullPath)
If FoundIndex = -1 Then
VsFg1.Rows = 1
      Call MsgBox("The file  " & strFile _
                  & vbCrLf & "was not found." _
                  , vbExclamation, App.Title)
      
Exit Sub
End If

VsFg1.Rows = 1

Set Lfile = Zip.GetFileInfo(FoundIndex)
SB1.Panels(2).Text = Lfile.Filename
strTemp = Str(FoundIndex) & vbTab & Lfile.Filename & vbTab & Lfile.CompressionSize & vbTab & Lfile.UncompressedSize & vbTab
strTemp = strTemp & "     " & Lfile.ModificationDate
Label1:
VsFg1.AddItem strTemp



End Sub

Private Sub mnuFindFiles_Click()
Dim strFile As String
Dim indexes() As Variant
Dim i As Long, strTemp As String
On Error GoTo ErrTrap

If Zip Is Nothing Then
Call MsgBox(Lang(543) & vbCrLf & Lang(546), vbExclamation, "Error")

Exit Sub
End If
strFile = InputBox("Pattern: ", "Search For", , 3360, 1080)
If strFile = vbNullString Then Exit Sub
indexes = Zip.FindFiles(strFile, False)
VsFg1.Rows = 1
For i = LBound(indexes) To UBound(indexes)
Set Lfile = Zip.GetFileInfo(indexes(i))
SB1.Panels(2).Text = Lfile.Filename
strTemp = Str(indexes(i)) & vbTab & Lfile.Filename & vbTab & Lfile.CompressionSize & vbTab & Lfile.UncompressedSize & vbTab
strTemp = strTemp & "     " & Lfile.ModificationDate
Label1:
VsFg1.AddItem strTemp

Next i

Exit Sub
ErrTrap:
GoTo Label1


End Sub


Private Sub mnuMulti_Click()
Dim PartLength As Long


CDL1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt

    CDL1.ShowSave
    
    If Len(CDL1.Filename) > 0 Then
    strZipName = CDL1.Filename
    If Right$(strZipName, 4) = ".zip" Then
    strZipName = Left$(strZipName, Len(strZipName) - 4)
    End If
    PartLength = Val(InputBox("Enter approx. section length in Mb", "Multi-Part Zip", "10", 3360, 1080))
    PartLength = PartLength * 1000000
        If Zip Is Nothing Then
            Set Zip = New Archive
           
        Else
            Zip.Close
        End If
        Zip.Create CDL1.Filename, CM_CREATE_SPAN, PartLength
        booMulti = True
        End If
End Sub

Private Sub mnuNew_Click()
CDL1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt
CDL1.Filter = "Zip|*.zip"

    CDL1.ShowSave
    If Len(CDL1.Filename) > 0 Then
        If Zip Is Nothing Then
            Set Zip = New Archive
           
        Else
            Zip.Close
        End If
        Zip.Create CDL1.Filename, CM_CREATE
        End If
        'zip.Password = "test"
End Sub

Private Sub mnuOpen_Click()
Dim x As Integer, i As Integer, strTempZip As String
Dim strTempPath As String, strInt As String, strInt2 As String

CDL1.ShowOpen

    If Len(CDL1.Filename) > 0 Then
        If Zip Is Nothing Then
            Set Zip = New Archive
        Else
            Zip.Close
        End If
        
        Zip.Open CDL1.Filename, OM_OPEN
       
        frmNewZip.Caption = "Archive - " & CDL1.Filename
        x = InStrRev(CDL1.Filename, "\")
        strTempPath = Left$(CDL1.Filename, x)
        strTempZip = Mid$(CDL1.Filename, x + 1)
        strTempZip = Left$(strTempZip, Len(strTempZip) - 3)
        For i = 1 To 99
        strInt = Trim$(Str(i))
        If Len(strInt) = 1 Then
        strInt = "0" & strInt
        End If
        If FileExists(strTempPath & strTempZip & "z" & strInt) Then
        strInt2 = Trim$(Str(i - 1))
        If Len(strInt2) = 1 Then
        strInt2 = "0" & strInt2
        End If
        Name strTempPath & strTempZip & "z" & strInt As strTempPath & strTempZip & "0" & strInt2
        Else
        Exit For
        End If
        
        Next i
        
        DoEvents
        Call FillGrid
        End If
        
End Sub





Private Sub mnuProcess_Click()
'Check1(0).Value = 1

Check1(1).Value = 2
Check1(2).Value = 2
End Sub

Private Sub Slider1_Change()
Label2.Caption = "Compression Level = " & Str(Slider1.Value)
CompLevel = Slider1.Value

End Sub

Private Sub Text1_Change()
strWildcards = Text1
End Sub


Private Sub Zip_OnAdd(ByVal Filename As String, ByVal soFar As Long, ByVal ToDo As Long, Cancel As Boolean)
SB1.Panels(2).Text = Filename
End Sub

Private Sub Zip_OnExtract(ByVal Filename As String, ByVal soFar As Long, ByVal ToDo As Long, Cancel As Boolean)


SB1.Panels(2).Text = Filename


If (ExtFile * 100) / Zip.FileCount <= 100 Then
PB1.Value = (ExtFile * 100) / Zip.FileCount
End If
ExtFile = ExtFile + 1
End Sub


