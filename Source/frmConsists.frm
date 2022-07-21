VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form frmConsists 
   Caption         =   "Stock used by Selected Consists."
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save List"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print List"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VSFlex8LCtl.VSFlexGrid Grid3 
      Height          =   4335
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   9015
      _cx             =   15901
      _cy             =   7646
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
      Rows            =   5000
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmConsists.frx":0000
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
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove Duplicates"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   5280
      Width           =   1575
   End
End
Attribute VB_Name = "frmConsists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim nextrow As Integer
Dim nextcol As Integer
Dim cn As Integer, cnn As Integer
Const REF_CHUNK = 100

Private Sub Command1_Click()
Dim i As Integer, strPath(0 To 1), strEng(0 To 1)

On Error GoTo ErrTrap

MousePointer = 11
i = 1
Do
Grid3.Select i, 0
strPath(0) = Grid3.Cell(flexcpText)
Grid3.Select i + 1, 0
strPath(1) = Grid3.Cell(flexcpText)
Grid3.Select i, 1
strEng(0) = Grid3.Cell(flexcpText)
Grid3.Select i + 1, 1
strEng(1) = Grid3.Cell(flexcpText)
If strPath(0) = strPath(1) And strEng(0) = strEng(1) Then
Grid3.RemoveItem i + 1
Else
i = i + 1
End If
Loop While i < Grid3.Rows
EndIt:
Grid3.Refresh
MousePointer = 0
Exit Sub
ErrTrap:
If Err = 381 Then GoTo EndIt
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
flagPrint = 17
fEZPrint.Show
End Sub

Private Sub Command4_Click()
Dim tit1 As String

CommonDialog1.Filter = "Comma Separated (*.csv)|*.csv|Tab Separated (*.txt)|*.txt"""
CommonDialog1.DialogTitle = "Save Grid as File"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 2
tit1 = CommonDialog1.Filename
If tit1 = vbNullString Then Exit Sub
If Right$(tit1, 3) = "csv" Then
Grid3.SaveGrid tit1, flexFileCommaText
ElseIf Right$(tit1, 3) = "txt" Then
Grid3.SaveGrid tit1, flexFileTabText
End If
End Sub

Private Sub Form_Load()
Dim i As Integer, Filpath$, Filpath1$, strCon As String, booSelected As Boolean

ReDim strConsists(0 To REF_CHUNK)

Grid3.AllowUserResizing = flexResizeBoth
  
   
   Grid3.ExplorerBar = flexExSort
   Grid3.BackColor = vbWhite
   Grid3.Rows = 1
nextrow = 1
nextcol = 0
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
Show
MousePointer = 11
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
  If frmUtils.File1(cursouind).Selected(i) Then
  booSelected = True
    Filpath$ = frmUtils.File1(cursouind).Path
   
   strCon = Filpath$ & "\" & frmUtils.File1(cursouind).List(i)
   If Right$(strCon, 3) <> "con" Then GoTo GetAnother
 
   Call CheckForConsistGrid(strCon, nextrow)
   
   
   End If
GetAnother:
   Next i
   If booSelected = False Then
   Call MsgBox("No Consists have been selected.", vbExclamation, App.Title)
   MousePointer = 0
   Exit Sub
   End If
   Grid3.Select 1, 0, 1, 1
   Grid3.Sort = flexSortStringAscending

   MousePointer = 0
End Sub


Private Sub CheckForConsistGrid(CFilepath As String, nextrow As Integer)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Engpath As String, Engname As String
Dim strNew2 As String, Wagonpath As String, Wagname As String
Dim TrainsetPath As String, ConName As String, strTemp As String, booCon As Boolean
Dim booEntry As Boolean

On Error GoTo ErrTrap


Fnumber = FreeFile
x = InStrRev(CFilepath, "\")
ConName = Mid$(CFilepath, x + 1)
strConsistNames = strConsistNames & ConName & vbCrLf
TrainsetPath = MSTSPath & "\Trains\Trainset\"

Open CFilepath For Input As Fnumber
 
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   
   strNew = Trim$(strNew)

   x = InStr(strNew, "EngineData")
   
   If x > 0 Then
   Call CheckEngineData(strNew, Engname, Engpath, booEntry)


   
   If booEntry = True Then
   If Not FileExists(TrainsetPath & Engpath & "\" & Engname & ".eng") Then
Grid3.AddItem Engpath & vbTab & Engname & ".eng" & vbTab & "*"
strTemp = Engpath & "\" & Engname & ".eng"
    For cnn = 0 To cn
    If strTemp = strConsists(cnn) Then
        booCon = True
        Exit For
    End If
    Next
    If booCon = False Then
    cn = cn + 1
            If cn > UBound(strConsists) Then
            ReDim Preserve strConsists(0 To cn + REF_CHUNK)
            End If
    strConsists(cn) = strTemp
    End If
booCon = False
nextrow = nextrow + 1
        If Grid3.Rows < nextrow Then
        Grid3.Rows = nextrow
        End If
   Else

Grid3.AddItem Engpath & vbTab & Engname & ".eng"
strTemp = Engpath & "\" & Engname & ".eng"
    For cnn = 0 To cn
    If strTemp = strConsists(cnn) Then
        booCon = True
        Exit For
    End If
    Next
    If booCon = False Then
    cn = cn + 1
            If cn > UBound(strConsists) Then
            ReDim Preserve strConsists(0 To cn + REF_CHUNK)
            End If
    strConsists(cn) = strTemp
    End If
booCon = False
nextrow = nextrow + 1
        If Grid3.Rows < nextrow Then
        Grid3.Rows = nextrow
        End If

   End If
   strNew2 = vbNullString
   End If
   End If
   booEntry = False
   x = InStr(strNew, "wagonData")
      If x > 0 Then
   Call CheckWagonData(strNew, Wagname, Wagonpath, booEntry)
   If booEntry = True Then
  If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname & ".wag") Then
Grid3.AddItem Wagonpath & vbTab & Wagname & ".wag" & vbTab & "*"
strTemp = Wagonpath & "\" & Wagname & ".wag"
For cnn = 0 To cn
If strTemp = strConsists(cnn) Then
booCon = True
Exit For
End If
Next
If booCon = False Then
cn = cn + 1
If cn > UBound(strConsists) Then
ReDim Preserve strConsists(0 To cn + REF_CHUNK)
End If
strConsists(cn) = strTemp
End If
booCon = False
nextrow = nextrow + 1
If Grid3.Rows < nextrow Then
Grid3.Rows = nextrow
End If
   Else

Grid3.AddItem Wagonpath & vbTab & Wagname & ".wag"
strTemp = Wagonpath & "\" & Wagname & ".wag"
For cnn = 0 To cn
If strTemp = strConsists(cnn) Then
booCon = True
Exit For
End If
Next
If booCon = False Then
cn = cn + 1
If cn > UBound(strConsists) Then
ReDim Preserve strConsists(0 To cn + REF_CHUNK)
End If
strConsists(cn) = strTemp
End If
booCon = False
nextrow = nextrow + 1
If Grid3.Rows < nextrow Then
Grid3.Rows = nextrow
End If

   End If
   
  End If
   End If
   strNew = vbNullString
 
   Loop
   Close #Fnumber
 
Exit Sub
ErrTrap:

If Err = 381 Then
Resume Next
Else
Call MsgBox("An error #" & Err & " occurred in subroutine 'CheckForConsistGrid' while checking" _
            & vbCrLf & CFilepath _
            , vbExclamation, frmConsists)

Resume Next
End If
End Sub




