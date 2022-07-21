VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmGrid 
   Caption         =   "Activities & Their Associated Files."
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   14250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Show Errors Only"
      Height          =   615
      Left            =   9960
      TabIndex        =   11
      Top             =   9000
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13680
      Top             =   9720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Save as CSV"
      Height          =   615
      Left            =   8760
      TabIndex        =   10
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Error Report"
      Height          =   615
      Left            =   7560
      TabIndex        =   9
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Print"
      Height          =   615
      Left            =   12720
      TabIndex        =   8
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Check Consists"
      Height          =   615
      Left            =   6360
      TabIndex        =   7
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Unused Services && Consists"
      Height          =   615
      Left            =   5160
      TabIndex        =   6
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print Sheet"
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Sheet Format"
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   9000
      Width           =   1095
   End
   Begin VSFlex8LCtl.VSFlexGrid Grid2 
      Height          =   8535
      Left            =   11160
      TabIndex        =   2
      Top             =   240
      Width           =   2895
      _cx             =   5106
      _cy             =   15055
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   3
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmGrid.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
      ExplorerBar     =   3
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   11400
      TabIndex        =   1
      Top             =   9000
      Width           =   1095
   End
   Begin VSFlex8LCtl.VSFlexGrid Grid1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10935
      _cx             =   19288
      _cy             =   15055
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmGrid.frx":0038
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Items with a RED background are missing ! Or Faulty"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   9000
      Width           =   2535
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim booSelect As Boolean

Dim ConDown As Boolean
Dim flagPathRed As Boolean
Dim flagTfcRed As Boolean
Dim flagEngRed As Boolean
Private Sub LookForWag(strWag As String, strFoundPath As String)
Dim TrainsetPath As String, i As Integer
On Error GoTo Errtrap
TrainsetPath = MSTSPath & "\Trains\Trainset\"
For i = 0 To lngWagons - 1
If strWag = Wagons(i) Then

If FileExists(TrainsetPath & Wagpath(i) & "\" & strWag) Then
strFoundPath = Wagpath(i) & "\" & strWag
End If
End If
Next i
Exit Sub
Errtrap:

Call MsgBox("An Error number " & Err & " occurred in LookForWag" _
            & vbCrLf & "Error description: " & Err.Description _
            , vbExclamation, frmGrid)
End Sub

Private Sub LookForLoco(strEng As String, strFoundPath As String)
Dim TrainsetPath As String
On Error GoTo Errtrap
TrainsetPath = MSTSPath & "\Trains\Trainset\"
For i = 0 To lngLoco - 1
If strEng = Locomotives(i) Then

If FileExists(TrainsetPath & LocoPath(i) & "\" & strEng) Then
strFoundPath = LocoPath(i) & "\" & strEng
End If
End If
Next i
Exit Sub
Errtrap:


Call MsgBox("An Error number " & Err & " occurred in LookForLoco" _
            & vbCrLf & "Error description: " & Err.Description _
            , vbExclamation, frmGrid)

End Sub

Private Sub SetLang()
Dim i As Integer

On Error GoTo Errtrap
Me.Caption = Lang(236)
Label1.Caption = Lang(183)
Command2.Caption = Lang(237)
Command2.ToolTipText = Lang(238)
Command3.Caption = Lang(239)
Command3.ToolTipText = Lang(240)
Command4.Caption = Lang(241)
Command4.ToolTipText = Lang(242)
Command5.Caption = Lang(243)
Command5.ToolTipText = Lang(244)
Command7.Caption = Lang(245)
Command7.ToolTipText = Lang(246)
Command8.Caption = Lang(247)
Command8.ToolTipText = Lang(248)
Command1.Caption = Lang(38)

Command6.Caption = Lang(367)
For i = 0 To 5
Grid1.Select 0, i
Grid1.Cell(flexcpText) = Lang(591 + i)
DoEvents
Next

Grid2.Select 0, 0
Grid2.Cell(flexcpText) = Lang(597)
Exit Sub
Errtrap:
Call MsgBox("An Error number " & Err & " occurred in SetLang" _
            & vbCrLf & "Error description: " & Err.Description _
            , vbExclamation, frmGrid)
End Sub

Private Sub Command1_Click()
Grid1.Clear
Grid2.Clear
Unload Me

End Sub


Private Sub Command2_Click()

Grid1.MergeCol(0) = True
Grid1.MergeCol(1) = True
Grid1.MergeCol(2) = True

Grid1.BackColor = vbWhite
Grid2.BackColor = vbWhite

If Grid1.MergeCells = 0 Then
        Grid1.MergeCells = 2
    Else
        Grid1.MergeCells = 0
    End If
End Sub

Private Sub Command3_Click()
flagPrint = 4
fEZPrint.Show

End Sub

Private Sub Command4_Click()
If booUnusedChanged = True Then
booUnusedChanged = False
Call MsgBox(Lang(505) & vbCrLf & Lang(506), vbExclamation, App.Title)
Exit Sub
End If
Screen.MousePointer = 11
frmUnusedSrv.Show

End Sub

Private Sub Command5_Click()
frmUtils.Command15.value = True

'frmStock.Show
End Sub

Private Sub Command6_Click()
flagPrint = 10
fEZPrint.Show

End Sub

Private Sub Command7_Click()
frmReport.Rich1.Text = strbadbits & strForPrint & strReport
frmReport.Show 1
End Sub

Private Sub Command8_Click()
Dim tit1 As String

CommonDialog1.Filter = "Comma Separated (*.csv)|*.csv"
CommonDialog1.DialogTitle = "Save Grid as CSV File"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 2
tit1 = CommonDialog1.Filename
If tit1 <> vbNullString Then
Grid1.SaveGrid tit1, flexFileCommaText
End If
End Sub

Private Sub Command9_Click()
Dim i As Long

If Command9.Caption = "Show Errors Only" Then
Command9.Caption = "Show All"
For i = Grid1.FixedRows To Grid1.Rows - 1


Grid1.Select i, 1
If Grid1.CellBackColor = &H8080FF Then
Grid1.RowHidden(i) = False
GoTo NextOne
End If
Grid1.Select i, 2
If Grid1.CellBackColor = &H8080FF Then
Grid1.RowHidden(i) = False
GoTo NextOne
End If
Grid1.Select i, 3
If Grid1.CellBackColor = &H8080FF Then
Grid1.RowHidden(i) = False
GoTo NextOne
End If
Grid1.Select i, 4
If Grid1.CellBackColor = &H8080FF Then
Grid1.RowHidden(i) = False
GoTo NextOne
End If
Grid1.Select i, 5
If Grid1.CellBackColor = &H8080FF Then
Grid1.RowHidden(i) = False
GoTo NextOne
End If
Grid1.RowHidden(i) = True
NextOne:
Next i

ElseIf Command9.Caption = "Show All" Then
Command9.Caption = "Show Errors Only"
For i = Grid1.FixedRows To Grid1.Rows - 1



Grid1.Select i, 0

Grid1.RowHidden(i) = False

Next i
End If
End Sub

Private Sub Form_Load()
Dim i%, tempPath As String, x As Integer, Y As Integer, j As Integer, SPath As String
Dim missSrv As Boolean

On Error GoTo Errtrap

'strbadbits = vbNullString
Call SetLang

   Grid1.AllowUserResizing = flexResizeBoth
   Grid1.ExtendLastCol = True
    
   Grid1.Rows = 1
   Grid1.Cell(flexcpFontBold, 0, 0, 0, 5) = True
   Grid2.Cell(flexcpFontBold, 0, 0) = True
   Grid1.ExplorerBar = flexExSort
   Grid1.BackColor = vbWhite
   Grid2.BackColor = vbWhite
   
   For i = 0 To lngAct - 1
   j = 0
  
   'If i = 150 Then Err.Raise 9
   frmUtils.SB2.Panels(2).Text = Activities(i)
   frmUtils.Refresh
     
      x = InStrRev(ActPath(i), "\")
      Y = InStrRev(ActPath(i), "\", x - 1)
      tempPath = Mid$(ActPath(i), Y + 1, (x - 1) - Y)
      SPath = Left$(ActPath(i), Len(ActPath(i)) - 10)
      
      If Not FileExists(MSTSPath & "\trains\consists\" & PConName(i, j)) Then
          FlagColRed = True
      End If
      If PConName(i, j) <> vbNullString Then
      flagEngRed = False
         Call CheckForConsistCol(PConName(i, j), Activities(i), PSvcName(i, j), flagEngRed)
      End If
      If Not FileExists(SPath & "Services\" & PSvcName(i, j)) And PSvcName(i, j) <> vbNullString Then
        
          If missSrv = False Then
            strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(524) & vbCrLf
            missSrv = True
        End If
         strbadbits = strbadbits & vbCrLf & Lang(524) & PSvcName(i, j) & Lang(627) & Activities(i) & vbCrLf
      End If
      If Not FileExists(SPath & "paths\" & pPathName(i, j)) And pPathName(i, j) <> vbNullString Then flagPathRed = True
      If Not FileExists(SPath & "Traffic\" & PTfcName(i)) And PTfcName(i) <> vbNullString Then flagTfcRed = True
      Call LooseActivities(ActPath(i) & "\" & Activities(i))
      Grid1.AddItem tempPath & vbTab & Activities(i) & vbTab & PTfcName(i) & vbTab & PSvcName(i, j) & vbTab & PConName(i, j) & vbTab & pPathName(i, j)
      DoEvents
      flagTfcRed = False
      flagPathRed = False
AnyMore:
        missSrv = False
      j = j + 1
      
      If Not FileExists(SPath & "Services\" & PSvcName(i, j)) And PSvcName(i, j) <> vbNullString Then
        flagSvcRed = True
                  If missSrv = False Then
            strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(524) & vbCrLf
            missSrv = True
        End If
        strbadbits = strbadbits & vbCrLf & Lang(524) & PSvcName(i, j) & Lang(627) & Activities(i) & vbCrLf
        missSrv = True
        Grid1.AddItem tempPath & vbTab & Activities(i) & vbTab & PTfcName(i) & vbTab & PSvcName(i, j) & vbTab & PConName(i, j) & vbTab & pPathName(i, j)
      End If
      If PSvcName(i, j) <> vbNullString And missSrv = False Then
      
      If Not FileExists(MSTSPath & "\trains\consists\" & PConName(i, j)) Then
       FlagColRed = True
       'strbadbits = strbadbits & vbCrLf & "Missing Consist " & PConName(i, j)
      End If
      If PConName(i, j) <> vbNullString Then
      flagEngRed = False
       Call CheckForConsistCol(PConName(i, j), Activities(i), PSvcName(i, j), flagEngRed)
       Grid1.AddItem tempPath & vbTab & Activities(i) & vbTab & PTfcName(i) & vbTab & PSvcName(i, j) & vbTab & PConName(i, j) & vbTab & pPathName(i, j)
      End If
      GoTo AnyMore
      End If
TryAnother:
    Next
Label1:
Grid1.col = 0
Grid1.Sort = flexSortStringAscending

If strbadbits <> vbNullString Or strForPrint <> vbNullString Or strReport <> vbNullString Then
frmReport.Rich1.Text = strbadbits & vbCrLf & vbCrLf & strForPrint & vbCrLf & strReport
frmReport.Show 1
frmReport.ZOrder

End If

'strbadbits = vbNullString
'strForPrint = vbNullString

'End If
Exit Sub
Errtrap:

Call MsgBox("An error " & Err & " occurred in subroutine 'frmGrid - Load' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
GoTo TryAnother
End Sub

Private Sub LooseActivities(ActPath As String)
Dim x As Integer, tempService As String
Dim ActName As String, missAct As Boolean, booEntry As Boolean
Dim strNew As String, Engname As String, Engpath As String, Wagname As String
Dim Wagonpath As String, strFoundPath As String

On Error GoTo Errtrap

MousePointer = 11
RoutePath = MSTSPath & "\Routes"
TrainsetPath = MSTSPath & "\Trains\Trainset\"

x = InStrRev(ActPath, "\")
ActName = Mid$(ActPath, x + 1)


NewFile = FreeFile
 Open ActPath For Input As #NewFile
 Do While Not EOF(NewFile)
Line Input #NewFile, A$
 
 Rem ************* Find loose consists - Locos
 tempService = vbNullString
   x = InStr(A$, "EngineData")
   strNew = A$
      If x > 0 Then
   Call CheckEngineData(strNew, Engname, Engpath, booEntry)

  If Right(Engname, 3) <> "eng" Then
Engname = Engname & ".eng"
End If
If booEntry = True Then

If Not FileExists(TrainsetPath & Engpath & "\" & Engname) Then
strFoundPath = vbNullString
Call LookForLoco(Engname, strFoundPath)
If strFoundPath = vbNullString Then
  flagActBad = True
 
  strbadbits = strbadbits & vbCrLf & Lang(525) & ActName & ", " & Lang(526) & Engpath & "\" & Engname & vbCrLf
Else
flagActBad = True
strbadbits = strbadbits & vbCrLf & Lang(525) & ActName & ", " & Lang(526) & Engpath & "\" & Engname & vbCrLf & Lang(631) & strFoundPath & vbCrLf
 End If
 End If
 End If
 End If
 Rem ************* Find loose consists - Wagons
 tempService = vbNullString
   x = InStr(A$, "WagonData")
   strNew = A$
      If x > 0 Then
   Call CheckWagonData(strNew, Wagname, Wagonpath, booEntry)
Wagname = Wagname & ".wag"
 If booEntry = True Then
 If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname) Then
 strFoundPath = vbNullString
Call LookForWag(Wagname, strFoundPath)
If strFoundPath = vbNullString Then
  flagActBad = True
  strbadbits = strbadbits & vbCrLf & Lang(525) & ActName & ", " & Lang(527) & Wagonpath & "\" & Wagname & vbCrLf
Else
 flagActBad = True
           If missAct = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(525) & vbCrLf
missAct = True
End If
 strbadbits = strbadbits & vbCrLf & Lang(525) & ActName & ", " & Lang(527) & Wagonpath & "\" & Wagname & vbCrLf & Lang(632) & strFoundPath & vbCrLf
 End If
 End If
 End If
 End If
 Loop
 Close #NewFile

MousePointer = 0
Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'LooseActivities' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub
Private Sub LooseConsistsGrid(ActPath As String)
Dim x As Integer, tempService As String, strNew As String
Dim Engname As String, Engpath As String, Wagname As String, Wagonpath As String
Dim booEntry As Boolean

On Error GoTo Errtrap

MousePointer = 11
RoutePath = MSTSPath & "\Routes"
TrainsetPath = MSTSPath & "\Trains\Trainset\"

ActPath = RoutePath & "\" & ActPath


NewFile = FreeFile
 Open ActPath For Input As #NewFile
 Do While Not EOF(NewFile)
Line Input #NewFile, A$
 
 Rem ************* Find loose consists - Locos
 tempService = vbNullString
   x = InStr(A$, "EngineData")
   strNew = A$
       If x > 0 Then
   Call CheckEngineData(strNew, Engname, Engpath, booEntry)

  
Engname = Engname & ".eng"
If booEntry = True Then
If Not FileExists(TrainsetPath & Engpath & "\" & Engname) Then
  FlagColRed = True
  If missEng = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(528) & vbCrLf
missEng = True
End If
  strbadbits = strbadbits & vbCrLf & Lang(528) & Engname
   Grid2.AddItem Engname
   
'   Grid2.CellBackColor = &HFF
   Else
   FlagColRed = False
   Grid2.AddItem Engname
   End If
   
 
 End If
 End If
 Rem ************* Find loose consists - Wagons
 tempService = vbNullString
   x = InStr(A$, "WagonData")
   strNew = A$
       If x > 0 Then
   Call CheckWagonData(strNew, Wagname, Wagonpath, booEntry)

  
Wagname = Wagname & ".wag"
 If booEntry = True Then
 If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname) Then
 FlagColRed = True
 If missWag = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(529) & vbCrLf
missWag = True
End If
 strbadbits = strbadbits & vbCrLf & Lang(529) & Wagname
 ' Grid2.CellForeColor = vbRed
   Grid2.AddItem Wagname
'   Grid2.CellBackColor = &HFF
   Else
   FlagColRed = False
   Grid2.AddItem Wagname
   End If
  End If
 End If
 
 Loop
 Close #NewFile

MousePointer = 0
Exit Sub
Errtrap:

Call MsgBox("An Error number " & Err & " occurred in LooseConsistsGrid" _
            & vbCrLf & "Error description: " & Err.Description _
            , vbExclamation, frmGrid)

End Sub

Private Sub Form_Resize()
On Error GoTo Errtrap

Grid1.width = Me.width * 0.75
Grid1.Left = 120
Grid2.width = Me.width * 0.2
Grid2.Left = Grid1.Left + Grid1.width + 50
Grid1.Top = 240
Grid2.Top = Grid1.Top
Grid1.height = Me.height * 0.8
Grid2.height = Grid1.height
Label1.Top = Grid1.Top + Grid1.height + 250
Command1.Top = Label1.Top
Command2.Top = Label1.Top
Command3.Top = Label1.Top
Command4.Top = Label1.Top
Command5.Top = Label1.Top
Command6.Top = Label1.Top
Command7.Top = Label1.Top
Command8.Top = Label1.Top
Command9.Top = Label1.Top
Command1.Left = Grid2.Left
Label1.Left = Grid1.Left
Command2.Left = Label1.Left + Label1.width + 50
Command3.Left = Command2.Left + Command2.width + 50
Command4.Left = Command3.Left + Command3.width + 50
Command5.Left = Command4.Left + Command4.width + 50
Command7.Left = Command5.Left + Command5.width + 50
Command8.Left = Command7.Left + Command7.width + 50
Command9.Left = Command8.Left + Command8.width + 50
Command1.Left = Grid2.Left
Command6.Left = Command1.Left + Command1.width + 100


If booNoButtons = True Then
Command4.Visible = False
Command5.Visible = False
End If
Exit Sub
Errtrap:
If Err = 380 Then
Exit Sub
End If
End Sub

Private Sub Grid1_CellChanged(ByVal row As Long, ByVal col As Long)
If FlagColRed = True And col = 4 Then
Grid1.Select row, col
Grid1.FillStyle = flexFillSingle
Grid1.CellBackColor = &H8080FF
FlagColRed = False
End If
If flagSvcRed = True And col = 3 Then
Grid1.Select row, col
Grid1.FillStyle = flexFillSingle
Grid1.CellBackColor = &H8080FF
flagSvcRed = False
End If
If flagConBad = True And col = 4 Then
Grid1.Select row, col
Grid1.FillStyle = flexFillSingle
Grid1.CellBackColor = &H8080FF
flagConBad = False
End If
If flagEngRed = True And col = 4 Then
Grid1.Select row, col
Grid1.FillStyle = flexFillSingle
Grid1.CellBackColor = &H8080FF
flagEngRed = False
End If
If flagActBad = True And col = 1 Then
Grid1.Select row, col
Grid1.FillStyle = flexFillSingle
Grid1.CellBackColor = &H8080FF
flagActBad = False
End If
If flagPathRed = True And col = 5 Then
Grid1.Select row, col
Grid1.FillStyle = flexFillSingle
Grid1.CellBackColor = &H8080FF
flagPathRed = False
End If
If flagTfcRed = True And col = 2 Then
Grid1.Select row, col
Grid1.FillStyle = flexFillSingle
Grid1.CellBackColor = &H8080FF
flagTfcRed = False
End If
End Sub

Private Sub Grid1_Click()
Dim ActName As String

If Grid1.col = 1 Then
ConDown = False

flagGrid = 3
Grid1.col = 0
ActName = Grid1.Cell(flexcpText)
Grid1.col = 1
ActName = ActName & "\Activities\" & Grid1.Cell(flexcpText)
Grid2.Rows = 0
Grid2.AddItem Lang(634)
Grid2.Rows = 1
Grid2.ExplorerBar = flexExSort
Call LooseConsistsGrid(ActName)
End If


If Grid1.col = 4 Then

ConDown = True

ConsistPath = MSTSPath & "\Trains\Consists"
Grid2.Rows = 0
Grid2.AddItem Lang(633)
Grid2.Rows = 1
Grid2.ExplorerBar = flexExSort

Call CheckForConsistGrid(ConsistPath & "\" & Grid1.Cell(flexcpText))
End If
End Sub


Private Sub CheckForConsistCol(CFilepath As Variant, strAct As Variant, strSvc As Variant, flagEngRed As Boolean)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Engpath As String, Engname As String
Dim strNew2 As String, Wagonpath As String, Wagname As String
Dim TrainsetPath As String, ConName As String, ConP As String
Dim booEntry As Boolean, strFoundPath As String


On Error GoTo Errtrap
ConName = CFilepath
ConP = MSTSPath & "\Trains\Consists\" & CFilepath
Fnumber = FreeFile


TrainsetPath = MSTSPath & "\Trains\Trainset\"
If Not FileExists(ConP) Then
strbadbits = strbadbits & vbCrLf & Lang(530) & ConName & Lang(531) & strAct & Lang(532) & strSvc & vbCrLf
Exit Sub
End If

Open ConP For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
    x = InStr(strNew, "EngineData")
        If x > 0 Then
   Call CheckEngineData(strNew, Engname, Engpath, booEntry)

      
Engname = Engname & ".eng"
   If booEntry = True Then
   If Not FileExists(TrainsetPath & Engpath & "\" & Engname) Then
   strFoundPath = vbNullString
Call LookForLoco(Engname, strFoundPath)
If strFoundPath = vbNullString Then
  flagConBad = True
  flagEngRed = True
  strbadbits = strbadbits & vbCrLf & Lang(533) & " " & ConName & Lang(526) & Engpath & "\" & Engname & vbCrLf
  Else
  flagEngRed = True
strbadbits = strbadbits & vbCrLf & Lang(533) & ConName & ", " & Lang(526) & Engpath & "\" & Engname & vbCrLf & Lang(631) & strFoundPath & vbCrLf
  End If
   End If
   End If
   strNew2 = vbNullString
   End If
    x = InStr(strNew, "WagonData")
   
      If x > 0 Then
   Call CheckWagonData(strNew, Wagname, Wagonpath, booEntry)

Wagname = Wagname & ".wag"
If booEntry = True Then
  If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname) Then
  
     strFoundPath = vbNullString
Call LookForWag(Wagname, strFoundPath)
If strFoundPath = vbNullString Then
 flagConBad = True
 flagEngRed = True
  strbadbits = strbadbits & vbCrLf & Lang(533) & " " & ConName & Lang(527) & Wagonpath & "\" & Wagname & vbCrLf
  Else
  flagEngRed = True
  strbadbits = strbadbits & vbCrLf & Lang(533) & ConName & ", " & Lang(527) & Wagonpath & "\" & Wagname & vbCrLf & Lang(632) & strFoundPath & vbCrLf
  End If
 Close #Fnumber
   Exit Sub
  
   End If
   End If
  End If
   
   strNew = vbNullString
 ' itExists = False
   Loop
   Close #Fnumber
  
Exit Sub
Errtrap:

Call MsgBox("An Error number " & Err & " occurred in CheckForConsistCol" _
            & vbCrLf & "Error description: " & Err.Description _
            , vbExclamation, frmGrid)

End Sub


Private Sub CheckForConsistGrid(CFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Engpath As String, Engname As String
Dim strNew2 As String, Wagonpath As String, Wagname As String
Dim TrainsetPath As String, ConName As String, booEntry As Boolean

On Error GoTo Errtrap

Fnumber = FreeFile
x = InStrRev(CFilepath, "\")
ConName = Mid$(CFilepath, x + 1)

TrainsetPath = MSTSPath & "\Trains\Trainset\"
If Not FileExists(CFilepath) Then

'Call MsgBox(Lang(352) & ConName & Lang(353), vbExclamation, App.Title)
Exit Sub
End If
Open CFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 
   x = InStr(strNew, "EngineData")
   
   If x > 0 Then
   Call CheckEngineData(strNew, Engname, Engpath, booEntry)

'
   If booEntry = True Then
   If Not FileExists(TrainsetPath & Engpath & "\" & Engname & ".eng") Then
  
  FlagColRed = True
   Grid2.AddItem Engname & ".eng"
   If missEng = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(528) & vbCrLf
missEng = True
End If
   strbadbits = strbadbits & vbCrLf & Lang(528) & Engname & ".eng"
   Else
   Grid2.AddItem Engname & ".eng"
   End If
   End If
   strNew2 = vbNullString
   End If
    x = InStr(strNew, "WagonData")
   
      If x > 0 Then
   Call CheckWagonData(strNew, Wagname, Wagonpath, booEntry)



  If booEntry = True Then
  If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname & ".wag") Then
 FlagColRed = True
 If missWag = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(529) & vbCrLf
missWag = True
End If
   Grid2.AddItem Wagname & ".wag"
   strbadbits = strbadbits & vbCrLf & Lang(529) & Wagname & ".wag"
   Else
   Grid2.AddItem Wagname & ".wag"
   End If
   
  End If
  End If
   strNew = vbNullString
 ' itExists = False
   Loop
   Close #Fnumber
  
Exit Sub
Errtrap:



Call MsgBox("An error #" & Err & " occurred in subroutine 'CheckForConsistGrid' while checking" _
            & vbCrLf & CFilepath _
            , vbExclamation, frmGrid)


End Sub



Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 And Grid1.col = 1 Then

If Grid1.MergeCells = 2 Then
Select Case MsgBox(Lang(507) & vbCrLf & Lang(508), vbYesNo + vbExclamation + vbDefaultButton1, App.Title)

    Case vbYes
Grid1.MergeCells = 0

DoEvents
    Case vbNo
Exit Sub
End Select

End If
frmDelete.Show
ElseIf Button = 2 And Grid1.col = 4 Then
flagGrid = 1
frmRepCon.Show

'PopupMenu mnuPop
End If
End Sub

Private Sub Grid2_CellChanged(ByVal row As Long, ByVal col As Long)
If FlagColRed = True Then
Grid2.Select row, col
Grid2.FillStyle = flexFillSingle
Grid2.CellBackColor = vbRed
FlagColRed = False
End If
End Sub




Private Sub Grid2_Click()
booSelect = True
End Sub


Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
booSelect = True
End If
If Button = 2 And booSelect = True Then
booSelect = False

If ConDown = True Then
flagGrid = 4
End If

ThisRow = Grid2.row
frmRepStock.Show

End If
End Sub


