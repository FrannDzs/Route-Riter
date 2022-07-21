VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form frmUnusedSrv 
   Caption         =   "Unused Items"
   ClientHeight    =   9135
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   15330
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   495
      Left            =   12600
      TabIndex        =   10
      ToolTipText     =   "Saves the 4 unused grids"
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Unused Paths"
      Height          =   495
      Index           =   2
      Left            =   7320
      TabIndex        =   9
      Top             =   7800
      Width           =   1695
   End
   Begin VSFlex8LCtl.VSFlexGrid GridPaths 
      Height          =   7215
      Left            =   6480
      TabIndex        =   8
      Top             =   360
      Width           =   2895
      _cx             =   5106
      _cy             =   12726
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
      SelectionMode   =   3
      GridLines       =   1
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
      FormatString    =   $"frmUnused.frx":0000
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   495
      Index           =   4
      Left            =   13920
      TabIndex        =   7
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Unused Consists"
      Height          =   495
      Index           =   3
      Left            =   10320
      TabIndex        =   6
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Unused Traffic"
      Height          =   495
      Index           =   1
      Left            =   4080
      TabIndex        =   5
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Unused Services"
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   7800
      Width           =   1695
   End
   Begin VSFlex8LCtl.VSFlexGrid GridLoco 
      Height          =   7215
      Left            =   12600
      TabIndex        =   3
      Top             =   360
      Width           =   2535
      _cx             =   4471
      _cy             =   12726
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
      SelectionMode   =   3
      GridLines       =   1
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
      FormatString    =   $"frmUnused.frx":0038
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
   Begin VSFlex8LCtl.VSFlexGrid GridCon 
      Height          =   7215
      Left            =   9480
      TabIndex        =   2
      Top             =   360
      Width           =   3015
      _cx             =   5318
      _cy             =   12726
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
      SelectionMode   =   3
      GridLines       =   1
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
      FormatString    =   $"frmUnused.frx":0070
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
   Begin VSFlex8LCtl.VSFlexGrid GridTfc 
      Height          =   7215
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   3015
      _cx             =   5318
      _cy             =   12726
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
      SelectionMode   =   3
      GridLines       =   1
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
      FormatString    =   $"frmUnused.frx":00AB
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
   Begin VSFlex8LCtl.VSFlexGrid GridUnused 
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3015
      _cx             =   5318
      _cy             =   12726
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
      SelectionMode   =   3
      GridLines       =   1
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
      FormatString    =   $"frmUnused.frx":00E5
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
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuDelSvc 
         Caption         =   "Delete Service"
      End
      Begin VB.Menu mnuMovSvc 
         Caption         =   "Move Service"
      End
   End
   Begin VB.Menu mnuPop2 
      Caption         =   "Popup Menu2"
      Visible         =   0   'False
      Begin VB.Menu mnuDelTfc 
         Caption         =   "Delete Traffic"
      End
      Begin VB.Menu mnuMovTfc 
         Caption         =   "Move Traffic"
      End
   End
   Begin VB.Menu mnuPop3 
      Caption         =   "Popup Menu3"
      Visible         =   0   'False
      Begin VB.Menu mnuDelCon 
         Caption         =   "Delete Consist"
      End
      Begin VB.Menu mnuMovCon 
         Caption         =   "Move Consist"
      End
   End
   Begin VB.Menu mnupop4 
      Caption         =   "Popup Menu4"
      Visible         =   0   'False
      Begin VB.Menu mnuDelPath 
         Caption         =   "Delete Path"
      End
      Begin VB.Menu mnuMovPath 
         Caption         =   "Move Path"
      End
   End
End
Attribute VB_Name = "frmUnusedSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim FlagColRed As Boolean
Dim booBadConsist As Boolean
Dim booSelect As Boolean
Dim SelRows() As Integer

Dim booTraffic As Boolean

Private Sub Command1_Click(Index As Integer)
Select Case Index

    
Case 0
flagPrint = 8
fEZPrint.Show
Case 1
flagPrint = 11
fEZPrint.Show
Case 2
flagPrint = 13
fEZPrint.Show
Case 3
flagPrint = 9
fEZPrint.Show
Case 4
Unload Me
End Select
End Sub

Private Sub Command2_Click()
Dim tit1 As String, tit2 As String, x As Integer, tit3 As String



CommonDialog1.Filter = "Comma Separated (*.csv)|*.csv"
CommonDialog1.DialogTitle = "Save Grid as CSV File"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 2
tit1 = CommonDialog1.Filename
tit2 = CommonDialog1.FileTitle
x = InStrRev(tit1, "\")
tit3 = Left$(tit1, x)

If tit1 <> vbNullString Then
GridUnused.SaveGrid tit3 & "SVC_" & tit2, flexFileCommaText
DoEvents
GridTfc.SaveGrid tit3 & "TFC_" & tit2, flexFileCommaText
DoEvents
GridPaths.SaveGrid tit3 & "PAT_" & tit2, flexFileCommaText
DoEvents
GridCon.SaveGrid tit3 & "CON_" & tit2, flexFileCommaText
DoEvents
End If
End Sub


Private Sub Form_Load()


Dim i As Integer, ii As Integer, j As Integer, tempPath As String, x As Integer, Y As Integer
Dim Z As Integer, q As Integer

Me.Caption = Lang(319)
GridUnused.Select 0, 0
GridUnused.Cell(flexcpText) = Lang(320)
GridTfc.Select 0, 0
GridTfc.Cell(flexcpText) = Lang(321)
GridPaths.Select 0, 0
GridPaths.Cell(flexcpText) = Lang(322)
GridCon.Select 0, 0
GridCon.Cell(flexcpText) = Lang(323)
GridLoco.Select 0, 0
GridLoco.Cell(flexcpText) = Lang(324)
For i = 0 To 3
Command1(i).Caption = Lang(379 + i)
Next i
Command1(4).Caption = Lang(38)
GridUnused.Rows = 1
GridUnused.ExplorerBar = flexExSort
GridUnused.BackColor = vbWhite
GridCon.Rows = 1
GridLoco.Rows = 1
GridCon.BackColor = vbWhite
GridLoco.BackColor = vbWhite
GridTfc.BackColor = vbWhite
GridTfc.Rows = 1
GridPaths.Rows = 1
GridPaths.BackColor = vbWhite


For j = 0 To lngSrv - 1
For i = 0 To lngAct - 1
For ii = 0 To 500
booFound = False
Rem ****************

If PSvcName(i, ii) = vbNullString Then Exit For 'GoTo CarryOn
If Service(j) = PSvcName(i, ii) Then
booFound = True

Exit For
End If
CarryON:
Next ii
If booFound = True Then
Exit For
End If

Next i

If booFound = False And Trim$(Service(j)) <> vbNullString Then


Y = Len(SrvPath(j)) - 1
x = InStrRev(SrvPath(j), "\", Y)
Z = InStrRev(SrvPath(j), "\", x - 1)

tempPath = Mid$(SrvPath(j), Z + 1, x - Z)
tempPath = Left$(tempPath, Len(tempPath) - 1)
GridUnused.AddItem tempPath & "\" & Service(j)   ' & vbtab & PathsUsed
Else
booFound = False
End If
Next j

Rem ************** Get Traffic
For j = 0 To lngTfc - 1
For i = 0 To lngAct - 1
booFound = False
If PTfcName(i) = vbNullString Then GoTo GetAnother
If Traffic(j) = PTfcName(i) Then
booFound = True
Exit For
End If
GetAnother:
Next i

If booFound = False And Trim$(Traffic(j)) <> vbNullString Then
x = InStrRev(TfcPath(j), "\traffic")
xx = InStrRev(TfcPath(j), "\", x - 1)
strTemp = Mid$(TfcPath(j), xx + 1, x - xx)
strTemp = strTemp & Traffic(j)
GridTfc.AddItem strTemp
q = q + 1
Else
booFound = False
End If
Next j
Rem ************** Get Paths
For j = 0 To lngPaths - 1
For i = 0 To PathUsedNumb - 1

booFound = False

If Paths(j) = PathUsed(i) Then
booFound = True
Exit For
End If
Next i


If booFound = False And Trim$(Paths(j)) <> vbNullString Then

x = InStrRev(PathsPath(j), "\Paths")
xx = InStrRev(PathsPath(j), "\", x - 1)
strTemp = Mid$(PathsPath(j), xx + 1, x - xx)
strTemp = strTemp & Paths(j)
GridPaths.AddItem strTemp
'q = q + 1
Else
booFound = False
End If
Next j


Rem ****************Get Consists
GridCon.ExplorerBar = flexExSort
GridCon.Rows = 1
For j = 0 To lngCon - 1
For i = 0 To lngAct - 1
For ii = 0 To 500
booFound = False
If PConName(i, ii) = vbNullString Then Exit For 'GoTo CarryOn2
If Consists(j) = PConName(i, ii) Then
booFound = True

Exit For
End If
CarryOn2:
Next ii
If booFound = True Then
Exit For
End If

Next i
ConsistPath = MSTSPath & "\Trains\Consists"

If booFound = False And Trim$(Consists(j)) <> vbNullString Then

FlagColRed = False
Call CheckConsists(ConsistPath & "\" & Consists(j))

GridCon.AddItem Consists(j)
Else
booFound = False
End If
Next j
If booBadConsist = True Then
booBadConsist = False
MousePointer = 0
DoEvents


End If
Rem *****************
GridUnused.col = 0
GridUnused.Sort = flexSortStringAscending
GridCon.col = 0
GridCon.Sort = flexSortStringAscending
GridTfc.col = 0
GridTfc.Sort = flexSortStringAscending
GridPaths.col = 0
GridPaths.Sort = flexSortStringAscending
Screen.MousePointer = 0
End Sub


Private Sub CheckConsists(CFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Engpath As String, Engname As String
Dim strNew2 As String, Wagonpath As String, Wagname As String
Dim TrainsetPath As String, ConName As String, booEntry As Boolean

On Error GoTo ErrTrap

Fnumber = FreeFile
x = InStrRev(CFilepath, "\")
ConName = Mid$(CFilepath, x + 1)

TrainsetPath = MSTSPath & "\Trains\Trainset\"
If Not FileExists(CFilepath) Then

'Call MsgBox(Lang(352) & ConName & Lang(353), vbExclamation, "Missing Consist")
Exit Sub
End If
Open CFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 
   x = InStr(strNew, "EngineData")
   
         If x > 0 Then
   Call CheckEngineData(strNew, Engname, Engpath, booEntry)
 


   
   If booEntry = True Then
   If Not FileExists(TrainsetPath & Engpath & "\" & Engname & ".eng") Then
  FlagColRed = True
  If missEng = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(528) & vbCrLf
missEng = True
End If
  strbadbits = strbadbits & vbCrLf & Lang(528) & Engname & ".eng"
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
 strbadbits = strbadbits & vbCrLf & Lang(529) & Wagname & ".wag"
   End If
   
  End If
   End If
   strNew = vbNullString
 ' itExists = False
   Loop
   Close #Fnumber
 
Exit Sub
ErrTrap:

Resume Next

End Sub


Private Sub CheckForConsistGrid(CFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Y As Integer, Engpath As String, Engname As String
Dim strNew2 As String, Wagonpath As String, Wagname As String
Dim TrainsetPath As String, ConName As String, yy As Integer

On Error GoTo ErrTrap

Fnumber = FreeFile
x = InStrRev(CFilepath, "\")
ConName = Mid$(CFilepath, x + 1)

TrainsetPath = MSTSPath & "\Trains\Trainset\"
If Not FileExists(CFilepath) Then

'Call MsgBox(Lang(352) & ConName & Lang(353), vbExclamation, "Missing Consist")
Exit Sub
End If
Open CFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 
   x = InStr(strNew, "EngineData")
   
   If x > 0 Then

   Y = InStr(x, strNew, "(")
   yy = InStr(Y, strNew, ")")
   strNew2 = Trim$(Mid$(strNew, Y + 1, yy - Y - 1))
   'strNew2 = left$(strNew2, Len(strNew2) - 1)
   strNew2 = Trim$(strNew2)
   x = InStr(strNew2, " ")
   Engname = Left$(strNew2, x - 1)
   Engpath = Mid$(strNew2, x + 1)
   If Left$(Engname, 1) = ChrW$(34) Then
   Engname = Mid$(Engname, 2)
   Y = InStr(Engname, ChrW$(34))
   If Y > 0 Then
   Engname = Left$(Engname, Y - 1)
   End If
   End If
   If Left$(Engpath, 1) = ChrW$(34) Then
   Engpath = Mid$(Engpath, 2)
   Y = InStr(Engpath, ChrW$(34))
    If Y > 0 Then
    Engpath = Left$(Engpath, Y - 1)
    End If
   End If

   
   
   If Not FileExists(TrainsetPath & Engpath & "\" & Engname & ".eng") Then
    FlagColRed = True
  If missEng = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(528) & vbCrLf
missEng = True
End If
  strbadbits = strbadbits & vbCrLf & Lang(528) & Engname & ".eng"
   GridLoco.AddItem Engname & ".eng"
   FlagColRed = False
   Else
   GridLoco.AddItem Engname & ".eng"
   End If
   strNew2 = vbNullString
   End If
    x = InStr(strNew, "WagonData")
   
   If x > 0 Then
 
   Y = InStr(x, strNew, "(")
   yy = InStr(Y, strNew, ")")
   strNew2 = Trim$(Mid$(strNew, Y + 1, yy - Y - 1))
   strNew2 = Trim$(strNew2)
   x = InStr(strNew2, " ")
   Wagname = Left$(strNew2, x - 1)
   Wagname = Trim$(Wagname)
   Wagonpath = Mid$(strNew2, x + 1)
   If Left$(Wagname, 1) = ChrW$(34) Then
   Wagname = Mid$(Wagname, 2)
   Y = InStr(Wagname, ChrW$(34))
    If Y > 0 Then
    Wagname = Left$(Wagname, Y - 1)
    End If
   End If
   If Left$(Wagonpath, 1) = ChrW$(34) Then
   Wagonpath = Mid$(Wagonpath, 2)
   Y = InStr(Wagonpath, ChrW$(34))
    If Y > 0 Then
    Wagonpath = Left$(Wagonpath, Y - 1)
    End If
   End If

   
  If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname & ".wag") Then
   FlagColRed = True
 If missWag = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(529) & vbCrLf
missWag = True
End If
 strbadbits = strbadbits & vbCrLf & Lang(529) & Wagname & ".wag"
   GridLoco.AddItem Wagname & ".wag"
  FlagColRed = False
   Else
   GridLoco.AddItem Wagname & ".wag"
   End If
   
  End If
   
   strNew = vbNullString
 ' itExists = False
   Loop
   Close #Fnumber
  
Exit Sub
ErrTrap:
Call MsgBox("An error #" & Err & " occurred in subroutine 'CheckForConsistGrid' while checking" _
            & vbCrLf & CFilepath _
            , vbExclamation, App.Title)

'Resume Next

End Sub










Private Sub Form_Resize()
Dim i As Integer

GridUnused.Top = 240
GridUnused.Left = 120
GridCon.Top = GridUnused.Top
GridLoco.Top = GridUnused.Top
GridPaths.Top = GridUnused.Top
GridTfc.Top = GridUnused.Top
GridUnused.height = Me.height * 0.8
GridCon.height = GridUnused.height
GridLoco.height = GridUnused.height
GridPaths.height = GridUnused.height
GridTfc.height = GridUnused.height
For i = 0 To 4
Command1(i).Top = GridUnused.Top + GridUnused.height + 100
Next
Command2.Top = Command1(4).Top
GridUnused.width = Me.width * 0.2
GridCon.width = Me.width * 0.2
GridLoco.width = Me.width * 0.17
GridPaths.width = Me.width * 0.2
GridTfc.width = Me.width * 0.2
GridTfc.Left = GridUnused.Left + GridUnused.width + 50
GridPaths.Left = GridTfc.Left + GridTfc.width + 50
GridCon.Left = GridPaths.Left + GridPaths.width + 50
GridLoco.Left = GridCon.Left + GridCon.width + 50

Command1(0).Left = GridUnused.Left + GridUnused.width / 2 - (Command1(0).width / 2)
Command1(1).Left = GridTfc.Left + GridTfc.width / 2 - (Command1(1).width / 2)
Command1(2).Left = GridPaths.Left + GridPaths.width / 2 - (Command1(2).width / 2)
Command1(3).Left = GridCon.Left + GridCon.width / 2 - (Command1(3).width / 2)
Command2.Left = GridLoco.Left
Command1(4).Left = Command2.Left + Command2.width + 100
GridUnused.ColWidth(0) = GridUnused.width
GridCon.ColWidth(0) = GridCon.width
GridLoco.ColWidth(0) = GridLoco.width
GridPaths.ColWidth(0) = GridPaths.width
GridTfc.ColWidth(0) = GridTfc.width
End Sub


Private Sub GridCon_CellChanged(ByVal row As Long, ByVal col As Long)
If FlagColRed = True Then
booBadConsist = True
GridCon.Select row, col
GridCon.FillStyle = flexFillSingle
GridCon.CellBackColor = vbRed
FlagColRed = False
End If
End Sub

Private Sub GridCon_Click()

ConsistPath = MSTSPath & "\Trains\Consists"


GridLoco.Rows = 1
GridLoco.ExplorerBar = flexExSort

Call CheckForConsistGrid(ConsistPath & "\" & GridCon.Cell(flexcpText))
End Sub

Private Sub GridCon_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuPop3
End If
End Sub

Private Sub GridLoco_CellChanged(ByVal row As Long, ByVal col As Long)
If FlagColRed = True Then
GridLoco.Select row, col
GridLoco.FillStyle = flexFillSingle
GridLoco.CellBackColor = vbRed
FlagColRed = False
End If
End Sub

Private Sub GridLoco_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
booSelect = True
End If
If Button = 2 And booSelect = True Then
booSelect = False
flagGrid = 5
frmRepStock.Show
End If
End Sub

Private Sub GridPaths_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
PopupMenu mnupop4
End If
End Sub


Private Sub GridTfc_Click()
booTraffic = True
End Sub

Private Sub GridTfc_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 And booTraffic = True Then
PopupMenu mnupop2
End If
End Sub


Private Sub GridUnused_CellChanged(ByVal row As Long, ByVal col As Long)
If FlagColRed = True And col = 2 Then
GridUnused.Select row, col
GridUnused.FillStyle = flexFillSingle
GridUnused.CellBackColor = vbCyan
FlagColRed = False
End If
End Sub

Private Sub GridUnused_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuPopup
End If
End Sub


Private Sub mnuDelCon_Click()
Dim i As Integer, intRows As Integer
Dim ConPath As String
On Error GoTo ErrTrap

intRows = GridCon.SelectedRows - 1
ReDim SelRows(0 To intRows)

For i = 0 To GridCon.SelectedRows - 1
       SelRows(i) = GridCon.SelectedRow(i)
    Next i

For i = 0 To intRows
GridCon.col = 0
GridCon.row = SelRows(i)
ConPath = MSTSPath & "\trains\consists\" & GridCon.Cell(flexcpText)

If FileExists(ConPath) Then
   GridCon.FillStyle = flexFillSingle
    GridCon.CellBackColor = vbGreen
    Kill ConPath
 booUnusedChanged = True
End If


Next
Exit Sub
ErrTrap:

If Err = 75 Then
SetAttr ConPath, vbNormal
Kill ConPath
Resume Next
End If

End Sub

Private Sub mnuDelPath_Click()
Dim Rname As String, PathPath As String, i As Integer, intRows As Integer

On Error GoTo ErrTrap
intRows = GridPaths.SelectedRows - 1
ReDim SelRows(0 To intRows)

For i = 0 To GridPaths.SelectedRows - 1
       SelRows(i) = GridPaths.SelectedRow(i)
    Next


For i = 0 To intRows
GridPaths.row = SelRows(i)
GridPaths.col = 0

strTemp = GridPaths.Cell(flexcpText)
x = InStr(strTemp, "\")
Rname = Left$(strTemp, x - 1)
PathName = Mid$(strTemp, x + 1)
PathPath = MSTSPath & "\Routes\" & Rname & "\Paths\" & PathName
 If FileExists(PathPath) Then
   GridPaths.FillStyle = flexFillSingle
    GridPaths.CellBackColor = vbGreen
    Kill PathPath
       End If
Next i
booUnusedChanged = True
Exit Sub
ErrTrap:

If Err = 75 Then
SetAttr PathPath, vbNormal
Kill PathPath
Resume Next
End If
End Sub

Private Sub mnuDelSvc_Click()
Dim ThisServicePath As String, ThisService As String, i As Integer, intRows As Integer
Dim strTemp As String

On Error GoTo ErrTrap
intRows = GridUnused.SelectedRows - 1
ReDim SelRows(0 To intRows)

For i = 0 To GridUnused.SelectedRows - 1
       SelRows(i) = GridUnused.SelectedRow(i)
    Next


For i = 0 To intRows
GridUnused.row = SelRows(i)
GridUnused.col = 0

strTemp = GridUnused.Cell(flexcpText)
x = InStr(strTemp, "\")
Rname = Left$(strTemp, x - 1)
ThisService = Mid$(strTemp, x + 1)
ThisServicePath = MSTSPath & "\Routes\" & Rname & "\Services\"

If FileExists(ThisServicePath & "\" & ThisService) Then
   GridUnused.FillStyle = flexFillSingle
    GridUnused.CellBackColor = vbGreen
    'FileCopy ThisServicePath & "\" & ThisService, ThisServicePath & "\SpareServices\" & ThisService
    Kill ThisServicePath & "\" & ThisService
       End If
       Next i
       booUnusedChanged = True
       
Exit Sub

ErrTrap:

If Err = 75 Then
SetAttr ThisServicePath & "\" & ThisService, vbNormal
Kill ThisServicePath & "\" & ThisService
Resume Next
End If


End Sub


Private Sub mnuDelTfc_Click()
Rem not done yet
Dim TfcPath As String, Rname As String, i As Integer, intRows As Integer, x As Integer
Dim TfcName As String, strTemp As String

On Error GoTo ErrTrap
intRows = GridTfc.SelectedRows - 1
ReDim SelRows(0 To intRows)

For i = 0 To GridTfc.SelectedRows - 1
       SelRows(i) = GridTfc.SelectedRow(i)
    Next


For i = 0 To intRows
GridTfc.row = SelRows(i)
GridTfc.col = 0
strTemp = GridTfc.Cell(flexcpText)
x = InStr(strTemp, "\")
Rname = Left$(strTemp, x - 1)
TfcName = Mid$(strTemp, x + 1)
TfcPath = MSTSPath & "\Routes\" & Rname & "\Traffic\" & TfcName

Select Case MsgBox(Lang(552) & vbCrLf & TfcPath, vbYesNo + vbExclamation + vbDefaultButton1, App.Title)

    Case vbYes


GridTfc.col = 0
'PathPath = MSTSPath & "\Routes\" & Rname & "\Paths\" & GridTfc.Cell(flexcpText)
If FileExists(TfcPath) Then
   GridTfc.FillStyle = flexFillSingle
    GridTfc.CellBackColor = vbGreen
    Kill TfcPath
       End If
           Case vbNo
Rem ************* Exit now.
End Select
Next i
booUnusedChanged = True
Exit Sub
ErrTrap:

If Err = 75 Then
SetAttr TfcPath, vbNormal
Kill TfcPath
Resume Next
End If
End Sub

Private Sub mnuMovCon_Click()
Dim i As Integer, intRows As Integer
Dim ConPath As String

On Error GoTo ErrTrap
intRows = GridCon.SelectedRows
ReDim SelRows(0 To intRows)

For i = 0 To GridCon.SelectedRows - 1
       SelRows(i) = GridCon.SelectedRow(i)
    Next i

For i = 0 To intRows - 1
  
GridCon.col = 0
GridCon.row = SelRows(i)
ConPath = MSTSPath & "\trains\consists\" & GridCon.Cell(flexcpText)
If Not DirExists(MSTSPath & "\trains\consists\SpareCon") Then
  MkDir (MSTSPath & "\trains\consists\SpareCon")
  End If
If FileExists(ConPath) Then
   GridCon.FillStyle = flexFillSingle
    GridCon.CellBackColor = vbGreen
    FileCopy ConPath, MSTSPath & "\trains\consists\SpareCon\" & GridCon.Cell(flexcpText)
    Kill ConPath
   
    End If

Next i


booUnusedChanged = True
Exit Sub
ErrTrap:

If Err = 75 Then
SetAttr ConPath, vbNormal
Kill ConPath
Resume Next
End If
End Sub

Private Sub mnuMovPath_Click()
Rem - Still to be implemented
Dim ThisPathPath As String, i As Integer, intRows As Integer

On Error GoTo ErrTrap
intRows = GridPaths.SelectedRows - 1
ReDim SelRows(0 To intRows)

For i = 0 To GridPaths.SelectedRows - 1
       SelRows(i) = GridPaths.SelectedRow(i)
    Next


For i = 0 To intRows
GridPaths.row = SelRows(i)
GridPaths.col = 0

strTemp = GridPaths.Cell(flexcpText)
x = InStr(strTemp, "\")
Rname = Left$(strTemp, x - 1)
PathName = Mid$(strTemp, x + 1)
PathPath = MSTSPath & "\Routes\" & Rname & "\Paths\" & PathName

ThisPathPath = MSTSPath & "\routes\" & Rname & "\Paths"

'GridPaths.Col = 2
'ThisPath = GridPaths.Cell(flexcpText)

If Not DirExists(ThisPathPath & "\SparePat\") Then
  MkDir (ThisPathPath & "\SparePat\")
End If
If FileExists(ThisPathPath & "\" & PathName) Then
   GridPaths.FillStyle = flexFillSingle
    GridPaths.CellBackColor = vbGreen
    FileCopy ThisPathPath & "\" & PathName, ThisPathPath & "\SparePat\" & PathName
    Kill ThisPathPath & "\" & PathName
End If
Next i
booUnusedChanged = True
Exit Sub
ErrTrap:

If Err = 75 Then
SetAttr ThisPathPath & "\" & PathName, vbNormal
Kill ThisPathPath & "\" & PathName
Resume Next
End If
End Sub

Private Sub mnuMovSvc_Click()
Dim ThisServicePath As String, ThisService As String, i As Integer, intRows As Integer
Dim strTemp As String

On Error GoTo ErrTrap
intRows = GridUnused.SelectedRows - 1
ReDim SelRows(0 To intRows)

For i = 0 To GridUnused.SelectedRows - 1
       SelRows(i) = GridUnused.SelectedRow(i)
    Next


For i = 0 To intRows
GridUnused.row = SelRows(i)
GridUnused.col = 0

strTemp = GridUnused.Cell(flexcpText)
x = InStr(strTemp, "\")
Rname = Left$(strTemp, x - 1)
ThisService = Mid$(strTemp, x + 1)
ThisServicePath = MSTSPath & "\Routes\" & Rname & "\Services\"
If Not DirExists(ThisServicePath & "\SpareSvc\") Then
  MkDir (ThisServicePath & "\SpareSvc\")
  End If
If FileExists(ThisServicePath & "\" & ThisService) Then
   GridUnused.FillStyle = flexFillSingle
    GridUnused.CellBackColor = vbGreen
    FileCopy ThisServicePath & "\" & ThisService, ThisServicePath & "\SpareSvc\" & ThisService
    Kill ThisServicePath & "\" & ThisService
       End If
       Next i
       booUnusedChanged = True
       
Exit Sub
ErrTrap:

If Err = 75 Then
SetAttr ThisServicePath & "\" & ThisService, vbNormal
Kill ThisServicePath & "\" & ThisService
Resume Next
End If
End Sub


Private Sub mnuMovTfc_Click()

Dim TfcPath As String, Rname As String, i As Integer, intRows As Integer, x As Integer
Dim TfcName As String, strTemp As String

On Error GoTo ErrTrap
intRows = GridTfc.SelectedRows - 1
ReDim SelRows(0 To intRows)

For i = 0 To GridTfc.SelectedRows - 1
       SelRows(i) = GridTfc.SelectedRow(i)
    Next


For i = 0 To intRows
GridTfc.row = SelRows(i)
GridTfc.col = 0
strTemp = GridTfc.Cell(flexcpText)
x = InStr(strTemp, "\")
Rname = Left$(strTemp, x - 1)
TfcName = Mid$(strTemp, x + 1)
TfcPath = MSTSPath & "\Routes\" & Rname & "\traffic\" & TfcName



If Not DirExists(MSTSPath & "\Routes\" & Rname & "\traffic\SpareTfc\") Then
  MkDir (MSTSPath & "\Routes\" & Rname & "\traffic\SpareTfc\")
End If
If FileExists(TfcPath) Then
   GridTfc.FillStyle = flexFillSingle
    GridTfc.CellBackColor = vbGreen
    FileCopy TfcPath, MSTSPath & "\Routes\" & Rname & "\traffic\SpareTfc\" & TfcName
    Kill TfcPath
End If
Next i
booUnusedChanged = True
Exit Sub

ErrTrap:

If Err = 75 Then
SetAttr TfcPath, vbNormal
Kill TfcPath
Resume Next
End If
End Sub


