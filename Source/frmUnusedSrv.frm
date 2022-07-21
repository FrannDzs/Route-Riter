VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmUnusedSrvOld 
   Caption         =   "Unused Services"
   ClientHeight    =   8145
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14730
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   14730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Change Sheet Format"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   6960
      Width           =   1335
   End
   Begin VSFlex7LCtl.VSFlexGrid GridLoco 
      Height          =   6495
      Left            =   12000
      TabIndex        =   6
      Top             =   240
      Width           =   2535
      _cx             =   4471
      _cy             =   11456
      _ConvInfo       =   1
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
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmUnusedSrv.frx":0000
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
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   495
      Left            =   9240
      TabIndex        =   4
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   12480
      TabIndex        =   1
      Top             =   6960
      Width           =   1095
   End
   Begin VSFlex7LCtl.VSFlexGrid GridUnused 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      _cx             =   13361
      _cy             =   11456
      _ConvInfo       =   1
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmUnusedSrv.frx":0038
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
   End
   Begin VSFlex7LCtl.VSFlexGrid GridCon 
      Height          =   6495
      Left            =   7800
      TabIndex        =   3
      Top             =   240
      Width           =   4095
      _cx             =   7223
      _cy             =   11456
      _ConvInfo       =   1
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
      FormatString    =   $"frmUnusedSrv.frx":00A3
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
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Paths in Cyan are used by another Service"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   6960
      Width           =   3255
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp Menu"
      Begin VB.Menu mnuDelSrv 
         Caption         =   "Delete Service"
      End
      Begin VB.Menu mnuDelBoth 
         Caption         =   "Delete Service && Path"
      End
      Begin VB.Menu mnuMoveSvc 
         Caption         =   "Move Service"
      End
      Begin VB.Menu mnuMoveSel 
         Caption         =   "Move Selected Services"
      End
   End
   Begin VB.Menu mnuPop2 
      Caption         =   "PopUp Menu2"
      Begin VB.Menu mnuDelCon 
         Caption         =   "Delete Consist"
      End
      Begin VB.Menu mnuDelSelCon 
         Caption         =   "Delete Selected Consists"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "&Move Consist"
      End
      Begin VB.Menu mnuSelCon 
         Caption         =   "Move Selected Consists"
      End
   End
End
Attribute VB_Name = "frmUnusedSrvOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim FlagColRed As Boolean
Dim booBadConsist As Boolean
Dim booSelect As Boolean
Dim SelRows() As Integer
Private Sub Command1_Click()
Unload Me

End Sub


Private Sub Command2_Click()
flagPrint = 9
fEZPrint.Show
End Sub

Private Sub Command3_Click()
flagPrint = 8
fEZPrint.Show
End Sub


Private Sub Command4_Click()

    GridUnused.MergeCol(0) = True
    GridUnused.MergeCol(1) = True
    
If GridUnused.MergeCells = 0 Then
        GridUnused.MergeCells = 2
    Else
        GridUnused.MergeCells = 0
    End If
    

End Sub

Private Sub Form_Load()
Dim i As Integer, ii As Integer, j As Integer, tempPath As String, X As Integer, Y As Integer
Dim z As Integer, PathUsed As String, p As Integer, pp As Integer
Label1.BackColor = vbCyan
Label1.Caption = Lang(89)

GridUnused.Rows = 1
GridUnused.ExplorerBar = flexExSort
GridUnused.BackColor = vbWhite
GridCon.BackColor = vbWhite
GridLoco.BackColor = vbWhite

For j = 1 To lngSrv
For i = 1 To lngAct
For ii = 1 To 500
booFound = False
Rem ****************
'If UCase(Trim(Service(j))) = "ACT#1_0E12.SRV" Then Stop
'If UCase(Trim(PSvcName(i, ii))) = "ACT#1_0E12.SRV" Then Stop

Rem**************
If PSvcName(i, ii) = "" Then Exit For 'GoTo CarryOn
If Service(j) = PSvcName(i, ii) Then
booFound = True

Exit For
End If
CarryOn:
Next ii
If booFound = True Then
Exit For
End If

Next i

If booFound = False And Trim(Service(j)) <> "" Then
Call CheckService(SrvPath(j) & Service(j), PathUsed)
FlagColRed = False
For p = 1 To lngAct
For pp = 1 To 500
If PathUsed = pPathName(p, pp) Then
FlagColRed = True
Exit For
End If
Next pp
If FlagColRed = True Then Exit For
Next p

Y = Len(SrvPath(j)) - 1
X = InStrRev(SrvPath(j), "\", Y)
z = InStrRev(SrvPath(j), "\", X - 1)

tempPath = Mid(SrvPath(j), z + 1, X - z)
tempPath = Left(tempPath, Len(tempPath) - 1)
GridUnused.AddItem tempPath & Chr$(9) & Service(j) & Chr$(9) & PathUsed
Else
booFound = False
End If
Next j

Rem ****************Get Consists
GridCon.ExplorerBar = flexExSort
GridCon.Rows = 1
For j = 1 To lngCon
For i = 1 To lngAct
For ii = 1 To 500
booFound = False
If PConName(i, ii) = "" Then Exit For 'GoTo CarryOn2
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

If booFound = False And Trim(Consists(j)) <> "" Then

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

frmWarning.Show 1
End If
Rem *****************
GridUnused.Col = 0
GridUnused.Sort = flexSortStringAscending
GridCon.Col = 0
GridCon.Sort = flexSortStringAscending
Screen.MousePointer = 0
End Sub

Private Sub CheckService(tempPath As String, PathUsed As String)
Dim tempSvcPath As String

svcExists = True
NewFile = FreeFile
Open tempPath For Input As #NewFile
 Do While Not EOF(NewFile)
 Line Input #NewFile, a$
 
 
 X = InStr(a$, "PathID")
 If X > 0 Then
 tempSvcPath = Trim(Mid(a$, X + 7))
 If Left(tempSvcPath, 1) = "(" Then
 tempSvcPath = Mid(tempSvcPath, 2)
 End If
 If Right(tempSvcPath, 1) = ")" Then
 tempSvcPath = Left(tempSvcPath, Len(tempSvcPath) - 1)
 End If
 tempSvcPath = Trim(tempSvcPath)
 If Left(tempSvcPath, 1) = Chr(34) Then
 tempSvcPath = Mid(tempSvcPath, 2)
 Y = InStr(tempSvcPath, Chr(34))
 If Y > 0 Then
 tempSvcPath = Left(tempSvcPath, Y - 1)
 End If
 End If
 
 
 tempSvcPath = Trim(tempSvcPath)
 
PathUsed = tempSvcPath & ".pat"
 End If
 Loop
 Close #NewFile


End Sub



Private Sub Form_Resize()
GridUnused.Top = 240
GridUnused.Left = 120
GridCon.Top = GridUnused.Top
GridLoco.Top = GridUnused.Top
GridUnused.Height = Me.Height * 0.8
GridCon.Height = GridUnused.Height
GridLoco.Height = GridUnused.Height
Command1.Top = GridUnused.Top + GridUnused.Height + 100
Command2.Top = Command1.Top
Command3.Top = Command1.Top
Label1.Top = Command1.Top
GridUnused.Width = Me.Width * 0.5
GridCon.Width = Me.Width * 0.23
GridLoco.Width = Me.Width * 0.23
GridCon.Left = GridUnused.Left + GridUnused.Width + 50
GridLoco.Left = GridCon.Left + GridCon.Width + 50
Command1.Left = GridLoco.Left
Command2.Left = GridCon.Left + GridCon.Width / 2 - (Command2.Width / 2)
Command3.Left = GridUnused.Left + GridUnused.Width / 2 - (Command3.Width / 2)
Command4.Top = Command2.Top
Command4.Left = GridUnused.Left + 300
Label1.Left = Command3.Left + Command3.Width + 200
End Sub



Private Sub GridCon_CellChanged(ByVal Row As Long, ByVal Col As Long)

If FlagColRed = True Then
booBadConsist = True
GridCon.Select Row, Col
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

Private Sub CheckConsists(CFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim X As Integer, Y As Integer, EngPath As String, EngName As String
Dim strNew2 As String, WagonPath As String, WagName As String
Dim TrainSetPath As String, ConName As String

On Error GoTo errtrap

Fnumber = FreeFile
X = InStrRev(CFilepath, "\")
ConName = Mid(CFilepath, X + 1)

TrainSetPath = MSTSPath & "\Trains\Trainset\"
If Not FileExists(CFilepath) Then

Call MsgBox(Lang(90) & ConName & Lang(91), vbExclamation, "Missing Consist")
Exit Sub
End If
Open CFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim(strNew)
 
   X = InStr(strNew, "EngineData")
   
   If X > 0 Then
 
   Y = InStr(X, strNew, "(")
   strNew2 = Trim(Mid(strNew, Y + 1))
   strNew2 = Left(strNew2, Len(strNew2) - 1)
   strNew2 = Trim(strNew2)
   X = InStr(strNew2, " ")
   EngName = Left(strNew2, X - 1)
   EngPath = Mid(strNew2, X + 1)
   If Left(EngName, 1) = Chr(34) Then
   EngName = Mid(EngName, 2)
   Y = InStr(EngName, Chr(34))
   If Y > 0 Then
   EngName = Left(EngName, Y - 1)
   End If
   End If
   If Left(EngPath, 1) = Chr(34) Then
   EngPath = Mid(EngPath, 2)
   Y = InStr(EngPath, Chr(34))
    If Y > 0 Then
    EngPath = Left(EngPath, Y - 1)
    End If
   End If

   
   
   If Not FileExists(TrainSetPath & EngPath & "\" & EngName & ".eng") Then
  FlagColRed = True
  If missEng = False Then
strBadBits = strBadBits & vbCrLf & vbCrLf & "MISSING LOCOMOTIVES" & vbCrLf
missEng = True
End If
  strBadBits = strBadBits & vbCrLf & "Missing engine " & EngName & ".eng"
   End If
   strNew2 = ""
   End If
    X = InStr(strNew, "WagonData")
   
   If X > 0 Then

   Y = InStr(X, strNew, "(")
   strNew2 = Trim(Mid(strNew, Y + 1))
   strNew2 = Left(strNew2, Len(strNew2) - 1)
   strNew2 = Trim(strNew2)
   X = InStr(strNew2, " ")
   WagName = Left(strNew2, X - 1)
   WagName = Trim(WagName)
   WagonPath = Mid(strNew2, X + 1)
   If Left(WagName, 1) = Chr(34) Then
   WagName = Mid(WagName, 2)
   Y = InStr(WagName, Chr(34))
    If Y > 0 Then
    WagName = Left(WagName, Y - 1)
    End If
   End If
   If Left(WagonPath, 1) = Chr(34) Then
   WagonPath = Mid(WagonPath, 2)
   Y = InStr(WagonPath, Chr(34))
    If Y > 0 Then
    WagonPath = Left(WagonPath, Y - 1)
    End If
   End If
'   If Right(WagonPath, 1) = Chr(34) Then
'   WagonPath = Left(WagonPath, Len(WagonPath) - 1)
'   End If

   
  If Not FileExists(TrainSetPath & WagonPath & "\" & WagName & ".wag") Then

 FlagColRed = True
 If missWag = False Then
strBadBits = strBadBits & vbCrLf & vbCrLf & "MISSING WAGONS" & vbCrLf
missWag = True
End If
 strBadBits = strBadBits & vbCrLf & "Missing wagon " & WagName & ".wag"
   End If
   
  End If
   
   strNew = ""
 ' itExists = False
   Loop
   Close #Fnumber
 
Exit Sub
errtrap:

Resume Next

End Sub

Private Sub CheckForConsistGrid(CFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim X As Integer, Y As Integer, EngPath As String, EngName As String
Dim strNew2 As String, WagonPath As String, WagName As String
Dim TrainSetPath As String, ConName As String

On Error GoTo errtrap

Fnumber = FreeFile
X = InStrRev(CFilepath, "\")
ConName = Mid(CFilepath, X + 1)

TrainSetPath = MSTSPath & "\Trains\Trainset\"
If Not FileExists(CFilepath) Then

Call MsgBox(Lang(90) & ConName & Lang(91), vbExclamation, "Missing Consist")
Exit Sub
End If
Open CFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim(strNew)
 
   X = InStr(strNew, "EngineData")
   
   If X > 0 Then
 
   Y = InStr(X, strNew, "(")
   strNew2 = Trim(Mid(strNew, Y + 1))
   strNew2 = Left(strNew2, Len(strNew2) - 1)
   strNew2 = Trim(strNew2)
   X = InStr(strNew2, " ")
   EngName = Left(strNew2, X - 1)
   EngPath = Mid(strNew2, X + 1)
   If Left(EngName, 1) = Chr(34) Then
   EngName = Mid(EngName, 2)
   Y = InStr(EngName, Chr(34))
   If Y > 0 Then
   EngName = Left(EngName, Y - 1)
   End If
   End If
   If Left(EngPath, 1) = Chr(34) Then
   EngPath = Mid(EngPath, 2)
   Y = InStr(EngPath, Chr(34))
    If Y > 0 Then
    EngPath = Left(EngPath, Y - 1)
    End If
   End If

   
   
   If Not FileExists(TrainSetPath & EngPath & "\" & EngName & ".eng") Then
  FlagColRed = True
  If missEng = False Then
strBadBits = strBadBits & vbCrLf & vbCrLf & "MISSING LOCOMOTIVES" & vbCrLf
missEng = True
End If
  strBadBits = strBadBits & vbCrLf & "Missing engine " & EngName & ".eng"
   GridLoco.AddItem EngName & ".eng"
   FlagColRed = False
   Else
   GridLoco.AddItem EngName & ".eng"
   End If
   strNew2 = ""
   End If
    X = InStr(strNew, "WagonData")
   
   If X > 0 Then
 
   Y = InStr(X, strNew, "(")
   strNew2 = Trim(Mid(strNew, Y + 1))
   strNew2 = Left(strNew2, Len(strNew2) - 1)
   strNew2 = Trim(strNew2)
   X = InStr(strNew2, " ")
   WagName = Left(strNew2, X - 1)
   WagName = Trim(WagName)
   WagonPath = Mid(strNew2, X + 1)
   If Left(WagName, 1) = Chr(34) Then
   WagName = Mid(WagName, 2)
   Y = InStr(WagName, Chr(34))
    If Y > 0 Then
    WagName = Left(WagName, Y - 1)
    End If
   End If
   If Left(WagonPath, 1) = Chr(34) Then
   WagonPath = Mid(WagonPath, 2)
   Y = InStr(WagonPath, Chr(34))
    If Y > 0 Then
    WagonPath = Left(WagonPath, Y - 1)
    End If
   End If

   
  If Not FileExists(TrainSetPath & WagonPath & "\" & WagName & ".wag") Then
 FlagColRed = True
 If missWag = False Then
strBadBits = strBadBits & vbCrLf & vbCrLf & "MISSING WAGONS" & vbCrLf
missWag = True
End If
 strBadBits = strBadBits & vbCrLf & "Missing wagon " & WagName & ".wag"
   GridLoco.AddItem WagName & ".wag"
  FlagColRed = False
   Else
   GridLoco.AddItem WagName & ".wag"
   End If
   
  End If
   
   strNew = ""
 ' itExists = False
   Loop
   Close #Fnumber
  
Exit Sub
errtrap:

Resume Next

End Sub





Private Sub GridCon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuPop2
End If
End Sub

Private Sub GridLoco_CellChanged(ByVal Row As Long, ByVal Col As Long)
If FlagColRed = True Then
GridLoco.Select Row, Col
GridLoco.FillStyle = flexFillSingle
GridLoco.CellBackColor = vbRed
FlagColRed = False
End If
End Sub

Private Sub GridLoco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
booSelect = True
End If
If Button = 2 And booSelect = True Then
booSelect = False
flagGrid = 5
frmRepStock.Show
End If
End Sub


Private Sub GridUnused_CellChanged(ByVal Row As Long, ByVal Col As Long)
If FlagColRed = True And Col = 2 Then
GridUnused.Select Row, Col
GridUnused.FillStyle = flexFillSingle
GridUnused.CellBackColor = vbCyan
FlagColRed = False
End If
End Sub

Private Sub GridUnused_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuPopup
End If
End Sub


Private Sub mnuDelBoth_Click()
Dim SvcPath As String, Rname As String, PathPath As String

GridUnused.Col = 0
Rname = GridUnused.Cell(flexcpText)
GridUnused.Col = 1
SvcPath = MSTSPath & "\Routes\" & Rname & "\Services\" & GridUnused.Cell(flexcpText)
GridUnused.Col = 2
PathPath = MSTSPath & "\Routes\" & Rname & "\Paths\" & GridUnused.Cell(flexcpText)
If FileExists(SvcPath) Then
   GridUnused.FillStyle = flexFillSingle
    GridUnused.CellBackColor = vbGreen
    Kill SvcPath
       End If
 If FileExists(PathPath) Then
   GridUnused.FillStyle = flexFillSingle
    GridUnused.CellBackColor = vbGreen
    Kill PathPath
       End If
       
End Sub

Private Sub mnuDelCon_Click()
Dim ConPath As String
Select Case MsgBox(Lang(92) & vbCrLf & Lang(93), vbYesNo + vbExclamation + vbDefaultButton1, "WARNING")

    Case vbYes
    
GridCon.Col = 0

ConPath = MSTSPath & "\trains\consists\" & GridCon.Cell(flexcpText)

If FileExists(ConPath) Then
   GridCon.FillStyle = flexFillSingle
    GridCon.CellBackColor = vbGreen
    Kill ConPath
   
    End If
    Case vbNo
Exit Sub
End Select

End Sub

Private Sub mnuDelSelCon_Click()
Dim i As Integer, intRows As Integer
Dim ConPath As String


intRows = GridCon.SelectedRows - 1
ReDim SelRows(0 To intRows)

For i = 0 To GridCon.SelectedRows - 1
       SelRows(i) = GridCon.SelectedRow(i)
    Next i

For i = 0 To intRows
GridCon.Col = 0
GridCon.Row = SelRows(i)
ConPath = MSTSPath & "\trains\consists\" & GridCon.Cell(flexcpText)

If FileExists(ConPath) Then
   GridCon.FillStyle = flexFillSingle
    GridCon.CellBackColor = vbGreen
    Kill ConPath
   
End If


Next
End Sub

Private Sub mnuDelSrv_Click()
Dim SvcPath As String, Rname As String, PathPath As String

GridUnused.Col = 0
Rname = GridUnused.Cell(flexcpText)
GridUnused.Col = 1
SvcPath = MSTSPath & "\Routes\" & Rname & "\Services\" & GridUnused.Cell(flexcpText)

Select Case MsgBox("Do you really wish to delete this service?" _
                   & vbCrLf & SvcPath _
                   , vbYesNo + vbExclamation + vbDefaultButton1, App.Title)

    Case vbYes


GridUnused.Col = 2
PathPath = MSTSPath & "\Routes\" & Rname & "\Paths\" & GridUnused.Cell(flexcpText)
If FileExists(SvcPath) Then
   GridUnused.FillStyle = flexFillSingle
    GridUnused.CellBackColor = vbGreen
    Kill SvcPath
       End If
           Case vbNo
Rem ************* Exit now.
End Select
End Sub


Private Sub mnuMove_Click()
Dim ConPath As String
Select Case MsgBox(Lang(94) _
                   & vbCrLf & Lang(95) _
                   , vbYesNo + vbExclamation + vbDefaultButton1, "Moving Consist")

    Case vbYes
    
GridCon.Col = 0

ConPath = MSTSPath & "\trains\consists\" & GridCon.Cell(flexcpText)
If Not DirExists(MSTSPath & "\trains\consists\FaultyConsists") Then
  MkDir (MSTSPath & "\trains\consists\FaultyConsists")
  End If
If FileExists(ConPath) Then
   GridCon.FillStyle = flexFillSingle
    GridCon.CellBackColor = vbGreen
    FileCopy ConPath, MSTSPath & "\trains\consists\FaultyConsists\" & GridCon.Cell(flexcpText)
    Kill ConPath
   
    End If
    Case vbNo
Exit Sub
End Select

End Sub


Private Sub mnuMoveSel_Click()
Dim ThisServicePath As String, ThisService As String, i As Integer, intRows As Integer

intRows = GridUnused.SelectedRows - 1
ReDim SelRows(0 To intRows)

For i = 0 To GridUnused.SelectedRows - 1
       SelRows(i) = GridUnused.SelectedRow(i)
    Next


For i = 0 To intRows
GridUnused.Row = SelRows(i)
GridUnused.Col = 0
ThisServicePath = MSTSPath & "\routes\" & GridUnused.Cell(flexcpText) & "\Services"

GridUnused.Col = 1
ThisService = GridUnused.Cell(flexcpText)

If Not DirExists(ThisServicePath & "\SpareServices\") Then
  MkDir (ThisServicePath & "\SpareServices\")
  End If
If FileExists(ThisServicePath & "\" & ThisService) Then
   GridUnused.FillStyle = flexFillSingle
    GridUnused.CellBackColor = vbGreen
    FileCopy ThisServicePath & "\" & ThisService, ThisServicePath & "\SpareServices\" & ThisService
    Kill ThisServicePath & "\" & ThisService
       End If
       Next i
End Sub

Private Sub mnuMoveSvc_Click()
Dim ThisServicePath As String, ThisService As String
GridUnused.Col = 0
ThisServicePath = MSTSPath & "\routes\" & GridUnused.Cell(flexcpText) & "\Services"

GridUnused.Col = 1
ThisService = GridUnused.Cell(flexcpText)

If Not DirExists(ThisServicePath & "\SpareServices\") Then
  MkDir (ThisServicePath & "\SpareServices\")
  End If
If FileExists(ThisServicePath & "\" & ThisService) Then
   GridUnused.FillStyle = flexFillSingle
    GridUnused.CellBackColor = vbGreen
    FileCopy ThisServicePath & "\" & ThisService, ThisServicePath & "\SpareServices\" & ThisService
    Kill ThisServicePath & "\" & ThisService
   
    End If

End Sub


Private Sub mnuSelCon_Click()
Dim i As Integer, intRows As Integer
Dim ConPath As String


intRows = GridCon.SelectedRows - 1
ReDim SelRows(0 To intRows)

For i = 0 To GridCon.SelectedRows - 1
       SelRows(i) = GridCon.SelectedRow(i)
    Next i

For i = 0 To intRows
  
GridCon.Col = 0
GridCon.Row = SelRows(i)
ConPath = MSTSPath & "\trains\consists\" & GridCon.Cell(flexcpText)
If Not DirExists(MSTSPath & "\trains\consists\FaultyConsists") Then
  MkDir (MSTSPath & "\trains\consists\FaultyConsists")
  End If
If FileExists(ConPath) Then
   GridCon.FillStyle = flexFillSingle
    GridCon.CellBackColor = vbGreen
    FileCopy ConPath, MSTSPath & "\trains\consists\FaultyConsists\" & GridCon.Cell(flexcpText)
    Kill ConPath
   
    End If

Next i




End Sub


