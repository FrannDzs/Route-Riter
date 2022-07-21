VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmStock 
   Caption         =   "Consists and Associated Rolling Stock"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16725
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   16725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Caption         =   "Next>>"
      Height          =   375
      Left            =   13200
      TabIndex        =   18
      Top             =   7200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10320
      TabIndex        =   16
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Use couple-chain in .sms file"
      Height          =   615
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "Changes Couple_auto to Couple_chain in .sms files"
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Move Selected Stock"
      Height          =   495
      Index           =   1
      Left            =   11760
      TabIndex        =   14
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Delete Selected Stock"
      Height          =   495
      Index           =   0
      Left            =   10440
      TabIndex        =   13
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Fix Air Brakes"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   8400
      Width           =   1215
   End
   Begin VSFlex8LCtl.VSFlexGrid GridUnused 
      Height          =   4215
      Left            =   480
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   12135
      _cx             =   21405
      _cy             =   7435
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStock.frx":0000
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
   Begin VB.CommandButton Command8 
      Caption         =   "Fix UK Handbrakes"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   8400
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   9120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14160
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Save as CSV"
      Height          =   495
      Left            =   9285
      TabIndex        =   8
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Print Missing"
      Height          =   495
      Left            =   8040
      TabIndex        =   6
      Top             =   8400
      Width           =   1215
   End
   Begin VSFlex8LCtl.VSFlexGrid Grid3 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   6720
      Width           =   9015
      _cx             =   15901
      _cy             =   2566
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
      BackColor       =   -2147483634
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483634
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStock.frx":0092
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
   Begin VB.CommandButton Command4 
      Caption         =   "Show Unused"
      Height          =   495
      Left            =   6750
      TabIndex        =   4
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   495
      Left            =   5475
      TabIndex        =   3
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Sheet Format"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   12960
      TabIndex        =   1
      Top             =   8400
      Width           =   1215
   End
   Begin VSFlex8LCtl.VSFlexGrid GridStock 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   16455
      _cx             =   29025
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStock.frx":0103
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
      ExplorerBar     =   3
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
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
   Begin VB.Label Label2 
      Caption         =   "Find Stock"
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   12240
      TabIndex        =   7
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Menu mnuPop1 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCouple 
         Caption         =   "Change Coupling(s)"
      End
   End
   Begin VB.Menu mnuPop2 
      Caption         =   "Popup Menu2"
      Visible         =   0   'False
      Begin VB.Menu mnuBrake 
         Caption         =   "Change Brake Type(s)"
      End
   End
   Begin VB.Menu mnuPop3 
      Caption         =   "Popup Menu3"
      Visible         =   0   'False
      Begin VB.Menu mnuFCouple 
         Caption         =   "Add Front Coupler"
      End
      Begin VB.Menu mnuChangeFC 
         Caption         =   "Change Front Coupler"
      End
   End
   Begin VB.Menu mnuPop4 
      Caption         =   "Popup Menu4"
      Visible         =   0   'False
      Begin VB.Menu mnuRigid 
         Caption         =   "Make Conn.  Rigid"
      End
   End
   Begin VB.Menu mnuPop5 
      Caption         =   "Popup Menu5"
      Visible         =   0   'False
      Begin VB.Menu mnuFRigid 
         Caption         =   "Make  Conn. Rigid"
      End
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Dim booNoFix As Boolean

Dim booUnused As Boolean
Dim intAirCap As Integer
Dim tempRow() As Integer
Dim foundrow As Integer
Private Function ReadUniFile(CompleteFilePath As String) As String

Dim length As Long, mytristate As Integer
Dim MyString As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean


Set File_obj = CreateObject("Scripting.FileSystemObject")
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, Me.Caption
  Exit Function
End If
mytristate = -1
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll
The_obj.Close
fileflag = False
ReadUniFile = MyString
End Function

Private Function WriteUniFile(CompleteFilePath As String, MyString As String) As String


Dim File_obj As Object, The_obj As Object
On Error GoTo Errtrap

Set File_obj = CreateObject("Scripting.FileSystemObject")

Set The_obj = File_obj.CreateTextFile(CompleteFilePath, True, True)
The_obj.Write (MyString)
The_obj.Close
Exit Function
Errtrap:
Call MsgBox(Err.Description & " occurred in WriteUniFile while processing" _
            & vbCrLf & "file " & CompleteFilePath _
            , vbExclamation, App.Title)


End Function

Private Sub ConvertSMS(strSMS As String, booFixed As Boolean)
Dim MyString As String, x As Long, xx As Long
Dim strStart As String, strEnd As String

xx = 1

MyString = ReadUniFile(strSMS)
Do
x = InStr(xx, MyString, "couple_auto")
If x = 0 Then Exit Do
strStart = Left(MyString, x - 1)
strEnd = Mid(MyString, x + 11)
MyString = strStart & "couple_chain" & strEnd
xx = x + 1
booFixed = True
Loop
If booFixed = True Then
Call WriteUniFile(strSMS, MyString)
End If

End Sub

Private Sub SetLang()
Me.Caption = Lang(305)
GridStock.Select 0, 0
GridStock.Cell(flexcpText) = Lang(306)
GridStock.Select 0, 1
GridStock.Cell(flexcpText) = Lang(307)
GridStock.Select 0, 2
GridStock.Cell(flexcpText) = Lang(308)
GridStock.Select 0, 3
GridStock.Cell(flexcpText) = Lang(309)
GridStock.Select 0, 4
GridStock.Cell(flexcpText) = Lang(310)
'Command7.Caption = Lang(311)
'Command7.ToolTipText = Lang(312)
Command8.Caption = Lang(313)
Command8.ToolTipText = Lang(314)
Command2.Caption = Lang(237)
Command2.ToolTipText = Lang(238)
Command3.Caption = Lang(367)
Command4.Caption = Lang(315)
Command4.ToolTipText = Lang(316)
Command5.Caption = Lang(378)
Command6.Caption = Lang(247)
Command1.Caption = Lang(38)







End Sub

Private Sub Command1_Click()
GridStock.Clear
frmStock.GridStock.col = 0

Grid3.Clear

DoEvents

Unload Me

End Sub


Private Sub Command10_Click()
Dim OldCol As Integer, OldRow As Integer, booChanged As Boolean
Dim strStock As String, strStockPath As String, flagway As Integer
Dim i As Integer, tempTtl As Integer, strSMSFile As String
Dim strStock2 As String

On Error GoTo Errtrap
ReDim tempRow(0 To REF_CHUNK)


For i = 0 To GridStock.SelectedRows - 1
tempRow(i) = GridStock.SelectedRow(i)
     If i > UBound(tempRow) Then
           ReDim Preserve tempRow(0 To i + REF_CHUNK)
           End If
Next i
tempTtl = GridStock.SelectedRows - 1
For i = 0 To tempTtl
OldCol = 11
OldRow = tempRow(i)
GridStock.Select OldRow, OldCol
strStock = GridStock.Cell(flexcpText)
If Left(strStock, 2) = ".." Then
strStock = Replace(strStock, "\\", "\")
strStock = Replace(strStock, "/", "\")
strStock2 = Mid(strStock, 6)
x = InStrRev(strStock2, "\")
strStock = Mid(strStock2, x + 1)
End If
GridStock.Select OldRow, 2
strStockPath = GridStock.Cell(flexcpText)
strStockPath = MSTSPath & "\trains\trainset\" & strStockPath & "\Sound"

SparePath = App.Path & "\" & "TempFiles"
strSMSFile = strStockPath & "\" & strStock
If Not FileExists(strSMSFile) Then
    strSMSFile = MSTSPath & "\Sound\" & strStock
    If Not FileExists(strSMSFile) Then
    strSMSFile = MSTSPath & "\trains\trainset" & strStock2
    If Not FileExists(strSMSFile) Then
    Call MsgBox("SMS file " & strStock & " not found", vbExclamation, App.Title)
    Exit For
    End If
End If
End If

FileCopy strSMSFile, SparePath & "\" & strStock

flagway = 0
Call ConvertSMS(SparePath & "\" & strStock, booChanged)
DoEvents
If booChanged = True Then
booChanged = False
FileCopy SparePath & "\" & strStock, strSMSFile
DoEvents
End If
Kill SparePath & "\" & strStock

Next i
Exit Sub
Errtrap:
If Err = 75 Then
Call MsgBox(strStock & Lang(511) & vbCrLf & Lang(512), vbExclamation, App.Title)
Exit Sub
End If

End Sub

Private Sub Command11_Click()

foundrow = GridStock.row

'GridStock.col = 3
'GridStock.row = foundrow
'GridStock.TopRow = foundrow
strTemp = Text1.Text

For i = foundrow + 1 To GridStock.Rows - 1
GridStock.Select i, 3
strStock = GridStock.Cell(flexcpText)
x = InStr(strStock, Text1)
If x > 0 Then
foundrow = i
GridStock.TopRow = foundrow
GridStock.row = foundrow
GridStock.TopRow = foundrow

Exit For
End If
Next i
'End If
If x = 0 Then
Call MsgBox("No more entries found", vbExclamation, App.Title)
Exit Sub
End If
End Sub

Private Sub Command2_Click()
If booStockOnly = False Then
    GridStock.MergeCol(0) = True
    GridStock.MergeCol(1) = True
    GridStock.MergeCol(2) = True
    Else
    GridStock.MergeCol(2) = True
    End If
    
    GridStock.BackColor = vbWhite
    
    
If GridStock.MergeCells = 0 Then
        GridStock.MergeCells = 2
    Else
        GridStock.MergeCells = 0
    End If
    

End Sub

Private Sub Command3_Click()
If booUnused = False Then
flagPrint = 5
Else
flagPrint = 16
End If
fEZPrint.Show
End Sub

Private Sub Command4_Click()


If Command4.Caption = Lang(550) Then
Command4.Caption = Lang(551)
booUnused = True
GridStock.Visible = False
GridUnused.Visible = True
Command9(0).Visible = True
Command9(1).Visible = True

ElseIf Command4.Caption = Lang(551) Then
Command4.Caption = Lang(550)
booUnused = False
GridStock.Visible = True
GridUnused.Visible = False
Command9(0).Visible = False
Command9(1).Visible = False
End If
End Sub


Private Sub Command5_Click()
flagPrint = 7
fEZPrint.Show
End Sub


Private Sub Command6_Click()
Dim tit1 As String

CommonDialog1.Filter = "Comma Separated (*.csv)|*.csv"
CommonDialog1.DialogTitle = "Save Grid as CSV File"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 2
tit1 = CommonDialog1.Filename
If tit1 <> vbNullString And booUnused = False Then
GridStock.SaveGrid tit1, flexFileCommaText
ElseIf tit1 <> vbNullString And booUnused = True Then
GridUnused.SaveGrid tit1, flexFileCommaText
End If
End Sub

Private Sub Command7_Click()
Dim OldCol As Integer, OldRow As Integer
Dim strStock As String, strStockPath As String, flagway As Integer
Dim tempRow(0 To 100) As Integer, tempTtl As Integer

On Error GoTo Errtrap

intAirCap = Val(InputBox("AuxiliaryResCapacity (e.g. 2 for Passenger, 5 for Freight)"))
If intAirCap < 2 Then
Call MsgBox("The value of the AuxiliaryResCapacity must be a positive number greater than 1.", vbExclamation, App.Title)
Exit Sub
End If
For i = 0 To GridStock.SelectedRows - 1
tempRow(i) = GridStock.SelectedRow(i)
Next i
tempTtl = GridStock.SelectedRows - 1
For i = 0 To tempTtl
OldCol = 3
OldRow = tempRow(i)
GridStock.Select OldRow, OldCol

strStock = GridStock.Cell(flexcpText)
GridStock.Select OldRow, OldCol - 1
strStockPath = GridStock.Cell(flexcpText)
strStockPath = MSTSPath & "\trains\trainset\" & strStockPath

SparePath = App.Path & "\" & "TempFiles"
FileCopy strStockPath & "\" & strStock, SparePath & "\" & strStock

flagway = 0
Call ConvertAirBrakes(SparePath & "\" & strStock, flagway)
If booNoFix = True Then
booNoFix = False
Exit Sub
End If
flagway = 1
Call ConvertAirBrakes(SparePath & "\" & strStock, flagway)
DoEvents
FileCopy SparePath & "\" & strStock, strStockPath & "\" & strStock
DoEvents
Kill SparePath & "\" & strStock

Next i
Exit Sub
Errtrap:
If Err = 75 Then
Call MsgBox(strStock & Lang(511) & vbCrLf & Lang(512), vbExclamation, App.Title)
Exit Sub
End If

End Sub

Private Sub Command8_Click()
Dim OldCol As Integer, OldRow As Integer
Dim strStock As String, strStockPath As String, flagway As Integer
Dim tempRow(0 To 100) As Integer, tempTtl As Integer

On Error GoTo Errtrap
Select Case MsgBox(Lang(513) & vbCrLf & Lang(514) & vbCrLf & " ", vbOKCancel + vbInformation + vbDefaultButton1, App.Title)

    Case vbOK
Rem***********
    Case vbCancel
Exit Sub
End Select


For i = 0 To GridStock.SelectedRows - 1
tempRow(i) = GridStock.SelectedRow(i)
Next i
tempTtl = GridStock.SelectedRows - 1
For i = 0 To tempTtl
OldCol = 3
OldRow = tempRow(i)
GridStock.Select OldRow, OldCol

strStock = GridStock.Cell(flexcpText)
GridStock.Select OldRow, OldCol - 1
strStockPath = GridStock.Cell(flexcpText)
strStockPath = MSTSPath & "\trains\trainset\" & strStockPath

SparePath = App.Path & "\" & "TempFiles"
FileCopy strStockPath & "\" & strStock, SparePath & "\" & strStock

flagway = 0
Call ConvertHandBrakes(SparePath & "\" & strStock, flagway)
If booNoFix = True Then
booNoFix = False
Exit Sub
End If
flagway = 1
Call ConvertHandBrakes(SparePath & "\" & strStock, flagway)
DoEvents
FileCopy SparePath & "\" & strStock, strStockPath & "\" & strStock
DoEvents
Kill SparePath & "\" & strStock
GridStock.Select OldRow, OldCol + 3
GridStock.Cell(flexcpText) = "Vacuum_piped"
Next i
Exit Sub
Errtrap:
If Err = 75 Then
Call MsgBox(strStock & Lang(511) & vbCrLf & Lang(512), vbExclamation, App.Title)
Exit Sub
End If

End Sub

Private Sub Command9_Click(Index As Integer)
Dim i As Integer, intRows As Integer
Dim TrainPath As String, strStock As String, strRRBackupPath As String
On Error GoTo Errtrap

intRows = GridUnused.SelectedRows - 1
ReDim SelRows(0 To intRows)

For i = 0 To GridUnused.SelectedRows - 1
       SelRows(i) = GridUnused.SelectedRow(i)
    Next i

For i = 0 To intRows
GridUnused.col = 0
GridUnused.row = SelRows(i)
TrainPath = MSTSPath & "\trains\trainset\" & GridUnused.Cell(flexcpText)
GridUnused.col = 1
strStock = GridUnused.Cell(flexcpText)
TrainPath = TrainPath & "\" & strStock

Select Case Index
Case 0
If FileExists(TrainPath) Then
   GridUnused.FillStyle = flexFillSingle
    GridUnused.CellBackColor = vbGreen
    Kill TrainPath
 booUnusedChanged = True
End If
Case 1
If Not DirExists(MSTSPath & "\trains\RRBackups") Then
MkDir MSTSPath & "\trains\RRBackups"
End If
strRRBackupPath = MSTSPath & "\trains\RRBackups\"
If FileExists(TrainPath) Then
   GridUnused.FillStyle = flexFillSingle
    GridUnused.CellBackColor = vbGreen
    FileCopy TrainPath, strRRBackupPath & strStock
    DoEvents
    Kill TrainPath
 booUnusedChanged = True
End If
End Select

Next
Exit Sub
Errtrap:

If Err = 75 Then
SetAttr TrainPath, vbNormal
Kill TrainPath
Resume Next
End If
End Sub

Private Sub Form_Load()
MousePointer = 11
Call SetLang
If booStockOnly = True Then
GridStock.ColHidden(0) = True
GridStock.ColHidden(1) = True
Command4.Visible = False
Me.Caption = Lang(310)
Else

'Command8.Visible = False
End If

GridStock.BackColor = vbWhite
Grid3.BackColor = vbWhite


 GridStock.Rows = 1
   GridStock.Cell(flexcpFontBold, 0, 0, 0, 2) = True
     GridStock.ExplorerBar = flexExSort
     Grid3.Cell(flexcpFontBold, 0, 0, 0, 1) = True
   Grid3.Rows = 1
   
   Grid3.ExplorerBar = flexExSort
   GridUnused.Rows = 1
   GridUnused.ExplorerBar = flexExSort
   
MousePointer = 0

End Sub


Private Sub Form_Resize()
GridStock.Top = 240
GridStock.Left = 120
GridStock.width = Me.width * 0.9
GridStock.height = Me.height * 0.6
'Grid2.Top = GridStock.Top
'Grid2.Height = GridStock.Height * 0.65
'Grid2.Width = Me.Width * 0.35
'Grid2.Left = GridStock.Left + GridStock.Width + 120
Grid3.Top = GridStock.Top + GridStock.height + 250

Grid3.width = GridStock.width * 0.7


Grid3.Left = GridStock.Left
Label2.Left = Grid3.Left + Grid3.width + 200
Label2.Top = Grid3.Top + 250
Text1.Left = Label2.Left + Label2.width + 50
Text1.Top = Label2.Top
Command11.Top = Label2.Top
Command11.Left = Text1.Left + Text1.width + 50
Command1.Top = Grid3.Top + Grid3.height + 200
Command2.Top = Command1.Top
'Command1.Left = Grid2.Left
Command2.Left = GridStock.Left
Command3.Top = Command2.Top
Command3.Left = Command2.Left + Command2.width + 50
Command4.Top = Command2.Top
Command4.Left = Command3.Left + Command3.width + 50
Command5.Top = Command2.Top
Command5.Left = Command4.Left + Command4.width + 50
Command6.Top = Command2.Top
Command6.Left = Command5.Left + Command5.width + 50
Command7.Left = Command6.Left + Command6.width + 50
Command7.Top = Command2.Top
Command8.Left = Command7.Left + Command7.width + 50

Command8.Top = Command2.Top
Command9(0).Top = Command2.Top
Command9(1).Top = Command2.Top
Command9(0).Left = Command8.Left + Command8.width + 50
Command9(1).Left = Command9(0).Left + Command9(0).width + 50
Command10.Left = Command9(1).Left + Command9(1).width + 50

Command10.Top = Command2.Top
Command1.Left = Command10.Left + Command10.width + 50
End Sub




Private Sub Form_Unload(Cancel As Integer)
GridStock.Clear
GridStock.col = 0

Grid3.Clear

DoEvents
End Sub

Private Sub Grid3_CellChanged(ByVal row As Long, ByVal col As Long)
If FlagColRed = True And col = 1 Then
Grid3.Select row, col
Grid3.FillStyle = flexFillSingle
Grid3.CellBackColor = vbCyan
'FlagColRed = False
End If
End Sub

Private Sub Grid3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
flagGrid = 6
frmRepStock.Show
End Sub


Private Sub GridStock_CellChanged(ByVal row As Long, ByVal col As Long)

If FlagColRed = True And col = 3 Then
GridStock.Select row, col
GridStock.FillStyle = flexFillSingle
GridStock.CellBackColor = vbRed
FlagColRed = False
End If
If FlagColGreen = True And (col = 2 Or col = 3) Then
GridStock.Select row, col
GridStock.FillStyle = flexFillSingle
GridStock.CellBackColor = vbGreen
FlagColRed = False
End If
End Sub

Private Function ConvertAirBrakes(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertAirBrakes = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Long
Dim xx As Long, strSearch As String, strBrake As String

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then

strSearch = "MaxAuxilaryChargingRate("
x = InStr(MyString, strSearch)
If x = 0 Then GoTo EndIt

xx = InStr(x, MyString, ")")
strBrake = "AuxilaryResCapacity( " & Trim$(Str(intAirCap)) & " )" & vbCrLf & "AuxilaryResMaxPressure (70)"
strStart = Left$(MyString, xx)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & vbCrLf & strBrake & vbCrLf & strEnd


EndIt:
End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertAirBrakes = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function


Private Function ConvertHandBrakes(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertHandBrakes = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer, strSearch As String, strNew As String

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then

strSearch = "BrakeEquipmentType( " & ChrW$(34) & "Handbrake" & ChrW$(34)
x = InStr(MyString, strSearch)
If x = 0 Then
strSearch = "BrakeEquipmentType(  " & ChrW$(34) & "Handbrake" & ChrW$(34)
x = InStr(MyString, strSearch)

End If

If x = 0 Then
Call MsgBox(Lang(515) & CompleteFilePath & vbCrLf & Lang(516), vbExclamation + vbDefaultButton1, App.Title)
booNoFix = True
Exit Function
End If
strNew = "BrakeEquipmentType( " & ChrW$(34) & "Handbrake, vacuum_brake" & ChrW$(34) & " )" & vbCrLf
strNew = strNew & vbTab & "BrakeSystemType( " & ChrW$(34) & "Vacuum_piped" & ChrW$(34) & " )" & vbCrLf
strNew = strNew & vbTab & "MaxBrakeForce( 0N )" & vbCrLf
strNew = strNew & vbTab & "MaxHandbrakeForce( 10000N )" & vbCrLf
strNew = strNew & vbTab & "NumberOfHandbrakeLeverSteps( 100 )" & vbCrLf

If x <> 0 Then
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strEnd
End If
strSearch = "MaxHandbrakeForce"
x = InStr(MyString, strSearch)
If x <> 0 Then
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strEnd
End If
strSearch = "MaxBrakeForce"
x = InStr(MyString, strSearch)
If x <> 0 Then
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strEnd
End If
strSearch = "NumberOfHandbrake"
x = InStr(MyString, strSearch)
If x <> 0 Then
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strNew & strEnd
End If


End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertHandBrakes = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function





Private Function ConvertBrakes(CompleteFilePath As String, flagway As Integer, strOld As String) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertBrakes = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer, strSearch As String, strNew As String

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then

strNew = "BrakeSystemType ( " & ChrW$(34) & strOld & ChrW$(34) & " )"
strSearch = "BrakeSystemType"

x = InStr(MyString, strSearch)
If x = 0 Then GoTo CarryON
xx = InStr(x, MyString, ")")

strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strNew & strEnd
CarryON:
End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertBrakes = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function














Private Function ConvertFRigid(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertFRigid = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer, Y As Integer


'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If

If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
MyString = Replace(MyString, "          ", " ")
MyString = Replace(MyString, "         ", " ")
MyString = Replace(MyString, "        ", " ")
MyString = Replace(MyString, "       ", " ")
MyString = Replace(MyString, "      ", " ")
MyString = Replace(MyString, "     ", " ")
MyString = Replace(MyString, "    ", " ")
MyString = Replace(MyString, "   ", " ")
MyString = Replace(MyString, "  ", " ")
MyString = Replace(MyString, " ", " ")
MyString = Replace(MyString, "(0.1m/s)", "( 0.1m/s )")
MyString = Replace(MyString, "(-0.1m/s)", "( -0.1m/s )")
MyString = Replace(MyString, "(0.12m/s)", "( 0.12m/s )")
MyString = Replace(MyString, "(-0.12m/s)", "( -0.12m/s )")
x = InStr(MyString, "Comment ( CouplingHasRigid")
If x = 0 Then
x = InStr(MyString, "# ( CouplingHasRigid")
End If
If x = 0 Then
x = InStr(MyString, "#(CouplingHasRigid")
End If
If x > 0 Then
Y = InStr(x, MyString, vbCrLf)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, Y)
MyString = strStart & strEnd
End If

x = InStr(MyString, "Coupling (")

Y = InStr(x, MyString, "r0 (")
If Y > 0 Then
yy = InStr(Y, MyString, ")")
End If
xx = InStr(Y, MyString, "Velocity ( 0.1")

x = InStr(xx, MyString, "Coupling (")

Y = InStr(x, MyString, "r0 (")
If Y > 0 Then
yy = InStr(Y, MyString, ")")
End If
xx = InStr(Y, MyString, "Velocity ( -0.1")
strStart = Left$(MyString, yy + 1)
strEnd = Mid$(MyString, xx - 1)
MyString = strStart & vbCrLf & ")" & vbCrLf & "CouplingHasRigidConnection ( 1 )" & vbCrLf & strEnd

EndIt:
End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertFRigid = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function
Private Function ConvertRigid(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertRigid = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer, Y As Integer
'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
MyString = Replace(MyString, "          ", " ")
MyString = Replace(MyString, "         ", " ")
MyString = Replace(MyString, "        ", " ")
MyString = Replace(MyString, "       ", " ")
MyString = Replace(MyString, "      ", " ")
MyString = Replace(MyString, "     ", " ")
MyString = Replace(MyString, "    ", " ")
MyString = Replace(MyString, "   ", " ")
MyString = Replace(MyString, "  ", " ")
MyString = Replace(MyString, " ", " ")
MyString = Replace(MyString, "(0.1m/s)", "( 0.1m/s )")
MyString = Replace(MyString, "(-0.1m/s)", "( -0.1m/s )")
MyString = Replace(MyString, "(0.12m/s)", "( 0.12m/s )")
MyString = Replace(MyString, "(-0.12m/s)", "( -0.12m/s )")

x = InStr(MyString, "Comment ( CouplingHasRigid")
If x = 0 Then
x = InStr(MyString, "# ( CouplingHasRigid")
End If
If x = 0 Then
x = InStr(MyString, "#(CouplingHasRigid")
End If
If x > 0 Then
Y = InStr(x, MyString, vbLf)

strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, Y)
MyString = strStart & strEnd
End If

x = InStr(MyString, "CouplingHasRigidConnection ()")
If x > 0 Then
MyString = Replace(MyString, "CouplingHasRigidConnection ()", "CouplingHasRigidConnection ( 1 )")
GoTo EndIt
End If
x = InStr(MyString, "CouplingHasRigidConnection ( )")
If x > 0 Then
MyString = Replace(MyString, "CouplingHasRigidConnection ( )", "CouplingHasRigidConnection ( 1 )")
GoTo EndIt
End If
x = InStr(MyString, "CouplingHasRigidConnection()")
If x > 0 Then
MyString = Replace(MyString, "CouplingHasRigidConnection()", "CouplingHasRigidConnection ( 1 )")
GoTo EndIt
End If
x = InStr(MyString, "CouplingHasRigidConnection ( 0 )")
If x > 0 Then
MyString = Replace(MyString, "CouplingHasRigidConnection ( 0 )", "CouplingHasRigidConnection ( 1 )")
GoTo EndIt
End If
x = InStr(MyString, "Coupling (")
Y = InStr(x, MyString, "r0 (")
If Y > 0 Then
yy = InStr(Y, MyString, ")")
End If
xx = InStr(Y, MyString, "Velocity ( 0.1")
strStart = Left$(MyString, yy + 1)
strEnd = Mid$(MyString, xx - 1)
MyString = strStart & vbCrLf & ")" & vbCrLf & "CouplingHasRigidConnection ( 1 )" & vbCrLf & strEnd


EndIt:
End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertRigid = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function

Private Function ConvertFCoupling(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertFCoupling = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer, Y As Integer
Dim strVelo As String, strFCouple As String

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
MyString = Replace(MyString, "          ", " ")
MyString = Replace(MyString, "         ", " ")
MyString = Replace(MyString, "        ", " ")
MyString = Replace(MyString, "       ", " ")
MyString = Replace(MyString, "      ", " ")
MyString = Replace(MyString, "     ", " ")
MyString = Replace(MyString, "    ", " ")
MyString = Replace(MyString, "   ", " ")
MyString = Replace(MyString, "  ", " ")
MyString = Replace(MyString, " ", " ")


x = InStr(MyString, "Coupling (")
xx = InStr(x + 5, MyString, "Coupling (")
If xx > x Then
strReport = strReport & CompleteFilePath & " Already had a Front coupler, no changes made" & vbCrLf & vbCrLf
GoTo EndIt
End If
xx = InStr(x + 5, MyString, "Velocity (")
If xx > 0 Then
If Mid$(MyString, xx - 3, 3) = "Max" Then
xx = 0
End If
End If
Rem *************** Has Velocity Entry ****************
If xx > 0 Then
Y = InStr(xx, MyString, ")")
strVelo = Mid$(MyString, xx + 10, Y - (xx + 10))
strVelo = Trim$(strVelo)
yy = InStr(Y + 2, MyString, ")")
strFCouple = Mid$(MyString, x, (yy + 1) - x)
strFCouple = Replace(strFCouple, strVelo, "-" & strVelo)
strStart = Left$(MyString, yy)
strEnd = Mid$(MyString, yy + 1)
MyString = strStart & vbCrLf & strFCouple & vbCrLf & strEnd
GoTo EndIt
ElseIf xx = 0 Then              ' ***************** No Velocity Entry *********
xx = InStr(x + 5, MyString, "r0 (")
Y = InStr(xx, MyString, ")")
Y = InStr(Y + 1, MyString, ")")
yy = InStr(Y + 1, MyString, ")")
strVelo = "Velocity ( 0.1m/s )"
strStart = Left$(MyString, Y)
strEnd = Mid$(MyString, yy)
MyString = strStart & vbCrLf & strVelo & vbCrLf & strEnd
Rem ************ Now copy in Front Coupler
x = InStr(MyString, "Coupling (")
xx = InStr(x + 5, MyString, "Velocity (")
Y = InStr(xx, MyString, ")")
strVelo = Mid$(MyString, xx + 10, Y - (xx + 10))
strVelo = Trim$(strVelo)
yy = InStr(Y + 1, MyString, ")")
strFCouple = Mid$(MyString, x, (yy + 1) - x)
strFCouple = Replace(strFCouple, strVelo, "-" & strVelo)
strStart = Left$(MyString, yy)
strEnd = Mid$(MyString, yy + 1)
MyString = strStart & vbCrLf & strFCouple & vbCrLf & strEnd

End If



EndIt:
End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertFCoupling = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function






Private Function ConvertFrontCoupling(CompleteFilePath As String, flagway As Integer, strOld As String) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertFrontCoupling = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer, Y As Integer, strSearch As String, strNew As String

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then

xx = InStr(MyString, "Coupling (")
Y = InStr(xx + 2, MyString, "Coupling (")
If Y = 0 Then GoTo EndIt
If strOld = "Automatic" Then
strSearch = "Type ( Automatic"
strNew = "Type ( Chain"
Else
strSearch = "Type ( Chain"
strNew = "Type ( Automatic"
End If

x = InStr(Y, MyString, strSearch)
If x = 0 Then
GoTo EndIt
End If

strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, x + Len(strSearch))
MyString = strStart & strNew & strEnd
EndIt:
End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertFrontCoupling = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function

Private Function ConvertCoupling(CompleteFilePath As String, flagway As Integer, strOld As String) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertCoupling = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim strSearch As String, strNew As String

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
If strOld = "Automatic" Then
strSearch = "Type ( Automatic"
strNew = "Type ( Chain"
Else
strSearch = "Type ( Chain"
strNew = "Type ( Automatic"
End If

x = InStr(MyString, strSearch)
If x = 0 Then
GoTo EndIt
End If

strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, x + Len(strSearch))
MyString = strStart & strNew & strEnd
EndIt:
End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertCoupling = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function

         
         

Private Sub GridStock_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)


If Button = 2 And GridStock.col = 3 Then
flagGrid = 2
MousePointer = 11
frmRepStock.Show
ElseIf Button = 2 And GridStock.col = 5 Then

PopupMenu mnuPop1
ElseIf Button = 2 And GridStock.col = 6 Then

PopupMenu mnuPop2
ElseIf Button = 2 And GridStock.col = 8 Then
PopupMenu mnuPop3
ElseIf Button = 2 And GridStock.col = 9 Then
PopupMenu mnupop4
ElseIf Button = 2 And GridStock.col = 10 Then
PopupMenu mnuPop5
End If
End Sub


Private Sub mnuBrake_Click()
Dim strBrake As String, strNewBrake As String, OldCol As Integer, OldRow As Integer
Dim strStock As String, strStockPath As String, flagway As Integer
Dim tempRow(0 To 100) As Integer, tempTtl As Integer

On Error GoTo Errtrap
frmBrake.Show 1
strNewBrake = Label1.Caption
If strNewBrake = vbNullString Then Exit Sub
For i = 0 To GridStock.SelectedRows - 1
tempRow(i) = GridStock.SelectedRow(i)
Next i
tempTtl = GridStock.SelectedRows - 1
For i = 0 To tempTtl
OldCol = 6
OldRow = tempRow(i)
GridStock.Select OldRow, OldCol
strBrake = GridStock.Cell(flexcpText)
GridStock.Select OldRow, OldCol - 3
strStock = GridStock.Cell(flexcpText)
GridStock.Select OldRow, OldCol - 4
strStockPath = GridStock.Cell(flexcpText)
strStockPath = MSTSPath & "\trains\trainset\" & strStockPath

SparePath = App.Path & "\" & "TempFiles"
FileCopy strStockPath & "\" & strStock, SparePath & "\" & strStock

flagway = 0
Call ConvertBrakes(SparePath & "\" & strStock, flagway, strNewBrake)

flagway = 1
Call ConvertBrakes(SparePath & "\" & strStock, flagway, strNewBrake)
DoEvents
FileCopy SparePath & "\" & strStock, strStockPath & "\" & strStock
DoEvents
Kill SparePath & "\" & strStock
GridStock.Select OldRow, OldCol
GridStock.Cell(flexcpText) = strNewBrake
Next i
Exit Sub
Errtrap:
If Err = 75 Then
Call MsgBox(strStock & Lang(511) & vbCrLf & Lang(512), vbExclamation, App.Title)
Exit Sub
End If

End Sub

Private Sub mnuChangeFC_Click()
Dim strCouple As String, strNewCouple As String, OldCol As Integer, OldRow As Integer
Dim strStock As String, strStockPath As String, flagway As Integer
Dim tempRow(0 To 100) As Integer, tempTtl As Integer
On Error GoTo Errtrap
For i = 0 To GridStock.SelectedRows - 1
tempRow(i) = GridStock.SelectedRow(i)
Next i
tempTtl = GridStock.SelectedRows - 1
For i = 0 To tempTtl
OldCol = 8
OldRow = tempRow(i)
GridStock.Select OldRow, OldCol
strCouple = GridStock.Cell(flexcpText)
If strCouple = "Bar" Then
Call MsgBox("This option does not allow you to change Bar couplings.", vbExclamation, App.Title)

Exit Sub
End If
If strCouple = "Chain" Then
strNewCouple = "Automatic"
ElseIf strCouple = "Automatic" Then
strNewCouple = "Chain"
End If
GridStock.Select OldRow, OldCol - 5
strStock = GridStock.Cell(flexcpText)
GridStock.Select OldRow, OldCol - 6
strStockPath = GridStock.Cell(flexcpText)
strStockPath = MSTSPath & "\trains\trainset\" & strStockPath

SparePath = App.Path & "\" & "TempFiles"
FileCopy strStockPath & "\" & strStock, SparePath & "\" & strStock

flagway = 0
Call ConvertFrontCoupling(SparePath & "\" & strStock, flagway, strCouple)
flagway = 1
Call ConvertFrontCoupling(SparePath & "\" & strStock, flagway, strCouple)
DoEvents

Close
FileCopy SparePath & "\" & strStock, strStockPath & "\" & strStock
DoEvents
Kill SparePath & "\" & strStock
GridStock.Select OldRow, OldCol
GridStock.Cell(flexcpText) = strNewCouple
Next i
Exit Sub
Errtrap:
If Err = 75 Then
Call MsgBox("File " & strStock & " is Read Only, you must" _
            & vbCrLf & "make this file Read/Write before you can edit it." _
            , vbExclamation, App.Title)
Exit Sub
End If

End Sub

Private Sub mnuCouple_Click()
Dim strCouple As String, strNewCouple As String, OldCol As Integer, OldRow As Integer
Dim strStock As String, strStockPath As String, flagway As Integer
Dim tempRow(0 To 100) As Integer, tempTtl As Integer
On Error GoTo Errtrap
For i = 0 To GridStock.SelectedRows - 1
tempRow(i) = GridStock.SelectedRow(i)
Next i
tempTtl = GridStock.SelectedRows - 1
For i = 0 To tempTtl
OldCol = 5
OldRow = tempRow(i)
GridStock.Select OldRow, OldCol
strCouple = GridStock.Cell(flexcpText)
If strCouple = "Bar" Then
Call MsgBox("This option does not allow you to change Bar couplings.", vbExclamation, App.Title)

Exit Sub
End If
If strCouple = "Chain" Then
strNewCouple = "Automatic"
ElseIf strCouple = "Automatic" Then
strNewCouple = "Chain"
End If
GridStock.Select OldRow, OldCol - 2
strStock = GridStock.Cell(flexcpText)
GridStock.Select OldRow, OldCol - 3
strStockPath = GridStock.Cell(flexcpText)
strStockPath = MSTSPath & "\trains\trainset\" & strStockPath

SparePath = App.Path & "\" & "TempFiles"
FileCopy strStockPath & "\" & strStock, SparePath & "\" & strStock

flagway = 0
Call ConvertCoupling(SparePath & "\" & strStock, flagway, strCouple)
flagway = 1
Call ConvertCoupling(SparePath & "\" & strStock, flagway, strCouple)
DoEvents

Close
FileCopy SparePath & "\" & strStock, strStockPath & "\" & strStock
DoEvents
Kill SparePath & "\" & strStock
GridStock.Select OldRow, OldCol
GridStock.Cell(flexcpText) = strNewCouple
Next i
Exit Sub
Errtrap:
If Err = 75 Then
Call MsgBox("File " & strStock & " is Read Only, you must" _
            & vbCrLf & "make this file Read/Write before you can edit it." _
            , vbExclamation, App.Title)
Exit Sub
End If

End Sub


Private Sub mnuFCouple_Click()
Dim OldCol As Integer, OldRow As Integer
Dim strStock As String, strStockPath As String, flagway As Integer
Dim tempRow(0 To 100) As Integer, tempTtl As Integer

On Error GoTo Errtrap
For i = 0 To GridStock.SelectedRows - 1
tempRow(i) = GridStock.SelectedRow(i)
Next i

tempTtl = GridStock.SelectedRows - 1
For i = 0 To tempTtl
OldCol = 8
OldRow = tempRow(i)

GridStock.Select OldRow, OldCol - 5
strStock = GridStock.Cell(flexcpText)
If strStock = "Default.wag" Then GoTo CarryON
GridStock.Select OldRow, OldCol - 6
strStockPath = GridStock.Cell(flexcpText)
strStockPath = MSTSPath & "\trains\trainset\" & strStockPath

SparePath = App.Path & "\" & "TempFiles"
FileCopy strStockPath & "\" & strStock, SparePath & "\" & strStock

flagway = 0
Call ConvertFCoupling(SparePath & "\" & strStock, flagway)

flagway = 1
Call ConvertFCoupling(SparePath & "\" & strStock, flagway)
DoEvents
FileCopy SparePath & "\" & strStock, strStockPath & "\" & strStock
DoEvents
Kill SparePath & "\" & strStock
'GridStock.Select OldRow, OldCol
'GridStock.Cell(flexcpText) = strNewBrake
CarryON:
Next i
Exit Sub
Errtrap:
If Err = 75 Then
Call MsgBox(strStock & Lang(511) & vbCrLf & Lang(512), vbExclamation, App.Title)
Exit Sub
End If

End Sub


Private Sub mnuFRigid_Click()
Dim OldCol As Integer, OldRow As Integer
Dim strStock As String, strStockPath As String, flagway As Integer
Dim tempRow(0 To 100) As Integer, tempTtl As Integer

On Error GoTo Errtrap
For i = 0 To GridStock.SelectedRows - 1
tempRow(i) = GridStock.SelectedRow(i)
Next i

tempTtl = GridStock.SelectedRows - 1
For i = 0 To tempTtl
OldCol = 10
OldRow = tempRow(i)

GridStock.Select OldRow, OldCol - 7
strStock = GridStock.Cell(flexcpText)
If strStock = "Default.wag" Then GoTo CarryON
GridStock.Select OldRow, OldCol - 8
strStockPath = GridStock.Cell(flexcpText)
strStockPath = MSTSPath & "\trains\trainset\" & strStockPath

SparePath = App.Path & "\" & "TempFiles"
FileCopy strStockPath & "\" & strStock, SparePath & "\" & strStock

flagway = 0
Call ConvertFRigid(SparePath & "\" & strStock, flagway)

flagway = 1
Call ConvertFRigid(SparePath & "\" & strStock, flagway)
DoEvents
FileCopy SparePath & "\" & strStock, strStockPath & "\" & strStock
DoEvents
Kill SparePath & "\" & strStock
GridStock.Select OldRow, OldCol
GridStock.Cell(flexcpText) = "Rigid"

CarryON:
Next i
Exit Sub
Errtrap:
If Err = 75 Then
Call MsgBox(strStock & Lang(511) & vbCrLf & Lang(512), vbExclamation, App.Title)
Exit Sub
End If
End Sub

Private Sub mnuRigid_Click()
Dim OldCol As Integer, OldRow As Integer
Dim strStock As String, strStockPath As String, flagway As Integer
Dim tempRow(0 To 100) As Integer, tempTtl As Integer

On Error GoTo Errtrap
For i = 0 To GridStock.SelectedRows - 1
tempRow(i) = GridStock.SelectedRow(i)
Next i

tempTtl = GridStock.SelectedRows - 1
For i = 0 To tempTtl
OldCol = 9
OldRow = tempRow(i)

GridStock.Select OldRow, OldCol - 6
strStock = GridStock.Cell(flexcpText)
If strStock = "Default.wag" Then GoTo CarryON
GridStock.Select OldRow, OldCol - 7
strStockPath = GridStock.Cell(flexcpText)
strStockPath = MSTSPath & "\trains\trainset\" & strStockPath

SparePath = App.Path & "\" & "TempFiles"
FileCopy strStockPath & "\" & strStock, SparePath & "\" & strStock

flagway = 0
Call ConvertRigid(SparePath & "\" & strStock, flagway)

flagway = 1
Call ConvertRigid(SparePath & "\" & strStock, flagway)
DoEvents
FileCopy SparePath & "\" & strStock, strStockPath & "\" & strStock
DoEvents
Kill SparePath & "\" & strStock
GridStock.Select OldRow, OldCol
GridStock.Cell(flexcpText) = "Rigid"

CarryON:
Next i
Exit Sub
Errtrap:
If Err = 75 Then
Call MsgBox(strStock & Lang(511) & vbCrLf & Lang(512), vbExclamation, App.Title)
Exit Sub
End If

End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, x As Integer, strTemp As String, strStock As String
If KeyCode = 13 Then

strTemp = Text1.Text

For i = 1 To GridStock.Rows - 1
GridStock.Select i, 3
strStock = GridStock.Cell(flexcpText)
x = InStr(strStock, Text1)
If x > 0 Then
foundrow = i
GridStock.TopRow = foundrow
GridStock.row = foundrow
GridStock.TopRow = foundrow

Exit For
End If
Next i
ElseIf KeyCode = 114 Then
Command11.value = True

End If
End Sub

