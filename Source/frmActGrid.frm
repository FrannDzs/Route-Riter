VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmActGrid 
   Caption         =   "Activity Details for:-"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   15330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "List Rolling-Stock used"
      Height          =   615
      Left            =   11880
      TabIndex        =   16
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Package Activity"
      Height          =   615
      Left            =   10800
      TabIndex        =   15
      Top             =   9000
      Width           =   975
   End
   Begin VSFlex8LCtl.VSFlexGrid Grid3 
      Height          =   4815
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   10095
      _cx             =   17806
      _cy             =   8493
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   2000
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13080
      Top             =   9000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Save as CSV"
      Height          =   615
      Left            =   6840
      TabIndex        =   7
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Error Report"
      Height          =   615
      Left            =   5400
      TabIndex        =   6
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Print Loose Consists"
      Height          =   615
      Left            =   9600
      TabIndex        =   5
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print Consists"
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   9000
      Width           =   1215
   End
   Begin VSFlex8LCtl.VSFlexGrid Grid2 
      Height          =   7695
      Left            =   10440
      TabIndex        =   2
      Top             =   840
      Width           =   3255
      _cx             =   5741
      _cy             =   13573
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
      BackColor       =   -2147483639
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483639
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
      FormatString    =   $"frmActGrid.frx":0000
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
      Left            =   8400
      TabIndex        =   1
      Top             =   9000
      Width           =   1095
   End
   Begin VSFlex8LCtl.VSFlexGrid Grid1 
      Height          =   2775
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   7575
      _cx             =   13361
      _cy             =   4895
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
      BackColor       =   -2147483639
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483639
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmActGrid.frx":003A
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
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   8280
      TabIndex        =   13
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   12
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   11
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Traffic:"
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   10
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Activity:"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Route:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Items with a RED background are missing ! Or Faulty"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   9000
      Width           =   3255
   End
End
Attribute VB_Name = "frmActGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim booSelect As Boolean

Dim cn As Integer, cnn As Integer
Dim booTraffic As Boolean
Dim strConsistNames As String

Dim strStockName() As String


Private Sub SetLan()
Label1.Caption = Lang(183)
Command3.Caption = Lang(184)
Command3.ToolTipText = Lang(185)
Command7.Caption = Lang(186)
Command7.ToolTipText = Lang(187)
Command8.Caption = Lang(188)
Command8.ToolTipText = Lang(189)
Command1.Caption = Lang(38)
Command6.Caption = Lang(190)
Command6.ToolTipText = Lang(191)
Command2.Caption = Lang(192)
Command2.ToolTipText = Lang(193)
Command4.Caption = Lang(194)
Command4.ToolTipText = Lang(195)
frmActGrid.Caption = Lang(196)
For i = 0 To 2
Label2(i).Caption = Lang(197 + i)
Next


End Sub

Private Function ReadUniFile(CompleteFilePath As String) As String

Dim length As Long, mytristate As Integer
Dim MyString As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean


Set File_obj = CreateObject("Scripting.FileSystemObject")
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox "" & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & "" & Lang(401), vbInformation, Me.Caption
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

Private Sub GetStockName(LocoPath As String, strName As String)
Dim NewFile As Integer, A$, x As Long, xx As Long

On Error GoTo Errtrap

NewFile = FreeFile

A$ = ReadUniFile(LocoPath)


x = InStr(A$, "Name (")
If x = 0 Then
x = InStr(A$, "Name  (")
End If
If x = 0 Then
x = InStr(A$, "Name   (")
End If
If x = 0 Then
x = InStr(A$, "Name(")
End If
If x > 0 Then
xx = InStr(x, A$, "(")
xy = InStr(xx, A$, ")")
strName = Mid$(A$, xx + 1, xy - xx - 1)
strName = Trim$(strName)
If Left$(strName, 1) = ChrW$(34) Then
strName = Mid$(strName, 2, Len(strName) - 2)
End If

End If

Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'GetCoupling' please advise" _
            & vbCrLf & "while checking " & LocoPath _
            , vbExclamation, App.Title)
Resume Next

End Sub


Private Sub Command1_Click()
'Grid1.Clear
'Grid2.Clear
'Me.Hide
Unload Me
End Sub


Private Sub Command2_Click()
Dim MyCol As Integer, strSpare As String, intFlag As Integer
Dim i As Integer, NewFile As Integer, ActName As String
Dim ThisDrive As String, x As Integer, strName As String

On Error GoTo Errtrap

ReDim Preserve strConsists(0 To cn)
ReDim Preserve strStockName(0 To cn)

FromAct = True
strSpare = App.Path & "\TempFiles\Train Simulator"

Rem ************ Clean up SetupFiles folder ******************
DoEvents
ThisDrive = Left$(strSpare, 1)
ChDrive ThisDrive
If DirExists(strSpare & "\Routes\" & Label3(0).Caption & "\Activities\") Then
ChDir strSpare & "\Routes\" & Label3(0).Caption & "\Activities\"
Kill "*.*"
ChDir ".."
RmDir "Activities"
End If
If DirExists(strSpare & "\Routes\" & Label3(0).Caption & "\Services\") Then
ChDir strSpare & "\Routes\" & Label3(0).Caption & "\Services\"
Kill "*.*"
ChDir ".."
RmDir "Services"
End If
If DirExists(strSpare & "\Routes\" & Label3(0).Caption & "\Paths\") Then
ChDir strSpare & "\Routes\" & Label3(0).Caption & "\Paths\"
Kill "*.*"
ChDir ".."
RmDir "Paths"
End If
If DirExists(strSpare & "\Routes\" & Label3(0).Caption & "\Traffic\") Then
ChDir strSpare & "\Routes\" & Label3(0).Caption & "\Traffic\"
Kill "*.*"
ChDir ".."
RmDir "Traffic"
Kill "*.*"
End If
ChDir ".."
If DirExists(Label3(0).Caption) Then
RmDir Label3(0).Caption
End If
ChDir ".."
If DirExists("Routes") Then
RmDir "Routes"
End If
If DirExists(strSpare & "\trains\consists\") Then
ChDir strSpare & "\trains\consists\"
Kill "*.*"
ChDir ".."
RmDir "consists"
ChDir ".."
RmDir "Trains"
End If
ChDir ".."
'RmDir "Train Simulator"
Rem *********************************************************

intFlag = 0
x = 1
If Not DirExists(strSpare) Then
MkDir strSpare
End If
x = 2
If Not DirExists(strSpare & "\Routes") Then
MkDir strSpare & "\Routes\"
End If
x = 3

If Not DirExists(strSpare & "\Routes\" & Label3(0).Caption) Then
MkDir strSpare & "\Routes\" & Label3(0).Caption
End If
x = 4
If Not DirExists(strSpare & "\Routes\" & Label3(0).Caption & "Activities") Then
MkDir strSpare & "\Routes\" & Label3(0).Caption & "\Activities"
End If
x = 5
If Not DirExists(strSpare & "\Routes\" & Label3(0).Caption & "Services") Then
MkDir strSpare & "\Routes\" & Label3(0).Caption & "\Services"
End If
x = 6
If Not DirExists(strSpare & "\Routes\" & Label3(0).Caption & "Paths") Then
MkDir strSpare & "\Routes\" & Label3(0).Caption & "\Paths"
End If
x = 7
If Label3(2).Caption <> vbNullString Then
If Not DirExists(strSpare & "\Routes\" & Label3(0).Caption & "Traffic") Then
MkDir strSpare & "\Routes\" & Label3(0).Caption & "\Traffic"
End If
End If
x = 8
If Not DirExists(strSpare & "\Trains") Then
MkDir strSpare & "\Trains"
End If
x = 9
If Not DirExists(strSpare & "\Trains\Consists") Then
MkDir strSpare & "\Trains\Consists"
End If
ActName = Label3(1).Caption
ActName = Left$(ActName, Len(ActName) - 4)

x = 10
FileCopy MSTSPath & "\Routes\" & Label3(0).Caption & "\Activities\" & Label3(1).Caption, strSpare & "\Routes\" & Label3(0).Caption & "\Activities\" & Label3(1).Caption
If FileExists(MSTSPath & "\Routes\" & Label3(0).Caption & "\Activities\" & ActName & ".txt") Then
FileCopy MSTSPath & "\Routes\" & Label3(0).Caption & "\Activities\" & ActName & ".txt", strSpare & "\Routes\" & Label3(0).Caption & "\" & ActName & ".txt"

End If
x = 11
If FileExists(MSTSPath & "\Routes\" & Label3(0).Caption & "\Activities\" & ActName & ".asv") Then
FileCopy MSTSPath & "\Routes\" & Label3(0).Caption & "\Activities\" & ActName & ".asv", strSpare & "\Routes\" & Label3(0).Caption & "\Activities\" & ActName & ".asv"

End If
x = 12

If Label3(2).Caption <> vbNullString Then
booTraffic = True
FileCopy MSTSPath & "\Routes\" & Label3(0).Caption & "\Traffic\" & Label3(2).Caption, strSpare & "\Routes\" & Label3(0).Caption & "\Traffic\" & Label3(2).Caption
End If
x = 13
For i = 1 To intGrid
MyCol = 0
Grid1.Select i, MyCol

FileCopy MSTSPath & "\Routes\" & Label3(0).Caption & "\Services\" & Grid1.Cell(flexcpText), strSpare & "\Routes\" & Label3(0).Caption & "\Services\" & Grid1.Cell(flexcpText)
MyCol = 2
Grid1.Select i, MyCol
FileCopy MSTSPath & "\Routes\" & Label3(0).Caption & "\Paths\" & Grid1.Cell(flexcpText), strSpare & "\Routes\" & Label3(0).Caption & "\paths\" & Grid1.Cell(flexcpText)
Next i
x = 14
For i = 1 To intGrid
MyCol = 1
Grid1.Select i, MyCol

FileCopy MSTSPath & "\Trains\Consists\" & Grid1.Cell(flexcpText), strSpare & "\Trains\Consists\" & Grid1.Cell(flexcpText)
Next i
x = 15

Call QSort(strConsists(), 0, UBound(strConsists))
x = 16

Rem ************** New bit *********
For i = 1 To cn
Call GetStockName(MSTSPath & "\trains\trainset\" & strConsists(i), strName)
If strName <> vbNullString Then
strStockName(i) = "( " & strName & " )"
Else
strStockName(i) = vbNullString
End If
strName = vbNullString
Next i

Rem ********************************
NewFile = FreeFile
Open strSpare & "\Routes\" & Label3(0).Caption & "\" & ActName & "_RollingStockNeeded.txt" For Output As #NewFile
Print #NewFile, "The following Rolling-Stock is needed to run this activity:-" & vbCrLf & vbCrLf
For i = 1 To cn
If strStockName(i) <> vbNullString Then
Print #NewFile, strConsists(i) & vbTab & vbTab & strStockName(i)
Else
Print #NewFile, strConsists(i)
End If
Next i
Close #NewFile
x = 17
NewZipPath = strSpare & "\"
FromZip = 2
frmNewZip.Show 1

x = 18
DoEvents
ThisDrive = Left$(strSpare, 1)
ChDrive ThisDrive


ChDir strSpare & "\Routes\" & Label3(0).Caption & "\Activities\"
Kill "*.*"
ChDir ".."
RmDir "Activities"
x = 19
ChDir strSpare & "\Routes\" & Label3(0).Caption & "\Services\"
Kill "*.*"
ChDir ".."
RmDir "Services"
x = 20
ChDir strSpare & "\Routes\" & Label3(0).Caption & "\Paths\"
Kill "*.*"
ChDir ".."
RmDir "Paths"
x = 21
If booTraffic = True Then
ChDir strSpare & "\Routes\" & Label3(0).Caption & "\Traffic\"
Kill "*.*"
ChDir ".."
RmDir "Traffic"
End If
x = 22
Kill "*.*"

ChDir ".."
RmDir Label3(0).Caption

ChDir ".."
RmDir "Routes"
ChDir strSpare & "\trains\consists\"
Kill "*.*"
x = 23
ChDir ".."
RmDir "consists"
x = 24
ChDir ".."
RmDir "Trains"
x = 25

ChDir ".."
RmDir "Train Simulator"
If DirExists(strSpare) Then
Call MsgBox(strSpare & Lang(491) & vbCrLf & Lang(492), vbExclamation, Lang(493))

End If
x = 26
Exit Sub
Errtrap:
If Err = 75 Then
Resume Next
End If
Call MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'Package Activity' please advise" _
            & vbCrLf & "that x=" & Str(x) _
            , vbExclamation, App.Title)


Resume Next
End Sub

Private Sub Command3_Click()
'flagPrint = 4
'fEZPrint.Show
Grid3.PrintGrid , True

End Sub

Private Sub Command4_Click()
Dim strStock As String, strName As String
ReDim Preserve strConsists(0 To cn)
ReDim Preserve strStockName(0 To cn)

Call QSort(strConsists(), 0, UBound(strConsists))
For i = 1 To cn
Call GetStockName(MSTSPath & "\trains\trainset\" & strConsists(i), strName)
If strName <> vbNullString Then
strStockName(i) = "( " & strName & " )"
Else
strStockName(i) = vbNullString
End If
strName = vbNullString
Next i

For i = 1 To cn
strStock = strStock & strConsists(i) & vbTab & vbTab & strStockName(i) & vbCrLf
Next i
strActReport = strActReport & vbCrLf & Label3(1).Caption & vbCrLf & strConsistNames & vbCrLf & vbCrLf & strStock & vbCrLf & vbCrLf
frmReport.Rich1.Text = strActReport
     frmReport.Show 1
     'cn = 0
End Sub

Private Sub Command6_Click()
'flagPrint = 12
'fEZPrint.Show
Grid2.PrintGrid , True
End Sub

Private Sub Command7_Click()
strActReport = strActReport & Label3(1).Caption & " Errors " & vbCrLf & strbadbits & vbCrLf

'frmReport.Show 1

End Sub

Private Sub Command8_Click()
Dim tit1 As String, tit2 As String, tit3 As String
Dim strTemp As String, strTemp2 As String

Call MsgBox(Lang(494) & vbCrLf & Lang(495), vbInformation, Lang(496))

CommonDialog1.Filter = "Comma Separated (*.csv)|*.csv|Tab Separated (*.txt)|*.txt"
CommonDialog1.DialogTitle = "Save Grid to File"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 2
CommonDialog1.DefaultExt = "csv"
tit1 = CommonDialog1.Filename
If tit1 = vbNullString Then Exit Sub
strTemp = Left$(tit1, Len(tit1) - 4)
strTemp2 = Right$(tit1, 4)
tit1 = strTemp & "a" & strTemp2
tit2 = strTemp & "b" & strTemp2
tit3 = strTemp & "c" & strTemp2

If tit1 <> vbNullString Then
If Right$(tit1, 3) = "csv" Then
Grid1.SaveGrid tit1, flexFileCommaText
Grid2.SaveGrid tit2, flexFileCommaText
Grid3.SaveGrid tit3, flexFileCommaText
Else
Grid1.SaveGrid tit1, flexFileTabText
Grid2.SaveGrid tit2, flexFileTabText
Grid3.SaveGrid tit3, flexFileTabText
End If
End If
End Sub

Private Sub Form_Load()
Dim i%, tempPath As String, x As Integer, Y As Integer, j As Integer, SPath As String
Dim missSrv As Boolean, ActName As String
On Error GoTo Errtrap

frmActGrid.Caption = frmUtils.Caption

DoEvents
ReDim strConsists(0 To 1500)
Call SetLan
cn = 0
strConsistNames = "Consists used:-" & vbCrLf

'strbadbits = vbNullString
intGrid = 0
Label1.Caption = Lang(183)

   Grid1.AllowUserResizing = flexResizeBoth
    Grid1.ExtendLastCol = True
    
   Grid1.Rows = 1
   Grid1.Cell(flexcpFontBold, 0, 0, 0, 2) = True
   Grid2.Cell(flexcpFontBold, 0, 0) = True
   Grid1.ExplorerBar = flexExSort
   Grid1.BackColor = vbWhite
   Grid2.BackColor = vbWhite
   
If BooCheckAct = True Then

'Checking single activities
i = 0
   j = 0
      'x = InStrRev(ActPath(i), "\")
      x = InStrRev(ActPath(i), "\")
      Y = InStrRev(ActPath(i), "\", x - 1)
      tempPath = Mid$(ActPath(i), Y + 1, (x - 1) - Y)
      SPath = Left$(ActPath(i), Len(ActPath(i)) - 10)
      Label3(0).Caption = tempPath
      Label3(1).Caption = Activities(i)
      Label3(2).Caption = PTfcName(i)
      Rem ***********Loose Consists *************


RoutePath = ActPath(i)
ActName = Label3(1).Caption

Grid2.Rows = 0
Grid2.AddItem Lang(634)
Grid2.Rows = 1
Grid2.ExplorerBar = flexExSort

Call LooseConsistsGrid(ActName)

      
     Rem ******************************************
      If Not FileExists(MSTSPath & "\trains\consists\" & PConName(i, j)) Then

          FlagColRed = True

      End If
      If PConName(i, j) <> vbNullString Then
     
      Call CheckForConsistCol(PConName(i, j), Activities(i), PSvcName(i, j))
      End If
      If Not FileExists(SPath & "Services\" & PSvcName(i, j)) And PSvcName(i, j) <> vbNullString Then
        'flagSvcRed = True
       
          If missSrv = False Then
            strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(524) & vbCrLf
            missSrv = True
        End If
         strbadbits = strbadbits & vbCrLf & Lang(524) & PSvcName(i, j) & Lang(627) & Activities(i) & vbCrLf
      End If
     
      Call LooseActivities(ActPath(i) & "\" & Activities(i))
      Grid1.AddItem PSvcName(i, j) & vbTab & PConName(i, j) & vbTab & pPathName(i, j)
      intGrid = intGrid + 1
AnyMore2:
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
        Grid1.AddItem PSvcName(i, j) & vbTab & PConName(i, j) & vbTab & pPathName(i, j)
      End If
      If PSvcName(i, j) <> vbNullString And missSrv = False Then
      
      If Not FileExists(MSTSPath & "\trains\consists\" & PConName(i, j)) Then
       FlagColRed = True
       'strbadbits = strbadbits & vbCrLf & "Missing Consist " & PConName(i, j)
      End If
       Call CheckForConsistCol(PConName(i, j), Activities(i), PSvcName(i, j))
       'Grid1.AddItem tempPath & vbtab & Activities(i) & vbtab & PTfcName(i) & vbtab & PSvcName(i, j) & vbtab & PConName(i, j) & vbtab & pPathName(i, j)
       Grid1.AddItem PSvcName(i, j) & vbTab & PConName(i, j) & vbTab & pPathName(i, j)
       intGrid = intGrid + 1
      GoTo AnyMore2
      End If
   ' Next

Grid1.col = 0
Grid1.Sort = flexSortStringAscending
frmActGrid.Caption = frmUtils.Caption

DoEvents
If strbadbits <> vbNullString Then
frmReport.Rich1.Text = strbadbits
frmReport.Show 1
frmReport.ZOrder

End If
End If
Exit Sub
Errtrap:
If Err = 53 Then

Call MsgBox(Lang(497), vbExclamation, Lang(347))

Resume Next
Else

Call MsgBox("An error " & Err.Description & " occurred in subroutine 'Load ActGrid' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)

Resume Next
End If

End Sub

Private Sub LooseActivities(ActPath As String)
Dim x As Integer, tempService As String
Dim ActName As String, missAct As Boolean
Dim booEntry As Boolean, strNew As String
Dim Engpath As String, Engname As String
Dim Wagname As String, Wagonpath As String, strFoundPath As String
On Error GoTo Errtrap

MousePointer = 11
RoutePath = MSTSPath & "\Routes"
TrainsetPath = MSTSPath & "\Trains\Trainset\"

x = InStrRev(ActPath, "\")
ActName = Mid$(ActPath, x + 1)


NewFile = FreeFile
 Open ActPath For Input As #NewFile
 Do While Not EOF(NewFile)
Line Input #NewFile, strNew
 
 Rem ************* Find loose consists - Locos
 tempService = vbNullString
   x = InStr(strNew, "EngineData")
   
      If x > 0 Then
     
   Call CheckEngineData(strNew, Engname, Engpath, booEntry)

   
  Engname = Engname & ".eng"
  
If booEntry = True Then
If Not FileExists(TrainsetPath & Engpath & "\" & Engname) Then
strFoundPath = vbNullString
Call LookForLoco(Engname, strFoundPath)
If strFoundPath = vbNullString Then
  flagActBad = True

  strbadbits = strbadbits & vbCrLf & Lang(525) & ActName & ", " & Lang(526) & Engname
Else
flagActBad = True
strbadbits = strbadbits & vbCrLf & Lang(525) & ActName & ", " & Lang(526) & Engname & vbCrLf & Lang(631) & strFoundPath
 End If
 End If
 End If
 End If
 Rem ************* Find loose consists - Wagons
 tempService = vbNullString
   x = InStr(strNew, "WagonData")
   
   booEntry = False
   x = InStr(strNew, "wagonData")
      If x > 0 Then
   Call CheckWagonData(strNew, Wagname, Wagonpath, booEntry)
Wagname = Wagname & ".wag"
 If booEntry = True Then
 If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname) Then
 
 strFoundPath = vbNullString
Call LookForWag(Wagname, strFoundPath)
If strFoundPath = vbNullString Then
  flagActBad = True
  strbadbits = strbadbits & vbCrLf & Lang(525) & ActName & ", " & Lang(527) & Wagname
Else
 flagActBad = True
           If missAct = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(525) & vbCrLf
missAct = True
End If
 strbadbits = strbadbits & vbCrLf & Lang(525) & ActName & ", " & Lang(527) & Wagname & vbCrLf & Lang(632) & strFoundPath
 End If
 End If
 End If
 End If
 
 Loop
 Close #NewFile

MousePointer = 0
Exit Sub
Errtrap:
Call MsgBox("An error occurred in subroutine 'LooseActivities' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)

'Resume Next

End Sub

Private Sub LookForWag(strWag As String, strFoundPath As String)
Dim TrainsetPath As String, i As Integer

TrainsetPath = MSTSPath & "\Trains\Trainset\"
For i = 0 To lngWagons - 1
If strWag = Wagons(i) Then

If FileExists(TrainsetPath & Wagpath(i) & "\" & strWag) Then
strFoundPath = TrainsetPath & Wagpath(i) & "\" & strWag
End If
End If
Next i
End Sub

Private Sub LookForLoco(strEng As String, strFoundPath As String)
Dim TrainsetPath As String

TrainsetPath = MSTSPath & "\Trains\Trainset\"
For i = 0 To lngLoco - 1
If strEng = Locomotives(i) Then

If FileExists(TrainsetPath & LocoPath(i) & "\" & strEng) Then
strFoundPath = TrainsetPath & LocoPath(i) & "\" & strEng
End If
End If
Next i
End Sub
Private Sub LooseConsistsGrid(ActPath As String)
Dim x As Integer
Dim Z As Integer, strTemp As String, booCon As Boolean
Dim booEntry As Boolean, strNew As String, Engname As String, Engpath As String
Dim Wagname As String, Wagonpath As String
On Error GoTo Errtrap

MousePointer = 11
'RoutePath = MSTSPath & "\Routes"
TrainsetPath = MSTSPath & "\Trains\Trainset\"
Z = 1
ActPath = RoutePath & "\" & ActPath


NewFile = FreeFile
 Open ActPath For Input As #NewFile
 Do While Not EOF(NewFile)
Line Input #NewFile, strNew
 
 Rem ************* Find loose consists - Locos

   x = InStr(strNew, "EngineData")
      If x > 0 Then
   Call CheckEngineData(strNew, Engname, Engpath, booEntry)
 

   
  
If booEntry = True Then
Engname = Engname & ".eng"
   If Not FileExists(TrainsetPath & Engpath & "\" & Engname) Then

  FlagColRed = True
  If missEng = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(528) & vbCrLf
missEng = True
End If
  strbadbits = strbadbits & vbCrLf & Lang(528) & Engname
   Grid2.AddItem Engpath & "\" & Engname
   strTemp = Engpath & "\" & Engname
   For cnn = 0 To cn
If strTemp = strConsists(cnn) Then
booCon = True
Exit For
End If
Next
If booCon = False Then
cn = cn + 1
strConsists(cn) = strTemp
End If
booCon = False
   Z = Z + 1
   If Grid2.Rows < Z Then
   Grid2.Rows = Z
   End If
'   Grid2.CellBackColor = &HFF
   Else
   FlagColRed = False
   Grid2.AddItem Engpath & "\" & Engname
      strTemp = Engpath & "\" & Engname
   For cnn = 0 To cn
If strTemp = strConsists(cnn) Then
booCon = True
Exit For
End If
Next
If booCon = False Then
cn = cn + 1
strConsists(cn) = strTemp
End If
booCon = False
   If Grid2.Rows < Z Then
   Grid2.Rows = Z
   End If
   End If
   
 End If
 End If
 Rem ************* Find loose consists - Wagons

   booEntry = False
   x = InStr(strNew, "wagonData")
      If x > 0 Then
   Call CheckWagonData(strNew, Wagname, Wagonpath, booEntry)
 

 If booEntry = True Then
 Wagname = Wagname & ".wag"
  If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname) Then
 FlagColRed = True
 If missWag = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(529) & vbCrLf
missWag = True
End If
 strbadbits = strbadbits & vbCrLf & Lang(529) & Wagname
 ' Grid2.CellForeColor = vbRed
   Grid2.AddItem Wagonpath & "\" & Wagname
      strTemp = Wagonpath & "\" & Wagname
   For cnn = 0 To cn
If strTemp = strConsists(cnn) Then
booCon = True
Exit For
End If
Next
If booCon = False Then
cn = cn + 1
strConsists(cn) = strTemp
End If
booCon = False
   If Grid2.Rows < Z Then
   Grid2.Rows = Z
   End If
'   Grid2.CellBackColor = &HFF
   Else
   FlagColRed = False
   Grid2.AddItem Wagonpath & "\" & Wagname
         strTemp = Wagonpath & "\" & Wagname
   For cnn = 0 To cn
If strTemp = strConsists(cnn) Then
booCon = True
Exit For
End If
Next
If booCon = False Then
cn = cn + 1
strConsists(cn) = strTemp
End If
booCon = False
   If Grid2.Rows < Z Then
   Grid2.Rows = Z
   End If
   End If
   
 End If
 End If
 Loop
 Close #NewFile

MousePointer = 0
Exit Sub
Errtrap:
Call MsgBox("An error occurred in subroutine 'LooseConsistsGrid' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)

Resume Next

End Sub

Private Sub Form_Resize()
On Error GoTo Errtrap

Label3(0).Top = 240
Label3(1).Top = 240
Label3(2).Top = 240
Label2(0).Top = 240
Label2(1).Top = 240
Label2(2).Top = 240

Grid1.width = Me.width * 0.55
Grid1.Left = 1680
Grid3.Left = 120
Grid3.Top = Grid1.Top + Grid1.height + 100
Grid3.height = Me.height * 0.5
Grid3.width = Me.width * 0.73
Grid2.width = Me.width * 0.25
Grid2.Left = Grid3.width + Grid3.Left + 50  'Grid1.Left + Grid1.Width + 50
Grid1.Top = Label3(0).Top + Label3(0).height + 150
Grid2.Top = Grid1.Top
Grid1.height = Me.height * 0.25
Grid2.height = Me.height * 0.75


Label1.Top = Grid3.Top + Grid3.height + 150
Command1.Top = Label1.Top
'Command2.Top = Label1.Top
Command3.Top = Label1.Top
Command4.Top = Label1.Top

Command6.Top = Label1.Top
Command7.Top = Label1.Top
Command8.Top = Label1.Top
Command2.Top = Label1.Top
Command1.Left = Grid2.Left
Label1.Left = Grid1.Left
'Command2.Left = Label1.Left + Label1.Width + 100
'Command3.Left = Command2.Left + Command2.Width + 100
Command3.Left = Label1.Left + Label1.width + 100

Command7.Left = Command3.Left + Command3.width + 100
Command8.Left = Command7.Left + Command7.width + 100
Command2.Left = Command8.Left + Command8.width + 100
Command4.Left = Command2.Left + Command2.width + 100

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
Dim nextrow As Integer, nextcol As Integer

If flagSvcRed = True And col = 0 Then
Grid1.Select row, col
Grid1.FillStyle = flexFillSingle
Grid1.CellBackColor = &H8080FF
flagSvcRed = False
End If
If col = 1 And row > 0 Then
Grid1.Select row, col


nextcol = row - 1
If Grid3.Cols < nextcol Then
Grid3.Cols = nextcol + 2
End If
If Grid3.Cols <= row Then
Grid3.Cols = row + 2
End If
DoEvents
Grid3.col = row
Grid3.row = 0
ConsistPath = MSTSPath & "\Trains\Consists"
'Grid3.Rows = 0
'Grid3.AddItem "Consist Items:-"
'Grid3.Rows = 1
Grid3.ExplorerBar = flexExSort

Grid3.Select 0, nextcol
Grid3.Cell(flexcpText) = Grid1.Cell(flexcpText)
nextrow = 1
Call CheckForConsistGrid(ConsistPath & "\" & Grid1.Cell(flexcpText), nextrow, nextcol)
End If
End Sub




Private Sub CheckForConsistCol(CFilepath As Variant, strAct As Variant, strSvc As Variant)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Engpath As String, Engname As String
Dim Wagonpath As String, Wagname As String
Dim TrainsetPath As String, ConName As String, ConP As String
Dim booEntry As Boolean

On Error GoTo Errtrap
ConName = CFilepath
ConP = MSTSPath & "\Trains\Consists\" & CFilepath
Fnumber = FreeFile

'strConsistNames = strConsistNames & ConName & vbCrLf
TrainsetPath = MSTSPath & "\Trains\Trainset\"
If Not FileExists(ConP) Then
strbadbits = strbadbits & vbCrLf & Lang(530) & ConName & Lang(531) & strAct & Lang(532) & strSvc & vbCrLf
Exit Sub
End If

Open ConP For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   x = InStr(strNew, "EngineData")
      If x > 0 Then
   Call CheckEngineData(strNew, Engname, Engpath, booEntry)

   If booEntry = True Then
   If Not FileExists(TrainsetPath & Engpath & "\" & Engname & ".eng") Then

  flagConBad = True
  strbadbits = strbadbits & vbCrLf & Lang(533) & ConName & " " & Lang(528) & Engname & ".eng"
 ' Close #Fnumber
  ' Exit Sub
   End If
   End If
   End If
   booEntry = False
   x = InStr(strNew, "wagonData")
      If x > 0 Then
   Call CheckWagonData(strNew, Wagname, Wagonpath, booEntry)

If booEntry = True Then
  If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname & ".wag") Then
 flagConBad = True
  strbadbits = strbadbits & vbCrLf & Lang(533) & ConName & " " & Lang(529) & Wagname & ".wag"

   End If
   
  End If
   End If
   strNew = vbNullString
 
   Loop
   Close #Fnumber

Exit Sub
Errtrap:
Call MsgBox("An error occurred in subroutine 'CheckForConsistCol' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)

'Resume Next

End Sub

Private Sub CheckForConsistGrid(CFilepath As String, nextrow As Integer, nextcol As Integer)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Engpath As String, Engname As String
Dim strNew2 As String, Wagonpath As String, Wagname As String
Dim TrainsetPath As String, ConName As String, strTemp As String, booCon As Boolean
Dim booEntry As Boolean

On Error GoTo Errtrap
If Grid3.Cols < nextcol Then
Grid3.Cols = nextcol
End If
Fnumber = FreeFile
x = InStrRev(CFilepath, "\")
ConName = Mid$(CFilepath, x + 1)
strConsistNames = strConsistNames & ConName & vbCrLf
TrainsetPath = MSTSPath & "\Trains\Trainset\"
If Not FileExists(CFilepath) Then

'Call MsgBox(Lang(352) & ConName & Lang(353), vbExclamation, Lang(530))
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
  Grid3.Select nextrow, nextcol
  Grid3.FillStyle = flexFillSingle
Grid3.CellBackColor = vbRed
FlagColRed = False
Grid3.Cell(flexcpText) = Engpath & "\" & Engname & ".eng"

strTemp = Engpath & "\" & Engname & ".eng"

For cnn = 0 To cn
If strTemp = strConsists(cnn) Then
booCon = True
Exit For
End If
Next
If booCon = False Then
cn = cn + 1
strConsists(cn) = strTemp
End If
booCon = False
nextrow = nextrow + 1
If Grid3.Rows < nextrow Then
Grid3.Rows = nextrow
End If
'  Grid3.Col = NextCol
'   Grid3.AddItem EngName & ".eng"
   If missEng = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(528) & vbCrLf
missEng = True
End If
   strbadbits = strbadbits & vbCrLf & Lang(528) & Engname & ".eng"
   Else
     Grid3.Select nextrow, nextcol
Grid3.Cell(flexcpText) = Engpath & "\" & Engname & ".eng"

strTemp = Engpath & "\" & Engname & ".eng"
For cnn = 0 To cn
If strTemp = strConsists(cnn) Then
booCon = True
Exit For
End If
Next
If booCon = False Then
cn = cn + 1
strConsists(cn) = strTemp
End If
booCon = False
nextrow = nextrow + 1
If Grid3.Rows < nextrow Then
Grid3.Rows = nextrow
End If
'  Grid3.Col = NextCol
'
'   Grid3.AddItem EngName & ".eng"
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
 
 FlagColRed = True
 If missWag = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(529) & vbCrLf
missWag = True
End If
  Grid3.Select nextrow, nextcol
    Grid3.FillStyle = flexFillSingle
Grid3.CellBackColor = vbRed
FlagColRed = False
Grid3.Cell(flexcpText) = Wagonpath & "\" & Wagname & ".wag"
strTemp = Wagonpath & "\" & Wagname & ".wag"
For cnn = 0 To cn
If strTemp = strConsists(cnn) Then
booCon = True
Exit For
End If
Next
If booCon = False Then
cn = cn + 1
strConsists(cn) = strTemp
End If
booCon = False
nextrow = nextrow + 1
If Grid3.Rows < nextrow Then
Grid3.Rows = nextrow
End If
'Grid3.Col = NextCol
'   Grid3.AddItem WagName & ".wag"
   strbadbits = strbadbits & vbCrLf & Lang(529) & Wagname & ".wag"
   Else
     Grid3.Select nextrow, nextcol
Grid3.Cell(flexcpText) = Wagonpath & "\" & Wagname & ".wag"
strTemp = Wagonpath & "\" & Wagname & ".wag"
For cnn = 0 To cn
If strTemp = strConsists(cnn) Then
booCon = True
Exit For
End If
Next
If booCon = False Then
cn = cn + 1
strConsists(cn) = strTemp
End If
booCon = False
nextrow = nextrow + 1
If Grid3.Rows < nextrow Then
Grid3.Rows = nextrow
End If
'   Grid3.Col = NextCol
'   Grid3.AddItem WagName & ".wag"
   End If
   
  End If
   End If
   strNew = vbNullString
 ' itExists = False
   Loop
   Close #Fnumber
  
Exit Sub
Errtrap:

If Err = 381 Then
Resume Next
Else
Call MsgBox("An error #" & Err & " occurred in subroutine 'CheckForConsistGrid' while checking" _
            & vbCrLf & CFilepath _
            , vbExclamation, App.Title)

Resume Next
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





