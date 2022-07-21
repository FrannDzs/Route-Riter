VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmLinks 
   Caption         =   "Linked Files"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   10080
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8LCtl.VSFlexGrid GridLinks 
      Height          =   7935
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   10455
      _cx             =   18441
      _cy             =   13996
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLinks.frx":0000
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
   Begin VB.CommandButton Command3 
      Caption         =   "Delete all Selected Files"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete all Unlinked Files"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8640
      TabIndex        =   0
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Files"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   9120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Files with 1 link are NOT linked to any other file."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "frmLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Command2_Click()
Dim j As Long

GridLinks.col = 0

Select Case MsgBox("Confirm you wish to delete all unlinked files from the Common Folders" _
                   & vbCrLf & "This operation can not be reversed....." _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, "WARNING...")

    Case vbYes
For j = 1 To GridLinks.Rows - 1
GridLinks.Select j, 0

If Val(GridLinks.Cell(flexcpText)) = 1 Then

GridLinks.Select j, 1
Kill GridLinks.Cell(flexcpText, j, 1)

DoEvents

End If
Next j
Case vbNo
Exit Sub
End Select
End Sub

Private Sub Command3_Click()

Dim iCnt As Long, i As Long
Select Case MsgBox("Confirm you wish to delete all selected files from the Common Folders" _
                   & vbCrLf & "This operation can not be reversed....." _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, "WARNING...")

    Case vbYes
    iCnt = GridLinks.SelectedRows

For i = 0 To iCnt - 1
      
     Kill GridLinks.Cell(flexcpText, GridLinks.SelectedRow(i), 1)
     DoEvents
    Next
    Case vbNo
Exit Sub
End Select





End Sub


Private Sub Form_Load()

If booLink = False Then
Command2.Visible = False
Command3.Visible = False
ElseIf booLink = True Then
Command2.Visible = True
Command3.Visible = True
End If
GridLinks.BackColor = vbWhite
GridLinks.Rows = 1
GridLinks.Cell(flexcpFontBold, 0, 0, 0, 1) = True
GridLinks.ExplorerBar = flexExSort

DoEvents


End Sub




Private Sub GridLinks_GotFocus()
Label3.Caption = GridLinks.Rows
End Sub


