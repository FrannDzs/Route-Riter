VERSION 5.00
Begin VB.Form frmDelete 
   Caption         =   "Select Files to Delete or Move"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   8475
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "&Move"
      Height          =   495
      Left            =   5040
      TabIndex        =   13
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Select All"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin VB.ListBox lstFoundFiles 
      Height          =   1185
      Index           =   2
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   7
      ToolTipText     =   "Paths"
      Top             =   3840
      Width           =   6135
   End
   Begin VB.ListBox lstFoundFiles 
      Height          =   1185
      Index           =   1
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   6
      ToolTipText     =   "Consists"
      Top             =   2400
      Width           =   6135
   End
   Begin VB.ListBox lstFoundFiles 
      Height          =   1185
      Index           =   0
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   5
      ToolTipText     =   "Services"
      Top             =   960
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000A&
      Caption         =   "Traffic - "
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   7215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000A&
      Caption         =   "Activity - "
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Paths:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Consists:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Services:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   5760
      Width           =   3855
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SetLang()
Me.Caption = Lang(220)
Command3.Caption = Lang(216)
Command4.Caption = Lang(217)
Command1.Caption = Lang(218)
Command5.Caption = Lang(219)
Command2.Caption = Lang(38)
End Sub

Private Sub Command1_Click()
Dim MyCol As Integer, MyRow As Integer, tempText As String, tempPath As String
Dim Rname As String, booKilled As Boolean

MyRow = frmGrid.Grid1.row
MyCol = frmGrid.Grid1.col
If MyCol = 1 Then
tempText = frmGrid.Grid1.Cell(flexcpText)

frmGrid.Grid1.Select MyRow, 0
Rname = frmGrid.Grid1.Cell(flexcpText)
tempPath = MSTSPath & "\Routes\" & Rname & "\Activities\" & tempText
End If

If Check1(0).Value = 1 Then
  
    frmGrid.Grid1.Select MyRow, 1
    frmGrid.Grid1.FillStyle = flexFillSingle
    frmGrid.Grid1.CellBackColor = vbGreen
    If FileExists(tempPath) Then
    Kill tempPath
    
    booKilled = True
    End If
End If
For i = 0 To lstFoundFiles(1).ListCount - 1
frmGrid.Grid1.Select MyRow + i, 4
   If lstFoundFiles(1).Selected(i) Then
   If FileExists(MSTSPath & "\trains\consists\" & lstFoundFiles(1).List(i)) And lstFoundFiles(1).List(i) <> vbNullString Then
   frmGrid.Grid1.FillStyle = flexFillSingle
    frmGrid.Grid1.CellBackColor = vbGreen
   Kill MSTSPath & "\trains\consists\" & lstFoundFiles(1).List(i)
    
    End If
    End If
   Next
   For i = 0 To lstFoundFiles(0).ListCount - 1
frmGrid.Grid1.Select MyRow + i, 3
   If lstFoundFiles(0).Selected(i) Then
   If FileExists(MSTSPath & "\Routes\" & Rname & "\Services\" & lstFoundFiles(0).List(i)) And lstFoundFiles(0).List(i) <> vbNullString Then
   frmGrid.Grid1.FillStyle = flexFillSingle
    frmGrid.Grid1.CellBackColor = vbGreen
    Kill MSTSPath & "\Routes\" & Rname & "\Services\" & lstFoundFiles(0).List(i)
   
    End If
    End If
   Next
   For i = 0 To lstFoundFiles(2).ListCount - 1
frmGrid.Grid1.Select MyRow + i, 5
   If lstFoundFiles(2).Selected(i) Then
   If FileExists(MSTSPath & "\Routes\" & Rname & "\Paths\" & lstFoundFiles(2).List(i)) And lstFoundFiles(2).List(i) <> vbNullString Then
   frmGrid.Grid1.FillStyle = flexFillSingle
    frmGrid.Grid1.CellBackColor = vbGreen
    Kill MSTSPath & "\Routes\" & Rname & "\Paths\" & lstFoundFiles(2).List(i)
    
    End If
    End If
   Next



Rem ********** Traffic
If Check1(4).Value = 1 Then
frmGrid.Grid1.Select MyRow, 2

If Trim$(frmGrid.Grid1.Cell(flexcpText)) <> vbNullString Then
frmGrid.Grid1.FillStyle = flexFillSingle
frmGrid.Grid1.CellBackColor = vbGreen
    If FileExists(MSTSPath & "\Routes\" & Rname & "\Traffic\" & frmGrid.Grid1.Cell(flexcpText)) Then
    Kill MSTSPath & "\Routes\" & Rname & "\Traffic\" & frmGrid.Grid1.Cell(flexcpText)
   
    booKilled = True
    End If
End If
 End If
 If booKilled = True Then
 Label1.Caption = "Requested Files Deleted"
 End If
End Sub

Private Sub Command2_Click()
Unload Me

End Sub




Private Sub Command3_Click()
Dim i As Integer, j As Integer

Check1(0).Value = 1
Check1(4).Value = 1
For j = 0 To 2
For i = 0 To lstFoundFiles(j).ListCount - 1
    lstFoundFiles(j).Selected(i) = True
Next i
Next j

End Sub


Private Sub Command4_Click()
Dim i As Integer, j As Integer

Check1(0).Value = 0
Check1(4).Value = 0
For j = 0 To 2
For i = 0 To lstFoundFiles(j).ListCount - 1
    lstFoundFiles(j).Selected(i) = False
Next i
Next j

End Sub


Private Sub Command5_Click()
Dim MyCol As Integer, MyRow As Integer, tempText As String, tempPath As String
Dim Rname As String, booKilled As Boolean, strSvc As String, strPath As String, strTfc As String
Dim strSvcPath As String, strPathPath As String, strTfcPath As String
On Error GoTo Errtrap

MyRow = frmGrid.Grid1.row
MyCol = frmGrid.Grid1.col
If MyCol = 1 Then
tempText = frmGrid.Grid1.Cell(flexcpText)

frmGrid.Grid1.Select MyRow, 0
Rname = frmGrid.Grid1.Cell(flexcpText)
tempPath = MSTSPath & "\Routes\" & Rname & "\Activities\" & tempText
End If

If Check1(0).Value = 1 Then
  
    frmGrid.Grid1.Select MyRow, 1
    frmGrid.Grid1.FillStyle = flexFillSingle
    frmGrid.Grid1.CellBackColor = vbGreen
    If FileExists(tempPath) Then
    If Not DirExists(MSTSPath & "\Routes\" & Rname & "\Activities\SpareAct\") Then
  MkDir (MSTSPath & "\Routes\" & Rname & "\Activities\SpareAct\")
End If
    FileCopy tempPath, MSTSPath & "\Routes\" & Rname & "\Activities\SpareAct\" & tempText
    DoEvents
    Kill tempPath
    
    booKilled = True
    End If
End If
For i = 0 To lstFoundFiles(1).ListCount - 1
frmGrid.Grid1.Select MyRow + i, 4
   If lstFoundFiles(1).Selected(i) Then
   If FileExists(MSTSPath & "\trains\consists\" & lstFoundFiles(1).List(i)) And lstFoundFiles(1).List(i) <> vbNullString Then
   frmGrid.Grid1.FillStyle = flexFillSingle
    frmGrid.Grid1.CellBackColor = vbGreen
    If Not DirExists(MSTSPath & "\trains\consists\SpareCon") Then
    MkDir (MSTSPath & "\trains\consists\SpareCon")
    End If
    FileCopy MSTSPath & "\trains\consists\" & lstFoundFiles(1).List(i), MSTSPath & "\trains\consists\SpareCon\" & lstFoundFiles(1).List(i)
   DoEvents
   Kill MSTSPath & "\trains\consists\" & lstFoundFiles(1).List(i)
    
    End If
    End If
   Next
   For i = 0 To lstFoundFiles(0).ListCount - 1
frmGrid.Grid1.Select MyRow + i, 3
   If lstFoundFiles(0).Selected(i) Then
   strSvcPath = MSTSPath & "\Routes\" & Rname & "\Services\" & lstFoundFiles(0).List(i)
   strSvc = lstFoundFiles(0).List(i)
   If FileExists(strSvcPath) And strSvc <> vbNullString Then
   frmGrid.Grid1.FillStyle = flexFillSingle
    frmGrid.Grid1.CellBackColor = vbGreen
    If Not DirExists(MSTSPath & "\Routes\" & Rname & "\Services\SpareSvc\") Then
    MkDir (MSTSPath & "\Routes\" & Rname & "\Services\SpareSvc\")
    End If
    FileCopy strSvcPath, MSTSPath & "\Routes\" & Rname & "\Services\SpareSvc\" & strSvc
    DoEvents
    Kill strSvcPath
    End If
    End If
   Next
   For i = 0 To lstFoundFiles(2).ListCount - 1
frmGrid.Grid1.Select MyRow + i, 5
   If lstFoundFiles(2).Selected(i) Then
   strPathPath = MSTSPath & "\Routes\" & Rname & "\paths\" & lstFoundFiles(2).List(i)
   strPath = lstFoundFiles(2).List(i)
   
   If FileExists(strPathPath) And strPath <> vbNullString Then
   frmGrid.Grid1.FillStyle = flexFillSingle
    frmGrid.Grid1.CellBackColor = vbGreen
    If Not DirExists(MSTSPath & "\Routes\" & Rname & "\Paths\SparePat\") Then
  MkDir (MSTSPath & "\Routes\" & Rname & "\Paths\SparePat\")
End If
    FileCopy strPathPath, MSTSPath & "\Routes\" & Rname & "\Paths\SparePat\" & strPath
    DoEvents
    Kill strPathPath
    
    End If
    End If
   Next



Rem ********** Traffic
If Check1(4).Value = 1 Then
frmGrid.Grid1.Select MyRow, 2

If Trim$(frmGrid.Grid1.Cell(flexcpText)) <> vbNullString Then
frmGrid.Grid1.FillStyle = flexFillSingle
frmGrid.Grid1.CellBackColor = vbGreen
strTfc = frmGrid.Grid1.Cell(flexcpText)
strTfcPath = MSTSPath & "\Routes\" & Rname & "\Traffic\" & frmGrid.Grid1.Cell(flexcpText)
    If FileExists(strTfcPath) Then
        If Not DirExists(MSTSPath & "\Routes\" & Rname & "\Traffic\SpareTfc\") Then
  MkDir (MSTSPath & "\Routes\" & Rname & "\Traffic\SpareTfc\")
End If
FileCopy strTfcPath, MSTSPath & "\Routes\" & Rname & "\Traffic\SpareTfc\" & strTfc
    DoEvents

    Kill strTfcPath
   
    booKilled = True
    End If
End If
 End If
 If booKilled = True Then
 Label1.Caption = "Requested Files Moved"
 End If
 Exit Sub
Errtrap:
 
 If Err = 75 Then
 Call MsgBox("One or more of the files you attempted to move/delete was a Read-Only file." _
             & vbCrLf & "Select the route on the MSTS Route Utils screen and click Make Read/Write to fix." _
             , vbExclamation, App.Title)
 
 Exit Sub
 End If
End Sub

Private Sub Form_Load()
Dim MyCol As Integer, MyRow As Integer
Dim i As Integer, MainAct As String
Me.Caption = Lang(220)
MyRow = frmGrid.Grid1.row
MyCol = frmGrid.Grid1.col
Call SetLang
frmGrid.Grid1.Select MyRow, 1
MainAct = frmGrid.Grid1.Cell(flexcpText)
Check1(0).Caption = "Activity - " & frmGrid.Grid1.Cell(flexcpText) ' & "\" & tempText

frmGrid.Grid1.FillStyle = flexFillSingle
frmGrid.Grid1.Select MyRow, 3

'Check1(2).Caption = "Service - " & frmGrid.Grid1.Cell(flexcpText)
lstFoundFiles(0).AddItem frmGrid.Grid1.Cell(flexcpText)
frmGrid.Grid1.Select MyRow, 4
'Check1(1).Caption = "Consist - " & frmGrid.Grid1.Cell(flexcpText)
lstFoundFiles(1).AddItem frmGrid.Grid1.Cell(flexcpText)
frmGrid.Grid1.Select MyRow, 5
'Check1(3).Caption = "Path - " & frmGrid.Grid1.Cell(flexcpText)
lstFoundFiles(2).AddItem frmGrid.Grid1.Cell(flexcpText)
frmGrid.Grid1.Select MyRow, 2
Check1(4).Caption = "Traffic - " & frmGrid.Grid1.Cell(flexcpText)
Do
i = i + 1
frmGrid.Grid1.Select MyRow + i, 1
If frmGrid.Grid1.Cell(flexcpText) = MainAct Then
frmGrid.Grid1.FillStyle = flexFillSingle
frmGrid.Grid1.Select MyRow + i, 3

'Check1(2).Caption = "Service - " & frmGrid.Grid1.Cell(flexcpText)
lstFoundFiles(0).AddItem frmGrid.Grid1.Cell(flexcpText)
frmGrid.Grid1.Select MyRow + i, 4
'Check1(1).Caption = "Consist - " & frmGrid.Grid1.Cell(flexcpText)
lstFoundFiles(1).AddItem frmGrid.Grid1.Cell(flexcpText)
frmGrid.Grid1.Select MyRow + i, 5
'Check1(3).Caption = "Path - " & frmGrid.Grid1.Cell(flexcpText)
lstFoundFiles(2).AddItem frmGrid.Grid1.Cell(flexcpText)
Else
Exit Do
End If
Loop
frmGrid.Grid1.Select MyRow, 1
End Sub


