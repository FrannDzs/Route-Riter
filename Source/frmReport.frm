VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form frmReport 
   Caption         =   "Report Form"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "C"
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   8
      ToolTipText     =   "Enter 'CabView' in Search box."
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "S"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   7
      ToolTipText     =   "Enter 'Sound' in Search box."
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Alias Selected Text"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Search"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   6480
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9240
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "End"
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   6480
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   6480
      Width           =   735
   End
   Begin RichTextLib.RichTextBox Rich1 
      Height          =   6135
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10821
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"frmReport.frx":406A
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Dim Search As Long
Dim flagAlias As Boolean
Dim TrainsetPath As String



Private Sub SetLang()
Command1.Caption = Lang(367)
Command2.Caption = Lang(626)
Command3.Caption = Lang(38)

End Sub

Private Sub Command1_Click()
objList = Rich1.Text
selFlag = 5
'booListAce = True
fEZPrint.Show 1


End Sub


Private Sub Command2_Click()
Dim tit1 As String


If Not DirExists(App.Path & "\Reports") Then
MkDir App.Path & "\Reports"
End If

If booUniEdit = True Then
booUniEdit = False
Rich1.SaveFile strUniName, rtfText
DoEvents
Unload Me
Else
CommonDialog1.InitDir = App.Path & "\Reports"
CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
CommonDialog1.DialogTitle = "Save Object File"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 2
tit1 = CommonDialog1.Filename
Rich1.SaveFile tit1, rtfText
End If

End Sub

Private Sub Command3_Click()


Unload Me

End Sub


Private Sub Command4_Click()
Dim strSearch As String, x As Long
If Text1 <> vbNullString Then
strSearch = Text1
x = Search + 1
Search = Rich1.Find(strSearch, x)
End If
End Sub

Private Sub Command5_Click()
Dim strSnd As String, strPath As String
Dim strAlias As String


flagAlias = True
booUniEdit = True
If Rich1.SelLength = 0 Then
Call MsgBox("You do not appear to have selected a Sound or Cab statement.", vbExclamation, App.Title)

End If
strSnd = Rich1.SelText
If Left$(strSnd, 1) = ChrW$(34) And Right$(strSnd, 1) = ChrW$(34) Then
strSnd = ChrW$(34)
Else
strSnd = vbNullString
End If

CommonDialog1.InitDir = TrainsetPath
CommonDialog1.Filter = "Sound/Cab Files (*.sms;*.cvf)|*.sms;*.cvf"
CommonDialog1.DialogTitle = "Alias File"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 1
strPath = CommonDialog1.Filename
If strPath = vbNullString Then Exit Sub
 
   
   zz = InStr(strPath, "Trainset")
   If zz > 0 Then
   strAlias = Mid$(strPath, zz + 9)
   strAlias = Replace(strAlias, "\", "/")
   strAlias = strSnd & "../../" & strAlias & strSnd
   Else
   'zz = InStr(strPath, "Train Simulator")
   zz = InStr(strPath, MSTSPath)
   strAlias = Mid$(strPath, zz + Len(MSTSPath) + 1)
   strAlias = Replace(strAlias, "\", "/")
   strAlias = strSnd & "../../../../" & strAlias & strSnd
   End If
   
Rich1.SelText = strAlias

End Sub

Private Sub Command6_Click(Index As Integer)
If Index = 0 Then
Text1 = "Sound"
End If
If Index = 1 Then
Text1 = "CabView"
End If
Command4.Value = True
End Sub

Private Sub Form_Activate()
frmReport.ZOrder

End Sub

Private Sub Form_Load()
Call SetLang
Search = 0
Me.Caption = Lang(286)
TrainsetPath = MSTSPath & "\Trains\Trainset\"
End Sub


Private Sub Form_Resize()
Const Margin As Single = 150 'Twips (10 pixels)
On Error GoTo ErrTrap
Rich1.Move Margin, Margin, ScaleWidth - 2 * Margin, ScaleHeight - 3 * Margin - Command1.height
Text1.Top = 2 * Margin + Rich1.height
Command4.Move Text1.Left + Text1.width + 200, 2 * Margin + Rich1.height
Command5.Top = Command4.Top
Command5.Left = Command4.Left + Command4.width + 50
Command6(0).Top = Command4.Top
Command6(0).Left = Command5.Left + Command5.width + 50
Command6(1).Top = Command4.Top
Command6(1).Left = Command6(0).Left + Command6(0).width + 50
Command1.Top = Command4.Top
Command1.Left = Command6(1).Left + Command6(1).width + 50

Command2.Top = Command4.Top
Command2.Left = Command1.Left + Command1.width + 50
Command3.Top = Command4.Top
Command3.Left = Command2.Left + Command2.width + 50

Exit Sub
ErrTrap:
If Err = 380 Then Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
booReport = True
frmReport.Rich1.Text = vbNullString
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Text1 <> vbNullString And KeyCode = 13 Then
Command4.Value = True
End If
End Sub

