VERSION 5.00
Begin VB.Form TsUtil_CD 
   Caption         =   "Select Folder"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3240
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "TsUtil_CD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If TSUFlag = 1 Then
frmTsUtil.Label2.Caption = Text1 & "\Routes"

For i = frmTsUtil.List2.ListCount - 1 To 0 Step -1
frmTsUtil.List2.RemoveItem (i)
Next i
ElseIf TSUFlag = 2 Then
frmProcess.Textbox4.Text = Text1.Text
ElseIf TSUFlag = 3 Then
frmProcess.Textbox6.Text = Text1.Text
ElseIf TSUFlag = 4 Then
frmProcess.Textbox7.Text = Text1.Text
End If
DoEvents
Unload Me
End Sub

Private Sub Command2_Click()
Text1 = ""
frmTsUtil.Label2.Caption = Text1
Unload Me
End Sub


Private Sub Dir1_Change()
If Dir1.path <> Dir1.List(Dir1.ListIndex) Then
Dir1.path = Dir1.List(Dir1.ListIndex)
If Right$(Dir1.path, 1) <> "\" Then Dir1.path = Dir1.path & "\"
End If
Text1 = Dir1.path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
Text1 = Dir1.path
End Sub


