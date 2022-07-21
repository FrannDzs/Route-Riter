VERSION 5.00
Begin VB.Form frmTrainset 
   Caption         =   "Mini-Route Setup"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Drive and Path for your master Trains folder."
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmTrainset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Errtrap
frmUtils.Drive1(1).Drive = Drive1.Drive
frmUtils.Dir1(1).path = Dir1.path
DoEvents
Call frmUtils.MiniTrainsCopy
Unload Me
Exit Sub
Errtrap:
Call MsgBox("Error " & Err.Description & " occurred in frmTrainset", vbExclamation, App.Title)

End Sub




Private Sub Dir1_Change()
If Dir1.path <> Dir1.list(Dir1.ListIndex) Then
Dir1.path = Dir1.list(Dir1.ListIndex)
If Right$(Dir1.path, 1) <> "\" Then Dir1.path = Dir1.path & "\"
End If
Text1 = Dir1.path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
Text1 = Dir1.path
End Sub





Private Sub Form_Load()
Dim result As String, strMainPath As String

result = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "Path")
strMainPath = result
Drive1.Drive = Left$(strMainPath, 1)
Dir1.path = strMainPath & "\Trains"
Show

End Sub


