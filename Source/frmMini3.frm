VERSION 5.00
Begin VB.Form frmMini3 
   Caption         =   "Mini-Route Setup"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Create new Edit Folder"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add new folder if required"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Drive and Path for new Edit folder"
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Folder name for new Edit Folder"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMini3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Dim TrainsetPath As String
Private Sub Command1_Click()



If Not DirExists(Text1) Then
MkDir Text1
DoEvents
Else
Call MsgBox(Text1 & " already exists, you must create a new folder for your Edit Folder.", vbCritical, App.Title)
Exit Sub
End If

strEditPath = Text1
DoEvents
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
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


TrainsetPath = MSTSPath & "\Trains\Trainset\"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command1.value = True
End If
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text1.Text = Text1.Text & Text2
MkDir Dir1.path & Text2
End If
End Sub


