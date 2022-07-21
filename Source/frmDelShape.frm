VERSION 5.00
Begin VB.Form frmDelShape 
   Caption         =   "Delete Shape File"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Shape File Name   e.g. tree1.s"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmDelShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
strDelShape = Text1
If Right$(strDelShape, 2) <> ".s" And Right$(strDelShape, 4) <> ".haz" And Right$(strDelShape, 4) <> ".ace" Then
strDelShape = strDelShape & ".s"

End If
Unload Me

End Sub
Private Sub Command2_Click()
strDelShape = vbNullString
Unload Me

End Sub


Private Sub Form_Load()
Me.Caption = Lang(221)
Label1.Caption = Lang(222)
Command1.Caption = Lang(218)
Command2.Caption = Lang(203)
If booGetW = True Then
Command1.Caption = Lang(539)
frmDelShape.Caption = Lang(540)
End If
End Sub



Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
strDelShape = Text1
If Right$(strDelShape, 2) <> ".s" Then
strDelShape = strDelShape & ".s"
End If
Unload Me
End If
End Sub


