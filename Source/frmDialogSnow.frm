VERSION 5.00
Begin VB.Form frmDialogSnow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Missing from Route"
   ClientHeight    =   2475
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "No for All"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Yes for &All"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&No"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Yes"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "frmDialogSnow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
intResponseSnow = 3
Unload Me

End Sub

Private Sub Command1_Click()
intResponseSnow = 4
Unload Me
End Sub


Private Sub Form_Load()

Label1.Caption = strResponse
Me.Caption = Lang(227)
If booTerrtexSnow = False Then
OKButton(0).Caption = Lang(224)
OKButton(1).Caption = Lang(225)
ElseIf booTerrtexSnow = True Then
OKButton(0).Caption = "Summer"
OKButton(1).Caption = "Substitute Snow"
End If

CancelButton.Caption = Lang(226)
End Sub


Private Sub OKButton_Click(Index As Integer)
If Index = 0 Then
intResponseSnow = 1
ElseIf Index = 1 Then
intResponseSnow = 2
End If
Unload Me

End Sub


