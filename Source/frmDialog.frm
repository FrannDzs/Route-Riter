VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Missing from Route"
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "No &for All"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Yes for &All"
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&No"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Yes"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
intResponse = 3
Unload Me

End Sub

Private Sub Command1_Click()
intResponse = 4
Unload Me
End Sub



Private Sub Form_Load()

If booExact = False Then
Label1.Caption = strResponse
DoEvents
Me.Caption = Lang(223)
OKButton(0).Caption = Lang(224)
OKButton(1).Caption = Lang(225)
CancelButton.Caption = Lang(226)
Else
Label1.Caption = strResponse
DoEvents
Me.Caption = "Route_Riter"
OKButton(0).Caption = Lang(411)
OKButton(1).Caption = Lang(419)
CancelButton.Visible = True
CancelButton.Caption = "Keep All"
Command1.Visible = True
Command1.Caption = "Replace All"
End If
'If intResponse = 2 Then
'Unload Me
'End If
'If intResponse = 4 Then
'Unload Me
'End If
End Sub


Private Sub OKButton_Click(Index As Integer)
If Index = 0 Then
intResponse = 1
ElseIf Index = 1 Then
intResponse = 2
End If
Unload Me


End Sub


