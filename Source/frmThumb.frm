VERSION 5.00
Begin VB.Form frmThumb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".jpg Exists"
   ClientHeight    =   2025
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "No to All"
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "No"
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK To All"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmThumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()

End Sub

Private Sub Form_Load()
Label1.Caption = frmSearch.Label6.Caption & " Already exists, Do you wish to replace it?"
End Sub





Private Sub OKButton_Click(Index As Integer)

flagThumb = Index
Unload Me
End Sub


