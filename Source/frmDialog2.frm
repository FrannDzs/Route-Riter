VERSION 5.00
Begin VB.Form frmDialog2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmDialog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
intResponse2 = 5
Unload Me
End Sub

Private Sub OKButton_Click(Index As Integer)
If Index = 0 Then
intResponse2 = 1       ' 0
ElseIf Index = 1 Then
intResponse2 = 3        '257
ElseIf Index = 2 Then
intResponse2 = 2       ' All 0
ElseIf Index = 3 Then
intResponse2 = 4       ' All 257
End If

Unload Me
End Sub


