VERSION 5.00
Begin VB.Form dlgBrakes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Brake Types"
   ClientHeight    =   2775
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Load Sensing Device"
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   3375
      Begin VB.OptionButton Option2 
         Caption         =   "LSD Fitted"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         Caption         =   "LSD Not Fitted"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Brakes On This Stock Are:-"
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3375
      Begin VB.OptionButton Option1 
         Caption         =   "Cast Iron"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Composition"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   3015
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "dlgBrakes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub Form_Load()
If booProBrakes = True Then
Frame2.Visible = True
Else
Frame2.Visible = False
End If
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).Value = True Then
Option1(1).Value = False
booIron = False
Frame2.Visible = True
ElseIf Option1(1).Value = True Then
Option1(0).Value = False
booIron = True
Frame2.Visible = False
booLSD = False
End If
End Sub


Private Sub Option2_Click(Index As Integer)
If Option2(0).Value = True Then
Option2(1).Value = False
booLSD = False
ElseIf Option2(1).Value = True Then
Option2(0).Value = False
booLSD = True
End If
End Sub


