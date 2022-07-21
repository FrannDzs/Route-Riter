VERSION 5.00
Begin VB.Form frmAI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Wheel Radius"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Select Locomotive Type"
      Height          =   3135
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4215
      Begin VB.OptionButton Option1 
         Caption         =   "This is a MU (Multiple Unit) Locomotive i.e the second or third Loco in a multiple-headed train."
         Height          =   615
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   2040
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         Caption         =   $"frmAI.frx":0000
         Height          =   975
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "This is a basic AI Locomotive"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   3480
      Width           =   855
   End
End
Attribute VB_Name = "frmAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1(0).value = True Then
booAI = False
booMU = False
ElseIf Option1(1).value = True Then
booAI = True
booMU = False
ElseIf Option1(2).value = True Then
booAI = False
booMU = True
End If
Unload Me
End Sub


