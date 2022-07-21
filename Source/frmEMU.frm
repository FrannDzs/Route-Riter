VERSION 5.00
Begin VB.Form frmEMU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carriage Types"
   ClientHeight    =   2685
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Select Carriage Type"
      Height          =   1575
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   4935
      Begin VB.OptionButton Option1 
         Caption         =   "EMU unpowered carriage"
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   3975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "DMU unpowered carriage"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   3975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Loco Hauled Passenger Car"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   3975
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmEMU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
If Option1(0).Value = True Then
booEMU = False
booDMU = False
ElseIf Option1(1).Value = True Then
booEMU = False
booDMU = True
Else
booEMU = True
booDMU = False
End If
Unload Me
End Sub


