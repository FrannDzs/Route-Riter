VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTsect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Master tsection.dat"
   ClientHeight    =   2235
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   480
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   5175
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmTsect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Command1_Click()
CDL1.Filter = "Tsection Files (*.dat)|*.dat"
CDL1.DialogTitle = "Select Master Tsection File"
CDL1.FilterIndex = 1

CDL1.InitDir = MSTSPath & "\Global"

CDL1.Action = 1
DoEvents
Text1.Text = CDL1.FileName


End Sub


Private Sub OKButton_Click()
If Text1 <> vbNullString Then
strTPath = Text1
Unload Me
End If
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
strTPath = Text1
Unload Me
End If
End Sub


