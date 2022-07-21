VERSION 5.00
Begin VB.Form frmBrake 
   Caption         =   "Select Brake Type"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Change To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton Option1 
         Caption         =   "Vacuum_piped"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   11
         Top             =   3720
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Air_piped"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "EP"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ECP"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Vacuum_twin_pipe"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Vacuum_single_pipe"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Air_twin_pipe"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Air_single_pipe"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Change the Brake Type"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "frmBrake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
For i = 0 To 7
If Option1(i).Value = True Then
frmStock.Label1.Caption = Option1(i).Caption
Exit For
End If
Next i
Unload Me
End Sub


Private Sub Command2_Click()
frmStock.Label1.Caption = vbNullString
Unload Me
End Sub


Private Sub Form_Load()
Me.Caption = Lang(200)
Label1.Caption = Lang(201)
Command1.Caption = Lang(202)
Command2.Caption = Lang(203)
Frame1.Caption = Lang(204)

End Sub


