VERSION 5.00
Begin VB.Form frmPos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Position"
   ClientHeight    =   2055
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   5415
   Icon            =   "frmPos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFlip 
      Caption         =   "Flip"
      Height          =   375
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   780
      Width           =   1035
   End
   Begin VB.HScrollBar posX 
      Height          =   315
      LargeChange     =   100
      Left            =   1320
      Max             =   3000
      Min             =   -3000
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
   End
   Begin VB.HScrollBar posZ 
      Height          =   315
      LargeChange     =   100
      Left            =   1320
      Max             =   3000
      Min             =   -3000
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.VScrollBar posY 
      Height          =   1215
      LargeChange     =   100
      Left            =   1320
      Max             =   -3000
      Min             =   3000
      TabIndex        =   2
      Top             =   60
      Value           =   -3000
      Width           =   315
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Reset"
      Height          =   375
      Left            =   4140
      TabIndex        =   1
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4140
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label labx 
      Alignment       =   2  'Center
      Caption         =   "0.00m"
      Height          =   195
      Left            =   4440
      TabIndex        =   10
      Top             =   1740
      Width           =   915
   End
   Begin VB.Label labz 
      Alignment       =   2  'Center
      Caption         =   "0.00m"
      Height          =   195
      Left            =   4440
      TabIndex        =   9
      Top             =   1380
      Width           =   915
   End
   Begin VB.Label laby 
      Alignment       =   2  'Center
      Caption         =   "0.00m"
      Height          =   195
      Left            =   1740
      TabIndex        =   8
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Left/Right"
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Front/Rear"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Vertical"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   240
      Width           =   1395
   End
End
Attribute VB_Name = "frmPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    posX.Value = 0
    posY.Value = 0
    posZ.Value = 0
End Sub

Private Sub chkFlip_Click()
    gflip = (chkFlip = vbChecked)
End Sub

Private Sub Form_Load()
    posX.Value = gmfv.x * 100
    posY.Value = gmfv.y * 100
    posZ.Value = gmfv.z * 100
    gflip = (chkFlip = vbChecked)
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub

Private Sub posX_Change()
    gmfv.x = posX.Value / 100
    labx.Caption = gmfv.x & "m"
End Sub

Private Sub posY_Change()
    gmfv.y = posY.Value / 100
    laby.Caption = gmfv.y & "m"
End Sub

Private Sub posZ_Change()
    gmfv.z = posZ.Value / 100
    labz.Caption = gmfv.z & "m"
End Sub
