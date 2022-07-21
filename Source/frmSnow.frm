VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSnow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Ground Snow Texture."
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "....."
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1680
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CDSnow 
      Left            =   5280
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   4695
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmSnow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
strSnowName = vbNullString
booSnow = False
Unload Me
End Sub

Private Sub Command1_Click()

CDSnow.Filter = "Terrtex Files (*.ace)|*.ace"
CDSnow.DialogTitle = "Select ACE File"
CDSnow.FilterIndex = 1

CDSnow.FileName = MSTSPath & "\Routes\usa2\terrtex\snow\us2targnd.ace"

CDSnow.Action = 1
DoEvents
Text1.Text = CDSnow.FileName



End Sub

Private Sub Form_Load()
Dim strsnow As String
Me.Caption = Lang(304)

strsnow = Lang(374)
strsnow = strsnow & Lang(375)
strsnow = strsnow & Lang(376)
strsnow = strsnow & vbCrLf & Lang(377)
Label1.Caption = strsnow

End Sub


Private Sub OKButton_Click()

If CDSnow.FileName <> vbNullString Then
strSnowName = CDSnow.FileName
booSnow = True
Else
strSnowName = vbNullString
booSnow = False
End If
Unload Me
End Sub


