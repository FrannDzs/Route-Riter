VERSION 5.00
Begin VB.Form frmDialog4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change MSTS Registry Path"
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   5535
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmDialog4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
booRaildriver = False
Unload Me
End Sub

Private Sub Form_Load()
Dim strCap As String

strCap = "Warning: This option changes some registry settings, if you are not happy with this, then please click CANCEL" & vbCrLf & vbCrLf
strCap = strCap & "Otherwise, select the instance of MSTS you wish to use with Raildriver, then click OK"
Label1.Caption = strCap
booRaildriver = True


End Sub


Private Sub OKButton_Click()
Dim strNew As String

strNew = Text1.Text & Chr$(0)

Call SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "Path", strNew, REG_SZ)
DoEvents

Call SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "EXE Path", strNew, REG_SZ)

Unload Me
End Sub


Private Sub Text1_Change()
OKButton.Visible = True

End Sub


