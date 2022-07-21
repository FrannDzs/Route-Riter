VERSION 5.00
Begin VB.Form dlgComment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set comment"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "dlgComment"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtComment 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      Caption         =   "Enter some comment for ..."
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1890
   End
   Begin VB.Label lblFile 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "dlgComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public ok As Boolean

Private Sub CancelButton_Click()
    
    ok = False
    Hide

End Sub

Private Sub OKButton_Click()

    ok = True
    Hide

End Sub
