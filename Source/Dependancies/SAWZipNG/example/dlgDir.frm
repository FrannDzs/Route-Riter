VERSION 5.00
Begin VB.Form dlgDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add directory"
   ClientHeight    =   2040
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFullpath 
      Caption         =   "Store full path"
      Height          =   225
      Left            =   2340
      TabIndex        =   8
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1223
      TabIndex        =   7
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2655
      TabIndex        =   6
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CheckBox chkSubDir 
      Caption         =   "Include Subdirectories"
      Height          =   255
      Left            =   2340
      TabIndex        =   5
      Top             =   660
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelect 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   345
      Left            =   4290
      TabIndex        =   4
      Top             =   180
      Width           =   615
   End
   Begin VB.TextBox txtDir 
      Height          =   345
      Left            =   1020
      TabIndex        =   3
      Top             =   180
      Width           =   3255
   End
   Begin VB.TextBox txtWildCard 
      Height          =   315
      Left            =   1005
      TabIndex        =   2
      Text            =   "*.*"
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Wildcard"
      Height          =   195
      Left            =   45
      TabIndex        =   1
      Top             =   690
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Directory"
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   210
      Width           =   630
   End
End
Attribute VB_Name = "dlgDir"
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

Private Sub cmdSelect_Click()
    
    txtDir.Text = SelectDir(Me.hwnd, "Select a directory to add to the archive")

End Sub

Private Sub OKButton_Click()
    
    ok = True
    Hide

End Sub

Private Sub txtDir_Change()

    OKButton.Enabled = Len(Trim(txtDir.Text)) > 0

End Sub
