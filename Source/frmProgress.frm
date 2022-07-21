VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Working..."
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3720
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ProgressBar pbrProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblWorking 
      Alignment       =   2  'Center
      Caption         =   "Working..."
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Caption = Lang(265)
lblWorking.Caption = Lang(265)

End Sub

'==============================================================================
'Richsoft Computing 2001
'Richard Southey
'This code is e-mailware, if you use it please e-mail me and tell me about
'your program.
'
'For latest information about this and other projects please visit my website:
'www.richsoftcomputing.btinternet.co.uk
'
'If you would like to make any comments/suggestions then please e-mail them to
'richsoftcomputing@btinternet.co.uk
'==============================================================================
Private Sub lblWorking_Click()

End Sub


