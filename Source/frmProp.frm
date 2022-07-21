VERSION 5.00
Begin VB.Form frmProp 
   BackColor       =   &H00C0C0C0&
   Caption         =   "File Properties"
   ClientHeight    =   2745
   ClientLeft      =   4545
   ClientTop       =   2670
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2745
   ScaleWidth      =   4905
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Reset Attributes"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Archive"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "System"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hidden"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Read Only"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   7
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   6
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Attributes:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Date/Time:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Length:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Path:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim prop%
prop% = 0
If frmProp.Check1(0).value = 1 Then
prop = prop + vbReadOnly
End If
If frmProp.Check1(1).value = 1 Then
prop = prop + vbHidden
End If
If frmProp.Check1(2).value = 1 Then
prop = prop + vbSystem
End If
If frmProp.Check1(3).value = 1 Then
prop = prop + vbArchive
End If
SetAttr fullpath$, prop
Unload frmProp
End Sub


Private Sub Command2_Click()
Unload Me

End Sub


Private Sub Form_Load()
Dim i As Integer
Me.Caption = Lang(266)
For i = 0 To 3
Label1(i).Caption = Lang(267 + i)
Next i
Command1.Caption = Lang(271)
For i = 0 To 3
Check1(i).Caption = Lang(272 + i)
Next i
Command2.Caption = Lang(38)
End Sub


