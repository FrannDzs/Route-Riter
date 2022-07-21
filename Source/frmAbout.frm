VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   5490
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6225
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3789.296
   ScaleMode       =   0  'User
   ScaleWidth      =   5845.598
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   240
      Picture         =   "frmAbout.frx":406A
      ScaleHeight     =   758.52
      ScaleMode       =   0  'User
      ScaleWidth      =   758.52
      TabIndex        =   0
      Top             =   240
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5160
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404000&
      Index           =   2
      X1              =   0
      X2              =   5662.483
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Latest version at - http://www.rstools.info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   5535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "For Support: Contact me via the  link on my web site"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Penrith, NSW, Australia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "by Mike Simpson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404000&
      Index           =   1
      X1              =   0
      X2              =   5746.998
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A suite of utilities for Microsoft Train Simulator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   4680
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   720
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   4365
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":80D4
      ForeColor       =   &H000000FF&
      Height          =   1995
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   4815
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public registered As Boolean




Private Sub cmdOK_Click()

frmUtils.Show
Unload Me
  
End Sub



Private Sub Form_Load()


 
lblDisclaimer.Caption = "This program is Copyright 2002-11, T.M. Simpson. "
lblDisclaimer.Caption = lblDisclaimer.Caption & "Route_Riter is being released as freeware and may be "
lblDisclaimer.Caption = lblDisclaimer.Caption & "distributed to anyone interested in MSTS modelling.. "
lblDisclaimer.Caption = lblDisclaimer.Caption & "Please advise the author of any bugs or "
lblDisclaimer.Caption = lblDisclaimer.Caption & "problems observed, suggestions for enhancement are welcomed."
lblDisclaimer.Caption = lblDisclaimer.Caption & vbCrLf & vbCrLf



10        Me.Caption = "About " & App.Title
20        lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
30        lblTitle.Caption = App.Title & " for XP, Vista && Windows7"
40        If Command <> vbNullString Then cmdOK.value = True






End Sub


