VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form frmINITest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INI Loader Test Application"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load INI File..."
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   3720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "INI Files|*.ini|"
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5520
      Width           =   3975
   End
   Begin VB.ListBox lstKeys 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3975
   End
   Begin VB.ComboBox cboSections 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lblOpenINI 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblINI 
      Caption         =   "Open INI file:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblValue 
      Caption         =   "Value"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblKeys 
      Caption         =   "Keys:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblSections 
      Caption         =   "Sections:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmINITest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurINI As String

Private Sub cboSections_Click()
    LoadSection
End Sub

Private Sub cmdLoad_Click()
    CDialog.FileName = CurINI
    CDialog.ShowOpen
    CurINI = CDialog.FileName
    LoadINI
End Sub

Private Sub Form_Load()
    CurINI = "Win.ini"
    CDialog.InitDir = GetWindowsDir
    LoadINI
End Sub

Sub LoadINI()
Dim Sections() As String, i As Integer
    On Error Resume Next
    lblOpenINI = CurINI
    GetPrivateProfileSections Sections(), CurINI
    cboSections.Clear
    For i = 1 To StrArrCnt(Sections())
        cboSections.AddItem Sections(i)
    Next i
    cboSections.ListIndex = 0
End Sub

Function StrArrCnt(StrArr() As String) As Long
    On Error Resume Next
    StrArrCnt = UBound(StrArr)
End Function

Sub LoadSection()
Dim Keys() As String, i As Integer
    On Error Resume Next
    GetPrivateProfileKeys Keys(), cboSections.Text, CurINI
    lstKeys.Clear
    For i = 1 To StrArrCnt(Keys())
        lstKeys.AddItem Keys(i)
    Next i
    lstKeys.ListIndex = 0
End Sub

Private Sub lstKeys_Click()
    txtValue = GetPrivateProfileString(cboSections.Text, lstKeys.list(lstKeys.ListIndex), "(Error)", CurINI)
End Sub

