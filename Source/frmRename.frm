VERSION 5.00
Begin VB.Form frmRename 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rename Multiple Files"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmRename.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRename.frx":406A
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
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
      Left            =   1800
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "or"
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   10
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "With"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Replace"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Add Prefix"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "First filename is:-"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "You can either add a prefix to the selected file names, or change the suffix."
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
strReplace = vbNullString
strWith = vbNullString
strPrefix = vbNullString
Unload Me
End Sub

Private Sub Form_Load()
Dim x As Integer, strFileName As String, strPathName As String
Dim strFile As String

Me.Caption = Lang(276)
Label2.Caption = Lang(277)
Label3(0).Caption = Lang(278)
Label3(3).Caption = Lang(279)
Label3(1).Caption = Lang(280)
Label3(2).Caption = Lang(281)
Label1.Caption = Lang(373)

strFile = frmUtils.Label2(0).Caption

x = InStrRev(strFile, "\")
    strFileName = Mid$(strFile, x + 1)
    strPathName = Left$(strFile, x)
Label2.Caption = Lang(277) & strFileName


End Sub

Private Sub OKButton_Click()
If Text1 <> vbNullString And Text2 = vbNullString And Text3 = vbNullString Then
strPrefix = Trim$(Text1)
ElseIf Text1 <> vbNullString And Text2 <> vbNullString Then
Call MsgBox("You can either add a prefix, or replace the suffix," _
            & vbCrLf & "Not Both." _
            , vbExclamation, "Rename error")
            Exit Sub
ElseIf Text1 <> vbNullString And Text3 <> vbNullString Then
Call MsgBox("You can either add a prefix, or replace the suffix," _
            & vbCrLf & "Not Both." _
            , vbExclamation, "Rename error")
            Exit Sub
ElseIf Text1 = vbNullString And Text2 <> vbNullString And Text3 <> vbNullString Then
strReplace = Text2
strWith = Text3
End If
Unload Me
End Sub


