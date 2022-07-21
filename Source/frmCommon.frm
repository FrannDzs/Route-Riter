VERSION 5.00
Begin VB.Form frmCommon 
   Caption         =   "Select Common Path"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   5055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   1560
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Installer"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Common Path"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "frmCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text


Dim strInstaller As String
Private Sub Command1_Click()
On Error GoTo Errtrap
Dim x As Integer


strComPath = strGetShortFileName(strComPath)

 NewFile = FreeFile
 Open strInstaller For Input As #NewFile
  NewFile2 = FreeFile
  Open Text2 For Append As #NewFile2
   Do While Not EOF(NewFile)
     Line Input #NewFile, A$
     
     x = InStr(A$, "..\Europe1\")
   
     If x > 0 Then
     
       strStart = Left$(A$, x - 1)
       strEnd = Mid$(A$, x + 10)
       A$ = strStart & strComPath & strEnd
       Print #NewFile2, A$
       GoTo GetAnother
     End If
     x = InStr(A$, "..\Europe2\")
     If x > 0 Then
       strStart = Left$(A$, x - 1)
       strEnd = Mid$(A$, x + 10)
       A$ = strStart & strComPath & strEnd
       Print #NewFile2, A$
       GoTo GetAnother
     End If
     x = InStr(A$, "..\Japan1\")
     If x > 0 Then
       strStart = Left$(A$, x - 1)
       strEnd = Mid$(A$, x + 9)
       A$ = strStart & strComPath & strEnd
       Print #NewFile2, A$
       GoTo GetAnother
     End If
     x = InStr(A$, "..\Japan2\")
     If x > 0 Then
       strStart = Left$(A$, x - 1)
       strEnd = Mid$(A$, x + 9)
       A$ = strStart & strComPath & strEnd
       Print #NewFile2, A$
       GoTo GetAnother
     End If
     x = InStr(A$, "..\USA1\")
     If x > 0 Then
       strStart = Left$(A$, x - 1)
       strEnd = Mid$(A$, x + 7)
       A$ = strStart & strComPath & strEnd
       Print #NewFile2, A$
       GoTo GetAnother
     End If
     x = InStr(A$, "..\USA2\")
     If x > 0 Then
       strStart = Left$(A$, x - 1)
       strEnd = Mid$(A$, x + 7)
       A$ = strStart & strComPath & strEnd
       Print #NewFile2, A$
     End If
GetAnother:
If x = 0 Then
Print #NewFile2, A$

End If

x = 0
     DoEvents
   Loop
  Close NewFile2
 Close NewFile
 
Exit Sub
Errtrap:
Call MsgBox("An error occurred in subroutine 'Common Path' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)

'Resume Next
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Command3_Click(Index As Integer)
If Index = 0 Then
If booComDir = True Then
strComPath = Text1(0).Text
Unload Me
booComDir = False
Else
Command3(1).Visible = True
Text1(1).Visible = True
Text2.Visible = True
Label3.Visible = True
strComPath = Text1(0).Text
End If
End If
If Index = 1 Then
strInstaller = Text1(1).Text

Command1.Visible = True
End If
End Sub

Private Sub Form_GotFocus()
frmCommon.ZOrder
Me.Caption = Lang(205)
Label2.Caption = Lang(206)
Label3.Caption = Lang(207)
Command1.Caption = Lang(208)
Command2.Caption = Lang(203)
End Sub

Private Sub Form_Load()
Dim strLabel As String
Me.Top = 100
Me.Left = 100

If booComDir = True Then

strLabel = Lang(534)
strLabel = strLabel & Lang(535)
strLabel = strLabel & Lang(536)
Else
strLabel = Lang(535)
strLabel = strLabel & Lang(536)
strLabel = strLabel & Lang(537)
strLabel = strLabel & Lang(538)
End If
Label1.Caption = strLabel
frmCommon.ZOrder

Command1.Visible = False
Label3.Visible = False
Text1(1).Visible = False
Command3(1).Visible = False
Text2.Visible = False
End Sub


