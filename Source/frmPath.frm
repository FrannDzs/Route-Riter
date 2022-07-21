VERSION 5.00
Begin VB.Form frmPath 
   Caption         =   "Select MSTS Path"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MSTSPath = frmUtils.Dir1(cursouind).Path
If Not FileExists(MSTSPath & "\train.exe") Then
'Call MsgBox("The selected folder does not contain 'Train.exe' so can not be " _
'            & vbCrLf & "a main MSTS folder." _
'            , vbCritical, App.Title)
'Unload Me
'booWrongMSTS = True
'Exit Sub
Select Case MsgBox("The selected folder does not contain 'Train.exe' so is not" _
                   & vbCrLf & "a main MSTS folder. Do you wish to continue" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

    Case vbNo
Unload Me
booWrongMSTS = True
End Select


End If
frmUtils.Caption = "Path=" & MSTSPath
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim strLabel As String
Me.Caption = Lang(264)
strLabel = Lang(368)
strLabel = strLabel & Lang(369)
strLabel = strLabel & Lang(370)
strLabel = strLabel & Lang(371)
strLabel = strLabel & Lang(372)
Label1.Caption = strLabel
Command1.Caption = Lang(208)
Command2.Caption = Lang(203)
End Sub


