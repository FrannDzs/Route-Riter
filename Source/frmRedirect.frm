VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRedirect 
   Caption         =   "Redirect Test Form."
   ClientHeight    =   10110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin RichTextLib.RichTextBox RTBOutput 
      Height          =   8295
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   14631
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmRedirect.frx":0000
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   945
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   9855
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15901
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "&Write"
      Height          =   375
      Left            =   3660
      TabIndex        =   6
      Top             =   600
      Width           =   1155
   End
   Begin VB.TextBox txtCommand 
      Height          =   315
      Left            =   1140
      TabIndex        =   5
      Top             =   630
      Width           =   2415
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3660
      TabIndex        =   2
      Top             =   120
      Width           =   1155
   End
   Begin VB.TextBox txtApplication 
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   150
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Command"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   690
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Output"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1110
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Application"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   825
   End
End
Attribute VB_Name = "frmRedirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oLaunch As RedirectLib.Application
Attribute oLaunch.VB_VarHelpID = -1
Dim bExecute As Boolean

Private Sub cmdClear_Click()
    RTBOutput.Text = ""
    txtCommand = ""
    
End Sub

Private Sub cmdExecute_Click()
    
    If bExecute Then
        oLaunch.Stop
    Else
        oLaunch.Name = txtApplication.Text
        Select Case oLaunch.Start
            Case laAlreadyRunning
                sbStatus.Panels.Item(1).Text = "Already running !"
            Case laWindowsError
                sbStatus.Panels.Item(1).Text = "Windows error: " & CStr(oLaunch.LastErrorNumber) & "!"
            Case laOk
                bExecute = True
                txtApplication.Enabled = False
                txtCommand.Enabled = True
                cmdExecute.Caption = "&Stop"
                sbStatus.Panels.Item(1).Text = oLaunch.Name & " is running ..."
        End Select
    End If
    
End Sub

Private Sub cmdWrite_Click()

    oLaunch.Write txtCommand.Text + vbCrLf

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()

    Set oLaunch = New RedirectLib.Application
    oLaunch.BufferSize = 8192
    oLaunch.Wait = 1000
    bExecute = False
  
txtApplication.Text = "c:\windows\system32\cmd.exe"
DoEvents
cmdExecute.value = True
DoEvents
txtCommand.Text = strRedirect
DoEvents
Call Sleep(100)
cmdWrite.value = True
DoEvents

End Sub

Private Sub Form_Resize()
RTBOutput.width = Me.width - 300
RTBOutput.height = Me.height - 2815
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set oLaunch = Nothing
    
End Sub

Private Sub oLaunch_DataReceived(ByVal sData As String)
    
    RTBOutput.Text = RTBOutput.Text + sData
    RTBOutput.SelStart = Len(RTBOutput.Text)
    sbStatus.Panels.Item(2).Text = "Data received"
    
End Sub

Private Sub oLaunch_ProcessEnded()
    
    bExecute = False
    cmdExecute.Caption = "&Execute"
    txtApplication.Enabled = True
    cmdWrite.Enabled = False
    sbStatus.Panels.Item(1).Text = "Stopped !"
End Sub

Private Sub txtApplication_Change()
    
    cmdExecute.Enabled = (txtApplication.Text <> "")

End Sub

Private Sub txtApplication_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
cmdExecute.value = True
End If
End Sub


Private Sub txtCommand_Change()

    cmdWrite.Enabled = (txtCommand.Text <> "")

End Sub

Private Sub txtCommand_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
cmdWrite.value = True
End If
End Sub


Private Sub rtboutput_Change()

    cmdClear.Enabled = (RTBOutput.Text <> "")

End Sub
