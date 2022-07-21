VERSION 5.00
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Folder"
   ClientHeight    =   2640
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Backup Route"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   7095
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select folder for Backups in the RIGHT hand folder list, or if the folder does not exist, add it below. (e.g. C:\MSTS_Backups )"
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Text1 = vbNullString
Unload Me
End Sub

Private Sub Command1_Click()
Dim strTimePath As String, strDrv As String, strNewBat As String
Dim NewFile As Integer, Filpath1$
MousePointer = 11
Filpath1$ = App.Path & "\TempFiles"
If strBackupPath = vbNullString Then Exit Sub

If Right$(strBackupPath, 1) <> "\" Then
strBackupPath = strBackupPath & "\"
End If
SaveSetting "Route_Riter6", "Backup", "Path", strBackupPath

strTimePath = strBackupPath & RouteName & "_" & Format(Now, "d-mmm-yyyy-hhmm")
If Not DirExists(strBackupPath) Then
MkDir strBackupPath
End If
'If Not DirExists(strBackupPath & RouteName) Then
'MkDir strBackupPath & RouteName
'End If
If Not DirExists(strTimePath) Then
MkDir strTimePath
MkDir strTimePath & "\World"
MkDir strTimePath & "\Tiles"
MkDir strTimePath & "\TD"
End If
Rem ****************
strDrv = Left$(RoutePath, 2)

strNewBat = strDrv & vbCrLf
strNewBat = strNewBat & "chdir " & ChrW$(34) & RoutePath & ChrW$(34) & vbCrLf
strNewBat = strNewBat & "call xcopy " & "*.* " & ChrW$(34) & strTimePath & ChrW$(34) & " /y" & vbCrLf
strNewBat = strNewBat & "chdir " & ChrW$(34) & RoutePath & "\World" & ChrW$(34) & vbCrLf
strNewBat = strNewBat & "call xcopy " & "*.* " & ChrW$(34) & strTimePath & "\World" & ChrW$(34) & " /y" & vbCrLf
strNewBat = strNewBat & "chdir " & ChrW$(34) & RoutePath & "\TD" & ChrW$(34) & vbCrLf
strNewBat = strNewBat & "call xcopy " & "*.* " & ChrW$(34) & strTimePath & "\TD" & ChrW$(34) & " /y" & vbCrLf
strNewBat = strNewBat & "chdir " & ChrW$(34) & RoutePath & "\Tiles" & ChrW$(34) & vbCrLf
strNewBat = strNewBat & "call xcopy " & "*.* " & ChrW$(34) & strTimePath & "\Tiles" & ChrW$(34) & " /y" & vbCrLf

NewFile = FreeFile

  Open Filpath1$ & "\newbat.bat" For Output As #NewFile
  Print #NewFile, strNewBat
  Close #NewFile
  ChDrive Left$(Filpath1$, 1)
  ChDir Filpath1$
 
    Call ShellAndWait(ChrW$(34) & Filpath1$ & "\newbat.bat" & ChrW$(34), True, vbNormalFocus)
    DoEvents
    MousePointer = 0
Unload Me
End Sub

Private Sub Form_Load()
Text1 = strBackupPath

End Sub


Private Sub OKButton_Click()

strBackupPath = Text1
If Text1 = vbNullString Then
Unload Me
End If
If Right$(Text1, 1) = "\" Then
Text1 = Left$(Text1, Len(Text1) - 1)
End If
Command1.Visible = True


End Sub


