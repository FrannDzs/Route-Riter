VERSION 5.00
Begin VB.Form frmMini 
   Caption         =   "Mini-Route Setup"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Create new MiniRoute"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add new folder name if required and press ENTER"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Drive and Path for new Mini-Route"
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Folder PATH for Mini-Route"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim strMainPath As String, strBatFile As String, NewFile As Integer
Dim x As Integer
On Error GoTo ErrTrap

Close
If Not DirExists(App.Path & "\TempFiles") Then
MkDir App.Path & "\TempFiles"
DoEvents
End If
x = 1
If Not DirExists(Text1) Then
x = 2
MkDir Text1
MkDir Text1 & "\Train Simulator"
MkDir Text1 & "\Train Simulator\1033"
MkDir Text1 & "\Train Simulator\Fonts"
MkDir Text1 & "\Train Simulator\Global"
MkDir Text1 & "\Train Simulator\GUI"
MkDir Text1 & "\Train Simulator\Saves"
MkDir Text1 & "\Train Simulator\Routes"
MkDir Text1 & "\Train Simulator\Sound"
MkDir Text1 & "\Train Simulator\Trains"
MkDir Text1 & "\Train Simulator\Trains\Consists"
MkDir Text1 & "\Train Simulator\Trains\Trainset"
Else
x = 3
Call MsgBox(Text1 & " already exists, you must create a new folder for your Mini-Route.", vbCritical, App.Title)
Exit Sub
End If

strMainPath = MSTSPath
x = 4
If Right$(strMainPath, 1) = Chr$(0) Then
strMainPath = Left$(strMainPath, Len(strMainPath) - 1)
End If
x = 5
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strMainPath & "\*.*" & ChrW$(34) & " " & ChrW$(34) & Text1 & "\Train Simulator" & ChrW$(34) & " /Y" & vbCrLf
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strMainPath & "\1033\*.*" & ChrW$(34) & " " & ChrW$(34) & Text1 & "\Train Simulator\1033\" & ChrW$(34) & " /S /Y" & vbCrLf
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strMainPath & "\Fonts\*.*" & ChrW$(34) & " " & ChrW$(34) & Text1 & "\Train Simulator\Fonts\" & ChrW$(34) & " /S /Y" & vbCrLf
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strMainPath & "\Global\*.*" & ChrW$(34) & " " & ChrW$(34) & Text1 & "\Train Simulator\Global\" & ChrW$(34) & " /S /Y" & vbCrLf
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strMainPath & "\GUI\*.*" & ChrW$(34) & " " & ChrW$(34) & Text1 & "\Train Simulator\GUI\" & ChrW$(34) & " /S /Y" & vbCrLf
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strMainPath & "\Sound\*.*" & ChrW$(34) & " " & ChrW$(34) & Text1 & "\Train Simulator\Sound\" & ChrW$(34) & " /S /Y" & vbCrLf




If strBatFile <> vbNullString Then
x = 6
NewFile = FreeFile
Open App.Path & "\TempFiles\mini.bat" For Output As #NewFile
Print #NewFile, strBatFile
Close #NewFile
x = 7
ChDrive Left$(App.Path, 1)
x = 8
 ChDir App.Path & "\TempFiles"
x = 9
DoEvents
x = 10
Call ShellAndWait(App.Path & "\TempFiles\mini.bat", True, vbNormalFocus)
x = 11
DoEvents
End If
Close
x = 12
Call MsgBox("Your Mini-Route structure has now been set up. Now:" _
            & vbCrLf & "Use the Mini-Route Copy button to copy the required files from your Route folder into the Mini-Route 'Route' folder" _
                        , vbExclamation, App.Title)
Exit Sub
ErrTrap:
Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred while Creating a Mini-Route please advise" _
                       & vbCrLf & "Support that x=" & Str(x) _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
       
    Resume Next
        Case vbCancel
     'Resume Next
    Exit Sub
    End Select

End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Dir1_Change()
If Dir1.Path <> Dir1.List(Dir1.ListIndex) Then
Dir1.Path = Dir1.List(Dir1.ListIndex)
If Right$(Dir1.Path, 1) <> "\" Then Dir1.Path = Dir1.Path & "\"
End If
Text1 = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
Text1 = Dir1.Path
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command1.Value = True
End If
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

If Right$(Text1.Text, 1) <> "\" Then
Text1.Text = Text1.Text & "\"
End If
Text1.Text = Text1.Text & Text2
'MkDir Text1.Text
End If

End Sub


