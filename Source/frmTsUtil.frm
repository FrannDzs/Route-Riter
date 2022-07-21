VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form frmTsUtil 
   Caption         =   "TsUtils for Windows"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CDTU1 
      Left            =   240
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      ToolTipText     =   "Give up"
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      ToolTipText     =   "Carry On"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   6960
      Width           =   7455
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   600
      TabIndex        =   3
      ToolTipText     =   "Available Routes"
      Top             =   1200
      Width           =   8775
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "Available TsUtil Options"
      Top             =   4080
      Width           =   8775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      ToolTipText     =   "Click to select a different version of MSTS"
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Select Option:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Select Route (Only unstored routes shown) :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   960
      Width           =   5415
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   6600
      Width           =   7455
   End
   Begin VB.Label Label4 
      Caption         =   "Option Selected"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Route Selected"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   "Select a different MSTS path if necessary"
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "MSTS Path"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmTsUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckForTrk(strRoute As String, booFound As Boolean)
 frmUtils.Drive1(1).Drive = Left$(strRoute, 2)
frmUtils.Dir1(1).Path = strRoute
frmUtils.Text1(1).Text = "*.trk"
DoEvents
If frmUtils.File1(1).ListCount = 0 Then
booFound = False
Else
booFound = True
End If
End Sub

Private Sub FillListBox()
Dim MyPath As String, MyName As String, booFound As Boolean


MyPath = strTSRoutePath & "\"
MyName = Dir(MyPath, vbDirectory)
Do While MyName <> ""
If MyName <> "." And MyName <> ".." Then
    If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
    Call CheckForTrk(strTSRoutePath & "\" & MyName, booFound)
    If booFound = True Then
    List2.AddItem strTSRoutePath & "\" & MyName
    End If
    End If
End If
MyName = Dir
Loop

End Sub

Private Sub Command1_Click()
TSUFlag = 1

TsUtil_CD.Show 1
DoEvents
strTSRoutePath = Label2.Caption

Call FillListBox

End Sub

Private Sub Command2_Click()
        
        If TSOption = 11 Then GoTo CarryON
        If strSelectedPath = "" Or strTSoption = "" Then
        Call MsgBox("You have not selected either a route or an option.", vbExclamation, App.Title)
            Exit Sub
        End If
CarryON:
        Command3.Caption = "Exit"

        frmProcess.Show
        DoEvents
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim NewFile As Integer, strTemp As String
Dim result As String

If MSTSPath <> "" Then
strTSRoutePath = MSTSPath & "\Routes"
Else
result = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "Path")
MSTSPath = result
strTSRoutePath = MSTSPath & "\Routes"
End If
DoEvents
Label2.Caption = strTSRoutePath
DoEvents

Call FillListBox
NewFile = FreeFile
Open App.Path & "\options.txt" For Input As #NewFile
Do While Not EOF(NewFile)
Input #NewFile, strTemp
List1.AddItem strTemp
Loop
Close #NewFile

End Sub


Private Sub List1_Click()
 Dim x As Integer, strBox As String, Y As Integer

        For x = 0 To List1.ListCount - 1

            ' Determine if the item is selected.
            If List1.Selected(x) = True Then
                ' Deselect all items that are selected.
                strBox = List1.List(x)
                Text1.Text = strBox
                TSOption = x + 1
            
                Y = InStr(1, strBox, vbTab)
               
                strTSoption = Mid(strBox, Y + 1)



            End If
        Next x
        
End Sub

Private Sub List2_Click()
Dim x As Integer, strBox As String, Y As Integer

        For x = 0 To List2.ListCount - 1

            ' Determine if the item is selected.
            If List2.Selected(x) = True Then
                ' Deselect all items that are selected.
                strBox = List2.List(x)
                Label5.Caption = strBox
                strSelectedPath = strBox
                Y = InStrRev(strSelectedPath, "\")
                strRouteName = Mid(strSelectedPath, Y + 1)


            End If
        Next x
End Sub


