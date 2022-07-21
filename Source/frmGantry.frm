VERSION 5.00
Begin VB.Form frmGantry 
   Caption         =   "Remove Gantries"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove selected item"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   5640
      Width           =   1455
   End
   Begin VB.ListBox lstGantry 
      Height          =   4560
      Left            =   720
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select Gantry File to Remove"
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
      Left            =   1440
      TabIndex        =   1
      Top             =   4920
      Width           =   4695
   End
End
Attribute VB_Name = "frmGantry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
If lstGantry.SelCount > 1 Then
Call MsgBox("You have selected more than one shape to delete. You can only delete one shape at a time.", vbExclamation, App.Title)

Exit Sub
End If
For i = 0 To lstGantry.ListCount - 1
If lstGantry.Selected(i) = True Then
strDelShape = lstGantry.List(i)
End If
Next i
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim MyString As String, x As Long, xx As Long, strTemp As String

MyString = ReadUniFile(RoutePath & "\Gantry.dat")
x = 1
Do
x = InStr(x, MyString, "Filename (")
If x > 0 Then
xx = InStr(x, MyString, ")")
strTemp = Mid$(MyString, x + 10, xx - (x + 10))
strTemp = Trim$(strTemp)
If Left$(strTemp, 1) = ChrW$(34) Then
strTemp = Mid$(strTemp, 2)
End If
If Right$(strTemp, 1) = ChrW$(34) Then
strTemp = Left$(strTemp, Len(strTemp) - 1)
End If
lstGantry.AddItem strTemp
x = xx
End If
Loop While x > 0
End Sub

Private Function ReadUniFile(CompleteFilePath As String) As String

Dim length As Long, mytristate As Integer
Dim MyString As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean


Set File_obj = CreateObject("Scripting.FileSystemObject")
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, Me.Caption
  Exit Function
End If
mytristate = -1
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll
The_obj.Close
fileflag = False
ReadUniFile = MyString
End Function


