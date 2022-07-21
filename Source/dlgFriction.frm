VERSION 5.00
Begin VB.Form dlgFriction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Type of Rolling-Stock"
   ClientHeight    =   9720
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List2 
      Height          =   3180
      Left            =   360
      TabIndex        =   3
      Top             =   5520
      Width           =   6855
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   6855
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "FRICTION Bearing Rolling-Stock"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   5040
      Width           =   4095
   End
End
Attribute VB_Name = "dlgFriction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strStock(1 To 65) As String, strFriction(1 To 65) As String
Private Function ReadUniFile(CompleteFilePath As String) As String

Dim length As Long, mytristate As Integer
Dim MyString As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean


Set File_obj = CreateObject("Scripting.FileSystemObject")
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox "" & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & "" & Lang(401), vbInformation, Me.Caption
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



Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim MyString As String
Dim x As Integer, xx As Integer, y As Integer, yy As Integer, booFBR As Boolean
Dim i As Integer

MyString = ReadUniFile(App.path & "\BrakeFiles_Pro\Railcar Friction Values\Friction values.txt")
booFBR = False
i = 1
x = 1
Do
x = InStr(x, MyString, "#")
If x = 0 Then Exit Do
If Mid(MyString, x + 1, 2) = "##" Then
booFBR = True
x = x + 4
GoTo CarryOn
End If
y = InStr(x, MyString, ":")
strStock(i) = Mid(MyString, x + 1, y - (x + 1))
If booFBR = False Then
List1.AddItem strStock(i)
ElseIf booFBR = True Then
List2.AddItem strStock(i)
End If
xx = InStr(y, MyString, "Comment")
yy = InStr(xx, MyString, "$")
strFriction(i) = Mid(MyString, xx, yy - xx)

x = y + 1
i = i + 1
CarryOn:
Loop



End Sub


Private Sub OKButton_Click()
Dim i As Integer

For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
frmEngEdit.Text4.Text = strFriction(i + 1)
DoEvents
Exit For
End If
Next i
If List1.SelCount > 0 Then GoTo CarryOn
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
frmEngEdit.Text4.Text = strFriction(i + 1 + List1.ListCount)
DoEvents
Exit For
End If
Next i
CarryOn:
Unload Me
End Sub


