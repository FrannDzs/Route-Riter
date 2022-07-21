VERSION 5.00
Begin VB.Form frmRepShape 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace Shapes in .W files"
   ClientHeight    =   2415
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmRepShape.frx":0000
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "With this Shape file"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Replace this Shape"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmRepShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Label2.Visible = True
Text2.Visible = True
OKButton.Visible = True
IntRepW = 2

End Sub

Private Sub Command2_Click()
Text1 = vbNullString
Text2 = vbNullString
IntRepW = 1
 frmUtils.Drive1(1).Drive = Left$(RoutePath, 2)
frmUtils.Dir1(1).Path = RoutePath & "\shapes"
frmUtils.Text1(1).Text = "*.s"
DoEvents
flagChange = 1
End Sub

Private Sub Form_Activate()
frmRepShape.ZOrder

End Sub

Private Sub Form_Load()
If flagChange = 1 Then
Me.Caption = Lang(287)
Label3.Caption = Lang(288)
Label1.Caption = Lang(289)
Label2.Caption = Lang(290)
ElseIf flagChange = 2 Then
Me.Caption = Lang(657)
Label3.Caption = Lang(658)
Label1.Caption = Lang(659)
Label2.Caption = Lang(660)
End If
IntRepW = 1
Label2.Visible = False
Text2.Visible = False
OKButton.Visible = False
End Sub


Private Sub ReplaceInW(strOld As String, strNewShape As String)
Dim i As Integer, filepath1$, fullpath$
Dim NewFile As Integer, strPart As String, strNew As String
Dim strName As String
MousePointer = 11
Rem

cursouind = 1
filepath1$ = App.Path & "\TempFiles"
frmUtils.Drive1(1).Drive = Left$(filepath1$, 2)
frmUtils.Dir1(1).Path = filepath1$
frmUtils.Text1(1).Text = "*.W"
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i
For i = 0 To frmUtils.File1(cursouind).ListCount - 1

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = frmUtils.File1(cursouind).Path & "\" & frmUtils.File1(cursouind).List(i)
    strName = frmUtils.File1(cursouind).List(i)
    NewFile = FreeFile
   Call ConvertIt(fullpath$, 0)
  Open fullpath$ For Binary As #NewFile
 
strNew = String(lOf(NewFile), vbNullChar)
Get NewFile, , strNew

strPart = Replace(strNew, strOld, strNewShape, , , vbTextCompare)
  Close #NewFile

    End If
    
    
     NewFile = FreeFile
   
  Open fullpath$ For Output As #NewFile
  
Print #NewFile, strPart
  Close #NewFile
Call ConvertIt(fullpath$, 1)
DoEvents

FileCopy fullpath$, RoutePath & "\world\" & strName
DoEvents
   Next i
    Close
    MousePointer = 0
End Sub

Private Function ConvertIt(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertIt = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean


'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, Me.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, Me.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, Me.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, Me.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False

'If chkSave.Value <> 1 Then
'  Kill CompleteFilePath
'Else
'  'FIND A UNIQUE NAME FOR THE ORIGINAL FILE
'  UniqueFileName = 0
'  Do While File_obj.FileExists(CompleteFilePath & ".Original" & UniqueFileName)
'    UniqueFileName = UniqueFileName + 1
'  Loop
'  Name CompleteFilePath As CompleteFilePath & ".Original" & UniqueFileName
'End If

FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertIt = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, Me.Caption
  
End Function






Private Sub OKButton_Click()
Dim strOld As String, strNew As String
IntRepW = 0
CancelButton.Caption = "&Exit"
strOld = "( " & Text1 & " )"
strNew = "( " & Text2 & " )"
Call ReplaceInW(strOld, strNew)
End Sub


