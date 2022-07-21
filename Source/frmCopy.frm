VERSION 5.00
Begin VB.Form frmCopy 
   Caption         =   "Copy files between two Routes"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   9435
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Index           =   1
      Left            =   8760
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   8760
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "3. Select .S files in Left Hand SHAPES folder."
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "2. Select Source Route in LEFT Hand Window of Main Screen"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "1. Select Target Route in RIGHT Hand Window of Main Screen"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Text
Dim RefFile(0 To 5000) As String
Dim RefIndex(0 To 5000) As String
Dim strRefFile As String
Dim NumShapes As Long


Private Sub CopyRoute()
Dim strMiniRoute As String, strOrgRoute As String, x As Integer
Dim strRteName As String, strBatFile As String


strMiniRoute = Text1(0).Text
strOrgRoute = Text1(1).Text


x = InStrRev(strOrgRoute, "\")
strRteName = Mid$(strOrgRoute, x + 1)
MkDir strMiniRoute & "\" & strRteName
strBatFile = "call Xcopy " & ChrW$(34) & strOrgRoute & "\*.*" & ChrW$(34) & " " & ChrW$(34) & strMiniRoute & "\" & strRteName & ChrW$(34) & " /S /Y" & vbCrLf

If strBatFile <> vbNullString Then

Open App.Path & "\TempFiles\mini.bat" For Output As #1
Print #1, strBatFile
Close #1
ChDrive Left$(App.Path, 1)
 ChDir App.Path & "\TempFiles"

DoEvents
Call ShellAndWait("mini.bat", True, vbNormalFocus)

DoEvents
End If
MousePointer = 0
booMini = False

Unload Me

End Sub



Private Sub DoDeComp2(strFile As String, strFPath As String, strSparePath As String)
Dim strBatText As String, strSuffix As String

strSuffix = "-" & Right$(strFile, 1)


   ChDrive Left$(App.Path, 1)
 ''ChDir App.Path & "\TSUtil"
   strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_decomp.log" & ChrW$(34) & "  fmgr " & strSuffix & " -e -n" & ChrW$(34) & strFile & ChrW$(34) & " " & ChrW$(34) & strFPath & ChrW$(34) & " " & ChrW$(34) & strSparePath & ChrW$(34)


  Call ShellAndWait(strBatText, True, vbHide)

 DoEvents
 
End Sub


Private Sub FindCopy()
Dim fullpath$, varbat As String, x As Integer, strNew As String

Call UncompressSFiles2

cursouind = 1
SparePath = App.Path & "\TempFiles"
frmUtils.Drive1(1).Drive = Left$(SparePath, 2)
frmUtils.Dir1(1).Path = SparePath
frmUtils.Text1(1).Text = "*.s"
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
   If frmUtils.File1(cursouind).Selected(i) Then
   fullpath$ = frmUtils.File1(cursouind).Path & "\" & frmUtils.File1(cursouind).List(i)
   x = InStrRev(fullpath$, "\")
   strNew = Mid$(fullpath$, x + 1)
   varbat = varbat & "copy " & ChrW$(34) & OrigRoutePath & "\shapes\" & strNew & ChrW$(34) & " " & ChrW$(34) & DestRoutePath & "\shapes\" & strNew & ChrW$(34) & vbCrLf
   varbat = varbat & "copy " & ChrW$(34) & OrigRoutePath & "\shapes\" & strNew & "d " & ChrW$(34) & " " & ChrW$(34) & DestRoutePath & "\shapes\" & strNew & "d " & ChrW$(34) & vbCrLf
Call FindShapeRef(strNew)

Call CheckForAce3(fullpath$, varbat)
End If
Next i
Call MakeNewRef(strRefFile)


If FileExists(SparePath & "\do_ffeditc.bat") Then
Kill SparePath & "\do_ffeditc.bat"
End If
DoEvents
 Newfile3 = FreeFile
   Open SparePath & "\do_ffeditc.bat" For Append As #Newfile3
   Print #Newfile3, varbat
   Close Newfile3
   DoEvents
   strDrive = Left$(SparePath, 1)
ChDrive strDrive
ChDir SparePath
mydir = CurDir
  DoEvents
Call ShellAndWait("do_ffeditc.bat", True, vbNormalFocus)

End Sub

Private Sub FindShapeRef(ShapeName As String)
Dim x As Long

For x = 0 To NumShapes
If RefIndex(x) = ShapeName Then
strRefFile = strRefFile & RefFile(x)
Exit For
End If
Next x

End Sub

Private Sub GetRef()
Dim myfile As Integer, MyRef As String, flagway As Integer, NewFile As Integer
Dim strTemp As String

On Error GoTo Errtrap
frmUtils.Text1(0) = "*.ref"
frmUtils.Dir1(0).Path = OrigRoutePath
For i = 0 To frmUtils.File1(0).ListCount - 1
    frmUtils.File1(0).Selected(i) = True
Next i
MyRef = frmUtils.File1(0).List(ListIndex)

NewFile = FreeFile
   Open OrigRoutePath & "\" & MyRef For Binary As #NewFile
    strTemp = String(2, " ")
    Get #NewFile, , strTemp
 Close #NewFile
 
If Asc(Mid$(strTemp, 1, 1)) <> 255 And Asc(Mid$(strTemp, 2, 1)) <> 254 Then
Call ConvertIt(OrigRoutePath & "\" & MyRef, 1)
DoEvents
End If

SparePath = App.Path & "\TempFiles"
FileCopy OrigRoutePath & "\" & MyRef, SparePath & "\" & MyRef
flagway = 0
Call ConvertIt(SparePath & "\" & MyRef, flagway)
myfile = FreeFile
Open SparePath & "\" & MyRef For Input As myfile
i = 0
Do While Not EOF(myfile)
   Line Input #myfile, strNew
   strTemp = Trim$(strNew)
        If Left$(strTemp, 4) = "skip" Then
        strTemp = vbNullString
        End If
   If Left$(strTemp, 4) = "skip" Or Left$(strTemp, 4) = "stat" Or Left$(strTemp, 4) = "fore" Or Left$(strTemp, 4) = "haza" Or Left$(strTemp, 4) = "leve" Or Left$(strTemp, 4) = "tran" Then
    strTemp2 = strTemp2 & strTemp & vbCrLf
        Do
        Line Input #myfile, strNew
       
        strTemp = Trim$(strNew)
              If strTemp <> ")" Then
              strTemp2 = strTemp2 & strTemp & vbCrLf
              Else
              strTemp2 = strTemp2 & strTemp & vbCrLf
              Rem***************
              x = InStr(strTemp2, "FileName")
            
                     If x > 0 Then
                     Y = InStr(x, strTemp2, "(")
                     yy = InStr(Y, strTemp2, ")")
                     strNew = Mid$(strTemp2, Y + 1, yy - (Y + 1))
                     strNew = Trim$(strNew)
                     
                   
                              
                              If Left$(strNew, 1) = ChrW$(34) Then
                                strNew = Mid$(strNew, 2)
                                Y = InStr(strNew, ChrW$(34))
                                 If Y > 0 Then
                                 strNew = Left$(strNew, Y - 1)
                                 End If
                                End If
                    End If
              Rem****************
            
              RefIndex(i) = strNew
              RefFile(i) = strTemp2
              
              strTemp = vbNullString
              strTemp2 = vbNullString
              i = i + 1
              End If
        Loop
   End If
Loop
   
  
Label1:
   Close myfile
 NumShapes = i - 1
 Kill SparePath & "\" & MyRef
    Exit Sub
Errtrap:

If Err = 62 Then
GoTo Label1
ElseIf Err = 53 Then
Resume Next

Else

Resume Next

End If
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






Private Sub MakeNewRef(strRef As String)
Dim MyRef As String, i As Integer, NewFile As Integer, Newfile3 As Integer, flagway As Integer
Dim strTemp As String

SparePath = App.Path & "\TempFiles"
frmUtils.Text1(0) = "*.ref"
frmUtils.Dir1(0).Path = DestRoutePath
For i = 0 To frmUtils.File1(0).ListCount - 1
    frmUtils.File1(0).Selected(i) = True
Next i
MyRef = frmUtils.File1(0).List(ListIndex)
NewFile = FreeFile
   Open DestRoutePath & "\" & MyRef For Binary As #NewFile
    strTemp = String(2, " ")
    Get #NewFile, , strTemp
 Close #NewFile
 
If Asc(Mid$(strTemp, 1, 1)) <> 255 And Asc(Mid$(strTemp, 2, 1)) <> 254 Then
Call ConvertIt(DestRoutePath & "\" & MyRef, 1)
DoEvents
End If
FileCopy DestRoutePath & "\" & MyRef, SparePath & "\" & MyRef
flagway = 0
Call ConvertIt(SparePath & "\" & MyRef, flagway)

Newfile3 = FreeFile
   Open SparePath & "\temp.ref" For Output As #Newfile3
   Print #Newfile3, strRef
   Close Newfile3
'flagway = 1
'Call ConvertIt(SparePath & "\temp.ref", flagway)

NewFile = FreeFile
Open SparePath & "\temp.ref" For Input As #NewFile
Newfile3 = FreeFile
   Open SparePath & "\" & MyRef For Append As #Newfile3
   Do While Not EOF(NewFile)
   Line Input #NewFile, strNew
   Print #Newfile3, strNew
   Loop
   Close Newfile3
Close NewFile
flagway = 1
Call ConvertIt(SparePath & "\" & MyRef, flagway)
FileCopy SparePath & "\" & MyRef, DestRoutePath & "\" & MyRef
DoEvents
Kill SparePath & "\" & MyRef
End Sub

Private Sub UncompressSFiles2()

Dim i As Integer
Dim strSpc As String
Dim strOrigFile As String
'Dim tfh As TokenFileHandler, result As Boolean
'
'Set tfh = New TokenFileHandler

On Error GoTo Errtrap
Rem ********** Kill .s in the temp directory
MousePointer = 11
    cursouind = 1
SparePath = App.Path & "\TempFiles"
frmUtils.Drive1(1).Drive = Left$(SparePath, 2)
frmUtils.Dir1(1).Path = SparePath
frmUtils.Text1(1).Text = "*.S"
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i
TokMode = 0
'Call InitialiseSTokens
''Call Init_MyTokens
For i = 0 To frmUtils.File1(cursouind).ListCount - 1

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = frmUtils.File1(cursouind).Path
   strOrigFile = frmUtils.File1(cursouind).List(i)
      If Right$(strOrigFile, 2) <> ".s" Then
   MousePointer = 0
   Call MsgBox(Lang(393) & strOrigFile & " is not a Shape file.", vbExclamation, Lang(404))
GoTo NextOne
End If
   '***************

Open fullpath$ For Binary As #5
    strSpc = String(2, " ")
    Get #5, , strSpc
 Close #5
 
 If Not (Asc(Mid$(strSpc, 1, 1)) = 255 And Asc(Mid$(strSpc, 2, 1)) = 254) Then
   ' result = tfh.decompress(fullpath$, fullpath$)
   booWriteFile = True
   Call DoDeComp2(strOrigFile, fullpath$, fullpath$)
    
   End If
  End If
NextOne:
   Next i
   
  


Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If
Resume Next

End Sub

Private Sub Command1_Click(Index As Integer)
Dim Filpath2$, Filpath$
On Error GoTo Errtrap

SparePath = App.Path & "\TempFiles"
If Index = 0 Then
Command1(1).Visible = True
Text1(1).Visible = True
Label1(1).Visible = True
DestRoutePath = Text1(0).Text
 If booMini = True And Right$(DestRoutePath, 6) <> "Routes" Then
Call MsgBox("The selected folder is not a 'Routes' folder." _
            & vbCrLf & "Select the correct mini-routes 'Routes' folder and try again." _
            , vbExclamation, App.Title)

Exit Sub
End If
End If
If Index = 1 Then
OrigRoutePath = Text1(1).Text

If booMini = True Then
Call CopyRoute
DoEvents
Call MsgBox("Your Route has now been copied to your Mini-Route folder. Now:-" _
            & vbCrLf & "1. Go to the Mini-Route folder in the Left-hand window and double clicking on Train Simulator" _
            & vbCrLf & "2. Go to the Files\MSTS Path menu and click on Select to set your Mini-Route as the Default Path" _
            & vbCrLf & "3. Select your Mini-Route and click the Confirm Route button" _
            & vbCrLf & "4. Exit Route_Riter, then restart it to reset everything." _
            & vbCrLf & "5. Click on the Mini-Route Get Stock button to transfer the rolling stock across." _
            & vbCrLf & "" _
            , vbInformation, "Setting up a Mini-Route")
booMini = False

Exit Sub
End If
frmUtils.Text1(0) = "*.s"
frmUtils.Dir1(0).Path = OrigRoutePath & "\Shapes"
Label1(2).Visible = True
Command1(2).Visible = True

End If
If Index = 2 Then

cursouind = 1
SparePath = App.Path & "\TempFiles"
frmUtils.Drive1(1).Drive = Left$(SparePath, 2)
frmUtils.Dir1(1).Path = SparePath
frmUtils.Text1(1).Text = "*.s"
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
 Kill frmUtils.File1(cursouind).Path & "\" & frmUtils.File1(cursouind).List(i)
 Next i


cursouind = 0

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
  If frmUtils.File1(cursouind).Selected(i) Then
    Filpath$ = frmUtils.File1(cursouind).Path
   If Right$(Filpath$, 1) = "\" Then
   Filpath$ = Left$(Filpath$, Len(Filpath$) - 1)
   End If
Filpath2$ = frmUtils.File1(curtarind).Path
   If Right$(Filpath2$, 1) = "\" Then
   Filpath2$ = Left$(Filpath2$, Len(Filpath2$) - 1)
   End If
   FileCopy Filpath$ & "\" & frmUtils.File1(cursouind).List(i), SparePath & "\" & frmUtils.File1(cursouind).List(i)
   End If
   Next i
 
   Call GetRef
   Call FindCopy
  
frmUtils.File1(cursouind).Refresh
frmUtils.File1(curtarind).Refresh
Unload Me
End If

Exit Sub
Errtrap:

Resume Next
End Sub
Private Sub Command2_Click()
Unload Me

End Sub


Private Sub Form_GotFocus()
Dim i As Integer

If booMini = False Then
strRefFile = vbNullString
frmCopy.ZOrder
Me.Caption = Lang(209)
For i = 0 To 2
Label1(i).Caption = Lang(210 + i)
Next i
Command2.Caption = Lang(203)
ElseIf booMini = True Then
Label1(0).Caption = "1. Select Mini-Route \Routes folder in the LEFT Hand Window of Main Screen"
Label1(1).Caption = "2. Select Route to be copied in the RIGHT hand Folders Window"
Label1(2).Caption = "Click OK Now"

End If
End Sub

Public Sub CheckForAce3(SFilepath As String, varbat As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Y As Integer, Z As Integer
Dim flagACE As Integer, strS As String

On Error GoTo Errtrap
If Not FileExists(SFilepath) Then
Call MsgBox(SFilepath & " Was not found.", vbExclamation, App.Title)
Exit Sub
End If
Fnumber = FreeFile
x = InStrRev(SFilepath, "\")
strS = Mid$(SFilepath, x + 1)
flagACE = 1
Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 
   x = InStr(strNew, ".ace")
   
   If x > 0 Then
 
   Y = InStrRev(strNew, "(", x)
  Z = InStrRev(strNew, ChrW$(34), x)
   If Z > Y Then
   strNew = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   Else
   strNew = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   
   End If
   
   strNew = Trim$(strNew)
   If Left$(strNew, 1) = ChrW$(34) Then
  
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   varbat = varbat & "call xcopy " & ChrW$(34) & OrigRoutePath & "\textures\" & strNew & ChrW$(34) & " " & ChrW$(34) & DestRoutePath & "\textures\" & ChrW$(34) & " /s/y" & vbCrLf
   
    
   End If
   Loop
   
   Close #Fnumber
   Exit Sub
Errtrap:
'      If Err = 53 Then
'
'     MsgBox Lang(342) & SFilepath & Lang(343) & vbcr & Lang(344), 48, Lang(345)
''********************
'   End If


End Sub


Private Sub Form_Load()
Me.Top = 100
Me.Left = 100
frmCopy.ZOrder
Dim i As Integer

If booMini = False Then
strRefFile = vbNullString
frmCopy.ZOrder
Me.Caption = Lang(209)
For i = 0 To 2
Label1(i).Caption = Lang(210 + i)
Next i
Command2.Caption = Lang(203)
ElseIf booMini = True Then
Label1(0).Caption = "1. Select Mini-Route \Routes folder in the LEFT Hand Window of Main Screen"
Label1(1).Caption = "2. Select Route to be copied in the RIGHT hand Folders Window"
Label1(2).Caption = "Click OK Now"

End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
booCopy = False

End Sub


