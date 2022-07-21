VERSION 5.00
Begin VB.Form frmUpdate 
   Caption         =   "Produce Update Patch for Route"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   9435
   Begin VB.CommandButton Command3 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
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
      Visible         =   0   'False
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copying Files"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "3. Make Update Patch from 1. to 2."
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
      Caption         =   "2. Select New Version of Route in Either Dir-List  Window"
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
      Caption         =   "1. Select Early Version of Route in either Dir-List Window"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Text


Dim PatchPath As String
Dim PatchName As String
Dim OldRoutePath As String
Dim NewRoutePath As String
Dim OldFiles() As Variant
Dim NewFiles() As Variant
Dim OldShortPath() As String
Dim OldCRC() As String
Dim NewShortPath() As String
Dim NewCRC() As String
Dim lngOldFiles As Long, lngNewFiles As Long
Dim OldRouteName As String
Dim NewRouteName As String
Dim Patch(1 To 35) As String
Dim PatchRoot As String
Const CHUNK = 500
Function GetCRC(strFileName As String) As String
Dim strTemp As String

 Set m_CRC = New clsCRC
 m_CRC.Algorithm = 1
 strTemp = Hex(m_CRC.CalculateFile(strFileName))
 GetCRC = strTemp
 
End Function

Private Sub WriteUpdateMe()
Dim strBat As String, NewFile As Integer

strBat = "cd .." & vbCrLf
strBat = strBat & "rename " & OldRouteName & " " & NewRouteName & vbCrLf
strBat = strBat & "cd " & NewRouteName & vbCrLf
strBat = strBat & "del ?*.rdb" & vbCrLf
strBat = strBat & "del ?*.rit" & vbCrLf
strBat = strBat & "del ?*.ref" & vbCrLf
strBat = strBat & "del ?*.tdb" & vbCrLf
strBat = strBat & "del ?*.tit" & vbCrLf
strBat = strBat & "del ?*.trk" & vbCrLf
'strBat = strBat & "del activities\?*.*" & vbCrLf
'strBat = strBat & "del services\?*.*" & vbCrLf
'strBat = strBat & "del paths\?*.*" & vbCrLf
'strBat = strBat & "del traffic\?*.*" & vbCrLf
strBat = strBat & "@Echo off" & vbCrLf
strBat = strBat & "Echo." & vbCrLf
strBat = strBat & "Echo " & ChrW$(34) & "Clean-up Finished - Press Any Key to start copying files" & ChrW$(34) & vbCrLf
strBat = strBat & "Echo " & ChrW$(34) & "Installme.bat will then run automatically" & ChrW$(34) & vbCrLf
strBat = strBat & "Echo." & vbCrLf
strBat = strBat & "Echo." & vbCrLf
strBat = strBat & "@pause" & vbCrLf
strBat = strBat & "cd .." & vbCrLf
strBat = strBat & "call xcopy " & PatchName & "\*.*" & " " & NewRouteName & "\ /s /y" & vbCrLf
strBat = strBat & "cd " & NewRouteName & vbCrLf
strBat = strBat & "start installme.bat" & vbCrLf
strBat = strBat & "@Echo off" & vbCrLf
strBat = strBat & "Echo." & vbCrLf
strBat = strBat & "Echo." & vbCrLf
strBat = strBat & "Echo All done, Close DOS box to finish." & vbCrLf
strBat = strBat & "@pause" & vbCrLf
NewFile = FreeFile
Open PatchPath & "\UpdateMe.bat" For Output As #NewFile

Print #NewFile, strBat
Close #NewFile

NewFile = FreeFile
Open PatchPath & "\Installme.bat" For Append As #NewFile
Print #NewFile, "DEL .\Tiles\*e.raw"
Print #NewFile, "DEL .\Tiles\*n.raw"
Print #NewFile, "If not exist c:\windows\command\deltree.exe goto tryxp"
Print #NewFile, "Deltree /y " & ChrW$(34) & "..\" & PatchName & ChrW$(34)
Print #NewFile, "GoTo CarryOn"
Print #NewFile, ":tryxp"
Print #NewFile, "RD /s /q " & ChrW$(34) & "..\" & PatchName & ChrW$(34)
Print #NewFile, ":CarryOn"
Print #NewFile, "echo *"
Print #NewFile, "echo *"
Print #NewFile, "echo All done - Please close any open DOS windows"
Print #NewFile, "echo *"
Close #NewFile
End Sub

Private Sub Command1_Click(Index As Integer)
Dim X As Integer
On Error GoTo ErrTrap


SparePath = App.Path & "\TempFiles"

Select Case Index
Case 0
'Command1(1).Visible = True
Text1(1).Visible = True
Label1(1).Visible = True
OldRoutePath = Text1(0).Text
X = InStrRev(OldRoutePath, "\")
OldRouteName = Mid$(OldRoutePath, X + 1)
intUpdate = 1

Case 1
NewRoutePath = Text1(1).Text
X = InStrRev(NewRoutePath, "\")
NewRouteName = Mid$(NewRoutePath, X + 1)

Label1(2).Visible = True
Command1(2).Visible = True
PatchPath = Text1(1) & "_Patch"
X = InStrRev(PatchPath, "\")
PatchName = Mid$(PatchPath, X + 1)
PatchRoot = Left$(PatchPath, X - 1)

If Not DirExists(PatchPath) Then MkDir PatchPath
If Not DirExists(PatchPath & "\Activities") Then MkDir PatchPath & "\Activities"
If Not DirExists(PatchPath & "\Envfiles") Then MkDir PatchPath & "\Envfiles"
If Not DirExists(PatchPath & "\Envfiles\Textures") Then MkDir PatchPath & "\Envfiles\Textures"
If Not DirExists(PatchPath & "\Lo_Tiles") Then MkDir PatchPath & "\Lo_Tiles"
If Not DirExists(PatchPath & "\Paths") Then MkDir PatchPath & "\Paths"
If Not DirExists(PatchPath & "\Services") Then MkDir PatchPath & "\Services"
If Not DirExists(PatchPath & "\Shapes") Then MkDir PatchPath & "\Shapes"
If Not DirExists(PatchPath & "\Sound") Then MkDir PatchPath & "\Sound"
If Not DirExists(PatchPath & "\Td") Then MkDir PatchPath & "\Td"
If Not DirExists(PatchPath & "\Terrtex") Then MkDir PatchPath & "\Terrtex"
If Not DirExists(PatchPath & "\Terrtex\Snow") Then MkDir PatchPath & "\Terrtex\Snow"
If Not DirExists(PatchPath & "\Textures") Then MkDir PatchPath & "\Textures"
If Not DirExists(PatchPath & "\Textures\Autumn") Then MkDir PatchPath & "\Textures\Autumn"
If Not DirExists(PatchPath & "\Textures\AutumnSnow") Then MkDir PatchPath & "\Textures\AutumnSnow"
If Not DirExists(PatchPath & "\Textures\Night") Then MkDir PatchPath & "\Textures\Night"
If Not DirExists(PatchPath & "\Textures\Snow") Then MkDir PatchPath & "\Textures\Snow"
If Not DirExists(PatchPath & "\Textures\Spring") Then MkDir PatchPath & "\Textures\Spring"
If Not DirExists(PatchPath & "\Textures\SpringSnow") Then MkDir PatchPath & "\Textures\SpringSnow"
If Not DirExists(PatchPath & "\Textures\Winter") Then MkDir PatchPath & "\Textures\Winter"
If Not DirExists(PatchPath & "\Textures\WinterSnow") Then MkDir PatchPath & "\Textures\WinterSnow"
If Not DirExists(PatchPath & "\Tiles") Then MkDir PatchPath & "\Tiles"
If Not DirExists(PatchPath & "\Traffic") Then MkDir PatchPath & "\Traffic"
If Not DirExists(PatchPath & "\World") Then MkDir PatchPath & "\World"
If Not DirExists(PatchPath & "\Global") Then MkDir PatchPath & "\Global"
If Not DirExists(PatchPath & "\Global\Shapes") Then MkDir PatchPath & "\Global\Shapes"
If Not DirExists(PatchPath & "\Global\Textures") Then MkDir PatchPath & "\Global\Textures"
DoEvents
Command3.Visible = True


End Select
Exit Sub
ErrTrap:

Resume Next
End Sub
Private Sub CompareArrays()
Dim X As Long, Y As Long, booFound As Boolean

On Error GoTo ErrTrap

For X = 1 To lngNewFiles
frmUtils.SB1.Panels(2).Text = NewFiles(X)
For Y = 1 To lngOldFiles

booFound = False

If NewFiles(X) = OldFiles(Y) Then


    If NewCRC(X) = OldCRC(Y) Then
   
    booFound = True
    Exit For
    End If
End If

Next Y
If booFound = False Then            'New file not in old route
If Not DirExists(PatchPath & "\" & NewShortPath(X)) Then MkDir PatchPath & "\" & NewShortPath(X)
FileCopy NewRoutePath & "\" & NewShortPath(X) & NewFiles(X), PatchPath & "\" & NewShortPath(X) & NewFiles(X)
DoEvents
ElseIf NewShortPath(X) = "Global\Shapes\" Or NewShortPath(X) = "Global\Textures\" Then
If Not DirExists(PatchPath & "\" & NewShortPath(X)) Then MkDir PatchPath & "\" & NewShortPath(X)
FileCopy NewRoutePath & "\" & NewShortPath(X) & NewFiles(X), PatchPath & "\" & NewShortPath(X) & NewFiles(X)
DoEvents
End If
Next X

Exit Sub
ErrTrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'CompareArrays' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub

Private Function DirDiverNew(NewPath As String, DirCount As Integer, Backup As String) As Integer
'  Recursively search directories from NewPath down...
'  NewPath is searched on this recursion.
'  BackUp is origin of this recursion.
'  DirCount is number of subdirectories in this directory.

Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, Entry As String
Dim retval As Integer, strShortPath As String
Dim strCRC As String, X As Integer



    SearchFlag = True           ' Set flag so the user can interrupt.
    DirDiverNew = False            ' Set to True if there is an error.
    retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    If SearchFlag = False Then
        DirDiverNew = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = frmUtils.Dir1(cursouind).ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = frmUtils.Dir1(cursouind).Path                      ' Save old path for next recursion.
        frmUtils.Dir1(cursouind).Path = NewPath
        If frmUtils.Dir1(cursouind).ListCount > 0 Then
            ' Get to the node bottom.
            frmUtils.Dir1(cursouind).Path = frmUtils.Dir1(cursouind).List(DirsToPeek - 1)
            AbandonSearch = DirDiverNew((frmUtils.Dir1(cursouind).Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    
    ' Call function to enumerate files.
    If frmUtils.File1(cursouind).ListCount Then
        If Len(frmUtils.Dir1(cursouind).Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = frmUtils.Dir1(cursouind).Path                  ' If at root level, leave as is...
        Else
            ThePath = frmUtils.Dir1(cursouind).Path + "\"            ' Otherwise put "\" before the filename.
        End If
        
        X = InStrRev(ThePath, NewRouteName)
        strShortPath = Mid$(ThePath, X + Len(NewRouteName) + 1)
        For ind = 0 To frmUtils.File1(cursouind).ListCount - 1
        ' Add conforming files in this directory to the list box.
            Entry = frmUtils.File1(cursouind).List(ind)
            frmUtils.SB1.Panels(2).Text = Entry
           ' strShortPath = Mid$(ThePath, Len(BackUp) + 2)
            strCRC = GetCRC(ThePath & Entry)
            
             ' Add conforming files in this directory to the list box.
           lngNewFiles = lngNewFiles + 1
           If lngNewFiles > UBound(NewFiles) Then
           ReDim Preserve NewFiles(1 To lngNewFiles + CHUNK)
           ReDim Preserve NewShortPath(1 To lngNewFiles + CHUNK)
           ReDim Preserve NewCRC(1 To lngNewFiles + CHUNK)
           End If
           NewFiles(lngNewFiles) = Entry
           NewShortPath(lngNewFiles) = strShortPath
           NewCRC(lngNewFiles) = strCRC
          
        Next ind
    End If
    If Backup <> vbNullString Then        ' If there is a superior directory, move it.
        frmUtils.Dir1(cursouind).Path = Backup
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiverNew = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    Else                        ' Otherwise display error message and quit.
        MsgBox Error
        End
    End If
End Function



Private Function DirDiverOld(NewPath As String, DirCount As Integer, Backup As String) As Integer
'  Recursively search directories from NewPath down...
'  NewPath is searched on this recursion.
'  BackUp is origin of this recursion.
'  DirCount is number of subdirectories in this directory.

Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, Entry As String
Dim retval As Integer, strShortPath As String
Dim strCRC As String, X As Integer




    SearchFlag = True           ' Set flag so the user can interrupt.
    DirDiverOld = False            ' Set to True if there is an error.
    retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    If SearchFlag = False Then
        DirDiverOld = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = frmUtils.Dir1(cursouind).ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = frmUtils.Dir1(cursouind).Path                      ' Save old path for next recursion.
        frmUtils.Dir1(cursouind).Path = NewPath
        If frmUtils.Dir1(cursouind).ListCount > 0 Then
            ' Get to the node bottom.
            frmUtils.Dir1(cursouind).Path = frmUtils.Dir1(cursouind).List(DirsToPeek - 1)
            AbandonSearch = DirDiverOld((frmUtils.Dir1(cursouind).Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    
    ' Call function to enumerate files.
    If frmUtils.File1(cursouind).ListCount Then
        If Len(frmUtils.Dir1(cursouind).Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = frmUtils.Dir1(cursouind).Path                  ' If at root level, leave as is...
        Else
            ThePath = frmUtils.Dir1(cursouind).Path + "\"            ' Otherwise put "\" before the filename.
        End If
        
        X = InStrRev(ThePath, OldRouteName)
        strShortPath = Mid$(ThePath, X + Len(OldRouteName) + 1)
        For ind = 0 To frmUtils.File1(cursouind).ListCount - 1
        ' Add conforming files in this directory to the list box.
            Entry = frmUtils.File1(cursouind).List(ind)
            frmUtils.SB1.Panels(2).Text = Entry
            'strShortPath = Mid$(ThePath, Len(BackUp) + 2)
            strCRC = GetCRC(ThePath & Entry)
            
             ' Add conforming files in this directory to the list box.
           lngOldFiles = lngOldFiles + 1
           If lngOldFiles > UBound(OldFiles) Then
           ReDim Preserve OldFiles(1 To lngOldFiles + CHUNK)
           ReDim Preserve OldShortPath(1 To lngOldFiles + CHUNK)
           ReDim Preserve OldCRC(1 To lngOldFiles + CHUNK)
           End If
           
           OldFiles(lngOldFiles) = Entry
           OldShortPath(lngOldFiles) = strShortPath
           
           OldCRC(lngOldFiles) = strCRC
       
        Next ind
        
    End If
    If Backup <> vbNullString Then        ' If there is a superior directory, move it.
        frmUtils.Dir1(cursouind).Path = Backup
    End If
   
    Exit Function
DirDriverHandler:
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiverOld = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    Else                        ' Otherwise display error message and quit.
        MsgBox Error
        
        End
        Resume Next
    End If
End Function







Private Sub Command2_Click()
If Me.Caption = "&Cancel" Then
Unload Me
Else
Unload Me
End If

End Sub


Private Sub Command3_Click()
Dim FirstPath As String, DirCount As Integer, NewFile As Integer, i As Integer

On Error GoTo ErrTrap


SparePath = App.Path & "\TempFiles"
Close

MousePointer = 11
cursouind = 1

frmUtils.Dir1(cursouind).Path = OldRoutePath
    frmUtils.File1(cursouind).Pattern = frmUtils.Text1(cursouind).Text
   
    FirstPath = frmUtils.Dir1(cursouind).Path
    DirCount = frmUtils.Dir1(cursouind).ListCount
    result = DirDiverOld(FirstPath, DirCount, "")
    frmUtils.File1(cursouind).Path = frmUtils.Dir1(cursouind).Path
    DoEvents
   
    cursouind = 0

frmUtils.Dir1(cursouind).Path = NewRoutePath
    frmUtils.File1(cursouind).Pattern = frmUtils.Text1(cursouind).Text
    FirstPath = frmUtils.Dir1(cursouind).Path
    DirCount = frmUtils.Dir1(cursouind).ListCount
    result = DirDiverNew(FirstPath, DirCount, "")
    frmUtils.File1(cursouind).Path = frmUtils.Dir1(cursouind).Path
    
MousePointer = 0

            ReDim Preserve OldFiles(1 To lngOldFiles)
           ReDim Preserve OldShortPath(1 To lngOldFiles)
           ReDim Preserve OldCRC(1 To lngOldFiles)
            ReDim Preserve NewFiles(1 To lngNewFiles)
           ReDim Preserve NewShortPath(1 To lngNewFiles)
           ReDim Preserve NewCRC(1 To lngNewFiles)
Label2.Caption = "Copying Files"
Label2.Visible = True
Call CompareArrays
Call WriteUpdateMe

NewFile = FreeFile

Open PatchRoot & "\" & NewRouteName & "_Installation.txt" For Output As #NewFile
Print #NewFile, Patch(1) & OldRouteName & " To " & NewRouteName
Print #NewFile, Patch(2)
For i = 3 To 16
Print #NewFile, Patch(i)
Next i
Print #NewFile, "a folder named " & NewRouteName & "_Patch"
For i = 17 To 35
Print #NewFile, Patch(i)
Next i
Close #NewFile

Label2.BackColor = vbGreen
Label2.Caption = Lang(383)
Command2.Caption = Lang(38)
MousePointer = 0
Exit Sub
ErrTrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'frmUpdate - Command3' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub
Private Sub Form_GotFocus()
frmUpdate.ZOrder
End Sub

Private Sub Form_Load()
Dim i As Integer
ReDim NewFiles(1 To CHUNK)
ReDim NewShortPath(1 To CHUNK)
ReDim NewCRC(1 To CHUNK)
ReDim OldFiles(1 To CHUNK)
ReDim OldShortPath(1 To CHUNK)
ReDim OldCRC(1 To CHUNK)


Me.Caption = Lang(325)
For i = 0 To 2
Label1(i).Caption = Lang(326 + i)
Next i
Label2.Caption = Lang(329)

Me.Top = 100
Me.Left = 100
frmUpdate.ZOrder
intUpdate = 0

Patch(1) = "This .zip file will patch Version "
Patch(2) = vbNullString
Patch(3) = "Note that after conversion, the original version will NO LONGER be available on"
Patch(4) = "your system. If you wish to retain the original version, use Windows Explorer or"
Patch(5) = "other file utility to copy the route to a temporary folder until the conversion"
Patch(6) = "has completed. The new version will be renamed, so once it is installed, it is"
Patch(7) = "quite safe to move the old version back and then both versions will be available."
Patch(8) = vbNullString
Patch(9) = "Installation of Patch"
Patch(10) = vbNullString
Patch(11) = "Requirements:-"
Patch(12) = vbNullString
Patch(13) = "You must have the original version of the route and the 6 default MSTS routes"
Patch(14) = "available on your PC for this installation to work."
Patch(15) = vbNullString
Patch(16) = "1. Unzip this Patch into your MSTS\Routes folder, it will be placed in"
Patch(17) = "Ensure that your copy of WinZip has 'Use Folder Names' checked."
Patch(18) = vbNullString
Patch(19) = "2. Navigate to the Patch folder and click on UpdateMe.bat - This batch"
Patch(20) = "file will copy all the necessary files from the Patch folder into the old route"
Patch(21) = "folder, delete some old files, and rename the folder to the new route name. "
Patch(22) = vbNullString
Patch(23) = "3. Once this has finished, you will be advised to 'Press Any Key' to run the"
Patch(24) = "installme.bat file which will then copy any necessary files from the Default routes."
Patch(25) = vbNullString
Patch(26) = "4. Upon completion the .bat will delete the Patch folder which is"
Patch(27) = "no longer necessary (and will confuse MSTS if not deleted as MSTS thinks it is another"
Patch(28) = "route)."
Patch(29) = vbNullString
Patch(30) = "5. Close any DOS windows which are still open then Run your new route."
Patch(31) = vbNullString
Patch(32) = "6. Once you are happy that everything is working as required, it is recommended that"
Patch(33) = "you Compact the route with Route_Riter (available for download from uktrainsim.com or"
Patch(34) = "train-sim.com along with many other MSTS library sites). This will remove any extra"
Patch(35) = "files left over from the old version of the route."

End Sub


Private Sub Form_Unload(Cancel As Integer)
booUpdate = False
End Sub


Private Sub Text1_Change(Index As Integer)
If Index = 0 Then
Command1(0).Visible = True
ElseIf Index = 1 Then
Command1(1).Visible = True
End If
End Sub


