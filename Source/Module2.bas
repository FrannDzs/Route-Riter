Attribute VB_Name = "Module2"
Option Explicit
Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public strTSRoutePath As String, FlagColGreen As Boolean
Public TSOption As Integer, strRouteName As String, strActReport As String
Public strTSoption As String, strOptionCode As String, booWrongMSTS As Boolean
Public strSelectedPath As String, strLogPath As String, MasterRoutePath As String
Public MyMainString As String, booTerrtexSnow As Boolean, strTrainStore As String, strConbuilder As String, flagNoTex As Boolean
Public PixRealName() As String, strSavePix As String, strPixPath As String, PixPath() As String
Public booActsChecked As Boolean, PixPicture() As String, PixName() As String, intNumPix As Long
Public booGetW As Boolean, lngZipSize As Long, lngZipFiles As Long, strINI As String, booFixFix As Boolean
Public flagSvcRed As Boolean, TokMode As Integer, booSpareTrack As Boolean
Public strPrefix As String, strReplace As String, strWith As String, booMulti As Boolean, flagInternet As Integer
Public nl As String, Lang() As String, Language(0 To 9) As String, numFiles As Long, strZipName As String
Public SizeZIP As Long, strDelShape As String, strConsists() As String, strPicView As String
Public NewZipPath As String, intGrid As Integer, pathTsection As String, intNextPix As Integer
Public maincendir As String, tit4$, OrigRoutePath As String, DestRoutePath As String, strTextEditor As String
Public strUniName As String, strComPath As String, conItem(0 To 200) As String, strBBoxFix As String
Public conNumber As Integer, conFlip(0 To 200) As Boolean, conWagon(0 To 200) As String
Public intResponse As Integer, intResponse2 As Integer, intResponseSnow As Integer
Public strResponse As String, ZipName As String, booListAce As Boolean, booMU As Boolean
Public RoutePath As String, BooCheckAct As Boolean, booExact As Boolean, booWriteFile As Boolean
Public flagCouple As Integer, FromAct As Boolean, FromZip As Integer, strMainReport As String
Public booFixAct As Boolean, booFixSrv As Boolean, booFixEng As Boolean, booFixSMS As Boolean
Public booFixSD As Boolean, booFixCon As Boolean, booWorldCount As Boolean, booMini As Boolean
Public strbadbits As String, ZIPfilename As String, booTsection As Boolean, strBackupPath As String
Public objList As String, ThisRow As Integer, booEnvFound As Boolean, IntRepW As Integer, strEditPath As String
Public booFixBB As Boolean, intAceResponse As Integer, strTPath As String, booFixCVF As Boolean
Public AllRoutes2() As String, booLink As Boolean, lLink As Long, booKillMove As Boolean, booRaildriver As Boolean
Public strOldRegPath As String, strScreenShotLocation As String, booEMU As Boolean, booDMU As Boolean
Public flagActBad As Boolean, intUpdate As Integer, booOKAll As Boolean, TSUFlag As Integer
Public strOldCon As String, strNewCon As String, flagConBad As Boolean, strKillFiles As String
Public flagGrid As Integer, booCancel As Boolean, booCopy As Boolean, strSnowName As String
Public booStockOnly As Boolean, booCommon As Boolean, booSnow As Boolean, booInsert As Boolean
Public selFlag As Integer, booUnusedChanged As Boolean, booComDir As Boolean, RouteName As String
Public strForPrint As String, booUniEdit As Boolean, booUpdate As Boolean, booCompressAll As Boolean
Public strForPrint2 As String, booAI As Boolean, booList As Boolean, booIron As Boolean
Public flagPrint As Integer, EngSize As Single, ThisCon As Integer, strReport As String, Trainspath As String
Public lngCon As Long, missEng As Boolean, missWag As Boolean, booAbort As Boolean, flagChange As Integer
Public lngSrv As Long, SrvPath() As String, Service() As String, Consists() As String, strWagonShape As String
Public Activities() As Variant, lngAct As Long, ActPath() As Variant, booBackup As Boolean
Public PConName() As Variant, PSvcName() As Variant, PTfcName() As Variant, booSouth As Boolean
Public pPathName() As Variant, booAllCompressed As Boolean, intPix As Integer
Public booNoButtons As Boolean, strKillPath As String
Public MSTSPath As String, PathUsed() As String, PathUsedNumb As Integer
Public FlagColRed As Boolean, flagThumb As Integer
Public LocoPath() As String, Wagpath() As String
Public Locomotives() As String, lngLoco As Long, LocoName() As String
Public Wagons() As String, lngWagons As Long, WagonName() As String
Public ConIntName() As String, ConIntWagName() As String
Public lngTfc As Long, Traffic() As Variant, TfcPath() As Variant
Public Paths() As Variant, PathsPath() As Variant, lngPaths As Long
Public RiseSet(1 To 8) As String, MoonSet(1 To 8) As String
Public RStart(1 To 4) As Double
Public LocoCoup() As Long, LocoFCoup() As Long, LocoBrake() As Long, LocoType() As Long, LocoRigid() As Integer, LocoFRigid() As Integer
Public WagCoup() As Long, WagFCoup() As Long, WagBrake() As Long, WagType() As Long, WagRigid() As Integer, WagFRigid() As Integer
Public Coupling(0 To 3) As String, Brake(0 To 8) As String, StockType(1 To 8) As String
Public booReport As Boolean, FCoupling(0 To 3) As String, Rigid(0 To 3) As String, FRigid(0 To 3)


 


Public LocoSMS() As String
Public booProBrakes As Boolean, booLSD As Boolean
Public Function ClearTmp()
Dim EnvString As String, Indx As Integer
Dim strTemp As String, strBatFile As String

Indx = 1   ' Initialize index to 1.
Do
   EnvString = Environ(Indx)   ' Get environment
   
            ' variable.
   If Left(EnvString, 5) = "TEMP=" Then   ' Check Temp entry.
      strTemp = Mid(EnvString, 6)
      Exit Do
   Else
      Indx = Indx + 1   ' Not PATH entry,
   End If   ' so increment.
Loop Until EnvString = ""
strBatFile = "DEL Tsutil*.tmp"
'strBatFile = "dir /p"
If strBatFile <> vbNullString Then

Open strTemp & "\debug.bat" For Output As #1
Print #1, strBatFile
Close #1
ChDrive Left$(strTemp, 1)
 ChDir strTemp

DoEvents
Call ShellAndWait("debug.bat", True, vbNormalFocus)

DoEvents
End If

'ChDrive Left$(strTEMP, 1)
' ChDir strTEMP

'DoEvents
'Call ShellAndWait(strBatFile, True, vbNormalFocus)

DoEvents

End Function

Sub forecap(fornam As String)
Dim E As Integer

If fornam = vbNullString Then Exit Sub

Mid$(fornam, 1, 1) = UCase(Mid$(fornam, 1, 1))
For E = 1 To Len(fornam) - 1
If Mid$(fornam, E, 1) = " " Then
Mid$(fornam, E + 1, 1) = UCase(Mid$(fornam, E + 1, 1))
End If
Next E


End Sub

Public Sub CheckWagonData(strNew As String, Wagname As String, Wagonpath As String, booFound As Boolean)
Dim x As Integer, Y As Integer, yy As Integer, Z As Integer, strNew2 As String

On Error GoTo Errtrap
strNew = Trim$(strNew)
Z = InStr(strNew, vbLf)
If Z > 0 Then
strNew = Left$(strNew, Z - 1)
End If
    x = InStr(1, strNew, "WagonData", vbTextCompare)
      If x > 0 Then
      
    Y = InStr(x, strNew, "(")
    
    yy = InStrRev(strNew, ")")
   
   strNew2 = Trim$(Mid$(strNew, Y + 1, yy - (Y + 1)))
   
   strNew2 = Left$(strNew2, Len(strNew2))
   strNew2 = Trim$(strNew2)
   If Left$(strNew2, 1) = ChrW$(34) Then
   x = InStr(2, strNew2, ChrW$(34))
   Else
   x = InStr(strNew2, " ")
   End If
  ' x = InStr(strNew2, " ")
   Wagname = Left$(strNew2, x - 1)
   Wagonpath = Mid$(strNew2, x + 1)
   If Left$(Wagname, 1) = ChrW$(34) Then
   Wagname = Mid$(Wagname, 2)
   Y = InStr(Wagname, ChrW$(34))
    If Y > 0 Then
    Wagname = Left$(Wagname, Y - 1)
    End If
   End If
   Wagonpath = Trim$(Wagonpath)
   If Left$(Wagonpath, 1) = ChrW$(34) Then
   Wagonpath = Mid$(Wagonpath, 2)
   Y = InStr(Wagonpath, ChrW$(34))
    If Y > 0 Then
    Wagonpath = Left$(Wagonpath, Y - 1)
    End If
    End If
    booFound = True
   End If
   Exit Sub
Errtrap:
Call MsgBox("An Error number " & Err & " occurred in CheckwagonData checking " & strNew _
            & vbCrLf & "Error description: " & Err.Description _
            , vbExclamation, frmGrid)

End Sub

Public Sub CheckEngineData(strNew As String, Engname As String, Engpath As String, booFound As Boolean)
Dim x As Integer, Y As Integer, yy As Integer, Z As Integer, strNew2 As String

On Error GoTo Errtrap
strNew = Trim$(strNew)
Z = InStr(strNew, vbLf)
If Z > 0 Then
strNew = Left$(strNew, Z - 1)
End If
    x = InStr(1, strNew, "EngineData", vbTextCompare)
      If x > 0 Then
      
    Y = InStr(x, strNew, "(")
    
    yy = InStrRev(strNew, ")")
   
   strNew2 = Trim$(Mid$(strNew, Y + 1, yy - (Y + 1)))
   
   strNew2 = Left$(strNew2, Len(strNew2))
   strNew2 = Trim$(strNew2)
   If Left$(strNew2, 1) = ChrW$(34) Then
   x = InStr(2, strNew2, ChrW$(34))
   Else
   x = InStr(strNew2, " ")
   End If
  ' x = InStr(strNew2, " ")
   Engname = Left$(strNew2, x - 1)
   Engpath = Mid$(strNew2, x + 1)
   If Left$(Engname, 1) = ChrW$(34) Then
   Engname = Mid$(Engname, 2)
   Y = InStr(Engname, ChrW$(34))
    If Y > 0 Then
    Engname = Left$(Engname, Y - 1)
    End If
   End If
   Engpath = Trim$(Engpath)
   If Left$(Engpath, 1) = ChrW$(34) Then
   Engpath = Mid$(Engpath, 2)
   Y = InStr(Engpath, ChrW$(34))
    If Y > 0 Then
    Engpath = Left$(Engpath, Y - 1)
    End If
    
   End If
   booFound = True
End If
Exit Sub
Errtrap:
Call MsgBox("An Error number " & Err & " occurred in CheckEngineData checking " & strNew _
            & vbCrLf & "Error description: " & Err.Description _
            , vbExclamation, frmGrid)

End Sub


Public Sub GetSize(strEng As String, sngEngSize As Single)
Dim NewFile As Integer, x As Integer, Y As Integer, xx As Integer, A$
If Not FileExists(strEng) Then
    Call MsgBox(strEng & Lang(353), vbExclamation, App.Title)
    Exit Sub
    End If
NewFile = FreeFile
 Open strEng For Input As #NewFile
 Do While Not EOF(NewFile)
Line Input #NewFile, A$
x = InStr(A$, "Size")
If x > 0 Then
Y = InStr(x, A$, ")")
A$ = Trim$(Mid$(A$, x, Y - x - 1))
xx = InStrRev(A$, " ")
A$ = Mid$(A$, xx + 1)
If Right$(A$, 1) = "m" Then
A$ = Left$(A$, Len(A$) - 1)
End If
sngEngSize = Val(A$)
Exit Do
End If
Loop
End Sub


Public Function QSort2(strList() As String, lLbound As Long, lUbound As Long)
    
    Dim strTemp As String
    Dim strBuffer As String
    Dim lngCurLow As Long
    Dim lngCurHigh As Long
    Dim lngCurMidpoint As Long
    
    lngCurLow = lLbound ' Start current low and high at actual low/high
    lngCurHigh = lUbound
    
    If lUbound <= lLbound Then Exit Function ' Error!
    lngCurMidpoint = (lLbound + lUbound) \ 2 ' Find the approx midpoint of the array
    
    strTemp = strList(lngCurMidpoint) ' Pick as a starting point (we are making
    ' an assumption that the data *might* be
    '
    ' in semi-sorted order already!
    


    Do While (lngCurLow <= lngCurHigh)


        Do While strList(lngCurLow) < strTemp
            lngCurLow = lngCurLow + 1
            If lngCurLow = lUbound Then Exit Do
        Loop
        


        Do While strTemp < strList(lngCurHigh)
            lngCurHigh = lngCurHigh - 1
            If lngCurHigh = lLbound Then Exit Do
        Loop


        If (lngCurLow <= lngCurHigh) Then ' if low is <= high then swap
            strBuffer = strList(lngCurLow)
            strList(lngCurLow) = strList(lngCurHigh)
            strList(lngCurHigh) = strBuffer
            '
            lngCurLow = lngCurLow + 1 ' CurLow++
            lngCurHigh = lngCurHigh - 1 ' CurLow--
        End If
        
    Loop
    


    If lLbound < lngCurHigh Then ' Recurse if necessary
        QSort2 strList(), lLbound, lngCurHigh
    End If
    


    If lngCurLow < lUbound Then ' Recurse if necessary
        QSort2 strList(), lngCurLow, lUbound
    End If
    
End Function

Public Function ReadFile3(CompleteFilePath As String, MyString As String) As String

Dim length As Long, mytristate As Integer

Dim File_obj As Object, The_obj As Object, fileflag As Boolean


Set File_obj = CreateObject("Scripting.FileSystemObject")
If Not File_obj.FileExists(CompleteFilePath) Then
MyString = ""
Exit Function
End If
length = FileLen(CompleteFilePath)

If Right(CompleteFilePath, 3) = "mkr" Then
mytristate = -1
Else
mytristate = 0
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll
The_obj.Close
fileflag = False

End Function

