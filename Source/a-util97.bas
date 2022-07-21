Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text
Public foundpos As Long

Public newpos As Long
Public options1 As Integer
Public options2 As Integer
Public newpos1 As Long
Public ftp As Integer
Public starting As Boolean

Public Type gedperson
    Name As String
    childof As Long
    spouse(7) As Long
End Type
Public Type family
    husband As Long
    wife As Long
    child(15) As Long
End Type


Public tit As String


Public cursouind As Integer
Public curtarind As Integer
Public fullpath$

Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long

Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long

Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" Alias "CreateToolhelp32Snapshot" ( _
ByVal lFlags As Long, lProcessID As Long) As Long

Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, _
ByVal uExitCode As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long



'-------------------------------------------------------
Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, _
ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

'Global vars
'


'

'
#If Win32 And LOGGING Then
Global Const resLOG_FILEUPTODATE = 2000
Global Const resLOG_FILECOPIED = 2001
Global Const resLOG_ERROR = 2002
Global Const resLOG_WARNING = 2003
Global Const resLOG_DURINGACTION = 2004
Global Const resLOG_CANNOTWRITE = 2005
Global Const resLOG_CANNOTCREATE = 2006
Global Const resLOG_DONOTMODIFY = 2007
Global Const resLOG_FILECONTENTS = 2008
Global Const resLOG_FILEUSEDFOR = 2009
Global Const resLOG_USERRESPONDEDWITH = 2012
Global Const resLOG_CANTRUNAPPREMOVER = 2013
Global Const resLOG_ABOUTTOREMOVEAPP = 2014

Global Const resLOG_IDOK = 2100
Global Const resLOG_IDCANCEL = 2101
Global Const resLOG_IDABORT = 2102
Global Const resLOG_IDRETRY = 2103
Global Const resLOG_IDIGNORE = 2104
Global Const resLOG_IDYES = 2105
Global Const resLOG_IDNO = 2106
Global Const resLOG_IDUNKNOWN = 2107
#End If




'Type Definitions

Type RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

'Global Variables
'
                             'name of app being installed
Global gstrTitle As String                                  '"setup" name of app being installed

#If Win32 And LOGGING Then
Global gstrAppRemovalLog As String                           'name of the app removal logfile
Global gstrAppRemovalEXE As String                           'name of the app removal executable
Global gfAppRemovalFilesMoved As Boolean                     'whether or not the app removal files have been moved to the application directory
#End If

'
'Form/Module Constants
'


'Special file names
#If Win16 Then

Global Const mstrFILE_RPCREG$ = "RPCREG.DAT"
#End If
#If Win32 And LOGGING Then

#End If




Global Const resDISKSPCERR% = 600




'Mouse Pointer Constants
Global Const gintMOUSE_DEFAULT% = 0


'MsgError() Constants


'MsgBox Constants
                    'Warning query
Global Const MB_ICONEXCLAMATION = 48                    'Warning message

'
'Type Definitions
'





                                'single line break
Global LS$                                              'double line break



Rem *****************
'API/DLL Declarations for 32 bit SetupToolkit
'
Declare Function DiskSpaceFree Lib "STKIT432.DLL" Alias "DISKSPACEFREE" () As Long

  






   

  
  
'Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Public pardatdev As String

Global tarname$, expr1&, expr2&, dist%

Global east%, north%
Global Const gstrSEP_DIR = "\"
Global Const gstrCOLON = ":"


' Data Access constants
'

'

'Database constants
Public Const MAX_FIELDS = 100


'Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long
Public Const mci_ovly_where_source = &H20000
Public Const mci_where = &H843

Public strThisPer As String

Public showflag As Boolean
Public printflag As Boolean
Public exportflag As Boolean



Public Type person2
    PID As Long
    AltId As Long
    sex As String * 1
    birthdate As String * 15
    deathdate As String * 15
    forenames As String * 50
    surname As String * 30
    birthplace As String * 50
    deathplace As String * 50
    childof As Long
End Type
Public Type family2
    FamID As Long
    AltId As Long
    husband As Long
    wife As Long
    date As String * 15
    place As String * 100
End Type
    
Public perno As Long, docno As Long

Public fullname As String
Public selper As Integer, seldoc As Boolean
Public datpath As String


Public datdev As String

Public Type BROWSEINFO
     hwndOwner As Long
     pidlRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long



Public Function ConvertAI(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertAI = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Long
Dim xx As Long, Y As Long, B$, c$, booEndLoop As Boolean, d$, strTemp As String

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
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If

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
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
MyString = Replace(MyString, vbTab, " ")
MyString = Replace(MyString, "          ", " ")
MyString = Replace(MyString, "         ", " ")
MyString = Replace(MyString, "        ", " ")
MyString = Replace(MyString, "       ", " ")
MyString = Replace(MyString, "      ", " ")
MyString = Replace(MyString, "     ", " ")
MyString = Replace(MyString, "    ", " ")
MyString = Replace(MyString, "   ", " ")
MyString = Replace(MyString, "  ", " ")
MyString = Replace(MyString, " ", " ")
DoEvents

x = 1
GetComment:
x = InStr(x, MyString, "Comment(")
If x > 0 Then
xx = InStr(x, MyString, vbCr)
If Mid$(MyString, xx - 1, 1) = ")" Then
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strEnd
x = x + 1
GoTo GetComment
End If
End If

x = InStr(MyString, "Wagon (")
If x > 0 Then
strStart = Left$(MyString, x + 6)
strEnd = Trim$(Mid$(MyString, x + 7))
If Left$(strEnd, 1) = ChrW$(34) Then
xx = InStr(2, strEnd, ChrW$(34))
strTemp = Mid$(strEnd, 2, xx - 2)
strEnd = Mid$(strEnd, xx + 1)
x = Len(strStart) + 10
End If
If booMU = False Then
MyString = strStart & " #" & strTemp & strEnd
Else
MyString = strStart & " MU_" & strTemp & strEnd
End If
strTemp = vbNullString
End If

xx = InStr(x + 10, MyString, "Wagon (")
If xx > 0 Then
strStart = Left$(MyString, xx + 6)
strEnd = Trim$(Mid$(MyString, xx + 7))
If Left$(strEnd, 1) = vbTab Then
strEnd = Mid$(strEnd, 2)
End If
If booMU = False Then
MyString = strStart & " #" & strEnd
Else
MyString = strStart & " MU_" & strEnd
End If

End If
x = InStr(MyString, "Engine (")
If x > 0 Then
strStart = Left$(MyString, x + 7)
strEnd = Trim$(Mid$(MyString, x + 8))
If Left$(strEnd, 1) = vbTab Then
strEnd = Mid$(strEnd, 2)
End If
If booMU = False Then
MyString = strStart & " #" & strEnd
Else
MyString = strStart & " MU_" & strEnd
End If
End If
'x = 1
MoreSounds:


Y = InStr(x, MyString, "Sound (")
If Y > 0 Then
xx = InStr(Y, MyString, ")")
strStart = Left$(MyString, Y - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strEnd
x = xx + 1
GoTo MoreSounds
End If

x = InStr(MyString, "Name (")
If x > 0 Then
xx = InStr(x, MyString, ChrW$(34))
strStart = Left$(MyString, xx)
strEnd = Trim$(Mid$(MyString, xx + 1))
If booMU = False Then
MyString = strStart & "#" & strEnd
Else
MyString = strStart & "MU_" & strEnd
End If
End If
x = InStr(MyString, "Cabview (")
If x > 0 Then
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strEnd
End If
x = InStr(MyString, "Headout (")
If x > 0 Then
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strEnd
End If
x = InStr(MyString, "Antislip (")
If x > 0 Then
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & "Antislip ( 1 )" & vbCrLf & strEnd
GoTo MissAnti
End If
x = InStr(MyString, "Adheasion")
xx = InStr(x, MyString, "(")
If xx - x > 18 Then GoTo MissAnti
If x > 0 Then
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, x)
MyString = strStart & "Antislip ( 1 )" & vbCrLf & strEnd
End If
MissAnti:
If booMU = True Then GoTo BitMore
If booAI = False Then
x = InStr(MyString, "WheelRadius (")
    If x > 0 Then
    Y = InStr(x, MyString, ")")
    B$ = Trim$(Mid$(MyString, x + 13, Y - (x + 13)))
    Call GetRadius(B$)
    strStart = Left$(MyString, x + 12)
    strEnd = Mid$(MyString, Y)
    MyString = strStart & " " & B$ & " " & strEnd
    End If
x = InStr(Y, MyString, "WheelRadius (")
    If x > 0 Then
    Y = InStr(x, MyString, ")")
    B$ = Trim$(Mid$(MyString, x + 13, Y - (x + 13)))
    Call GetRadius(B$)
    strStart = Left$(MyString, x + 12)
    strEnd = Mid$(MyString, Y)
    MyString = strStart & " " & B$ & " " & strEnd
    End If
End If
BitMore:
x = InStr(MyString, "Inside (")
If x > 0 Then
Y = x
Do

xx = InStr(Y, MyString, ")")
B$ = Mid$(MyString, xx - 1, 1)
c$ = Mid$(MyString, xx - 2, 1)
d$ = Mid$(MyString, xx - 3, 1)
If Asc(B$) > 32 Or Asc(c$) > 32 Or Asc(d$) > 32 Then

Y = Y + 1
Else
booEndLoop = True
End If


Loop While booEndLoop = False

strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strEnd
End If
x = InStr(MyString, "Comment")
If x > 0 Then
xx = InStr(x, MyString, "(")
Y = InStr(xx, MyString, ")")
If Y - xx < 4 Then
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, Y + 1)
MyString = strStart & strEnd
End If
End If
CarryON:
End If
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
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertAI = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function


Private Sub GetRadius(strRadius As String)
Dim TwoPI As Double, MyRad As Double, i As Integer, strTemp As String
Dim strTemp2 As String, strMyRad As String

TwoPI = 3.14159265358979 * 2
For i = Len(strRadius) To 1 Step -1
If IsNumeric(Mid$(strRadius, i, 1)) Then
strTemp = Left$(strRadius, i)
strTemp2 = Mid$(strRadius, i + 1)
Exit For
End If
Next i
MyRad = Val(strTemp)
MyRad = MyRad * TwoPI
strMyRad = Str(MyRad)
strMyRad = Trim$(strMyRad)
If Len(strMyRad) > 6 Then
strMyRad = Left$(strMyRad, 6)
strRadius = strMyRad & strTemp2
End If


End Sub




Public Function ConvertDummy2(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertDummy2 = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Long
Dim xx As Long, Y As Long, yy As Long
Dim strEngname As String, strTemp As String, y1 As Long, y2 As Long

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
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
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
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
MyString = Replace(MyString, vbTab, " ")
DoEvents
MyString = Replace(MyString, "Wagon(", "Wagon (")
DoEvents
MyString = Replace(MyString, "Engine(", "Engine (")
DoEvents
MyString = Replace(MyString, "Wagon (", "Wagon ( ")
DoEvents
MyString = Replace(MyString, "Engine (", "Engine ( ")
DoEvents
MyString = Replace(MyString, "  ", " ")
DoEvents
x = InStr(MyString, "Wagon (")
xx = InStr(x, MyString, vbCr)
'On Error GoTo Err2

strEngname = Mid$(MyString, x + 7, xx - (x + 7))
strEngname = Trim$(strEngname)
strStart = Left$(MyString, x + 6)
strEnd = Trim$(Mid$(MyString, x + 7))
If Left$(strEnd, 1) = ChrW$(34) Then
strEnd = Mid$(strEnd, 2)
MyString = strStart & " " & ChrW$(34) & "$" & strEnd
Else
MyString = strStart & " $" & strEnd
End If
x = x + 100
Label2:
If Left$(strEngname, 1) = ChrW$(34) Then
strEngname = Mid$(strEngname, 2)
End If
If Right$(strEngname, 1) = ChrW$(34) Then
strEngname = Left$(strEngname, Len(strEngname) - 1)
End If
xx = InStr(x, MyString, " " & strEngname)

If xx > 0 Then
strStart = Left$(MyString, xx)
strEnd = Mid$(MyString, xx + 1)
Y = InStrRev(MyString, vbCr, xx)
strTemp = Mid$(MyString, Y, xx - Y)
y1 = InStr(strTemp, "Name")
y2 = InStr(strTemp, "CouplingUniqueType")
yy = InStr(strTemp, "Comment")
If y1 > 0 Or y2 > 0 Or yy > 0 Then GoTo NotNeeded
If Mid$(strEnd, Len(strEngname) + 1, 4) <> ".cvf" And Mid$(strEnd, Len(strEngname) + 1, 1) <> ChrW$(34) And Mid$(strEnd, Len(strEngname) + 1, 1) <> "." Then
MyString = strStart & "$" & strEnd
End If
NotNeeded:
x = xx + 10
GoTo Label2
End If
xx = 1
x = InStr(xx, MyString, "Lights (")
yy = InStr(xx, MyString, "Engine (")

    If x < yy And x <> 0 Then
    strStart = Left$(MyString, x - 1)
    strEnd = vbCrLf & ")" & vbCrLf & Mid$(MyString, yy)
    MyString = strStart & strEnd
    End If
    
x = InStr(MyString, "Name (")
If x > 0 Then

    
    Rem ************ Find Engine Name *********
    Y = InStr(x, MyString, ChrW$(34))
        If Y < x + 12 Then
        strStart = Left$(MyString, Y)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & "$" & strEnd
        Else
        strStart = Left$(MyString, x + 5)
        strEnd = Mid$(MyString, x + 6)
            If Left$(strEnd, 1) <> " " Then
            strEnd = " " & strEnd
            End If
        MyString = strStart & " $" & strEnd
        End If
    
    End If
x = InStr(MyString, "Description")
If x > 0 Then
strStart = Left$(MyString, x - 1)
MyString = strStart & vbCrLf & ")" & vbCrLf
End If
x = InStr(MyString, "MaxPower")
If x = 0 Then GoTo Label3
Y = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, Y)
MyString = strStart & "MaxPower ( 0 " & strEnd
x = InStr(MyString, "MaxForce")
Y = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, Y)
MyString = strStart & "MaxForce ( 0 " & strEnd
x = InStr(MyString, "SteamSpecialEffects")
If x = 0 Then GoTo Label3
Y = InStr(x, MyString, "Wagon")
If Y = 0 Then GoTo Label3
strStart = Left$(MyString, x + 18)
strEnd = Mid$(MyString, Y)
MyString = strStart & vbCrLf & "(" & vbCrLf & ")" & vbCrLf & ")" & vbCrLf & strEnd
xx = 1
Label3:
x = InStr(xx, MyString, "Sound")
If x > 0 Then
Y = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, Y + 1)
MyString = strStart & "Sound ( " & ChrW$(34) & "GenFreightWag2.sms" & ChrW$(34) & " )" & strEnd
xx = x + 20
GoTo Label3
End If
xx = 1
Label4:
x = InStr(xx, MyString, "FreightAnim")
If x > 0 Then
Y = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, Y + 1)
MyString = strStart & strEnd
xx = x + 20
GoTo Label4
End If
x = InStr(MyString, "CabView")
If x > 0 Then
Y = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, Y + 1)
MyString = strStart & strEnd
End If
End If

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
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertDummy2 = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function
Public Function ConvertDummy(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertDummy = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Long
Dim xx As Long, yy As Long

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
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If

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
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
MyString = Replace(MyString, vbTab, " ")
MyString = Replace(MyString, "Wagon(", "Wagon (")
MyString = Replace(MyString, ")", " )")
DoEvents
MyString = Replace(MyString, ChrW$(34) & "Engine" & ChrW$(34), " Engine ")
DoEvents
MyString = Replace(MyString, "Type(", "Type ( ")
DoEvents
MyString = Replace(MyString, "  ", " ")
DoEvents
x = 1
GetComment:
x = InStr(x, MyString, "Comment (*")
If x > 0 Then
xx = InStr(x, MyString, "*)")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 2)
MyString = strStart & strEnd
x = xx + 2
GoTo GetComment
End If


x = InStr(MyString, "Wagon (")
If x > 0 Then
strStart = Left$(MyString, x + 6)
strEnd = Trim$(Mid$(MyString, x + 7))
If Left$(strEnd, 1) = ChrW$(34) Then
strEnd = Mid$(strEnd, 2)
MyString = strStart & " " & ChrW$(34) & "$" & strEnd
Else
MyString = strStart & " $" & strEnd
End If
xx = InStr(x + 10, MyString, "( Engine )")
If xx > 0 Then
strStart = Left$(MyString, xx - 1)
strEnd = Trim$(Mid$(MyString, xx + 10))
MyString = strStart & "( Carriage )" & strEnd
End If
x = InStr(xx, MyString, "Lights")
If x = 0 Then GoTo CarryON
If Mid$(MyString, x - 4, 4) = "head" Then
x = InStr(x + 5, MyString, "Lights")
End If
If x = 0 Then GoTo CarryON
yy = InStr(x, MyString, "Sound")
If yy = 0 Then GoTo CarryON
If x < yy Then
strStart = Left$(MyString, x - 1)
strEnd = vbCrLf & ")" & vbCrLf
ElseIf yy < x Then
strStart = Left$(MyString, yy - 1)
strEnd = vbCrLf & ")" & vbCrLf & ")" & vbCrLf
End If
GoTo FinishIt
CarryON:
x = InStr(xx, MyString, "Engine (")
strStart = Left$(MyString, x - 1)
strEnd = vbCrLf & ")" & vbCrLf

FinishIt:
MyString = strStart & "Sound ( " & ChrW$(34) & "GenFreightWag1.sms" & ChrW$(34) & " )" & strEnd

End If

End If
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
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertDummy = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function









  
Public Sub CheckWagPath(Wagname As String, Wagonpath As String, strWag As String, strPath As String)
Dim i As Long

For i = 0 To lngWagons - 1
If Wagname = Wagons(i) Then

strWag = Wagons(i)
Exit For
End If
Next i
For i = 0 To lngWagons - 1
If Wagonpath = Wagpath(i) Then
strPath = Wagpath(i)
Exit For
End If
Next i

End Sub
Public Sub CheckNamePath(Engname As String, Engpath As String, strEng As String, strPath As String)
Dim i As Long

For i = 0 To lngLoco - 1
If Engname = Locomotives(i) Then
strEng = Locomotives(i)
Exit For
End If
Next i
For i = 0 To lngLoco - 1
If Engpath = LocoPath(i) Then
strPath = LocoPath(i)
Exit For
End If
Next i

End Sub

 Sub QSort(sortme() As String, lowbound As Long, hibound As Long)
    Dim low As Long
    Dim high As Long
    Dim midval As String
    Dim Temp As String
      
    low = lowbound
    high = hibound
    midval = sortme((low + high) / 2)

    While (low <= high)
        While (sortme(low) < midval And _
            low < hibound)
        low = low + 1
        Wend
        While (midval < sortme(high) _
            And high > lowbound)
            high = high - 1
        Wend
        If (low <= high) Then
            Temp = sortme(low)
            sortme(low) = sortme(high)
            sortme(high) = Temp
            low = low + 1
            high = high - 1
        End If
    Wend

    If (lowbound < high) Then
        QSort sortme(), lowbound, high
    End If
    If (low < hibound) Then
        QSort sortme(), low, hibound
    End If

End Sub









Function FileExists(FullFilename As String) As Boolean
Dim NewFile As Integer

   
   On Error GoTo MakeF
   NewFile = FreeFile
        'If file does Not exist, there will be an Error
        Open FullFilename For Input As #NewFile
        Close #NewFile
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function
Function DirExists(Path As String) As Boolean
    On Error Resume Next
    If Right$(Path, 1) = "\" Then
    Path = Left$(Path, Len(Path) - 1)
    End If
    DirExists = (Dir$(Path & "\nul") <> vbNullString)
End Function


'-----------------------------------------------------------
' FUNCTION: ResolveResString
' Reads resource and replaces given macros with given values
'
' Example, given a resource number 14:
'    "Could not read '|1' in drive |2"
'   The call
'     ResolveResString(14, "|1", "TXTFILE.TXT", "|2", "A:")
'   would return the string
'     "Could not read 'TXTFILE.TXT' in drive A:"
'
' IN: [resID] - resource identifier
'     [varReplacements] - pairs of macro/replacement value
'-----------------------------------------------------------
'
Function ResolveResString(ByVal resID As Integer, ParamArray varReplacements() As Variant) As String
    Dim intMacro As Integer
    Dim strResString As String
    
    strResString = LoadResString(resID)
    
    ' For each macro/value pair passed in...
    For intMacro = LBound(varReplacements) To UBound(varReplacements) Step 2
        Dim strMacro As String
        Dim strValue As String
        
        strMacro = varReplacements(intMacro)
        On Error GoTo MismatchedPairs
        strValue = varReplacements(intMacro + 1)
        On Error GoTo 0
        
        ' Replace all occurrences of strMacro with strValue
        Dim intPos As Integer
        Do
            intPos = InStr(strResString, strMacro)
            If intPos > 0 Then
                strResString = Left$(strResString, intPos - 1) & strValue & Right$(strResString, Len(strResString) - Len(strMacro) - intPos + 1)
            End If
        Loop Until intPos = 0
    Next intMacro
    
    ResolveResString = strResString
    
    Exit Function
    
MismatchedPairs:
    Resume Next
End Function

 Public Sub KillProcess(NameProcess As String)
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const TH32CS_SNAPPROCESS As Long = 2&
Dim uProcess  As PROCESSENTRY32
Dim RProcessFound As Long
Dim hSnapshot As Long
Dim SzExename As String
Dim ExitCode As Long
Dim MyProcess As Long
Dim AppKill As Boolean
Dim AppCount As Integer
Dim i As Integer
Dim WinDirEnv As String
        
       If NameProcess <> "" Then
          AppCount = 0

          uProcess.dwSize = Len(uProcess)
          hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
          RProcessFound = ProcessFirst(hSnapshot, uProcess)
  
          Do
            i = InStr(1, uProcess.szexeFile, Chr(0))
            SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
            WinDirEnv = Environ("Windir") + "\"
            WinDirEnv = LCase$(WinDirEnv)
        
            If Right$(SzExename, Len(NameProcess)) = LCase$(NameProcess) Then
               AppCount = AppCount + 1
               MyProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
               AppKill = TerminateProcess(MyProcess, ExitCode)
               Call CloseHandle(MyProcess)
            End If
            RProcessFound = ProcessNext(hSnapshot, uProcess)
          Loop While RProcessFound
          Call CloseHandle(hSnapshot)
       End If

End Sub

 '-----------------------------------------------------------
' FUNCTION: GetDiskSpaceFree
' Get the amount of free disk space for the specified drive
'
' IN: [strDrive] - drive to check space for
'
' Returns: Amount of free disk space, or -1 if an error occurs
'-----------------------------------------------------------
'
Function GetDiskSpaceFree(ByVal strDrive As String) As Long
    Dim strCurDrive As String
    Dim lDiskFree As Long

    On Error Resume Next
    '
    'Save the current drive
    '
    strCurDrive = Left$(CurDir$, 2)

    '
    'Fixup drive so it includes only a drive letter and a colon
    '
    If InStr(strDrive, gstrCOLON) = 0 Or Len(strDrive) > 2 Then
        strDrive = Left$(strDrive, 1) & gstrCOLON
    End If

    '
    'Change to the drive we want to check space for.  The DiskSpaceFree() API
    'works on the current drive only.
    '
    ChDrive strDrive

    '
    'If we couldn't change to the request drive, it's an error, otherwise return
    'the amount of disk space free
    '
    If Err <> 0 Or (strDrive <> Left$(CurDir$, 2)) Then
        lDiskFree = -1
    Else

 lDiskFree = DiskSpaceFree()
        If Err <> 0 Then    'If Setup Toolkit's DLL couldn't be found
            lDiskFree = -1
        End If
    End If

    If lDiskFree = -1 Then
        MsgError Error$ & LS$ & ResolveResString(resDISKSPCERR) & strDrive, MB_ICONEXCLAMATION, gstrTitle
    End If

    GetDiskSpaceFree = lDiskFree

    '
    'Cleanup by setting the current drive back to the original
    '
    ChDrive strCurDrive

    Err = 0
End Function





'-----------------------------------------------------------
' FUNCTION: MsgError
'
' Forces mouse pointer to default, calls VB's MsgBox
' function, and logs this error and (32-bit only)
' writes the message and the user's response to the
' logfile (32-bit only)
'
' IN: [strMsg] - message to display
'     [intFlags] - MsgBox function type flags
'     [strCaption] - caption to use for message box
'     [intLogType] (optional) - The type of logfile entry to make.
'                   By default, creates an error entry.  Use
'                   the MsgWarning() function to create a warning.
'                   Valid types as MSGERR_ERROR and MSGERR_WARNING
'
' Returns: Result of MsgBox function
'-----------------------------------------------------------
'
Function MsgError(ByVal strmsg As String, ByVal intFlags As Integer, ByVal strCaption As String, Optional ByVal intLogType As Variant) As Integer
    Dim iRet As Integer
    
    iRet = MsgFunc(strmsg, intFlags, strCaption)
    MsgError = iRet
    
    #If Win32 And LOGGING Then
        ' We need to log this error and decode the user's response.
        Dim strID As String
        Dim strLogMsg As String

        Select Case iRet
        Case IDOK
            strID = ResolveResString(resLOG_IDOK)
        Case IDCANCEL
            strID = ResolveResString(resLOG_IDCANCEL)
        Case IDABORT
            strID = ResolveResString(resLOG_IDABORT)
        Case IDRETRY
            strID = ResolveResString(resLOG_IDRETRY)
        Case IDIGNORE
            strID = ResolveResString(resLOG_IDIGNORE)
        Case IDYES
            strID = ResolveResString(resLOG_IDYES)
        Case IDNO
            strID = ResolveResString(resLOG_IDNO)
        Case Else
            strID = ResolveResString(resLOG_IDUNKNOWN)
        End Select

        strLogMsg = strmsg & LF$ & "(" & ResolveResString(resLOG_USERRESPONDEDWITH, "|1", strID) & ")"
        If IsMissing(intLogType) Then
            intLogType = MSGERR_ERROR
        End If
        Select Case intLogType
        Case MSGERR_WARNING
            LogWarning strLogMsg
        Case MSGERR_ERROR
            LogError strLogMsg
        Case Else
            LogError strLogMsg
        End Select
    #End If
End Function

'-----------------------------------------------------------
' FUNCTION: MsgFunc
'
' Forces mouse pointer to default and calls VB's MsgBox
' function
'
' IN: [strMsg] - message to display
'     [intFlags] - MsgBox function type flags
'     [strCaption] - caption to use for message box
'     [fLogAsError] - If present and True (MSGBOX_ERR), the 32-bit
'                       version logs this message and the user's
'                       response in the logfile as an error.
'                       Otherwise it is presented to the user
'                       only.  (It is easier to use the MsgError()
'                       function.)
' Returns: Result of MsgBox function
'-----------------------------------------------------------
'
Function MsgFunc(ByVal strmsg As String, ByVal intFlags As Integer, ByVal strCaption As String) As Integer
    Dim intOldPointer As Integer
  
    intOldPointer = Screen.MousePointer

    Screen.MousePointer = gintMOUSE_DEFAULT
    MsgFunc = MsgBox(strmsg, intFlags, strCaption)
    Screen.MousePointer = intOldPointer
End Function

