VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmRepStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replace Rolling-Stock Item:-"
   ClientHeight    =   8250
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Substitute this item in ALL consists and activities ?"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   7680
      Width           =   975
   End
   Begin VSFlex8LCtl.VSFlexGrid GridRepCon 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   9015
      _cx             =   15901
      _cy             =   11880
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRepStock.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   8040
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   7680
      Width           =   3015
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmRepStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim MyRow As Integer, MyCol As Integer
Private Sub GetCoupling(LocoPath As String, flagCoup As Integer, strBrake As String, strType As String, strName As String)
Dim NewFile As Integer, A$, x As Long, xx As Long, strTender As String
On Error GoTo Errtrap

NewFile = FreeFile
Open LocoPath For Input As #NewFile
Do While Not EOF(NewFile)
 Line Input #NewFile, A$
 
 x = InStr(A$, "Type ( Engine )")
 If x > 0 Then strType = "Engine"
 x = InStr(A$, "Type ( Freight )")
 If x > 0 Then strType = "Freight"
 x = InStr(A$, "Type ( Tender )")
 If x > 0 Then strType = "Tender"
 x = InStr(A$, "Type ( Carriage )")
 If x > 0 Then strType = "Carriage"
 x = InStr(A$, "IsTenderRequired")
If x > 0 Then
xx = InStr(x, A$, "(")
xy = InStr(xx, A$, ")")
strTender = Trim$(Mid$(A$, xx + 1, xy - xx - 1))
If strTender = "1" Then
strType = "Engine *"
End If
End If

x = InStr(A$, "Type ( Automatic )")
xx = InStr(A$, "Type ( Chain )")
If x > 0 Then
flagCoup = 1
ElseIf xx > 0 Then
flagCoup = 2
End If

x = InStr(A$, "brakesystemtype")
If x > 0 Then

xx = InStr(x, A$, "(")
xy = InStr(xx, A$, ")")
strBrake = Mid$(A$, xx + 1, xy - xx - 1)
strBrake = Trim$(strBrake)
If Left$(strBrake, 1) = ChrW$(34) Then
strBrake = Mid$(strBrake, 2, Len(strBrake) - 2)
End If
End If

x = InStr(A$, "Name (")
If x > 0 Then
xx = InStr(x, A$, "(")
xy = InStr(xx, A$, ")")
strName = Mid$(A$, xx + 1, xy - xx - 1)
strName = Trim$(strName)
If Left$(strName, 1) = ChrW$(34) Then
strName = Mid$(strName, 2, Len(strName) - 2)
End If
End If
Loop
Close NewFile




Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'GetCoupling' please advise" _
            & vbCrLf & "while checking " & LocoPath _
            , vbExclamation, App.Title)
'Resume Next

End Sub

Private Function ConvertStock(CompleteFilePath As String, flagway As Integer, strOldCon As String, strNewCon As String, strStockType As String) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertStock = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean, xx As Long
Dim strStart As String, strEnd As String, Y As Long, Z As Long
Dim strSearch As String, x As Long
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
If strStockType = "eng" Then
strSearch = "EngineData"
Else
strSearch = "WagonData"
End If
If flagway = 0 Then

Z = InStr(MyString, strSearch)


TryAgain:
x = InStr(Z, MyString, strOldCon)

If x = 0 Then GoTo AllFound
If x - Z > 25 Then
Z = InStr(Z + 5, MyString, strSearch)
If Z = 0 Then GoTo AllFound
GoTo TryAgain
End If
strStart = Left$(MyString, x - 1)
xx = InStrRev(strStart, "(")
strStart = Left$(strStart, xx)
strEnd = Mid$(MyString, x + Len(strOldCon))
Y = InStr(strEnd, vbCr)
strEnd = Mid$(strEnd, Y - 1)
MyString = strStart & " " & strNewCon & " " & strEnd
Z = x + Len(strNewCon)
GoTo TryAgain
AllFound:
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
ConvertStock = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function



Private Sub Command1_Click()
Dim flagway As Integer, newRow As Integer, newCol As Integer
Dim strTempCon As String, strStockType As String
Dim strOldStock As String, strExisting As String, strAct As String
Dim strGridPath As String
On Error GoTo Errtrap
Command1.Enabled = False

If flagGrid = 2 And Check1.value = 0 Then
Rem ********* Change rolling stock item ************
newRow = GridRepCon.row
newCol = GridRepCon.col
strOldCon = Left$(Label2(0).Caption, Len(Label2(0).Caption) - 4)
strStockType = Right$(Label2(0).Caption, 3)
strNewCon = GridRepCon.Cell(flexcpText)
strNewCon = Left$(strNewCon, Len(strNewCon) - 4)
strTempCon = strNewCon
strNewCon = ChrW$(34) & strNewCon & ChrW$(34)
GridRepCon.Select newRow, newCol + 1
strNewPath = GridRepCon.Cell(flexcpText)
strGridPath = Trim(strNewPath)
If Left(strGridPath, 1) = ChrW$(34) Then
strGridPath = Mid(strGridPath, 2)
End If
If Right(strGridPath, 1) = ChrW$(34) Then
strGridPath = Left(strGridPath, Len(strGridPath) - 1)
End If
strNewPath = ChrW$(34) & strNewPath & ChrW$(34)
strNewCon = strNewCon & " " & strNewPath
Rem ****************

frmStock.GridStock.Select MyRow, 0
strCon = frmStock.GridStock.Cell(flexcpText)
frmStock.GridStock.Select MyRow, 0
If Right$(strCon, 3) <> "con" Then

      Call MsgBox("This option only works for Consists," _
                  & vbCrLf & "           not for Activities." _
                  , vbExclamation, App.Title)
Exit Sub
End If
'strRoute = frmGrid.Grid1.Cell(flexcpText)
'strRoute = MSTSPath & "\Routes\" & strRoute
FileCopy MSTSPath & "\trains\consists\" & strCon, App.Path & "\TempFiles\" & strCon
DoEvents

flagway = 0    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)

Kill MSTSPath & "\trains\consists\" & strCon
flagway = 1    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)
FileCopy App.Path & "\TempFiles\" & strCon, MSTSPath & "\trains\consists\" & strCon
Kill App.Path & "\TempFiles\" & strCon
Label3.Caption = "Modified Consist Activated"
FlagColGreen = True
frmStock.GridStock.Select MyRow, MyCol

frmStock.GridStock.Cell(flexcpText) = strTempCon & "." & strStockType
frmStock.GridStock.Select MyRow, 2
frmStock.GridStock.Cell(flexcpText) = strGridPath
FlagColGreen = False
Rem ***************** Activity Change *************

ElseIf flagGrid = 3 Then

newRow = GridRepCon.row
newCol = GridRepCon.col
strOldCon = Left$(Label2(0).Caption, Len(Label2(0).Caption) - 4)
strStockType = Right$(Label2(0).Caption, 3)
strNewCon = GridRepCon.Cell(flexcpText)
strTempCon = strNewCon
strNewCon = Left$(strNewCon, Len(strNewCon) - 4)

strNewCon = ChrW$(34) & strNewCon & ChrW$(34)
GridRepCon.Select newRow, newCol + 1
strNewPath = GridRepCon.Cell(flexcpText)
strNewPath = ChrW$(34) & strNewPath & ChrW$(34)
strNewCon = strNewCon & " " & strNewPath
frmGrid.Grid1.Select MyRow, 1
strCon = frmGrid.Grid1.Cell(flexcpText)
frmGrid.Grid1.Select MyRow, 0
strNewRoute = frmGrid.Grid1.Cell(flexcpText)
RoutePath = MSTSPath & "\Routes\"
FileCopy RoutePath & strNewRoute & "\Activities\" & strCon, App.Path & "\TempFiles\" & strCon
DoEvents
flagway = 0    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)
DoEvents
Kill RoutePath & strNewRoute & "\Activities\" & strCon
DoEvents
flagway = 1    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)
FileCopy App.Path & "\TempFiles\" & strCon, RoutePath & strNewRoute & "\Activities\" & strCon
DoEvents
Kill App.Path & "\TempFiles\" & strCon
Label3.Caption = "Activity Modified"
frmGrid.Grid1.Select MyRow, MyCol
frmGrid.Grid2.Select ThisRow, 0
frmGrid.Grid2.FillStyle = flexFillSingle
frmGrid.Grid2.CellBackColor = vbWhite
frmGrid.Grid2.Cell(flexcpText) = strTempCon
ElseIf flagGrid = 4 Then

newRow = GridRepCon.row
newCol = GridRepCon.col
strOldCon = Left$(Label2(0).Caption, Len(Label2(0).Caption) - 4)
strStockType = Right$(Label2(0).Caption, 3)
strNewCon = GridRepCon.Cell(flexcpText)
strNewCon = Left$(strNewCon, Len(strNewCon) - 4)
strTempCon = strNewCon
strNewCon = ChrW$(34) & strNewCon & ChrW$(34)
GridRepCon.Select newRow, newCol + 1
strNewPath = GridRepCon.Cell(flexcpText)
strNewPath = ChrW$(34) & strNewPath & ChrW$(34)
strNewCon = strNewCon & " " & strNewPath
frmGrid.Grid1.Select MyRow, 4
strCon = frmGrid.Grid1.Cell(flexcpText)
frmGrid.Grid1.Select MyRow, 0

FileCopy MSTSPath & "\trains\consists\" & strCon, App.Path & "\TempFiles\" & strCon
DoEvents

flagway = 0    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)

Kill MSTSPath & "\trains\consists\" & strCon
flagway = 1    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)
FileCopy App.Path & "\TempFiles\" & strCon, MSTSPath & "\trains\consists\" & strCon
Kill App.Path & "\TempFiles\" & strCon
Label3.Caption = Lang(548)
frmGrid.Grid1.Select MyRow, MyCol
frmGrid.Grid1.FillStyle = flexFillSingle
frmGrid.Grid1.CellBackColor = vbWhite

'frmGrid.Grid1.Cell(flexcpText) = strTempCon & ".con"
ElseIf flagGrid = 5 Then

newRow = GridRepCon.row
newCol = GridRepCon.col
strOldCon = Left$(Label2(0).Caption, Len(Label2(0).Caption) - 4)
strStockType = Right$(Label2(0).Caption, 3)
strNewCon = GridRepCon.Cell(flexcpText)
strNewCon = Left$(strNewCon, Len(strNewCon) - 4)
strTempCon = strNewCon
strNewCon = ChrW$(34) & strNewCon & ChrW$(34)
GridRepCon.Select newRow, newCol + 1
strNewPath = GridRepCon.Cell(flexcpText)
strNewPath = ChrW$(34) & strNewPath & ChrW$(34)
strNewCon = strNewCon & " " & strNewPath
frmUnusedSrv.GridCon.Select MyRow, 0
strCon = frmUnusedSrv.GridCon.Cell(flexcpText)
frmUnusedSrv.GridCon.Select MyRow, 0

FileCopy MSTSPath & "\trains\consists\" & strCon, App.Path & "\TempFiles\" & strCon
DoEvents

flagway = 0    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)

Kill MSTSPath & "\trains\consists\" & strCon
flagway = 1    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)
FileCopy App.Path & "\TempFiles\" & strCon, MSTSPath & "\trains\consists\" & strCon
Kill App.Path & "\TempFiles\" & strCon
Label3.Caption = Lang(355)
frmUnusedSrv.GridCon.Select MyRow, MyCol

Rem *************
ElseIf flagGrid = 6 Then
Rem ********* Change rolling stock item ************
newRow = GridRepCon.row
newCol = GridRepCon.col
strOldCon = Left$(Label2(0).Caption, Len(Label2(0).Caption) - 4)
strStockType = Right$(Label2(0).Caption, 3)
strNewCon = GridRepCon.Cell(flexcpText)
strNewCon = Left$(strNewCon, Len(strNewCon) - 4)
strTempCon = strNewCon
strNewCon = ChrW$(34) & strNewCon & ChrW$(34)
GridRepCon.Select newRow, newCol + 1
strNewPath = GridRepCon.Cell(flexcpText)
strNewPath = ChrW$(34) & strNewPath & ChrW$(34)
strNewCon = strNewCon & " " & strNewPath
Rem ****************

frmStock.Grid3.Select MyRow, 0
strCon = frmStock.Grid3.Cell(flexcpText)
frmStock.Grid3.Select MyRow, 0
If Right$(strCon, 3) <> "con" Then

      Call MsgBox("This option only works for Consists," _
                  & vbCrLf & "           not for Activities." _
                  , vbExclamation, App.Title)
Exit Sub
End If
'strRoute = frmGrid.Grid1.Cell(flexcpText)
'strRoute = MSTSPath & "\Routes\" & strRoute
FileCopy MSTSPath & "\trains\consists\" & strCon, App.Path & "\TempFiles\" & strCon
DoEvents

flagway = 0    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)

Kill MSTSPath & "\trains\consists\" & strCon
flagway = 1    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)
FileCopy App.Path & "\TempFiles\" & strCon, MSTSPath & "\trains\consists\" & strCon
Kill App.Path & "\TempFiles\" & strCon
Label3.Caption = "Modified Consist Activated"
frmStock.GridStock.Select MyRow, MyCol
frmStock.GridStock.Cell(flexcpText) = strTempCon & ".con"

ElseIf flagGrid = 2 And Check1.value = 1 Then
Rem *********Change in all Consists/Activities ************

newRow = GridRepCon.row
newCol = GridRepCon.col
strOldCon = Left$(Label2(0).Caption, Len(Label2(0).Caption) - 4)
strOldStock = Label2(0).Caption
strStockType = Right$(Label2(0).Caption, 3)
strNewCon = GridRepCon.Cell(flexcpText)
strNewCon = Left$(strNewCon, Len(strNewCon) - 4)
strTempCon = strNewCon
strNewCon = ChrW$(34) & strNewCon & ChrW$(34)
GridRepCon.Select newRow, newCol + 1
strNewPath = GridRepCon.Cell(flexcpText)
strNewPath = ChrW$(34) & strNewPath & ChrW$(34)
strNewCon = strNewCon & " " & strNewPath
Rem ****************


For i = 1 To frmStock.GridStock.Rows - 1
frmStock.GridStock.Select i, 3
strExisting = frmStock.GridStock.Cell(flexcpText)
If strExisting = strOldStock Then
frmStock.GridStock.Select i, 0
strCon = frmStock.GridStock.Cell(flexcpText)
If Right(strCon, 3) = "con" Then

FileCopy MSTSPath & "\trains\consists\" & strCon, App.Path & "\TempFiles\" & strCon
DoEvents

flagway = 0    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)

Kill MSTSPath & "\trains\consists\" & strCon
flagway = 1    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)
FileCopy App.Path & "\TempFiles\" & strCon, MSTSPath & "\trains\consists\" & strCon
Kill App.Path & "\TempFiles\" & strCon
Label3.Caption = "Modified Consist Activated"
frmStock.GridStock.Select i, 3

frmStock.GridStock.Cell(flexcpText) = strTempCon & "." & strStockType

ElseIf Right(strCon, 3) = "act" Then

frmStock.GridStock.Select i, 12

strAct = frmStock.GridStock.Cell(flexcpText)
FileCopy MSTSPath & "\Routes\" & strAct & "\Activities\" & strCon, App.Path & "\TempFiles\" & strCon
DoEvents

flagway = 0    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)

Kill MSTSPath & "\Routes\" & strAct & "\Activities\" & strCon
flagway = 1    'unicode to ascii
Call ConvertStock(App.Path & "\TempFiles\" & strCon, flagway, strOldCon, strNewCon, strStockType)
FileCopy App.Path & "\TempFiles\" & strCon, MSTSPath & "\Routes\" & strAct & "\Activities\" & strCon
Kill App.Path & "\TempFiles\" & strCon

frmStock.GridStock.Select i, 3

frmStock.GridStock.Cell(flexcpText) = strTempCon & "." & strStockType
End If
End If
Next i

End If
Exit Sub
Errtrap:
If Err = 381 Then
Call MsgBox(Lang(354), vbExclamation, App.Title)
Exit Sub
End If
End Sub

Private Sub Command2_Click()
Dim ActName As String

If frmGrid.Grid1.col = 1 Then
ConDown = False

flagGrid = 3
frmGrid.Grid1.col = 0
ActName = frmGrid.Grid1.Cell(flexcpText)
frmGrid.Grid1.col = 1
ActName = ActName & "\Activities\" & frmGrid.Grid1.Cell(flexcpText)
frmGrid.Grid2.Rows = 0
frmGrid.Grid2.AddItem Lang(634)
frmGrid.Grid2.Rows = 1
frmGrid.Grid2.ExplorerBar = flexExSort
Call LooseConsistsGrid(ActName)
End If
Command1.Enabled = True
Unload Me
End Sub


Private Sub LooseConsistsGrid(ActPath As String)
Dim x As Integer, tempService As String, strNew As String
Dim Engname As String, Engpath As String, Wagname As String
Dim Wagonpath As String, booEntry As Boolean


On Error GoTo Errtrap

MousePointer = 11
RoutePath = MSTSPath & "\Routes"
TrainsetPath = MSTSPath & "\Trains\Trainset\"

ActPath = RoutePath & "\" & ActPath


NewFile = FreeFile
 Open ActPath For Input As #NewFile
 Do While Not EOF(NewFile)
Line Input #NewFile, A$
 
 Rem ************* Find loose consists - Locos
 tempService = vbNullString
   x = InStr(A$, "EngineData")
   strNew = A$
         If x > 0 Then
   Call CheckEngineData(strNew, Engname, Engpath, booEntry)
 

  
Engname = Engname & ".eng"
If booEntry = True Then
If Not FileExists(TrainsetPath & Engpath & "\" & Engname) Then
  FlagColRed = True
  If missEng = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(528) & vbCrLf
missEng = True
End If
  strbadbits = strbadbits & vbCrLf & Lang(528) & Engname
   frmGrid.Grid2.AddItem Engname
   

   Else
   FlagColRed = False
   frmGrid.Grid2.AddItem Engname
   End If
   
 End If
 End If
 Rem ************* Find loose consists - Wagons
 tempService = vbNullString
   x = InStr(A$, "WagonData")
   strNew = A$
         If x > 0 Then
   Call CheckWagonData(strNew, Wagname, Wagonpath, booEntry)

  
Wagname = Wagname & ".wag"
 If booEntry = True Then
 If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname) Then
 FlagColRed = True
 If missWag = False Then
strbadbits = strbadbits & vbCrLf & vbCrLf & Lang(529) & vbCrLf
missWag = True
End If
 strbadbits = strbadbits & vbCrLf & Lang(529) & Wagname
 ' frmgrid.grid2.CellForeColor = vbRed
   frmGrid.Grid2.AddItem Wagname
'   frmgrid.grid2.CellBackColor = &HFF
   Else
   FlagColRed = False
   frmGrid.Grid2.AddItem Wagname
   End If
   
 End If
 End If
 Loop
 Close #NewFile

MousePointer = 0
Exit Sub
Errtrap:

Call MsgBox("An error " & Err & " occurred in subroutine 'LooseConsistGrid' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
'Resume Next

End Sub


Private Sub Form_Load()
Dim i As Integer, tempPath As String, strBrake As String, strType As String, strName As String
Dim Couple As String, x As Integer
'MousePointer = 11

GridRepCon.BackColor = vbWhite
Me.Caption = Lang(291)
If flagGrid = 2 Then
Check1.Visible = True
Else
Check1.Visible = False
End If


If flagGrid = 2 Then

MyRow = frmStock.GridStock.row
MyCol = frmStock.GridStock.col
strStockType = Right$(frmStock.GridStock.Cell(flexcpText), 3)
Label2(0).Caption = frmStock.GridStock.Cell(flexcpText)
GridRepCon.width = 9000
GridRepCon.Cols = 5

GridRepCon.ExplorerBar = flexExSort
If strStockType = "eng" Then
Label1.Caption = Lang(356)
frmRepStock.Caption = Lang(357)
GridRepCon.Select 0, 0
GridRepCon.Cell(flexcpText) = Lang(358)
GridRepCon.Select 0, 1
GridRepCon.Cell(flexcpText) = Lang(293)

GridRepCon.Rows = 1

For i = 0 To lngLoco - 1
tempPath = MSTSPath & "\trains\trainset\" & LocoPath(i) & "\" & Locomotives(i)
If Not FileExists(tempPath) Then GoTo TryAnotherLoco
Call GetCoupling(tempPath, flagCouple, strBrake, strType, strName)

Select Case flagCouple
Case 1
Couple = "Automatic"
Case 2
Couple = "Chain"
Case Else
Couple = "Bar"
End Select

If Label2(0).Caption = Locomotives(i) Then
Label2(1).Caption = Couple
Label2(2).Caption = strBrake
Label2(3).Caption = strType
End If
GridRepCon.AddItem Locomotives(i) & vbTab & LocoPath(i) & vbTab & Couple & vbTab & strBrake & vbTab & strType
TryAnotherLoco:
Next i
Else
Label1.Caption = Lang(359)
frmRepCon.Caption = Lang(360)
GridRepCon.Select 0, 0
GridRepCon.Cell(flexcpText) = Lang(361)
GridRepCon.Select 0, 1
GridRepCon.Cell(flexcpText) = Lang(293)

GridRepCon.Rows = 1

For i = 0 To lngWagons - 1
tempPath = MSTSPath & "\trains\trainset\" & Wagpath(i) & "\" & Wagons(i)
If Not FileExists(tempPath) Then GoTo TryAnotherWag
Call GetCoupling(tempPath, flagCouple, strBrake, strType, strName)
Select Case flagCouple
Case 1
Couple = "Automatic"
Case 2
Couple = "Chain"
Case Else
Couple = "Bar"
End Select
If Label2(0).Caption = Wagons(i) Then
Label2(1).Caption = Couple
Label2(2).Caption = strBrake
Label2(3).Caption = strType
End If
GridRepCon.AddItem Wagons(i) & vbTab & Wagpath(i) & vbTab & Couple & vbTab & strBrake & vbTab & strType
TryAnotherWag:
Next i
End If

ElseIf flagGrid = 3 Then ' ************************************    Activity
MyRow = frmGrid.Grid1.row
MyCol = frmGrid.Grid1.col
strStockType = Right$(frmGrid.Grid2.Cell(flexcpText), 3)
Label2(0).Caption = frmGrid.Grid2.Cell(flexcpText)
GridRepCon.width = 9000
GridRepCon.Cols = 5

GridRepCon.ExplorerBar = flexExSort
If strStockType = "eng" Then
Label1.Caption = Lang(356)
frmRepCon.Caption = Lang(362)
GridRepCon.Select 0, 0
GridRepCon.Cell(flexcpText) = Lang(358)
GridRepCon.Select 0, 1
GridRepCon.Cell(flexcpText) = Lang(293)

GridRepCon.Rows = 1

For i = 0 To lngLoco - 1
tempPath = MSTSPath & "\trains\trainset\" & LocoPath(i) & "\" & Locomotives(i)

If Not FileExists(tempPath) Then GoTo TryAnother
Call GetCoupling(tempPath, flagCouple, strBrake, strType, strName)
Select Case flagCouple
Case 1
Couple = "Automatic"
Case 2
Couple = "Chain"
Case Else
Couple = "Bar"
End Select
If Label2(0).Caption = Locomotives(i) Then
Label2(1).Caption = Couple
Label2(2).Caption = strBrake
Label2(3).Caption = strType
End If
GridRepCon.AddItem Locomotives(i) & vbTab & LocoPath(i) & vbTab & Couple & vbTab & strBrake & vbTab & strType
TryAnother:
Next i
Else
Label1.Caption = Lang(359)
frmRepCon.Caption = Lang(363)
GridRepCon.Select 0, 0
GridRepCon.Cell(flexcpText) = Lang(361)
GridRepCon.Select 0, 1
GridRepCon.Cell(flexcpText) = Lang(293)

GridRepCon.Rows = 1

For i = 0 To lngWagons - 1
tempPath = MSTSPath & "\trains\trainset\" & Wagpath(i) & "\" & Wagons(i)
If Not FileExists(tempPath) Then GoTo TryAnotherWag2
Call GetCoupling(tempPath, flagCouple, strBrake, strType, strName)
Select Case flagCouple
Case 1
Couple = "Automatic"
Case 2
Couple = "Chain"
Case Else
Couple = "Bar"
End Select
If Label2(0).Caption = Wagons(i) Then
Label2(1).Caption = Couple
Label2(2).Caption = strBrake
Label2(3).Caption = strType
End If
GridRepCon.AddItem Wagons(i) & vbTab & Wagpath(i) & vbTab & Couple & vbTab & strBrake & vbTab & strType
TryAnotherWag2:
Next i
End If

ElseIf flagGrid = 4 Then ' ********************************  Consist

MyRow = frmGrid.Grid1.row
MyCol = frmGrid.Grid1.col
strStockType = Right$(frmGrid.Grid2.Cell(flexcpText), 3)
Label2(0).Caption = frmGrid.Grid2.Cell(flexcpText)
'Call GetCoupling(tempPath, flagCouple, strBrake, strType, strName)
GridRepCon.width = 9000
GridRepCon.Cols = 5

GridRepCon.ExplorerBar = flexExSort
If strStockType = "eng" Then
Label1.Caption = Lang(356)
frmRepCon.Caption = Lang(357)
GridRepCon.Select 0, 0
GridRepCon.Cell(flexcpText) = Lang(358)
GridRepCon.Select 0, 1
GridRepCon.Cell(flexcpText) = Lang(293)

GridRepCon.Rows = 1

For i = 0 To lngLoco - 1
tempPath = MSTSPath & "\trains\trainset\" & LocoPath(i) & "\" & Locomotives(i)
If Not FileExists(tempPath) Then GoTo TryAnotherLoco3
Call GetCoupling(tempPath, flagCouple, strBrake, strType, strName)
Select Case flagCouple
Case 1
Couple = "Automatic"
Case 2
Couple = "Chain"
Case Else
Couple = "Bar"
End Select
If Label2(0).Caption = Locomotives(i) Then
Label2(1).Caption = Couple
Label2(2).Caption = strBrake
Label2(3).Caption = strType
End If
GridRepCon.AddItem Locomotives(i) & vbTab & LocoPath(i) & vbTab & Couple & vbTab & strBrake & vbTab & strType
TryAnotherLoco3:
Next i
Else
Label1.Caption = Lang(359)
frmRepCon.Caption = Lang(360)
GridRepCon.Select 0, 0
GridRepCon.Cell(flexcpText) = Lang(361)
GridRepCon.Select 0, 1
GridRepCon.Cell(flexcpText) = Lang(293)

GridRepCon.Rows = 1

For i = 0 To lngWagons - 1
tempPath = MSTSPath & "\trains\trainset\" & Wagpath(i) & "\" & Wagons(i)
If Not FileExists(tempPath) Then GoTo TryAnotherWag3
Call GetCoupling(tempPath, flagCouple, strBrake, strType, strName)
Select Case flagCouple
Case 1
Couple = "Automatic"
Case 2
Couple = "Chain"
Case Else
Couple = "Bar"
End Select
If Label2(0).Caption = Wagons(i) Then
Label2(1).Caption = Couple
Label2(2).Caption = strBrake
Label2(3).Caption = strType
End If
GridRepCon.AddItem Wagons(i) & vbTab & Wagpath(i) & vbTab & Couple & vbTab & strBrake & vbTab & strType
TryAnotherWag3:
Next i
End If
ElseIf flagGrid = 5 Then ' ********************************  Consist

MyRow = frmUnusedSrv.GridCon.row
MyCol = frmUnusedSrv.GridCon.col
strStockType = Right$(frmUnusedSrv.GridLoco.Cell(flexcpText), 3)
Label2(0).Caption = frmUnusedSrv.GridLoco.Cell(flexcpText)
GridRepCon.width = 7000
GridRepCon.Cols = 2

GridRepCon.ExplorerBar = flexExSort
If strStockType = "eng" Then
Label1.Caption = Lang(356)
frmRepCon.Caption = Lang(357)
GridRepCon.Select 0, 0
GridRepCon.Cell(flexcpText) = Lang(358)
GridRepCon.Select 0, 1
GridRepCon.Cell(flexcpText) = Lang(293)

GridRepCon.Rows = 1

For i = 0 To lngLoco - 1
tempPath = MSTSPath & "\trains\trainset\" & LocoPath(i) & "\" & Locomotives(i)
If Not FileExists(tempPath) Then GoTo TryAnotherLoco4
Call GetCoupling(tempPath, flagCouple, strBrake, strType, strName)
Select Case flagCouple
Case 1
Couple = "Automatic"
Case 2
Couple = "Chain"
Case Else
Couple = "Bar"
End Select
If Label2(0).Caption = Locomotives(i) Then
Label2(1).Caption = Couple
Label2(2).Caption = strBrake
Label2(3).Caption = strType
End If
GridRepCon.AddItem Locomotives(i) & vbTab & LocoPath(i) & vbTab & Couple & vbTab & strBrake & vbTab & strType
TryAnotherLoco4:
Next i
Else
Label1.Caption = Lang(359)
frmRepCon.Caption = Lang(360)
GridRepCon.Select 0, 0
GridRepCon.Cell(flexcpText) = Lang(361)
GridRepCon.Select 0, 1
GridRepCon.Cell(flexcpText) = Lang(293)

GridRepCon.Rows = 1

For i = 0 To lngWagons - 1
tempPath = MSTSPath & "\trains\trainset\" & Wagpath(i) & "\" & Wagons(i)
If Not FileExists(tempPath) Then GoTo TryAnotherWag4
Call GetCoupling(tempPath, flagCouple, strBrake, strType, strName)
Select Case flagCouple
Case 1
Couple = "Automatic"
Case 2
Couple = "Chain"
Case Else
Couple = "Bar"
End Select
If Label2(0).Caption = Wagons(i) Then
Label2(1).Caption = Couple
Label2(2).Caption = strBrake
Label2(3).Caption = strType
End If
GridRepCon.AddItem Wagons(i) & vbTab & Wagpath(i) & vbTab & Couple & vbTab & strBrake & vbTab & strType
TryAnotherWag4:
Next i
End If
Rem ********************* Replace Missing eng/wag in Consist ************
ElseIf flagGrid = 6 Then
MyRow = frmStock.Grid3.row
MyCol = frmStock.Grid3.col
strStockType = Right$(frmStock.Grid3.Cell(flexcpText), 3)
x = InStrRev(frmStock.Grid3.Cell(flexcpText), "\")
Label2(0).Caption = Mid$(frmStock.Grid3.Cell(flexcpText), x + 1)
'Label2(0).Caption = frmStock.Grid3.Cell(flexcpText)
GridRepCon.width = 9000
GridRepCon.Cols = 5

GridRepCon.ExplorerBar = flexExSort
If strStockType = "eng" Then
Label1.Caption = Lang(356)
frmRepStock.Caption = Lang(357)
GridRepCon.Select 0, 0
GridRepCon.Cell(flexcpText) = Lang(358)
GridRepCon.Select 0, 1
GridRepCon.Cell(flexcpText) = Lang(293)

GridRepCon.Rows = 1

For i = 0 To lngLoco - 1
tempPath = MSTSPath & "\trains\trainset\" & LocoPath(i) & "\" & Locomotives(i)
If Not FileExists(tempPath) Then GoTo TryAnotherLoco6
Call GetCoupling(tempPath, flagCouple, strBrake, strType, strName)

Select Case flagCouple
Case 1
Couple = "Automatic"
Case 2
Couple = "Chain"
Case Else
Couple = "Bar"
End Select

If Label2(0).Caption = Locomotives(i) Then
Label2(1).Caption = Couple
Label2(2).Caption = strBrake
Label2(3).Caption = strType
End If
GridRepCon.AddItem Locomotives(i) & vbTab & LocoPath(i) & vbTab & Couple & vbTab & strBrake & vbTab & strType
TryAnotherLoco6:
Next i
Else
Label1.Caption = Lang(359)
frmRepCon.Caption = Lang(360)
GridRepCon.Select 0, 0
GridRepCon.Cell(flexcpText) = Lang(361)
GridRepCon.Select 0, 1
GridRepCon.Cell(flexcpText) = Lang(293)

GridRepCon.Rows = 1

For i = 0 To lngWagons - 1
tempPath = MSTSPath & "\trains\trainset\" & Wagpath(i) & "\" & Wagons(i)
If Not FileExists(tempPath) Then GoTo TryAnotherWag6
Call GetCoupling(tempPath, flagCouple, strBrake, strType, strName)
Select Case flagCouple
Case 1
Couple = "Automatic"
Case 2
Couple = "Chain"
Case Else
Couple = "Bar"
End Select
If Label2(0).Caption = Wagons(i) Then
Label2(1).Caption = Couple
Label2(2).Caption = strBrake
Label2(3).Caption = strType
End If
GridRepCon.AddItem Wagons(i) & vbTab & Wagpath(i) & vbTab & Couple & vbTab & strBrake & vbTab & strType
TryAnotherWag6:
Next i
End If

End If
If flagGrid <> 1 Then
GridRepCon.ColAlignment(1) = flexAlignLeftCenter
GridRepCon.col = 0
GridRepCon.Sort = flexSortStringAscending
End If
frmStock.MousePointer = 0
End Sub


