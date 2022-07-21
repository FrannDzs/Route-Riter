VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmRepCon 
   Caption         =   "Replace Consist With:-"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8LCtl.VSFlexGrid GridStock 
      Height          =   5895
      Left            =   4680
      TabIndex        =   6
      Top             =   600
      Width           =   3135
      _cx             =   5530
      _cy             =   10398
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRepCon.frx":0000
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   7080
      Width           =   855
   End
   Begin VSFlex8LCtl.VSFlexGrid gridRepCon 
      Height          =   5895
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   4095
      _cx             =   7223
      _cy             =   10398
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
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   3000
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRepCon.frx":003F
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Replace Consist"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmRepCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Text
Dim MyRow As Integer, MyCol As Integer
Private Sub Command1_Click()
Dim strSrv As String, strRoute As String, flagway As Integer


On Error GoTo Errtrap

If flagGrid = 1 Then
strOldCon = Left$(Label2.Caption, Len(Label2.Caption) - 4)
strNewCon = GridRepCon.Cell(flexcpText)
strNewCon = Left$(strNewCon, Len(strNewCon) - 4)

frmGrid.Grid1.Select MyRow, 3
strSrv = frmGrid.Grid1.Cell(flexcpText)
frmGrid.Grid1.Select MyRow, 0
strRoute = frmGrid.Grid1.Cell(flexcpText)
strRoute = MSTSPath & "\Routes\" & strRoute
FileCopy strRoute & "\Services\" & strSrv, App.Path & "\TempFiles\" & strSrv
DoEvents

flagway = 0    'unicode to ascii
Call ConvertIt(App.Path & "\TempFiles\" & strSrv, flagway, strOldCon, strNewCon)

Kill strRoute & "\Services\" & strSrv
flagway = 1    'unicode to ascii
Call ConvertIt(App.Path & "\TempFiles\" & strSrv, flagway, strOldCon, strNewCon)
FileCopy App.Path & "\TempFiles\" & strSrv, strRoute & "\Services\" & strSrv
Kill App.Path & "\TempFiles\" & strSrv
Label3.Caption = Lang(548)
frmGrid.Grid1.Select MyRow, MyCol
frmGrid.Grid1.Cell(flexcpText) = strNewCon & ".con"


End If
Exit Sub
Errtrap:
If Err = 381 Then
Call MsgBox(Lang(354), vbExclamation, App.Title)
Exit Sub
End If
End Sub

Private Function ConvertIt(CompleteFilePath As String, flagway As Integer, strOldCon As String, strNewCon As String) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertIt = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, xx As Long

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
xx = 1
If flagway = 0 Then
Label1:
x = InStr(xx, MyString, strOldCon)
If x = 0 Then GoTo EndIt

strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, x + Len(strOldCon))
If Right$(strStart, 1) = ChrW$(34) Then
MyString = strStart & strNewCon & strEnd
Else
MyString = strStart & ChrW$(34) & strNewCon & ChrW$(34) & strEnd
End If
xx = x + 1
'GoTo Label1

End If
EndIt:
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
ConvertIt = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function








Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Dim i As Integer
GridRepCon.BackColor = vbWhite
GridStock.BackColor = vbWhite
Me.Caption = Lang(282)
If flagGrid = 1 Then
Label1.Caption = Lang(283)
Label2.BackColor = vbWhite

GridRepCon.Select 0, 0
GridRepCon.Cell(flexcpText) = Lang(284)
GridStock.Select 0, 0
GridStock.Cell(flexcpText) = Lang(285)
MyRow = frmGrid.Grid1.row
MyCol = frmGrid.Grid1.col
GridRepCon.Rows = 1
GridRepCon.ExplorerBar = flexExSort
For i = 0 To lngCon - 1
GridRepCon.AddItem Consists(i)
Next i
Label2.Caption = frmGrid.Grid1.Cell(flexcpText)

End If
End Sub


Private Sub GridRepCon_Click()

If flagGrid = 1 Then
ConsistPath = MSTSPath & "\Trains\Consists"

GridStock.Rows = 1
GridStock.ExplorerBar = flexExSort

Call CheckForConsistGrid(ConsistPath & "\" & GridRepCon.Cell(flexcpText))
End If
End Sub


Private Sub CheckForConsistGrid(CFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Engpath As String, Engname As String
Dim strNew2 As String, Wagonpath As String, Wagname As String
Dim TrainsetPath As String, ConName As String, booEntry As Boolean
On Error GoTo Errtrap

Fnumber = FreeFile
x = InStrRev(CFilepath, "\")
ConName = Mid$(CFilepath, x + 1)

TrainsetPath = MSTSPath & "\Trains\Trainset\"
If Not FileExists(CFilepath) Then
Call MsgBox(Lang(530) & CFilepath, vbExclamation, Lang(530))
Exit Sub
End If
Open CFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 
   x = InStr(strNew, "EngineData")
   
         If x > 0 Then
   Call CheckEngineData(strNew, Engname, Engpath, booEntry)
 

   
   If booEntry = True Then
   If Not FileExists(TrainsetPath & Engpath & "\" & Engname & ".eng") Then
  
   GridStock.AddItem Engname & ".eng"
   GridStock.CellBackColor = &HFF
   Else
   GridStock.AddItem Engname & ".eng"
   End If
   End If
   strNew2 = vbNullString
   End If
    x = InStr(strNew, "WagonData")
   
         If x > 0 Then
   Call CheckWagonData(strNew, Wagname, Wagonpath, booEntry)
 


   If booEntry = True Then
  If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname & ".wag") Then
 
   GridStock.AddItem Wagname & ".wag"
   GridStock.CellBackColor = &HFF
   Else
   GridStock.AddItem Wagname & ".wag"
   End If
   End If
  End If
   
   strNew = vbNullString
 ' itExists = False
   Loop
   Close #Fnumber
  
Exit Sub
Errtrap:
Call MsgBox("An error #" & Err & " occurred in subroutine 'CheckForConsistGrid' while checking" _
            & vbCrLf & CFilepath _
            , vbExclamation, App.Title)


'Resume Next

End Sub




