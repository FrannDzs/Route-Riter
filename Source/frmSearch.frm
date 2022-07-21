VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmSearch 
   Caption         =   "List of Filtered Files."
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   12930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFixFix 
      Caption         =   "Fix Broken .sms"
      Height          =   495
      Left            =   2280
      TabIndex        =   36
      Top             =   8760
      Width           =   1935
   End
   Begin VSFlex8LCtl.VSFlexGrid List1 
      Height          =   5295
      Left            =   240
      TabIndex        =   35
      Top             =   240
      Width           =   12375
      _cx             =   21828
      _cy             =   9340
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Ellipsis        =   1
      ExplorerBar     =   1
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
   Begin VB.CommandButton Command20 
      Caption         =   "Fix AI wheel radius for MSTSbin"
      Height          =   495
      Left            =   240
      TabIndex        =   34
      Top             =   8760
      Width           =   1935
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Change ESD val in all selected .SD files"
      Height          =   495
      Left            =   10440
      TabIndex        =   33
      ToolTipText     =   "Alter the ESD_Alternative_Texture value"
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Make Thumbnails of .eng and .wag files"
      Height          =   495
      Left            =   10440
      TabIndex        =   32
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Delete All Selected Files"
      Height          =   495
      Left            =   10440
      TabIndex        =   31
      ToolTipText     =   "Warning! This option completely deletes all files listed in the window above, they are not saved..."
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Remove AI/Dead Locos from List"
      Height          =   495
      Left            =   8400
      TabIndex        =   30
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Make Unpowered Locos"
      Height          =   495
      Left            =   6360
      TabIndex        =   29
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Make AI Locos"
      Height          =   495
      Left            =   4320
      TabIndex        =   28
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Auto-Convert selected .eng to Raildriver"
      Height          =   495
      Left            =   2280
      TabIndex        =   27
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Revert .SX/.ENX files back to .S/.ENG"
      Height          =   495
      Left            =   8400
      TabIndex        =   26
      ToolTipText     =   "Reverts files backed up as .sx and .enx files to their original state"
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton Command11 
      Caption         =   "List .S files used by selected Stock"
      Height          =   495
      Left            =   8400
      TabIndex        =   25
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Remove Animation from selected .S files"
      Height          =   495
      Left            =   8400
      TabIndex        =   24
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Convert Selected to Unicode"
      Height          =   495
      Left            =   6360
      TabIndex        =   23
      ToolTipText     =   "If the selected file is ASCII it is converted to Unicode."
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmdFixCVF 
      Caption         =   "Fix .CVF Files"
      Height          =   495
      Left            =   4320
      TabIndex        =   22
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmdFixBB 
      Caption         =   "Fix Bounding-Box Minimum"
      Height          =   495
      Left            =   2280
      TabIndex        =   21
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmdFixCon 
      Caption         =   "Fix .CON Files"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmdFixSrv 
      Caption         =   "Fix .srv names"
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdFixAct 
      Caption         =   "Fix .act names"
      Height          =   495
      Left            =   6360
      TabIndex        =   18
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdFixEng 
      Caption         =   "Fix .ENG/.WAG files"
      Height          =   495
      Left            =   4320
      TabIndex        =   17
      ToolTipText     =   "Fixes Case in Wagon Names, and Damper Units to N/m/s"
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdFixSD 
      Caption         =   "Fix All .SD Files"
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      ToolTipText     =   "To use this, you must first select Trainset in the left directory window, and *.sd in the Filters."
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdFixSms 
      Caption         =   "Fix .SMS Files"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "Where possible fixes errors in .sms files such as alias errors etc."
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Print/Save Pictures of .S files"
      Height          =   495
      Left            =   8400
      TabIndex        =   12
      Top             =   5760
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save Selected FileList"
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Print Selected FileList"
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Compress selected .ACE files"
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Select All"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Convert selected files to DXT1"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Compress selected .S files"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   495
      Index           =   2
      Left            =   10680
      TabIndex        =   2
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Unselect All"
      Height          =   495
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      ToolTipText     =   "Unselects all Selected files"
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Selected Files"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "If the file is a .S or .ACE it will be shown in the viewer, otherwise in the Text Editor."
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   11760
      TabIndex        =   14
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Number of Selected Files"
      Height          =   495
      Left            =   10440
      TabIndex        =   13
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "File being processed"
      Top             =   8880
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Total Number of Files in List."
      Height          =   495
      Left            =   10440
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   11760
      TabIndex        =   3
      Top             =   5760
      Width           =   855
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Text
Dim comp As CompressZIt
Dim lAnim As Long
Dim bxdata() As Byte
Const Wag_CHUNK = 500
Dim strShape() As Variant
Dim strShape2() As Variant
Dim booAbort As Boolean
Dim booAnimRemoved As Boolean

Dim booCheckSound As Boolean
Private WithEvents SP As cScanPath
Attribute SP.VB_VarHelpID = -1
Public strFilter As String
Private Sub DoComp(strFile As String, strFPath As String, strSparePath As String)
Dim strBatText As String, strSuffix As String

strSuffix = "-" & Right$(strFile, 1)


   ChDrive Left$(App.Path, 1)
 ''ChDir App.Path & "\TSUtil"
 If strSuffix = "-t" Then
 strBatText = "java TSUtil fmgr " & strSuffix & " -r -n" & ChrW$(34) & strFile & ChrW$(34) & " " & ChrW$(34) & strFPath & ChrW$(34) & " " & ChrW$(34) & strSparePath & ChrW$(34)
 ElseIf strSuffix <> "-t" Then
   strBatText = "java TSUtil fmgr " & strSuffix & " -c -n" & ChrW$(34) & strFile & ChrW$(34) & " " & ChrW$(34) & strFPath & ChrW$(34) & " " & ChrW$(34) & strSparePath & ChrW$(34)
End If

  Call ShellAndWait(strBatText, True, vbHide)

 DoEvents
 End Sub
Public Function RemD2(ByRef rArray() As Variant, xArray() As Variant) As Variant

    'Declare variables
    Dim ii As Long, jj As Long
 On Error GoTo Errtrap
    'Initialize variables
    count3 = 1
    high = UBound(rArray)
    'Declare temp array
    
    ReDim xArray(0 To high)
    
    'Start duplicates removal code

xArray(0) = rArray(0)
jj = 1
    For ii = 1 To high
        If rArray(ii) <> rArray(ii - 1) Then
        If rArray(ii) <> vbNullString Then
        xArray(jj) = rArray(ii)
        jj = jj + 1
        End If
End If
Next ii

If high > 0 Then
If rArray(high) <> rArray(high - 1) Then
xArray(high) = rArray(high)
End If
 End If
CarryON:

ReDim Preserve xArray(0 To jj - 1)

Exit Function
Errtrap:

Call MsgBox("An error " & Err & " occurred in subroutine 'RemD2' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Function

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

Private Function ReadASCIIFile(CompleteFilePath As String) As String

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
mytristate = 0
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll
The_obj.Close
fileflag = False
ReadASCIIFile = MyString
End Function

Private Sub CheckSound(CompleteFilePath As String, strSound As String)
Dim x As Integer, strTrainFolder As String, strSoundFolder As String, strGlobalSound As String


x = InStrRev(CompleteFilePath, "\")
strTrainFolder = Left(CompleteFilePath, x)
strSoundFolder = strTrainFolder & "Sound\"
strSound = Replace(strSound, "\\", "/")
strSound = Replace(strSound, ChrW$(34), "")

x = InStrRev(CompleteFilePath, "\Trains\")
strGlobalSound = Left(CompleteFilePath, x) & "Sound\"
x = InStr(strTrainFolder, "Trainset")
strTrainset = Left(strTrainFolder, x + 8)
If Left(strSound, 6) = "../../" Then
        strSound = Mid(strSound, 7)
        strSound = Replace(strSound, "/", "\")
        strSound = strTrainset & strSound

    If Not FileExists(strSound) Then
    
            strReport = strReport & CompleteFilePath & " is missing sound " & strSound & vbCrLf
    End If

ElseIf Not FileExists(strSoundFolder & strSound) Then
        If Not FileExists(strGlobalSound & strSound) Then
                strReport = strReport & CompleteFilePath & " is missing sound " & strSound & vbCrLf
        End If
End If

End Sub





Private Sub CountDots(strAlias As String, Y As Integer)
Dim x As Integer
x = 1
TryAgain:
x = InStr(x, strAlias, "../")
If x > 0 Then
x = x + 1
Y = Y + 1
GoTo TryAgain
End If

End Sub

Private Sub cmdFixAct_Click()
Dim ActPath As String, flagway As Integer
Dim MyString As String, booFailed As Boolean, x As Integer, strActPath As String, strActName As String
Dim strAct As String, strInitial As String

On Error GoTo Errtrap
MousePointer = 11
strReport = vbNullString
lblCount(0).Caption = Str(Val(lblCount(0).Caption) - 1)
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
ActPath = List1.TextMatrix(i, 0)


x = InStrRev(ActPath, "\")
strActPath = Left$(ActPath, x)

If Right$(strActPath, 11) <> "Activities\" Then
GoTo TryAgain
End If
strActName = Mid$(ActPath, x + 1)


   flagway = 0
   strAct = Left$(strActName, Len(strActName) - 4)
   MyString = ReadUniFile(ActPath)
      DoEvents
      strInitial = MyString
   Call FixCon(MyString, strActName, booFailed, ActPath)
   If booFailed = True Then
   booFailed = False
   GoTo TryAgain
   End If
   DoEvents
   If MyString <> strInitial Then
   Call WriteUniFile(ActPath, MyString)
   DoEvents
End If
   End If
TryAgain:
Next i
If strReport <> vbNullString Then
  frmReport.Rich1.Text = strReport
     frmReport.Show 1
  End If
  DoEvents
'  frmUtils.Text1(0) = "*.*"
'  Unload Me
 Call MsgBox("Activity Names have all been fixed, you may now close this screen.", vbInformation, App.Title)
 
MousePointer = 0
Exit Sub
Errtrap:
Call MsgBox(Err.Description & " occurred in cmdFixAct processing " & strActPath, vbExclamation, App.Title)
GoTo TryAgain
End Sub

Private Sub cmdFixBB_Click()
Dim i As Long, strPath As String, flagway As Integer, strCorrShape As String
Dim x As Integer

For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)
x = InStrRev(strPath, "\")
strCorrShape = Mid$(strPath, x + 1)
strCorrShape = Left$(strCorrShape, Len(strCorrShape) - 1)

If Right$(strPath, 3) <> ".sd" Then GoTo GetNext
flagway = 0
Call ConvertSD2(strPath, flagway)
flagway = 1
Call ConvertSD2(strPath, flagway)
DoEvents
End If
GetNext:
Next i
If strReport <> vbNullString Then
  frmReport.Rich1.Text = strReport
     frmReport.Show 1
  End If
  DoEvents
  
Call MsgBox("BoundingBox entries in .SD files have all been fixed, you may now close this screen.", vbInformation, App.Title)
End Sub

Private Sub cmdFixCon_Click()
Dim strConPath As String, strConName As String, strCon As String, flagway As Integer
Dim MyString As String, booFailed As Boolean, x As Integer, strInitial As String, booSame As Boolean
On Error GoTo Errtrap
strReport = vbNullString
MousePointer = 11

lblCount(0).Caption = Str(Val(lblCount(0).Caption) - 1)
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strConPath = List1.TextMatrix(i, 0)

x = InStrRev(strConPath, "\")
strCon = Left$(strConPath, x)

strConName = Mid$(strConPath, x + 1)
   flagway = 0
   strConName = Left$(strConName, Len(strConName) - 4)
   
   MyString = ReadUniFile(strConPath)
   If MyString = vbNullString Then
   strReport = strReport & "An error occurred while processing " & strConName & vbCrLf
   GoTo TryAgain
   End If
   DoEvents

   strInitial = MyString
   Call FixCon(MyString, strConName, booFailed, strConPath)
   If booFailed = True Then
   booFailed = False
   GoTo TryAgain
   End If
   DoEvents

   Call CompStrings(MyString, strInitial, booSame)
   If booSame = False Then
  

   Call WriteUniFile(strConPath, MyString)
   DoEvents
   End If
 End If
TryAgain:
Next i

If strReport <> vbNullString Then
  frmReport.Rich1.Text = strReport
     frmReport.Show 1
  End If
MousePointer = 0
 Call MsgBox("Consist Names have all been fixed, you may now close this screen.", vbInformation, App.Title)
Exit Sub
Errtrap:
strReport = strReport & "An error occurred while processing " & strConName & vbCrLf
GoTo TryAgain

End Sub



Private Sub cmdFixCVF_Click()
Dim strPath As String, strNew As String, x As Long, Y As Long, strCab As String
Dim NewFile As Integer, yy As Long, strCabName As String, strTrains As String
Dim Z As Long, booAlias As Boolean
Dim strThisPath As String, zz As Long
Dim FirstPass As Integer, strUni As String, zy As Long, strAliasPath As String
Dim strAlias As String, yz As Long, strDots As String
Dim strStart As String, strEnd As String, booCabView As Boolean


On Error GoTo Errtrap
lblCount(0).Caption = Str(Val(lblCount(0).Caption) - 1)
MousePointer = 11
strReport = vbNullString
strTrains = MSTSPath & "\trains\trainset\"


For i = 0 To List1.Rows - 1

If List1.IsSelected(i) = True Then

lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)
   'strPath = List1.Cell(flexcpText)
   Z = InStrRev(strPath, "\")
   x = InStr(strPath, "common")
   If Mid$(strPath, Z - 7, 7) = "CabView" Then
   booCabView = True
   strcabpath = Left$(strPath, Z)
   Else
   booCabView = False
   End If
   If Mid$(strPath, Z - 7, 7) <> "CabView" And x = 0 Then GoTo NextOne
   
   strThisPath = Left$(strPath, Z)
   zz = InStr(strThisPath, "trainset")
   zz = InStr(zz, strThisPath, "\")
   DoEvents
TryAgain:

    Open strPath For Binary As #2
    strUni = String(2, " ")
    Get #2, , strUni
     Close #2

 If Asc(Mid$(strUni, 1, 1)) <> 255 And Asc(Mid$(strUni, 2, 1)) <> 254 Then
 Call ConvertIt(strPath, 1)
 DoEvents
 End If
 
   If Right$(strPath, 4) <> ".cvf" Then GoTo NextOne
   
   strNew = ReadUniFile(strPath)
   strNew = Replace(strNew, "\\", "/")
   DoEvents
   strNew = Replace(strNew, "//", "/")
   DoEvents
   strNew = Replace(strNew, ChrW$(34) & "/../", ChrW$(34) & "../../")
   DoEvents
   'strNew = Replace(strNew, vbtab, " ")
   DoEvents
   strNew = Replace(strNew, "CabViewFile (" & ChrW$(34), "CabViewFile ( " & ChrW$(34))
   DoEvents
   Rem ********** Fix for Mike
   strNew = Replace(strNew, "CabViewFile ( " & ChrW$(34) & " " & ChrW$(34), "CabViewFile ( " & ChrW$(34))
   DoEvents
   strNew = Replace(strNew, "Graphic (" & ChrW$(34), "Graphic ( " & ChrW$(34))
   DoEvents
   x = 0
   FirstPass = 0
FindMore:
booAlias = False
FirstPass = FirstPass + 1
   x = InStr(x + 1, strNew, "CabViewFile")
  
   If x <> 0 Then
   If Mid$(strNew, x - 3, 3) = "Tr_" Then
   x = x + 1
   GoTo FindMore
   End If
   If Mid$(strNew, x, 13) <> "CabViewFile (" Then
   
   zy = InStr(x, strNew, "(")
   
   strNew = Left$(strNew, x + 11) & " " & Mid$(strNew, zy)
   
   End If
   If Mid$(strNew, x + 13, 1) <> " " Then
   strNew = Left$(strNew, x + 12) & " " & Mid$(strNew, x + 13)
   End If
   End If
         If x = 0 Then
         If FirstPass = 1 Then
         GoTo GetGraphics
         Else
         GoTo GetGraphics
         End If
         End If
         Y = InStr(x, strNew, ChrW$(34))
         If Y = 0 Or Y - x > 17 Then '***** No quotes ******
         
         yz = InStr(x, strNew, ".ace")
         If yz - x > 100 Then GoTo NextOne
         strCab = Mid$(strNew, x + 14, (yz + 4) - (x + 13))
         strCab = Trim$(strCab)
         strNew = Left$(strNew, x + 12) & " " & ChrW$(34) & strCab & ChrW$(34) & Mid$(strNew, yz + 4)
       
         GoTo NextBit
         End If

         yy = InStr(Y + 1, strNew, ChrW$(34))
         strCab = Trim$(Mid$(strNew, Y + 1, yy - (Y + 1)))

        strStart = Left$(strNew, Y)
        strEnd = Mid$(strNew, yy)
NextBit:
         strAlias = vbNullString
         If Left$(strCab, 12) = "../../../../" Then
         strDots = "..\..\"
             Z = InStrRev(strCab, "/")
           
            strCabName = Mid$(strCab, Z + 1)
           strcabpath = Mid$(strCab, 13, Z - 12)
           strAlias = Left$(strCab, Z)
            booAlias = True
            strAliasPath = MSTSPath & "\" & strcabpath
            strAliasPath = Replace(strAliasPath, "/", "\")
            ElseIf Left$(strCab, 6) = "../../" Then
            strDots = vbNullString
             Z = InStrRev(strCab, "/")
           
            strCabName = Mid$(strCab, Z + 1)
           strcabpath = Mid$(strCab, 7, Z - 6)
           
            strAlias = Left$(strCab, Z)
            booAlias = True
            strAliasPath = strTrains & strcabpath
            strAliasPath = Replace(strAliasPath, "/", "\")
            Else
            strCabName = strCab
           
            End If

        If booAlias = False Then
         If Not FileExists(strThisPath & strCabName) Then

     
                
              If FileExists(strTrains & "Dash9\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Dash9/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "380\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../380/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Scotsman\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Scotsman/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "GP38\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../GP38/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Kiha31\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Kiha31/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Series2000\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../SEries2000/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Series7000\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Series7000/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "SD402\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../SD402/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Acela\CabView\" & strCabName) Then
             
                strNew = strStart & strDots & "../../Acela/CabView/" & strCabName & strEnd
                Else
                
                strReport = strReport & strCabName & " Is missing from " & strPath & vbCrLf & vbCrLf
                End If
                End If
                End If
         If booAlias = True Then

                If Not FileExists(strAliasPath & strCabName) Then
                If FileExists(strTrains & "Dash9\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Dash9/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "380\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../380/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Scotsman\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Scotsman/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "GP38\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../GP38/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Kiha31\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Kiha31/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Series2000\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../SEries2000/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Series7000\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Series7000/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "SD402\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../SD402/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Acela\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Acela/CabView/" & strCabName & strEnd
                Else
                strReport = strReport & strCabName & " Is missing from " & strAliasPath & vbCrLf & "Called by " & strPath & vbCrLf & vbCrLf
                End If
          End If
        End If
         If x = 0 Then
         x = yy
         End If
         GoTo FindMore
         End If
Rem ***************** Find Graphics ****************************
GetGraphics:
    x = 0
   FirstPass = 0
FindMore2:
booAlias = False
FirstPass = FirstPass + 1
   x = InStr(x + 1, strNew, "Graphic")
  
   If x <> 0 Then
   
   If Mid$(strNew, x, 9) <> "Graphic (" Then
   
   zy = InStr(x, strNew, "(")
   
   strNew = Left$(strNew, x + 7) & " " & Mid$(strNew, zy)
   
   End If
   If Mid$(strNew, x + 9, 1) <> " " Then
   strNew = Left$(strNew, x + 8) & " " & Mid$(strNew, x + 9)
   End If
   End If
         If x = 0 Then
         If FirstPass = 1 Then
         GoTo NextOne
         Else
         GoTo CarryON
         End If
         End If
         Y = InStr(x, strNew, ChrW$(34))
         If Y = 0 Or Y - x > 17 Then '***** No quotes ******
         
         yz = InStr(x, strNew, ".ace")
         If yz - x > 100 Then GoTo NextOne
         strCab = Mid$(strNew, x + 10, (yz + 4) - (x + 9))
         strCab = Trim$(strCab)
         strNew = Left$(strNew, x + 8) & " " & ChrW$(34) & strCab & ChrW$(34) & Mid$(strNew, yz + 4)
      
         GoTo NextBit2
         End If

         yy = InStr(Y + 1, strNew, ChrW$(34))
         strCab = Trim$(Mid$(strNew, Y + 1, yy - (Y + 1)))

        strStart = Left$(strNew, Y)
        strEnd = Mid$(strNew, yy)
NextBit2:
         strAlias = vbNullString
         If Left$(strCab, 12) = "../../../../" Then
         strDots = "..\..\"
             Z = InStrRev(strCab, "/")
           
            strCabName = Mid$(strCab, Z + 1)
           strcabpath = Mid$(strCab, 13, Z - 12)
           strAlias = Left$(strCab, Z)
            booAlias = True
            strAliasPath = MSTSPath & "\" & strcabpath
            strAliasPath = Replace(strAliasPath, "/", "\")
            ElseIf Left$(strCab, 6) = "../../" Then
            strDots = vbNullString
             Z = InStrRev(strCab, "/")
           
            strCabName = Mid$(strCab, Z + 1)
           strcabpath = Mid$(strCab, 7, Z - 6)
           
            strAlias = Left$(strCab, Z)
            booAlias = True
            strAliasPath = strTrains & strcabpath
            strAliasPath = Replace(strAliasPath, "/", "\")
            Else
            strCabName = strCab
           
            End If

        If booAlias = False Then
         If Not FileExists(strThisPath & strCabName) Then

     
                
              If FileExists(strTrains & "Dash9\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Dash9/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "380\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../380/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Scotsman\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Scotsman/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "GP38\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../GP38/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Kiha31\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Kiha31/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Series2000\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../SEries2000/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Series7000\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Series7000/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "SD402\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../SD402/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Acela\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Acela/CabView/" & strCabName & strEnd
                Else
                strReport = strReport & strCabName & " Is missing from " & strPath & vbCrLf & vbCrLf
                End If
                End If
                End If
         If booAlias = True Then

                If Not FileExists(strAliasPath & strCabName) Then
                If FileExists(strTrains & "Dash9\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Dash9/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "380\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../380/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Scotsman\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Scotsman/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "GP38\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../GP38/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Kiha31\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Kiha31/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Series2000\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../SEries2000/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Series7000\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Series7000/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "SD402\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../SD402/CabView/" & strCabName & strEnd
                ElseIf FileExists(strTrains & "Acela\CabView\" & strCabName) Then
                strNew = strStart & strDots & "../../Acela/CabView/" & strCabName & strEnd
                Else
                strReport = strReport & strCabName & " Is missing from " & strAliasPath & vbCrLf & "Called by " & strPath & vbCrLf & vbCrLf
                End If
          End If
        End If
         If x = 0 Then
         x = yy
        ' End If
         GoTo FindMore2
        End If
   
   
Rem ************************************************************
CarryON:
If strNew = vbNullString Then GoTo NextOne
    NewFile = FreeFile
    Open strPath For Output As #NewFile
   Print #NewFile, strNew
    Close #NewFile
    DoEvents
    
   Call ConvertIt(strPath, 1)
   DoEvents

NextOne:

   Next i
   MousePointer = 0
   If strReport <> vbNullString Then
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   End If
   Call MsgBox(".CVF files have all been checked, you may now close this screen.", vbInformation, App.Title)
   Exit Sub
Errtrap:
 
   strReport = strReport & "An unresolved error occurred while processing " & strPath & vbCrLf
   GoTo NextOne
   
   
   
End Sub

Private Sub cmdFixEng_Click()
Dim strPath As String, x As Long, xx As Long, j As Long, TrainsetPath As String
Dim flagway As Integer, Y As Long, yy As Long, MyComp As Variant, strSound As String, booChanged As Boolean
Dim i As Long, MyString As String, strInitial As String, NewFile As Integer, strTemp As String
Dim booChanged2 As Boolean, strLoco As String, strOrig As String, booSame As Boolean

On Error GoTo Errtrap
lblCount(0).Caption = Str(Val(lblCount(0).Caption) - 1)
Select Case MsgBox("Do you also wish to check for missing Sounds in .eng/.wag files" _
                   & vbCrLf & "(This will slow things up somewhat)" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
booCheckSound = True
    Case vbNo
booCheckSound = False
End Select
MousePointer = 11
For i = 0 To List1.Rows - 1
booChanged = False
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)
 x = InStrRev(strPath, "\")
TrainsetPath = Left$(strPath, x)

   If Right$(strPath, 3) = "eng" Then
   MyComp = StrComp(Right$(strPath, 3), "eng", 0)
    If MyComp <> 0 Then
    
    strPath = Left$(strPath, Len(strPath) - 3)
    strPath = strPath & "eng"
    End If
   
   ElseIf Right$(strPath, 3) = "wag" Then
   MyComp = StrComp(Right$(strPath, 3), "wag", 0)
    If MyComp <> 0 Then
    
    strPath = Left$(strPath, Len(strPath) - 3)
    strPath = strPath & "wag"
    End If
   End If
   x = InStrRev(strPath, "\")

   flagway = 0
   x = InStrRev(strPath, "\")
   If Mid$(strPath, x + 1) = "Default.wag" Then GoTo CarryON
   
   Y = InStrRev(strPath, "\", x - 1)

   If Mid$(strPath, Y - 8, 8) <> "Trainset" Then GoTo CarryON
   Rem ***********************************************
   NewFile = FreeFile
   Open strPath For Binary As #NewFile
    strTemp = String(2, " ")
    Get #NewFile, , strTemp
 Close #NewFile
 
If Asc(Mid$(strTemp, 1, 1)) <> 255 And Asc(Mid$(strTemp, 2, 1)) <> 254 Then

   Call ReadFile3(strPath, MyString)
   booChanged2 = True
   Else
   MyString = ReadUniFile(strPath)
   End If
   strInitial = MyString
   If Len(MyString) < 10 Then

strReport = strReport & strPath & " could not be read" & vbCrLf & vbCrLf
GoTo CarryON
End If
'MyString = Replace(MyString, vbtab, " ")


Rem ***************** Start
x = InStrRev(strPath, "\")
strOrig = Mid$(strPath, x + 1)
strOrig = Left$(strOrig, Len(strOrig) - 4)
Y = InStrRev(strPath, "\", x - 1)
strOrigPath = Mid$(strPath, Y + 1, x - Y)
strEngpath = Left$(strPath, x)
MyString = Replace(MyString, "Wagon(", "Wagon (")
DoEvents
MyString = Replace(MyString, "Type(", "Type (")
DoEvents
MyString = Replace(MyString, "FreightAnim(", "FreightAnim (")
DoEvents
MyString = Replace(MyString, "Type ( Passenger )", "Type ( Carriage )")
DoEvents

x = InStr(MyString, "Wagon ")
Y = InStr(x, MyString, "(")
    If Y > x + 6 Then
    MyString = Left$(MyString, x + 4) & " " & Mid$(MyString, Y)
    End If
If Mid$(MyString, x + 7, 1) <> " " Then
MyString = Left$(MyString, x + 6) & " " & Mid$(MyString, x + 7)
End If
   Y = InStr(x, MyString, vbCr)
   strLoco = Mid$(MyString, x + 8, Y - (x + 8))
   strLoco = Trim$(strLoco)
   If Left$(strLoco, 1) = ChrW$(34) Then
   strLoco = Mid$(strLoco, 2)
   End If
   If Right$(strLoco, 1) = ChrW$(34) Then
   strLoco = Left$(strLoco, Len(strLoco) - 1)
   End If
   Call CompStrings(strLoco, strOrig, booSame)
If booSame = False Then

   If UCase(strLoco) = UCase(strOrig) Then
   booChanged2 = True
   GoTo Another
  
   Else
   strReport = strReport & strOrigPath & strOrig & " has a 'Wagon' entry of " & strLoco & " - These should be identical or errors in Activities may result." & vbCrLf & vbCrLf
   
   GoTo CarryOn2
   End If
 End If
   x = 1
Another:
   x = InStr(x, MyString, strLoco)
   If x = 0 Then GoTo CarryOn2
   
   strStart = Left$(MyString, x - 1)
   strEnd = Mid$(MyString, x + Len(strLoco))
   MyString = strStart & strOrig & strEnd
   x = x + 10
   GoTo Another
CarryOn2:

x = InStr(MyString, "WagonShape")
If x = 0 Then GoTo GetNext
j = InStr(x, MyString, "(")
MyString = Left$(MyString, x + 9) & " " & Mid$(MyString, j)

If x = 0 Then
    strReport = strReport & strPath & " could not be read" & vbCrLf & vbCrLf
GoTo CarryON
End If
If Mid$(MyString, x + 12, 1) <> " " Then
MyString = Left$(MyString, x + 11) & " " & Mid$(MyString, x + 12)
End If

   Y = InStr(x, MyString, ")")
   strSFile = Mid$(MyString, x + 12, Y - (x + 12))
   strStart = Left$(MyString, x + 12)
   strEnd = Mid$(MyString, Y)
   If Left$(strEnd, 1) <> " " Then
   strEnd = " " & strEnd
strSFile = Trim$(strSFile)
End If
If Left$(strSFile, 1) = ChrW$(34) Then
strSFile = Mid$(strSFile, 2)
strStart = strStart & ChrW$(34)
End If
If Right$(strSFile, 1) = ChrW$(34) Then
strSFile = Left$(strSFile, Len(strSFile) - 1)
strEnd = ChrW$(34) & strEnd
End If
If Not FileExists(TrainsetPath & "\" & strSFile) Then

strReport = strReport & "Shape file " & TrainsetPath & "\" & strSFile & " is missing" & vbCrLf & vbCrLf
GoTo CarryON
End If

frmUtils.Drive1(1).Drive = Left$(strEngpath, 1)
frmUtils.Dir1(1).Path = strEngpath
frmUtils.Text1(1) = strSFile
DoEvents
strSName = frmUtils.File1(1).List(0)
MyString = strStart & strSName & strEnd
DoEvents
Rem *************Check for Freight Anim

x = InStr(MyString, "FreightAnim")
If x > 0 Then
xx = InStrRev(MyString, "Comment", x)
If xx > 0 And (x - xx) < 40 Then GoTo TheRest

x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ".s")
strAnim = Mid$(MyString, x + 1, (xx + 2) - (x + 1))
strAnim = Trim$(strAnim)
If Left$(strAnim, 1) = ChrW$(34) Then
strAnim = Mid$(strAnim, 2)
End If
If Right$(strAnim, 1) = ChrW$(34) Then
strAnim = Left$(strAnim, Len(strAnim) - 1)
End If
If Left$(strAnim, 2) = ".." Then strAnim = vbNullString
End If
If strAnim <> "" And Not FileExists(strEngpath & strAnim) Then
strReport = strReport & "FreightAnim file " & strAnim & " is missing from " & strOrig & vbCrLf
End If
End If
strAnim = ""
Rem *****************End Fix Names
TheRest:
x = 1

strNew = "/s"

LookForMore:

xx = InStr(x, MyString, "Damping")
If xx = 0 Then GoTo CarryOn3
yy = InStr(xx, MyString, "(")
Y = InStr(yy, MyString, ")")
strDamp = Mid$(MyString, yy + 1, Y - yy - 1)
x = InStr(strDamp, "N/m/s")
    If x = 0 Then
    j = 1
TryAgain:
    x = InStr(j, strDamp, "N/m")
            If x > 0 Then
            strDamp = Left$(strDamp, x + 2) & "/s" & Mid$(strDamp, x + 3)
            j = x + 5
            GoTo TryAgain
            End If
    
    strStart = Left$(MyString, yy)
    strEnd = Mid$(MyString, Y)
    MyString = strStart & strDamp & strEnd

    End If
    x = yy + 5
    GoTo LookForMore
Rem ********** Place quotes around Sound entries
CarryOn3:

x = 1


LookForMore2:

xx = InStr(x, MyString, "Sound ")
If xx = 0 Then GoTo CarryON

yy = InStr(xx, MyString, "(")
If yy = 0 Then GoTo CarryON
If yy - xx > 20 Then
x = yy + 5
    GoTo LookForMore2
    End If
Y = InStr(yy, MyString, ")")
j = InStr(yy + 1, MyString, "(")
If j < Y And j <> 0 Then
Y = InStr(Y + 1, MyString, ")")
End If
strSound = Mid$(MyString, yy + 1, Y - yy - 1)
strSound = Trim$(strSound)
If Left(strSound, 1) = vbTab Then
strSound = Mid(strSound, 2)
End If
If Left$(strSound, 1) <> ChrW$(34) Then
strSound = ChrW$(34) & strSound
booChanged = True
End If
If Right$(strSound, 1) = vbLf Then
strSound = Left$(strSound, Len(strSound) - 1)
strSound = Trim$(strSound)
End If
If Right$(strSound, 1) = vbCr Then
strSound = Left$(strSound, Len(strSound) - 1)
strSound = Trim$(strSound)
End If
If Right$(strSound, 1) <> ChrW$(34) Then
strSound = strSound & ChrW$(34)
booChanged = True
End If

    If booChanged = True Then
  
    booChanged = False
    strStart = Left$(MyString, yy)
    strEnd = Mid$(MyString, Y)
    MyString = strStart & strSound & strEnd
End If

Rem ******************** Start Sound Check *********************
If booCheckSound = True Then

Call CheckSound(strPath, strSound)

End If

Rem ************************************************************
    x = yy + 5
    GoTo LookForMore2

CarryON:

If MyString <> strInitial Or booChanged2 = True Then
Call WriteUniFile(strPath, MyString)
DoEvents
End If
GetNext:
   Next i
   
 If strReport <> vbNullString Then
  frmReport.Rich1.Text = strReport
     frmReport.Show 1
  End If
  Call MsgBox("The CASE of Rolling-stock Names have all been fixed, you may now close this screen.", vbInformation, App.Title)
  MousePointer = 0
  
  Exit Sub
Errtrap:

  Resume Next
End Sub


Private Sub cmdFixFix_Click()
Dim strPath As String, strNew As String
Dim strTrains As String
Dim GlobalSoundPath As String, Z As Long
Dim jj As Integer

On Error GoTo Errtrap

MousePointer = 11
strReport = vbNullString
strTrains = MSTSPath & "\trains\trainset\"
GlobalSoundPath = MSTSPath & "\Sound"
Rem ******* Globalsoundpath

For i = 0 To List1.Rows - 1

If List1.IsSelected(i) = True Then

lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)

 
   If Right$(strPath, 4) <> ".sms" Then GoTo NextOne
   
   strNew = ReadUniFile(strPath)
  Z = InStr(strNew, "../../../../../../")
  If Z = 0 Then GoTo NextOne
   DoEvents
   For jj = 1 To 5
strNew = Replace(strNew, "../../../../../../", "../../")
DoEvents
strNew = Replace(strNew, "../../../", "../../")
DoEvents
Next jj

CarryON:

    
   Call WriteUniFile(strPath, strNew)
   DoEvents
 End If
NextOne:

   Next i
   MousePointer = 0
   If strReport <> vbNullString Then
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   End If
   Call MsgBox(".SMS files have all been checked, you may now close this screen.", vbInformation, App.Title)
   Exit Sub
Errtrap:
  
   strReport = strReport & "An unresolved error occurred while processing " & strPath & vbCrLf & vbCrLf
   GoTo NextOne
   
   
   
End Sub

Private Sub cmdFixSD_Click()
Dim i As Long, strPath As String, strCorrShape As String
Dim x As Integer, booLong As Boolean, MyString As String, strShapeFolder As String
Dim xx As Integer, strStart As String, strEnd As String, strTemp As String, booSame As Boolean

lblCount(0).Caption = Str(Val(lblCount(0).Caption) - 1)
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)

x = InStrRev(strPath, "\")
strCorrShape = Mid$(strPath, x + 1)
strCorrShape = Left$(strCorrShape, Len(strCorrShape) - 1)
x = InStrRev(strPath, "\", x - 2)
strShapeFolder = Mid(strPath, x + 1)

If Right$(strPath, 3) <> ".sd" Then GoTo GetNext
MyString = ReadUniFile(strPath)
Rem *********************

xx = InStr(MyString, "Shape")
If xx = 0 Then

strReport = strReport & strShapeFolder & "d does not have a valid 'Shape' entry and must be corrected" & vbCrLf
GoTo FinishIt
End If
x = InStr(xx, MyString, ".s")
If x = 0 Then

strReport = strReport & strShapeFolder & " does not have a valid 'Shape' entry and must be corrected" & vbCrLf
GoTo FinishIt
End If
strStart = Left$(MyString, xx - 1)
strEnd = Mid$(MyString, x + 2)
strTemp = Mid$(MyString, xx + 7, x + 2 - (xx + 7))

strTemp = Trim$(strTemp)
If Left$(strTemp, 1) = ChrW$(34) Then
strTemp = Mid$(strTemp, 2)
End If

Call CompStrings(strTemp, strCorrShape, booSame)
If booSame = False Then
booChanged = True
MyString = strStart & "Shape ( " & strCorrShape & strEnd

End If

xx = InStr(MyString, vbCrLf & vbCrLf & "Shape")
If xx = 0 Then
xx = InStr(MyString, vbCr)
x = InStr(MyString, "Shape")
strStart = Left$(MyString, xx - 1)
strEnd = Mid$(MyString, x)
MyString = strStart & vbCrLf & vbCrLf & strEnd
booChanged = True
End If

FinishIt:
End If


If booChanged = True Then
booChanged = False
Call WriteUniFile(strPath, MyString)

End If


Rem ***********************************

DoEvents
GetNext:
Next i
If strReport <> vbNullString Then
  frmReport.Rich1.Text = strReport
     frmReport.Show 1
  End If
  DoEvents
  If booLong = False Then
Call MsgBox(".SD files have all been fixed, you may now close this screen.", vbInformation, App.Title)
MousePointer = 0

End If
End Sub
Private Sub cmdFixSms_Click()
Dim strPath As String, strNew As String, x As Long, Y As Long, strSound As String
Dim NewFile As Integer, yy As Long, strSoundName As String, strTrains As String
Dim GlobalSoundPath As String, Z As Long, booAlias As Boolean, xq As Long
Dim strThisPath As String, strTemp As String, intLevel As Integer, zz As Long
Dim FirstPass As Integer, strUni As String, zy As Long, q As Long, strAliasPath As String
Dim strAlias As String, intDots As Integer, yz As Long, strDots As String, ix As Long
Dim strStart As String, strEnd As String, strAlias2 As String, booCommon As Boolean
Dim strAPath(0 To 4) As String, j As Integer
Dim jj As Integer, s As Long, booFix As Boolean, strBad As String


On Error GoTo Errtrap
lblCount(0).Caption = Str(Val(lblCount(0).Caption) - 1)
MousePointer = 11
strReport = vbNullString
strTrains = MSTSPath & "\trains\trainset\"
GlobalSoundPath = MSTSPath & "\Sound"
Rem ******* Globalsoundpath

For i = 0 To List1.Rows - 1

If List1.IsSelected(i) = True Then

intLevel = 0
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)

   Z = InStrRev(strPath, "\")
   x = InStr(strPath, "common")
   If x = 0 Then
   x = InStr(strPath, "EB_Sound")
   End If
   xq = InStr(strPath, "Ctn-a-Sound")
   If xq > 0 Then
   x = xq
   End If
   If x > 0 Then
   booCommon = True

   Else
   booCommon = False
   End If
   
   If Mid$(strPath, Z - 5, 5) <> "Sound" And x = 0 Then GoTo NextOne
   strThisPath = Left$(strPath, Z)
   zz = InStr(strThisPath, "trainset")
   zz = InStr(zz, strThisPath, "\")
   DoEvents
TryAgain:
   If zz > 0 Then
   intLevel = intLevel + 1
   zz = InStr(zz + 1, strThisPath, "\")
   GoTo TryAgain
   End If
  DoEvents
   intLevel = intLevel - 1
   If intLevel = 3 Then
   strTemp = "../"
   ElseIf intLevel = 4 Then
   strTemp = "../../"
   Else
   strTemp = vbNullString
   End If
   If booCommon = True Then

   For j = 0 To 3
   strAPath(j) = vbNullString
   Next j
   zz = InStrRev(strPath, "\")
   strAPath(0) = Left$(strPath, zz)
   zz = InStrRev(strPath, "\", zz - 1)
   strAPath(1) = Left$(strPath, zz)
   zz = InStrRev(strPath, "\", zz - 1)
   strAPath(2) = Left$(strPath, zz)
   zz = InStrRev(strPath, "\", zz - 1)
   strAPath(3) = Left$(strPath, zz)
   zz = InStrRev(strPath, "\", zz - 1)
   strAPath(4) = Left$(strPath, zz)
   
   End If
   
    Open strPath For Binary As #2
    strUni = String(2, " ")
    Get #2, , strUni
     Close #2

 If Asc(Mid$(strUni, 1, 1)) <> 255 And Asc(Mid$(strUni, 2, 1)) <> 254 Then
 Call ConvertIt(strPath, 1)
 DoEvents
 End If
 
   If Right$(strPath, 4) <> ".sms" Then GoTo NextOne
   
   strNew = ReadUniFile(strPath)
   strNew = Replace(strNew, "\\", "/")
   DoEvents
   strNew = Replace(strNew, "//", "/")
   DoEvents
   strNew = Replace(strNew, ChrW$(34) & "/../", ChrW$(34) & "../../")
   DoEvents
  ' strNew = Replace(strNew, vbtab, " ")
   DoEvents
   strNew = Replace(strNew, "     ", " ")
   DoEvents
   strNew = Replace(strNew, "File (" & ChrW$(34), "File ( " & ChrW$(34))
   DoEvents
   Rem ********** Fix for Mike
   strNew = Replace(strNew, "File ( " & ChrW$(34) & " " & ChrW$(34), "File ( " & ChrW$(34))
   DoEvents
   q = 1

   x = 0
   FirstPass = 0
FindMore:
booAlias = False

FirstPass = FirstPass + 1
   x = InStr(x + 1, strNew, "File")
   If x <> 0 Then
   If Mid$(strNew, x, 8) = "filename" Then
   x = x + 1
   GoTo FindMore
   End If
   Rem ******************* Skip ********
   If x <> 0 Then
   s = InStrRev(strNew, "Skip", x)
   If s > 0 And s < 200 Then
   GoTo FindMore
   End If
   End If
   Rem *********************************
   If Mid$(strNew, x, 6) <> "File (" Then
   
   zy = InStr(x, strNew, "(")
   
   strNew = Left$(strNew, x + 3) & " " & Mid$(strNew, zy)
   
   End If
   If Mid$(strNew, x + 6, 1) <> " " Then
   strNew = Left$(strNew, x + 5) & " " & Mid$(strNew, x + 6)
   End If
   End If
         If x = 0 Then
         If FirstPass = 1 Then
         GoTo NextOne
         Else
         GoTo CarryON
         End If
         End If


         Y = InStr(x, strNew, ChrW$(34))
         If Y = 0 Or Y - x > 10 Then '***** No quotes ******
         
         yz = InStr(x, strNew, ".wav")
         If yz - x > 100 Then GoTo NextOne
         strSound = Mid$(strNew, x + 7, (yz + 4) - (x + 7))
         strSound = Trim$(strSound)
         strNew = Left$(strNew, x + 5) & " " & ChrW$(34) & strSound & ChrW$(34) & Mid$(strNew, yz + 4)
         
         GoTo NextBit
         End If
        
NextBit:
         yy = InStr(Y + 1, strNew, ChrW$(34))
         strSound = Trim$(Mid$(strNew, Y + 1, yy - (Y + 1)))


strStart = Left$(strNew, Y)
strEnd = Mid$(strNew, yy)

   strAlias = vbNullString
    If Left$(strSound, 3) = "../" Then

       Z = InStrRev(strSound, "/")
            strSoundName = Mid$(strSound, Z + 1)
            strAlias = Left$(strSound, Z)
            booAlias = True
            Else
            strSoundName = strSound
            End If
If strAlias <> vbNullString And booCommon = False Then
        intDots = 0
        Call CountDots(strAlias, intDots)
        strDots = vbNullString
       For ix = 1 To intLevel
       strDots = strDots & "../"
       Next ix
Alias1:
       If Left$(strAlias, 3) = "../" Then
       strAlias = Mid$(strAlias, 4)
       GoTo Alias1
       End If
       strNew = strStart & strDots & strAlias & strSoundName & strEnd
        If Left$(strAlias, 5) = "Sound" Then
             strAliasPath = GlobalSoundPath
            Else
            strAlias2 = Replace(strAlias, "/", "\")
            strAliasPath = strTrains & strAlias2
        End If
End If
If strAlias <> vbNullString And booCommon = True Then

        intDots = 0
        Call CountDots(strAlias, intDots)
        
        strDots = vbNullString

Alias2:
       If Left$(strAlias, 3) = "../" Then
       strAlias = Mid$(strAlias, 4)
       GoTo Alias2
       End If
       For j = 0 To 4
       A$ = Replace(strAlias, "/", "\")
       strBad = strAPath(intDots) & A$ & strSoundName
       If FileExists(strAPath(j) & A$ & strSoundName) Then
        For jj = 1 To j
        strDots = strDots & "../"
        Next jj
       
       strNew = strStart & strDots & strAlias & strSoundName & strEnd
        
            strAlias2 = Replace(strAlias, "/", "\")
            strAliasPath = strTrains & strAlias2
            booFix = True
            Exit For
        End If
        Next j
        If booFix = False Then
        strReport = strReport & strSoundName & " Is missing from " & strBad & " - Called by " & strPath & vbCrLf & vbCrLf
        GoTo NextBit2
      
        End If
End If
        If booAlias = False Then
         If Not FileExists(strThisPath & strSoundName) Then

     
                If FileExists(GlobalSoundPath & "\" & strSoundName) Then
                strNew = strStart & strSoundName & strEnd
                ElseIf FileExists(strTrains & "Dash9\Sound\" & strSoundName) Then
                strNew = strStart & "../../Dash9/Sound/" & strSoundName & strEnd
                ElseIf FileExists(strTrains & "380\Sound\" & strSoundName) Then
                strNew = strStart & "../../380/Sound/" & strSoundName & strEnd
                ElseIf FileExists(strTrains & "Scotsman\Sound\" & strSoundName) Then
                strNew = strStart & "../../Scotsman/Sound/" & strSoundName & strEnd
                ElseIf FileExists(strTrains & "GP38\Sound\" & strSoundName) Then
                strNew = strStart & "../../GP38/Sound/" & strSoundName & strEnd
                ElseIf FileExists(strTrains & "Kiha31\Sound\" & strSoundName) Then
                strNew = strStart & "../../Kiha31/Sound/" & strSoundName & strEnd
                ElseIf FileExists(strTrains & "Series2000\Sound\" & strSoundName) Then
                strNew = strStart & "../../SEries2000/Sound/" & strSoundName & strEnd
                ElseIf FileExists(strTrains & "Series7000\Sound\" & strSoundName) Then
                strNew = strStart & "../../Series7000/Sound/" & strSoundName & strEnd
                ElseIf FileExists(strTrains & "SD402\Sound\" & strSoundName) Then
                strNew = strStart & "../../SD402/Sound/" & strSoundName & strEnd
                ElseIf FileExists(strTrains & "Acela\Sound\" & strSoundName) Then
               
                strNew = strStart & "../../Acela/Sound/" & strSoundName & strEnd
                Else
                
                strReport = strReport & strSoundName & " Is missing from " & strPath & vbCrLf & vbCrLf
                End If
                End If
                End If
         If booAlias = True And booCommon = False Then
                If Not FileExists(strAliasPath & strSoundName) Then
                If FileExists(GlobalSoundPath & "\" & strSoundName) Then
                
               
                strNew = strStart & strSoundName & strEnd
                Else
              
                strReport = strReport & strSoundName & " Is missing from " & strAliasPath & "Called by " & strPath & vbCrLf & vbCrLf
                End If
                End If
        
         End If
NextBit2:
         If x = 0 Then
         x = yy
         End If
         GoTo FindMore
         End If
        
   
CarryON:
If strNew = vbNullString Then GoTo NextOne
    NewFile = FreeFile
    
    
    Open strPath For Output As #NewFile
   Print #NewFile, strNew
    Close #NewFile
    DoEvents
    
   Call ConvertIt(strPath, 1)
   DoEvents
  
NextOne:

   Next i
   MousePointer = 0
   If strReport <> vbNullString Then
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   End If
   Call MsgBox(".SMS files have all been checked, you may now close this screen.", vbInformation, App.Title)
   Exit Sub
Errtrap:
  
   strReport = strReport & "An unresolved error occurred while processing " & strPath & vbCrLf & vbCrLf
   GoTo NextOne
   
   
   
   
End Sub

Private Sub cmdFixSrv_Click()
Dim SrvPath As String, flagway As Integer
Dim MyString As String, x As Integer, strSrvPath As String, strSrvName As String
Dim strSrv As String, booSame As Boolean, strStart As String, strEnd As String, i As Integer, ii As Integer
Dim strTemp As String, strTemp2 As String, booConexists As Boolean, j As Integer

lblCount(0).Caption = Str(Val(lblCount(0).Caption) - 1)
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
SrvPath = List1.TextMatrix(i, 0)
x = InStrRev(SrvPath, "\")
strSrvPath = Left$(SrvPath, x)
strSrvName = Mid$(SrvPath, x + 1)


If Right$(strSrvPath, 9) <> "Services\" Then
GoTo TryAgain
End If

   flagway = 0
   strSrv = Left$(strSrvName, Len(strSrvName) - 4)
   MyString = ReadUniFile(SrvPath)
   DoEvents

   x = InStr(MyString, "Train_Config")
            If x > 0 Then
            j = InStr(x, MyString, "(")
            xx = InStr(j, MyString, vbCr)
            xx = InStrRev(MyString, ")", xx)
           
            strSrv = Mid$(MyString, j + 1, xx - j - 1)
            End If
strSrv = Trim$(strSrv)
strSrv = Replace(strSrv, ChrW$(34), "")

strStart = Left$(MyString, j)
strEnd = Mid$(MyString, xx)
booConexists = False

    For ii = 0 To UBound(Consists)
            If strSrv & ".con" = Consists(ii) Then
            strTemp2 = Consists(ii)
            booConexists = True
            Exit For
            End If
    Next ii

If booConexists = False Then
strReport = strReport & strSrv & ".con called by " & SrvPath & " is missing from your Consists folder" & vbCrLf & vbCrLf
Else
booSame = False
Call FixSrv(strSrv & ".con", strTemp2, booSame)
    If booSame = False Then
        strTemp = Left$(strTemp2, Len(strTemp2) - 4)
        strTemp = Trim(strTemp)
'        j = InStr(strTemp, " ")
'        If j > 0 Then
'        j = 0
        MyString = strStart & " " & ChrW$(34) & strTemp & ChrW$(34) & " " & strEnd
'        Else
'        MyString = strStart & " " & strTemp & " " & strEnd
'        End If
        Call WriteUniFile(SrvPath, MyString)
        DoEvents
    End If
End If
TryAgain:
End If
Next i
If strReport <> vbNullString Then
  frmReport.Rich1.Text = strReport
     frmReport.Show 1
  End If
  DoEvents

  Call MsgBox("Service Names have all been fixed, you may now close this screen.", vbInformation, App.Title)
MousePointer = 0
End Sub


Private Sub Command1_Click(Index As Integer)
 Dim flagway As Integer, NewFile2 As Integer
Dim strFound As String, strTemp As String, strBatText As String
Dim i As Integer, x As Integer, strPath As String

Select Case Index
Case 0

For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
List1.TopRow = i

   strFound = List1.TextMatrix(i, 0)
   List1.IsSelected(i) = False


 x = InStrRev(strFound, "\")
 strPath = Left$(strFound, x - 1)
   If Right$(strFound, 2) = ".s" Then
  

  strPicView = strFound

  strBatText = ChrW$(34) & App.Path & "\sviewRR.exe" & ChrW$(34) & " " & ChrW$(34) & strPicView & ChrW$(34) & ";"
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)

  
  DoEvents
  
ElseIf Right$(strFound, 4) = ".ace" Then

   strPicView = strFound
  frmPicView.Show 1
  DoEvents
  
 Else
 Rem **************** Unicode ************************


Close
  
   
        fullpath$ = strFound
        
   NewFile2 = FreeFile
   Open fullpath$ For Binary As #NewFile2
    strTemp = String(2, " ")
    Get #NewFile2, , strTemp
 Close #NewFile2
 
 If Asc(Mid$(strTemp, 1, 1)) = 255 And Asc(Mid$(strTemp, 2, 1)) = 254 Then
 
   flagway = 0
   Call ConvertIt(fullpath$, flagway)
   DoEvents
   booUniEdit = True
   strUniName = fullpath$
   frmReport.Rich1.LoadFile fullpath$
   frmReport.Show 1
 DoEvents

 flagway = 1
   Call ConvertIt(fullpath$, flagway)
   
    Else
    Call MsgBox("File " & fullpath$ & vbCrLf & Lang(455), vbExclamation, App.Title)
                
    GoTo Another
    End If
    
Another:

 Rem **************** End Unicode ********************
   End If
   End If
   Next i
  
   frmSearch.WindowState = 0
   frmUtils.WindowState = 0
  frmSearch.ZOrder
  
Case 1
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
List1.IsSelected(i) = False
End If
Next i
lblCount(1).Caption = List1.SelectedRows
Case 2
'frmUtils.Text1(0).Text = "*.*"
If Command1(2).Caption = "&Exit" Or Command1(2).Caption = "Exit" Then
booAbort = False
Unload Me
ElseIf Command1(2).Caption = "Abort" Then
booAbort = True
Unload Me
End If
End Select

End Sub



Private Function ConvertSD2(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertSD2 = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer, Y As Integer, Z As Integer

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
  mytristate = 0
'    MsgBox chrw$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34) & Lang(403), vbInformation, frmUtils.Caption
'    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
xx = InStr(MyString, "ESD_Bounding_Box")
If xx = 0 Then
strReport = strReport & CompleteFilePath & " does not have a valid 'Bounding Box' entry" & vbCrLf
GoTo CarryON
End If
x = InStr(xx, MyString, "(")
If Mid$(MyString, x + 1, 1) = " " Then x = x + 2
Y = InStr(x, MyString, " ")
Z = InStr(Y + 1, MyString, " ")

strStart = Left$(MyString, Y)
strEnd = Mid$(MyString, Z)

MyString = strStart & strBBoxFix & strEnd
CarryON:
End If


'End If
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
ConvertSD2 = True

ERRHANDLER:

  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function


Private Function ConvertAnim(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertAnim = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String
Dim xx As Long

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
  mytristate = 0
'    MsgBox chrw$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34) & Lang(403), vbInformation, frmUtils.Caption
'    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
xx = InStr(MyString, "Animations")
If xx = 0 Then
'strReport = strReport & CompleteFilePath & " does not have an animation entry" & vbCrLf
GoTo FinishIt
Else
strReport = strReport & CompleteFilePath & " HAD an animation entry" & vbCrLf
lAnim = lAnim + 1
''End If
strStart = Left$(MyString, xx - 1)
strEnd = vbCrLf & ")"

MyString = strStart & strEnd
'strReport = strReport & "The Shape entry was corrected in " & CompleteFilePath & vbCrLf
booAnimRemoved = True
End If


FinishIt:
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

'If booChanged = True Then
FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertAnim = True
'ElseIf booChanged = False Then
'Kill tempfile
'ConvertAnim = False
'End If
ERRHANDLER:

  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
 ' Resume Next
End Function


Private Sub Command10_Click()
Dim strFound As String
Dim i As Integer, x As Integer, strPath As String, strName As String
Dim strTemp As String, flagway As Integer
Dim Ename As String, booDidNotComp As Boolean, strOrigFile As String


On Error GoTo Errtrap

Select Case MsgBox("DO NOT REMOVE THE ANIMATIONS FROM ELECTRIC LOCOS WITH PANTOGRAPHS" _
                   & vbCrLf & "OTHERWISE THE PANTOGRAPHS WILL CEASE WORKING." _
                   , vbOKCancel Or vbCritical Or vbDefaultButton1, "Warning")

    Case vbOK

    Case vbCancel
Exit Sub
End Select

MousePointer = 11
If Not DirExists(App.Path & "\Tempfiles") Then
MkDir App.Path & "\Tempfiles"
End If
strReport = vbNullString
Label6.Visible = True
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strFound = List1.TextMatrix(i, 0)
If strFound = vbNullString Then GoTo NextOne
 x = InStrRev(strFound, "\")
 strPath = Left$(strFound, x - 1)
 strName = Mid$(strFound, x + 1)
   If Right$(strFound, 2) = ".s" Then
      fullpath$ = strFound
   strOrigFile = strName
   Label6.Caption = "Checking: " & strOrigFile
   Label6.Refresh
   If Right$(strOrigFile, 2) <> ".s" Then
   MousePointer = 0
   Call MsgBox(Lang(393) & strOrigFile & Lang(410), vbExclamation, Lang(404))
    GoTo NextOne
    End If

Rem ************************************

   Open fullpath$ For Binary As #5
    strTemp = String(2, " ")
    Get #5, , strTemp
 Close #5

 If Asc(Mid$(strTemp, 1, 1)) = 255 And Asc(Mid$(strTemp, 2, 1)) = 254 Then

   Ename = App.Path & "\TempFiles\" & strOrigFile
FileCopy fullpath$, Ename
DoEvents

GoTo NextOne
End If
Rem *******
TokMode = 0
 DoEvents
 Ename = App.Path & "\TempFiles\" & strOrigFile
   booWriteFile = True
Call DoDeComp2(strOrigFile, strPath, App.Path & "\TempFiles")

'result = tfh.decompress(fullpath$, Ename)
'If result = False Then
'strReport = strReport & fullpath$ & " did not decompress" & vbCrLf
'End If
Rem *********** Process uncompressed file *************
flagway = 0
Call ConvertAnim(Ename, flagway)
flagway = 1
Call ConvertAnim(Ename, flagway)
Rem ***************************************************
If booAnimRemoved = True Then
booAnimRemoved = False
Name fullpath$ As fullpath$ & "x"
'result = tfh.compress(Ename, fullpath$)
booDidNotComp = False
   'Call CompressMe(fullpath$, Ename, booDidNotComp)
   Call DoComp(strOrigFile, strPath, App.Path & "\TempFiles")
Label6.Caption = "Compressing " & strOrigFile
DoEvents
Kill Ename
Else
Kill Ename
DoEvents
End If
   End If
   End If
NextOne:
   Next i
  
   If strReport <> vbNullString Then
frmReport.Rich1.Text = strReport & vbCrLf & vbCrLf & "Total = " & Str(lAnim) & vbCrLf
frmReport.Show 1
Else
strReport = "No selected .S files contained Animation entries"
frmReport.Rich1.Text = strReport
frmReport.Show 1

End If
Label6.Caption = "Finished"
 cursouind = 0
DoEvents
strReport = vbNullString

   frmSearch.WindowState = 0
   frmUtils.WindowState = 0
  frmSearch.ZOrder
  MousePointer = 0
  Exit Sub
  
Errtrap:
If Err = 75 Then
Resume Next
End If
  Select Case MsgBox("Error " & Err & " = " & Err.Description & " occurred in file" _
                     & vbCrLf & strFound _
                     & vbCrLf & "" _
                     , vbOKCancel Or vbExclamation Or vbDefaultButton1, App.Title)
  
    Case vbOK
  Resume Next
    Case vbCancel
    MousePointer = 0
    
  Exit Sub
  End Select
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


Private Sub Command11_Click()
Dim strFullPath As String, x As Integer, strNew2 As String, j As Integer, xx As Integer
Dim flagway As Integer, Y As Integer, strPath As String, strCfg As String, lngShp As Long
Dim i As Long

On Error GoTo Errtrap
ReDim strShape(0 To Wag_CHUNK)
ReDim strShape2(0 To Wag_CHUNK)

For i = 0 To List1.Rows - 1

If List1.IsSelected(i) = True Then
List1.IsSelected(i) = False
List1.TopRow = i
strFullPath = List1.TextMatrix(i, 0)
   flagway = 0
   x = InStrRev(strFullPath, "\")
   If Mid$(strFullPath, x + 1) = "Default.wag" Then GoTo CarryON
   strPath = Left$(strFullPath, x)
   Y = InStrRev(strFullPath, "\", x - 1)
  
   If Mid$(strFullPath, Y - 8, 8) <> "Trainset" Then GoTo CarryON
   strNew2 = ReadUniFile(strFullPath)
x = InStr(strNew2, "Wagonshape")
j = InStr(x, strNew2, "(")
xx = InStr(j, strNew2, ")")
strCfg = Mid$(strNew2, j + 1, xx - (j + 1))
strCfg = Trim$(strCfg)
strCfg = Replace(strCfg, ChrW$(34), " ")
strCfg = Trim$(strCfg)
strShape(lngShp) = strPath & strCfg
lngShp = lngShp + 1
If lngShp > UBound(strShape) Then
           ReDim Preserve strShape(0 To lngShp + Wag_CHUNK)
End If


DoEvents
End If
CarryON:
   Next i
   ReDim Preserve strShape(0 To lngShp - 1)
   
 QSort3 strShape(), 0, lngShp - 1
 DoEvents
 RemD2 strShape(), strShape2()
 lngShp = UBound(strShape2)
 For i = List1.Rows - 1 To 0 Step -1
 List1.RemoveItem i
 DoEvents
 Next i
 DoEvents
 For i = 0 To lngShp
 List1.AddItem strShape2(i)
 Next
 
 If strReport <> vbNullString Then
  frmReport.Rich1.Text = strReport
     frmReport.Show 1
  End If
  
  DoEvents
  strReport = vbNullString
Exit Sub
Errtrap:

Resume Next
End Sub

Public Function QSort3(strList() As Variant, lLbound As Long, lUbound As Long)
    
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
        QSort3 strList(), lLbound, lngCurHigh
    End If
    


    If lngCurLow < lUbound Then ' Recurse if necessary
        QSort3 strList(), lngCurLow, lUbound
    End If
   
End Function

Private Sub Command12_Click()
Dim strFound As String, i As Integer, strOld As String


Label6.Visible = True
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strFound = List1.TextMatrix(i, 0)
    If strFound = vbNullString Then GoTo NextOne
    If Right$(strFound, 2) = "sx" Then
    strOld = Left$(strFound, Len(strFound) - 1)
            If FileExists(strOld) Then
            Kill strOld
            DoEvents
            Name strFound As strOld
            End If
    ElseIf Right$(strFound, 3) = "enx" Then
    strOld = Left$(strFound, Len(strFound) - 1) & "g"
            If FileExists(strOld) Then
            Kill strOld
            DoEvents
            Name strFound As strOld
            End If
    End If
End If
NextOne:
Next
End Sub

Private Sub Command13_Click()
Dim i As Integer, x As Long, MyString As String, fullpath$, Y As Long
Dim strCab As String, strEngpath As String, strCSV As String
Dim booAliased As Boolean, intType As Integer, strTemp As String, strEngname As String
Dim q As Integer

On Error GoTo Errtrap
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
fullpath$ = List1.TextMatrix(i, 0)
 q = 1
       booAliased = False
        
        If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryON
        x = InStrRev(fullpath$, "\")
        strEngpath = Left$(fullpath$, x)
        strEngname = Mid$(fullpath$, x + 1)
        If Left$(strEngname, 1) = "#" Or Left$(strEngname, 1) = "$" Then
        GoTo AILoco
        End If
 q = 2
        MyString = ReadUniFile(fullpath$)
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
q = 3
        x = InStr(MyString, ".cvf")
        If x > 0 Then
        Y = InStrRev(MyString, "(", x)
        strCab = Mid$(MyString, Y + 1, x - (Y + 1) + 4)
        strCab = Trim$(strCab)
        If Left$(strCab, 2) = ".." Then
        booAliased = True
        GoTo AliasedCab
        End If
        q = 4
        If FileExists(strEngpath & "CabView\" & strCab) Then
        strCSV = ReadUniFile(strEngpath & "CabView\" & strCab)
        x = InStr(strCSV, "combinedcontrol")
        If x > 0 Then
        intType = 1
        GoTo CarryON
        End If
        Else
        booAliased = True
        End If
        End If
AliasedCab:
q = 5
        x = InStr(MyString, "GearBox")
        If x > 0 Then
        intType = 2
        GoTo CarryON
        End If
        x = InStr(MyString, "Type ( Electric")
        If x > 0 Then
        intType = 4
        GoTo CarryON
        End If
        x = InStr(MyString, "Type ( Steam")
        If x > 0 Then
        intType = 3
        GoTo CarryON
        End If
        x = InStr(MyString, "Type ( Diesel")
        If x > 0 Then
        intType = 5
        GoTo CarryON
        End If
        If x = 0 Then
        intType = 0
        End If
 q = 6
CarryON:
Select Case intType
Case 0
q = 7
strReport = strReport & fullpath$ & " not processed due to indetermined Loco type." & vbCrLf
Case 1   ' Combo
q = 8
If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryOn1
        MyString = ReadUniFile(fullpath$)
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

        MyASCFile = ReadASCIIFile(App.Path & "\RDCombo1.txt")
        Rem ********* Get Brake force *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strBrakeForce = Mid$(MyString, x, (xx + 1) - x)
        '*************Brake Equipt************************
        x = InStr(MyString, "BrakeEquipmentType")
        xx = InStr(MyString, "BrakeCylinderPressureForMaxBrakeBrakeForce")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '*************Vacuum****************************
        MyASCFile = ReadASCIIFile(App.Path & "\RDCombo2.txt")
        x = InStr(Y, MyString, "AirBrakesAirCompressorPowerRating")
        If x = 0 Then                 'This loco is not vacuum braked*****************
        x = InStr(Y, MyString, "TrainBrakesControllerMinPressureReduction")
        End If
        If x = 0 Then GoTo EngControl1
        xx = InStr(MyString, "BrakesEngineControllers")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '***************EngineControllers ***************
EngControl1:
q = 9
        MyASCFile = ReadASCIIFile(App.Path & "\RDCombo3.txt")
        Y = InStr(MyString, "BrakesEngineControllers")
        x = InStr(Y + 20, MyString, "EngineControllers")
        xx = InStr(x, MyString, "Headlights")
        If xx > 0 Then
        Y = InStr(xx, MyString, "Wipers (")
        If Y > 0 Then
        xx = Y
        End If
        End If
        If xx <> 0 Then
        xx = InStr(xx, MyString, ")")
        xx = InStr(xx + 1, MyString, ")")
        xx = xx + 1
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "Sound")
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "Name")
        End If
        If x > 0 And xx > x Then
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx)
        MyString = strStart & MyASCFile & strEnd
        '************ Replace MaxBrakeForce *************
        x = 1
CheckAgain1:
q = 10
        x = InStr(x, MyString, "MaxBrakeForce")
        If x = 0 Then GoTo GotIt1
        xx = InStr(x, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx + 1)
        MyString = strStart & strBrakeForce & strEnd
        x = x + 1
        GoTo CheckAgain1
GotIt1:
 q = 11
        '***********************************************
        strTemp = Left$(fullpath$, Len(fullpath$) - 1) & "x"
        If FileExists(strTemp) Then
        Kill strTemp
        DoEvents
        End If
        Name fullpath$ As strTemp
        DoEvents
        Call WriteUniFile(fullpath$, MyString)
   DoEvents
      End If
  ' End If
CarryOn1:
q = 12
If booWrong = True Then
booWrong = False
strReport = strReport & fullpath$ & " was not an Combo loco" & vbCrLf
End If
Case 2   ' Geared
q = 13
If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryOn2
        MyString = ReadUniFile(fullpath$)
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
        x = InStr(MyString, "GearBox")
        If x = 0 Then booWrong = True: GoTo CarryOn2
        MyASCFile = ReadASCIIFile(App.Path & "\RDGeared1.txt")
        Rem ********* Get Brake force *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strBrakeForce = Mid$(MyString, x, (xx + 1) - x)
        '*************Brake Equipt************************
        x = InStr(MyString, "BrakeEquipmentType")
        xx = InStr(MyString, "BrakeCylinderPressureForMaxBrakeBrakeForce")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '*************Brakes****************************
        MyASCFile = ReadASCIIFile(App.Path & "\RDGeared2.txt")
        x = InStr(Y, MyString, "TrainBrakesControllerMaxApplicationRate")
        If x = 0 Then
        x = InStr(Y, MyString, "TrainBrakesControllerMaxReleaseRate")
        End If
        If x = 0 Then GoTo EngControl2
        xx = InStr(MyString, "BrakesEngineControllers")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '***************EngineControllers ***************
EngControl2:
q = 14
        MyASCFile = ReadASCIIFile(App.Path & "\RDGeared3.txt")
        Y = InStr(MyString, "BrakesEngineControllers")
        x = InStr(Y + 20, MyString, "EngineControllers")
        xx = InStr(x, MyString, "Headlights")
        If xx <> 0 Then
        xx = InStr(xx, MyString, ")")
        xx = InStr(xx + 1, MyString, ")")
        xx = xx + 1
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "Name")
        End If
        If x > 0 And xx > x Then
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx)
        MyString = strStart & MyASCFile & strEnd
        '************ Replace MaxBrakeForce *************
        x = 1
CheckAgain2:
q = 15
        x = InStr(x, MyString, "MaxBrakeForce")
        If x = 0 Then GoTo GotIt2
        xx = InStr(x, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx + 1)
        MyString = strStart & strBrakeForce & strEnd
        x = x + 1
        GoTo CheckAgain2
GotIt2:
  q = 16
        '***********************************************
        strTemp = Left$(fullpath$, Len(fullpath$) - 1) & "x"
        If FileExists(strTemp) Then
        Kill strTemp
        DoEvents
        End If
        Name fullpath$ As strTemp
        DoEvents
        Call WriteUniFile(fullpath$, MyString)
   DoEvents
      End If
   'End If
CarryOn2:
q = 17
If booWrong = True Then
booWrong = False
strReport = strReport & fullpath$ & " was not an Geared loco" & vbCrLf
End If
Case 3   ' Steam
q = 18
If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryOn3
        MyString = ReadUniFile(fullpath$)
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
        x = InStr(MyString, "Type ( Steam")
        If x = 0 Then booWrong = True: GoTo CarryOn3
        MyASCFile = ReadASCIIFile(App.Path & "\RDSteam1.txt")
        Rem ********* Get Brake force *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strBrakeForce = Mid$(MyString, x, (xx + 1) - x)
        '*************Brake Equipt************************
        x = InStr(MyString, "BrakeEquipmentType")
        xx = InStr(MyString, "BrakeCylinderPressureForMaxBrakeBrakeForce")
        Y = InStr(xx, MyString, ")")
        
        
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '*************Vacuum****************************
        q = 19
        MyASCFile = ReadASCIIFile(App.Path & "\RDSteam2.txt")
        x = InStr(Y, MyString, "VacuumBrakes")
        If x = 0 Then
        x = InStr(Y, MyString, "AirBrakesAirCompressorPowerRating")
        End If
        xx = InStr(MyString, "BrakesEngineControllers")
        If xx = 0 Then
        xx = InStr(MyString, "EngineBrakesControllerMaxSystemPressure")
        End If
        If xx = 0 Then booWrong = True: GoTo CarryOn3
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '***************EngineControllers ***************
        MyASCFile = ReadASCIIFile(App.Path & "\RDSteam3.txt")
        Y = InStr(MyString, "BrakesEngineControllers")
        x = InStr(Y + 20, MyString, "EngineControllers")
        xx = InStr(x, MyString, "Brake_Hand")
        If xx <> 0 Then
        xx = InStr(xx, MyString, ")")
        xx = InStr(xx + 1, MyString, ")")
        xx = xx + 1
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "FireDoor")
        End If
        If x > 0 And xx > x Then
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx)
        MyString = strStart & MyASCFile & strEnd
        '************ Replace MaxBrakeForce *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx + 1)
        MyString = strStart & strBrakeForce & strEnd
        '***********************************************
        strTemp = Left$(fullpath$, Len(fullpath$) - 1) & "x"
        If FileExists(strTemp) Then
        Kill strTemp
        DoEvents
        End If
        Name fullpath$ As strTemp
        DoEvents
        Call WriteUniFile(fullpath$, MyString)
   DoEvents
      End If
  ' End If
CarryOn3:
q = 20
If booWrong = True Then
booWrong = False
strReport = strReport & fullpath$ & " had incorrect syntax - no changes were made" & vbCrLf
End If
Case 4   ' Electric
q = 21
If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryOn4
        MyString = ReadUniFile(fullpath$)
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
        x = InStr(MyString, "Type ( Electric")
        If x = 0 Then
        booWrong = True
        strWrong = " does not appear to be an Electric loco "
        GoTo CarryOn4
        End If
        MyASCFile = ReadASCIIFile(App.Path & "\RDElectric1.txt")
        Rem ********* Get Brake force *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strBrakeForce = Mid$(MyString, x, (xx + 1) - x)
        '*************Brake Equipt************************
        x = InStr(MyString, "BrakeEquipmentType")
        xx = InStr(MyString, "BrakeDistributorNormalFullReleasePressure")
        If xx = 0 Then
        xx = InStr(MyString, "BrakeCylinderPressureForMaxBrakeBrakeForce")
        End If
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '*************Vacuum****************************
        q = 22
        MyASCFile = ReadASCIIFile(App.Path & "\RDElectric2.txt")
        x = InStr(Y, MyString, "VacuumBrakes")
        If x = 0 Then                 'This loco is not vacuum braked*****************
        x = InStr(Y, MyString, "AirBrakesAirCompressorPowerRating")
        End If
        If x = 0 Then                 'This loco is not vacuum braked*****************
        x = InStr(Y, MyString, "EngineBrakesControllerDirectControlExponent")
        End If
        xx = InStr(MyString, "BrakesEngineControllers")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '***************EngineControllers ***************
        MyASCFile = ReadASCIIFile(App.Path & "\RDElectric3.txt")
        Y = InStr(MyString, "BrakesEngineControllers")
        x = InStr(Y + 20, MyString, "EngineControllers")
        xx = InStr(x, MyString, "Brake_Hand")
        If xx <> 0 Then
        xx = InStr(xx, MyString, ")")
        xx = InStr(xx + 1, MyString, ")")
        xx = xx + 1
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "DirControl")
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "EmergencyStopResetToggle")
        End If
        If xx = 0 Then
        booWrong = True
        strWrong = " could not be modified automatically"
        GoTo CarryOn4
        End If
        If x > 0 And xx > x Then
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx)
        MyString = strStart & MyASCFile & strEnd
        '************ Replace MaxBrakeForce *************
        x = 1
        q = 23
CheckAgain4:
        x = InStr(x, MyString, "MaxBrakeForce")
        If x = 0 Then GoTo GotIt4
        xx = InStr(x, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx + 1)
        MyString = strStart & strBrakeForce & strEnd
        x = x + 1
        GoTo CheckAgain4
GotIt4:
           
        '***********************************************
        strTemp = Left$(fullpath$, Len(fullpath$) - 1) & "x"
        If FileExists(strTemp) Then
        Kill strTemp
        DoEvents
        End If
        Name fullpath$ As strTemp
        DoEvents
        Call WriteUniFile(fullpath$, MyString)
   DoEvents
      End If
   'End If
CarryOn4:
If booWrong = True Then
booWrong = False
strReport = strReport & fullpath$ & strWrong & vbCrLf
End If
Case 5   ' Diesel
q = 24
If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryOn5

        MyString = ReadUniFile(fullpath$)
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
        x = InStr(MyString, "Type ( Diesel")
        If x = 0 Then booWrong = True: GoTo CarryOn5
        MyASCFile = ReadASCIIFile(App.Path & "\RDDiesel1.txt")
        Rem ********* Get Brake force *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strBrakeForce = Mid$(MyString, x, (xx + 1) - x)
        '*************Brake Equipt************************
        x = InStr(MyString, "BrakeEquipmentType")
        xx = InStr(MyString, "BrakeCylinderPressureForMaxBrakeBrakeForce")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '*************Vacuum****************************
        MyASCFile = ReadASCIIFile(App.Path & "\RDDiesel2.txt")
        x = InStr(Y, MyString, "AirBrakesAirCompressorPowerRating")
        If x = 0 Then                 'This loco is not vacuum braked*****************
        x = InStr(Y, MyString, "EngineBrakesControllerMinPressureReduction")
        End If
        If x = 0 Then GoTo EngControl5
        xx = InStr(MyString, "BrakesEngineControllers")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '***************EngineControllers ***************
EngControl5:
q = 25
        MyASCFile = ReadASCIIFile(App.Path & "\RDDiesel3.txt")
        Y = InStr(MyString, "BrakesEngineControllers")
        x = InStr(Y + 20, MyString, "EngineControllers")
        If x = 0 Then booWrong = True: GoTo CarryOn5
        xx = InStr(x, MyString, "BailOffButton")
        If xx <> 0 Then
        xx = InStr(xx, MyString, ")")
        xx = InStr(xx + 1, MyString, ")")
        xx = xx + 1
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "Sound")
        End If
        If x > 0 And xx > x Then
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx)
        MyString = strStart & MyASCFile & strEnd
        '************ Replace MaxBrakeForce *************
        x = 1
CheckAgain5:
        x = InStr(x, MyString, "MaxBrakeForce")
        If x = 0 Then GoTo GotIt5
        xx = InStr(x, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx + 1)
        MyString = strStart & strBrakeForce & strEnd
        x = x + 1
        GoTo CheckAgain5
GotIt5:
           
        '***********************************************
        strTemp = Left$(fullpath$, Len(fullpath$) - 1) & "x"
        If FileExists(strTemp) Then
        Kill strTemp
        DoEvents
        End If
        Name fullpath$ As strTemp
        DoEvents
        Call WriteUniFile(fullpath$, MyString)
   DoEvents
      End If
   'End If
CarryOn5:
If booWrong = True Then
booWrong = False
strReport = strReport & fullpath$ & " was not a Diesel loco" & vbCrLf
End If

End Select
AILoco:

  End If
  
Next i
q = 26
 If strReport <> vbNullString Then
frmReport.Rich1.Text = strReport
 
     frmReport.Show 1

End If
Exit Sub
Errtrap:
Call MsgBox("An error #" & Err & " " & Err.Description & " occurred when checking " _
            & vbCrLf & strEngname & "Q = " & Str(q) _
            , vbExclamation, App.Title)
Resume Next
        
End Sub


Private Sub Command14_Click()
Dim strPath As String, flagway As Integer, strAI As String, strEng As String
Dim x As Integer, strOld As String

If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
SparePath = App.Path & "\TempFiles"
frmAI.Show vbModal

     DoEvents
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)
x = InStrRev(strPath, "\")
strOld = Left$(strPath, x)
strEng = Mid$(strPath, x + 1)

   If Right$(strPath, 4) <> ".eng" Then
   Call MsgBox(Lang(393) & strPath & vbCrLf & Lang(456), vbExclamation, App.Title)
   
   GoTo CarryON
   End If

   FileCopy strPath, SparePath & "\#" & strEng
   strAI = SparePath & "\#" & strEng
   flagway = 0
   Call ConvertAI(strAI, flagway)
   DoEvents
   flagway = 1
   Call ConvertAI(strAI, flagway)
   DoEvents
 
   FileCopy strAI, strOld & "#" & strEng
   DoEvents
   Kill strAI
   End If
CarryON:
   Next i
   
  ' Text1(0) = "*.eng"
   DoEvents
  ' Text1(0) = "*.*"
   
End Sub

Private Sub Command15_Click()
Dim strPath As String, flagway As Integer, strAI As String, strDummy As String
Dim NewFile As Integer, strType As String, Y As Integer, yy As Integer, y1 As Integer
Dim y2 As Integer, strOld As String, strEng As String

yy = 1
If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
SparePath = App.Path & "\TempFiles"
Select Case MsgBox(Lang(500) & vbCrLf & Lang(501), vbOKCancel Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbOK
GoTo MakeWag
    Case vbCancel
Exit Sub
End Select

MakeWag:
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)
   x = InStrRev(strPath, "\")
strOld = Left$(strPath, x)
strEng = Mid$(strPath, x + 1)
   If Right$(strPath, 4) <> ".eng" Then
   Call MsgBox(Lang(393) & strEng & vbCrLf & Lang(456), vbExclamation, App.Title)
   
   GoTo CarryON
   End If
   NewFile = FreeFile
   Open strPath For Input As #NewFile
  Do While Not EOF(NewFile)
TryAgain:
  Line Input #NewFile, A$
A$ = Replace(A$, "  ", " ")

  Y = InStr(A$, "Type (")
  If Y > 0 Then
  
 y1 = Y + 5
  'y1 = InStr(y, a$, "(")
  y2 = InStr(Y, A$, ")")
  strType = Trim$(Mid$(A$, y1 + 1, (y2 - y1) - 1))
  
  If strType <> "Diesel" And strType <> "Electric" And strType <> "Steam" Then
  GoTo TryAgain
  Else
  Close #NewFile
  Exit Do
  End If
  End If
  Loop
   If strType = "Diesel" Or strType = "Electric" Then
   strDummy = "$" & strEng
   strDummy = Left$(strDummy, Len(strDummy) - 3) & "wag"
   strAI = SparePath & "\" & strDummy
   FileCopy strPath, strAI
   flagway = 0
   Call ConvertDummy(strAI, flagway)
   DoEvents
   flagway = 1
   Call ConvertDummy(strAI, flagway)
   DoEvents
   FileCopy strAI, strOld & "\" & strDummy
   DoEvents
   Kill strAI
   'Call MsgBox(Lang(502) & vbCrLf & Lang(503) & strDummy, vbExclamation, App.Title)
   
   ElseIf strType = "Steam" Then
   
   strDummy = "$" & strEng
   strAI = SparePath & "\" & strDummy
   FileCopy strPath, strAI
   flagway = 0
   Call ConvertDummy2(strAI, flagway)
   DoEvents
   
   flagway = 1
   Call ConvertDummy2(strAI, flagway)
   DoEvents
   FileCopy strAI, strOld & "\" & strDummy
   DoEvents
   Kill strAI
   'Call MsgBox(Lang(504) & vbCrLf & Lang(503) & strDummy, vbExclamation, App.Title)
   
   End If
   End If
CarryON:
   Next i
End Sub

Private Sub Command16_Click()
Dim i As Integer, strEng As String, strPath As String

For i = List1.Rows - 1 To 1 Step -1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)
x = InStrRev(strPath, "\")
strEng = Mid$(strPath, x + 1)
If Left$(strEng, 1) = "#" Or Left$(strEng, 1) = "$" Or Left$(strEng, 2) = "AI" Or Left$(strEng, 4) = "Dead" Then
List1.RemoveItem (i)
End If
End If
Next i
End Sub

Private Sub Command17_Click()
Dim i As Integer, strTemp As String

On Error GoTo Errtrap
Select Case MsgBox("Please confirm you really wish to DELETE all selected files." _
                   & vbCrLf & "This operation is irreversible." _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
For i = List1.Rows - 1 To 0 Step -1
If List1.IsSelected(i) = True Then
List1.Select i, 0
   strTemp = List1.Cell(flexcpText)
Kill strTemp
List1.IsSelected(i) = False
DoEvents
End If
CarryON:
Next i
    Case vbNo
Exit Sub
End Select
Exit Sub
Errtrap:
If Err = 75 Then GoTo CarryON
End Sub

Private Sub Command18_Click()
Dim i As Long, j As Long, x As Integer
Dim strPix As String, strBatText As String, strTemp As String, MyString As String
Dim xx As Integer, strShape As String, strWagName As String, strAnim As String, booIgnore As Boolean
Dim strStart As String, strTemp2 As String

lblCount(0).Caption = Str(Val(lblCount(0).Caption) - 1)
On Error GoTo Errtrap
Command1(2).Caption = "Abort"
booList = True
booAbort = False
strPicView = vbNullString
strForPrint = vbNullString
strReport = vbNullString
flagThumb = 0
If Not DirExists(App.Path & "\TempFiles") Then

MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents
Select Case MsgBox("Ignore items where a thumbnail already exists?", vbYesNo Or vbQuestion Or vbDefaultButton1, App.Title)

    Case vbYes
booIgnore = True
    Case vbNo
booIgnore = False
End Select
strSavePix = App.Path & "\Tempfiles"
strPixPath = strSavePix & "\"
SaveSetting "Decapod", "MSTS Shape Viewer", "screenshotLocation", strPixPath & "PIX"
DoEvents
ReDim PixPicture(0 To List1.SelectedRows - 1)
ReDim PixName(0 To List1.SelectedRows - 1)
ReDim PixRealName(0 To List1.SelectedRows - 1)
ReDim PixPath(0 To List1.SelectedRows - 1)
j = 0
intNumPix = List1.SelectedRows

For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strTemp2 = List1.TextMatrix(i, 0)
strAnim = vbNullString

If Right$(strTemp2, 4) <> ".eng" And Right$(strTemp2, 4) <> ".wag" Then
Call MsgBox(strTemp2 & vbCrLf & "Is not an .eng or .wag file and has been ignored.", vbExclamation, App.Title)
intNumPix = intNumPix - 1
GoTo CarryON
End If
Rem********************
If booIgnore = True Then
strStart = Left$(strTemp2, Len(strTemp2) - 3)
strStart = strStart & "jpg"
If FileExists(strStart) Then
GoTo CarryON
End If
End If
'**********************
x = InStr(strTemp2, "Common.")
If x > 0 Then GoTo CarryON
x = InStr(strTemp2, "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(strTemp2, "Invisocar")
If x > 0 Then GoTo CarryON
MyString = ReadUniFile(strTemp2)
x = InStr(MyString, "wagonshape")
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strShape = Mid$(MyString, x + 1, xx - (x + 1))
strShape = Trim$(strShape)
If Left$(strShape, 1) = ChrW$(34) Then
strShape = Mid$(strShape, 2)
End If
If Right$(strShape, 1) = ChrW$(34) Then
strShape = Left$(strShape, Len(strShape) - 1)
End If
If Right$(strShape, 2) <> ".s" Then
strReport = strReport & "File " & strTemp2 & " has an invalid WagonShape entry so could not be processed" & vbCrLf

GoTo CarryON
End If

x = InStrRev(strTemp2, "\")
strShapePath = Left$(strTemp2, x)
strWagName = Mid$(strTemp2, x + 1)
Rem ************* Look for freightanim
x = InStr(MyString, "FreightAnim")
If x > 0 Then
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ".s")
strAnim = Mid$(MyString, x + 1, (xx + 2) - (x + 1))
strAnim = Trim$(strAnim)
If Left$(strAnim, 1) = ChrW$(34) Then
strAnim = Mid$(strAnim, 2)
End If
If Right$(strAnim, 1) = ChrW$(34) Then
strAnim = Left$(strAnim, Len(strAnim) - 1)
End If
If Left$(strAnim, 2) = ".." Then strAnim = vbNullString
End If

Rem *********************** Continue from here.
PixPicture(j) = strShapePath & strWagName
x = InStr(PixPicture(j), "Common.")
If x > 0 Then GoTo CarryON
PixRealName(j) = PixPicture(j)
PixRealName(j) = Left$(PixRealName(j), Len(PixRealName(j)) - 4) & ".jpg"
x = InStrRev(PixPicture(j), "\", x - 1)
PixPath(j) = Mid$(PixPicture(j), x + 1)
intNextPix = CInt(GetSetting("Decapod", "3D Train Control", "Lastscreenshot", 0))
If intNextPix > 999 Then
SaveSetting "Decapod", "3D Train Control", "Lastscreenshot", 0
intNextPix = 0
End If
strPix = Trim$(Str(intNextPix))
If Len(strPix) < 3 Then
strPix = String(3 - Len(strPix), "0") & strPix
End If
PixName(j) = "Pix" & strPix & ".jpg"

strPicView = strShapePath & strShape
If strAnim <> vbNullString Then
strPicView = strPicView & ";" & strShapePath & strAnim
End If

strBatText = ChrW$(34) & App.Path & "\sviewRR4.exe" & ChrW$(34) & " " & ChrW$(34) & strPicView & ChrW$(34)
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)


TryAgain:
If Not FileExists(PixRealName(j)) Then
 FileCopy strPixPath & PixName(j), PixRealName(j)
 DoEvents
 Kill strPixPath & PixName(j)
 DoEvents
 ElseIf FileExists(PixRealName(j)) Then
 x = InStrRev(PixRealName(j), "\")
 strTemp = Mid$(PixRealName(j), x + 1)
 Label6.Caption = strTemp
 If flagThumb = 0 Or flagThumb = 2 Then
 frmThumb.Show 1
 End If

If flagThumb = 0 Or flagThumb = 1 Then
 Kill PixRealName(j)
 DoEvents
 FileCopy strPixPath & PixName(j), PixRealName(j)
 DoEvents
 Kill strPixPath & PixName(j)
 DoEvents
 ElseIf flagThumb = 2 Or flagThumb = 3 Then
 Kill strPixPath & PixName(j)
 DoEvents
 End If
 End If



 j = j + 1
 If j > 999 Then
 j = 0
 End If
End If

DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.Rows - 1
End If
DoEvents

Next i
MousePointer = 0
   If strReport <> vbNullString Then
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   End If
   Command1(2).Caption = "&Exit"
Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If
Command1(2).Caption = "&Exit"
End Sub

Private Sub Command19_Click()
Dim i As Long, strPath As String, strCorrShape As String
Dim x As Integer, MyString As String, xx As Integer
Dim strESD As String

strESD = InputBox("Enter new ESD_Alternative_Texture value", "Select new ESD")
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)

x = InStrRev(strPath, "\")
strCorrShape = Mid$(strPath, x + 1)
strCorrShape = Left$(strCorrShape, Len(strCorrShape) - 1)

If Right$(strPath, 3) <> ".sd" Then GoTo GetNext
MyString = ReadUniFile(strPath)
xx = InStr(MyString, "ESD_A")
x = InStr(xx, MyString, vbCr)
strStart = Left$(MyString, xx - 1)
strEnd = Mid$(MyString, x)
MyString = strStart & "ESD_Alternative_Texture ( " & strESD & " )" & strEnd
Call WriteUniFile(strPath, MyString)
DoEvents
End If
GetNext:
Next i

  DoEvents
  
Call MsgBox(".SD files have all been fixed, you may now close this screen.", vbInformation, App.Title)

Rem

Rem
End Sub

Private Sub Command2_Click()
Dim strFound As String, strOrigFile As String
Dim i As Integer, x As Integer, strPath As String, strName As String
Dim varString As String, Ename As String, booDidNotComp As Boolean

On Error GoTo Errtrap

MousePointer = 11
strReport = vbNullString
Label6.Visible = True
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
List1.IsSelected(i) = False
List1.TopRow = i
strFound = List1.TextMatrix(i, 0)
 x = InStrRev(strFound, "\")
 strPath = Left$(strFound, x - 1)
 strName = Mid$(strFound, x + 1)
   If Right$(strFound, 2) = ".s" Then
      fullpath$ = strFound
   strOrigFile = strName
   Label6.Caption = "Checking: " & strOrigFile
   Label6.Refresh
   If Right$(strOrigFile, 2) <> ".s" And Right$(strOrigFile, 2) <> ".t" And Right$(strOrigFile, 2) <> ".w" Then
   MousePointer = 0
   Call MsgBox(Lang(393) & strOrigFile & Lang(410), vbExclamation, Lang(404))
    GoTo NextOne
    End If
      
   Open fullpath$ For Binary As #5
    varString = String(2, " ")
    Get #5, , varString
 Close #5
Ename = App.Path & "\TempFiles\" & strOrigFile
 If Asc(Mid$(varString, 1, 1)) = 255 And Asc(Mid$(varString, 2, 1)) = 254 Then

booDidNotComp = False
Label6.Caption = "Compressing: " & strOrigFile
Label6.Refresh
   'Call CompressMe(fullpath$, Ename, booDidNotComp)
   Call DoComp(strOrigFile, strPath, App.Path & "\TempFiles")
   DoEvents
   If Not FileExists(Ename) Then
   strReport = strReport & strOrigFile & " did not compress" & vbCrLf
   GoTo NextOne
   End If
   Else

   booDidNotComp = True            '************************Check this works...
   
   End If
   
   
  
   If booDidNotComp = False Then
   Kill fullpath$
   DoEvents
   FileCopy Ename, fullpath$
   DoEvents
   Kill Ename
End If
End If
End If
NextOne:
   Next i
 If strReport <> vbNullString Then
  frmReport.Rich1.Text = strReport
     frmReport.Show 1
  End If
  strReport = vbNullString
   frmSearch.WindowState = 0
   frmUtils.WindowState = 0
  frmSearch.ZOrder
  MousePointer = 0
Exit Sub
Errtrap:

Resume Next
End Sub

Private Sub Command20_Click()
Dim strPath As String, strAI As String, strEng As String
Dim x As Integer, strOld As String, strOrig As String, strStart As String
Dim strEnd As String, MyString As String, Y As Long, yy As Long, strNew As String
Dim Z As Long, zz As Long

If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
SparePath = App.Path & "\TempFiles"
DoEvents
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)
x = InStrRev(strPath, "\")
strOld = Left$(strPath, x)
strEng = Mid$(strPath, x + 1)
If Left(strEng, 1) = "#" Then
strOrig = Mid(strEng, 2)
ElseIf Left(strEng, 3) = "AI_" Then
strOrig = Mid(strEng, 4)
ElseIf Left(strEng, 3) = "AI-" Then
strOrig = Mid(strEng, 4)
ElseIf Left(strEng, 2) = "AI" Then
strOrig = Mid(strEng, 3)
Else
GoTo CarryON
End If

   If Right$(strPath, 4) <> ".eng" Then
   Call MsgBox(Lang(393) & strPath & vbCrLf & Lang(456), vbExclamation, App.Title)
   
   GoTo CarryON
   End If
 
If Not FileExists(strOld & strOrig) Then
Call MsgBox(strOrig & " not found, could not adjust " & strEng, vbExclamation, App.Title)

GoTo CarryON
End If
   FileCopy strPath, SparePath & "\" & strEng
   DoEvents
   FileCopy strOld & strOrig, SparePath & "\" & strOrig
   strAI = ReadUniFile(SparePath & "\" & strEng)
   MyString = ReadUniFile(SparePath & "\" & strOrig)
   
   Y = InStr(MyString, "WheelRadius")
   yy = InStr(Y, MyString, ")")
   strTemp = Mid(MyString, Y, (yy - Y) + 1)
   Z = InStr(strAI, "WheelRadius")
   zz = InStr(Z, strAI, ")")
   strStart = Left(strAI, Z - 1)
   strEnd = Mid(strAI, zz + 1)
   strNew = strStart & strTemp & strEnd
   
   Y = InStr(yy, MyString, "WheelRadius")
   yy = InStr(Y, MyString, ")")
   strTemp = Mid(MyString, Y, (yy - Y) + 1)
   Z = InStr(zz, strAI, "WheelRadius")
   zz = InStr(Z, strAI, ")")
   strStart = Left(strAI, Z - 1)
   strEnd = Mid(strAI, zz + 1)
   strNew = strStart & strTemp & strEnd
   
   Call WriteUniFile(strPath, strNew)
   DoEvents
   Kill SparePath & "\" & strEng
   DoEvents
   Kill SparePath & "\" & strOrig
   DoEvents
   End If
CarryON:
   Next i
   
 
   DoEvents
 
End Sub

Private Sub Command3_Click()
 
Dim strFound As String
Dim i As Integer, x As Integer, strPath As String, strName As String
Dim jj As Integer, strAceView As String, strDir As String
Dim strTGAF As String, booFileName As Boolean, varBatText As String

On Error GoTo Errtrap

If Not DirExists(App.Path & "\TempFiles2") Then

MkDir App.Path & "\TempFiles2"
End If
Kill App.Path & "\TempFiles2\*.*"
DoEvents
Filpath1$ = App.Path & "\TempFiles2"
FileCopy App.Path & "\AceIt.exe", App.Path & "\TempFiles2\AceIt.exe"
MousePointer = 11
Label6.Visible = True
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
List1.IsSelected(i) = False
List1.TopRow = i
strFound = List1.TextMatrix(i, 0)
   strFound = List1.Cell(flexcpText)
 x = InStrRev(strFound, "\")
 strPath = Left$(strFound, x - 1)
 strName = Mid$(strFound, x + 1)
   If Right$(strFound, 4) = ".ace" Then
      fullpath$ = strFound
   strOrigFile = strName
   Else
   MousePointer = 0
   Call MsgBox(Lang(393) & strOrigFile & Lang(405), vbExclamation, Lang(404))
    GoTo NextOne
    End If
   x = InStrRev(strPath, "\")
 strDir = Mid$(strPath, x + 1)
x = InStr(strFound, "common")
If x > 0 Then GoTo NextOne  ' Common file

   strAceView = strFound
   
   Call readFile(strAceView, bxdata())
   If bxdata(16) = 17 Then GoTo NextOne    ' 32 bit
   If bxdata(16) = 18 Then GoTo NextOne   'Already DXT1
   strTGASave = Filpath1$ & "\" & Left$(strOrigFile, Len(strOrigFile) - 3) & "tga"
   
   Label6.Caption = "Converting: " & strDir & "\" & strOrigFile
   Label6.Refresh
  
  result = AceToTgaSquare(strAceView, strTGASave)
  DoEvents
  strTGAF = Left$(strOrigFile, Len(strOrigFile) - 3) & "tga"
  
  
   For jj = 1 To Len(strTGAF)
  
   If Mid$(strTGAF, jj, 1) = "&" Or Mid$(strTGAF, jj, 1) = " " Or Asc(Mid$(strTGAF, jj, 1)) > 122 Then
   booFileName = True
   Mid$(strTGAF, jj, 1) = "_"
   
   
   End If
   Next jj
  If booFileName = True Then
   Name strTGASave As Filpath1$ & "\" & strTGAF
   End If
 ChDrive Left$(Filpath1$, 1)
 ChDir Filpath1$
  
    varBatText = "AceIt.exe " & strTGAF & " " & Left$(strTGAF, Len(strTGAF) - 4) & ".ace -dxt /q"
  Call ShellAndWait(varBatText, True, vbHide)
  
  DoEvents
        If booFileName = True Then
        
        strOldAce = Left$(strTGAF, Len(strTGAF) - 4)
        strOldAce = strOldAce & ".ace"
        If FileExists(Filpath1$ & "\" & strOldAce) Then
        Name Filpath1$ & "\" & strOldAce As Filpath1$ & "\" & strOrigFile
        Else
        Kill Filpath1$ & "\" & strTGAF
DoEvents
        End If
        End If
        
DoEvents

If FileExists(Filpath1$ & "\" & strOrigFile) Then
Kill strFound
DoEvents
FileCopy Filpath1$ & "\" & strOrigFile, strFound
DoEvents
Kill Filpath1$ & "\" & strOrigFile
DoEvents
Kill Filpath1$ & "\" & strTGAF
DoEvents
Else
strReport = strReport & strDir & "\" & strOrigFile & Lang(549) & vbCrLf

End If

   End If
  
  
NextOne:
   Next i
If strReport <> vbNullString Then
  frmReport.Rich1.Text = strReport
     frmReport.Show 1
  End If
  strReport = vbNullString
  MousePointer = 0
  
Exit Sub
Errtrap:
If Err = 53 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'frmSearch' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)

End Sub


Private Sub Command4_Click()

List1.Select 0, 0, List1.Rows - 1, 0
lblCount(1).Caption = List1.SelectedRows
End Sub

Private Sub Command5_Click()
Dim saved As Long
MousePointer = 11
Label6.Visible = True
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
List1.IsSelected(i) = False
List1.TopRow = i
strFound = List1.TextMatrix(i, 0)
 x = InStrRev(strFound, "\")
 strPath = Left$(strFound, x - 1)
 strName = Mid$(strFound, x + 1)
 
   If Right$(strFound, 4) = ".ace" Then
      fullpath$ = strFound
   saved = saved + readFile2(fullpath$)
   Label6.Caption = "Compressing: " & strName
Label6.Refresh

   End If
   End If
   Next i
   MousePointer = 0
   
   
End Sub

Private Sub Command6_Click()
Dim i As Integer, strTemp As String

For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
List1.TopRow = i

   strTemp = List1.TextMatrix(i, 0)
 


strForPrint = strForPrint & strTemp & vbCrLf
List1.IsSelected(i) = False
End If
Next i

flagPrint = 19
fEZPrint.Show 1
End Sub

Private Sub Command7_Click()
Dim strList As String, i As Integer, strTemp As String

CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
CommonDialog1.DialogTitle = "Save List of Selected Files"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 2
strList = CommonDialog1.Filename

Open strList For Append As #5
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
List1.IsSelected(i) = False
List1.TopRow = i
strTemp = List1.TextMatrix(i, 0)
Print #5, strTemp
End If
Next i
Close #5
End Sub

Private Sub Command8_Click()
Dim i As Integer, j As Integer, x As Integer, q As Integer
Dim strPix As String, strBatText As String, strTemp As String


On Error GoTo Errtrap
'SaveSetting "Decapod", "3D Train Control", "Lastscreenshot", 998
booList = True
booAbort = False
strPicView = vbNullString
strForPrint = vbNullString
Rem ***********************
CommonDialog1.DialogTitle = "Select the Folder for your .jpg files"
        CommonDialog1.Flags = cdlOFNExplorer
        'CommonDialog1.Filter = "Enter the prefix|*.asdgfasd"
        CommonDialog1.Filename = "PIX"
        'On Error Resume Next
        CommonDialog1.ShowOpen

strSavePix = CommonDialog1.Filename
x = InStrRev(strSavePix, "\")
strPixPath = Left$(strSavePix, x)
SaveSetting "Decapod", "MSTS Shape Viewer", "screenshotLocation", strPixPath & "PIX"
DoEvents
ReDim PixPicture(0 To List1.SelectedRows - 1)
ReDim PixName(0 To List1.SelectedRows - 1)
ReDim PixRealName(0 To List1.SelectedRows - 1)
ReDim PixPath(0 To List1.SelectedRows - 1)
j = 0
intNumPix = List1.SelectedRows
'If intNumPix > 200 Then
'Call MsgBox("You have selected over 200 shapes. Please de-select some " _
'            & vbCrLf & "shapes to reduce the total to under 200." _
'            , vbCritical, App.Title)
'
'Exit Sub
'End If


For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
List1.IsSelected(i) = False
List1.TopRow = i
strTemp = List1.TextMatrix(i, 0)
If Right$(strTemp, 2) <> ".s" Then
Call MsgBox(strTemp & vbCrLf & "Is not an .S file and has been ignored.", vbExclamation, App.Title)
intNumPix = intNumPix - 1
GoTo CarryON
End If
'List1.IsSelected(i) = False
PixPicture(j) = strTemp
x = InStrRev(PixPicture(j), "\")
PixRealName(j) = Mid$(PixPicture(j), x + 1)
PixRealName(j) = Left$(PixRealName(j), Len(PixRealName(j)) - 2) & ".jpg"
x = InStrRev(PixPicture(j), "\", x - 1)
PixPath(j) = Mid$(PixPicture(j), x + 1)
intNextPix = CInt(GetSetting("Decapod", "3D Train Control", "Lastscreenshot", 0))
If intNextPix > 999 Then
SaveSetting "Decapod", "3D Train Control", "Lastscreenshot", 0
intNextPix = 0
End If
strPix = Trim$(Str(intNextPix))
If Len(strPix) < 3 Then
strPix = String(3 - Len(strPix), "0") & strPix
End If
PixName(j) = "Pix" & strPix & ".jpg"
strForPrint = strForPrint & strTemp & vbCrLf
strPicView = strTemp
'strBatText = chrw$(34) & "d:\vb projects\sview12source\sviewRR4.exe" & chrw$(34) & " " & chrw$(34) & strPicView & chrw$(34)
strBatText = ChrW$(34) & App.Path & "\sviewRR.exe" & ChrW$(34) & " " & ChrW$(34) & strPicView & ChrW$(34)
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)

 ' fLoad.Show 1
'  Do While booAbort = False
'  DoEvents
'  Loop
  
TryAgain:

 If Not FileExists(strPixPath & PixRealName(j)) Then
 Name strPixPath & PixName(j) As strPixPath & PixRealName(j)
 ElseIf Not FileExists(strPixPath & Left$(PixRealName(j), Len(PixRealName(j)) - 4) & "(" & q & ").jpg") Then
 Name strPixPath & PixName(j) As strPixPath & Left$(PixRealName(j), Len(PixRealName(j)) - 4) & "(" & q & ").jpg"
 PixRealName(j) = Left$(PixRealName(j), Len(PixRealName(j)) - 4) & "(" & q & ").jpg"
 Else
 q = q + 1
 GoTo TryAgain
 End If
 j = j + 1
 If j > 999 Then
 j = 0
 End If
End If
'Unload frmLoad
DoEvents
CarryON:
Set fLoad = Nothing

DoEvents

Next i
PrintNow:
flagPrint = 14
fEZPrint.Show 1

booList = False
Set fEZPrint = Nothing

Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If

End Sub






Private Sub Command9_Click()
Dim strPath As String, flagway As Integer

MousePointer = 11
For i = 0 To List1.Rows - 1
If List1.IsSelected(i) = True Then
lblCount(1).Caption = Str(Val(lblCount(0).Caption) - i)
List1.IsSelected(i) = False
List1.TopRow = i
strPath = List1.TextMatrix(i, 0)
   
   flagway = 1
   Call ConvertIt(strPath, flagway)
   End If
   Next i
MousePointer = 0
End Sub

Private Sub Form_Load()
Dim FirstPath As String, DirCount As Integer, s$
Me.Caption = Lang(297)
Command1(0).Caption = Lang(298)
Command1(0).ToolTipText = Lang(299)
Command1(1).Caption = Lang(300)
Command1(2).Caption = Lang(38)
Label1.Caption = Lang(301)
Label2.Caption = Lang(650)
Command2.Caption = Lang(302)
Command2.ToolTipText = Lang(653)
Command3.Caption = Lang(155)
Command3.ToolTipText = Lang(654)
Command4.Caption = Lang(216)

Command5.Caption = Lang(644)
Command5.ToolTipText = Lang(645)
Command6.Caption = Lang(646)
Command6.ToolTipText = Lang(647)
Command7.Caption = Lang(648)
Command7.ToolTipText = Lang(649)
Command8.Caption = Lang(651)
Command8.ToolTipText = Lang(652)
MousePointer = 11
DoEvents
s$ = String(280, " ")
List1.FormatString = s$
Set comp = New CompressZIt
If booLink = True Then Exit Sub

If frmUtils.Dir1(cursouind).Path <> frmUtils.Dir1(cursouind).List(frmUtils.Dir1(cursouind).ListIndex) Then
        frmUtils.Dir1(cursouind).Path = frmUtils.Dir1(cursouind).List(frmUtils.Dir1(cursouind).ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If


    frmUtils.File1(cursouind).Pattern = frmUtils.Text1(cursouind).Text
    FirstPath = frmUtils.Dir1(cursouind).Path
    DirCount = frmUtils.Dir1(cursouind).ListCount

    Rem ********************************************************************************
   Set SP = New cScanPath

strFilter = frmUtils.Text1(0).Text
With SP
            .Archive = True
            .Compressed = True
          '  .Hidden = False
            .Hidden = True
            .Normal = True
            .ReadOnly = False
            .System = False
           
            .Filter = strFilter
            .StartScan FirstPath, True, False And False, True, False
        End With
 
    
    
    
Rem *********************************************************************
    ' Start recursive direcory search.
                        ' Reset found files indicator.
  '  result = DirDiver(FirstPath, DirCount, "")
    If booAbort = True Then
'Unload frmSearch
Exit Sub
End If
frmSearch.Show
DoEvents

If booFixAct = True Then
List1.Select 0, 0, List1.Rows - 1, 0
lblCount(1).Caption = List1.SelectedRows
cmdFixAct.value = True
End If
If booFixSrv = True Then
'For i = 0 To List1.Rows - 1
'If List1.IsSelected(i) = False Then
'List1.IsSelected(i) = True
'End If
'Next i
List1.Select 0, 0, List1.Rows - 1, 0
lblCount(1).Caption = List1.SelectedRows

cmdFixSrv.value = True
End If

If booFixEng = True Then
'For i = 0 To List1.Rows - 1
'If List1.IsSelected(i) = False Then
'List1.IsSelected(i) = True
'End If
'Next i
List1.Select 0, 0, List1.Rows - 1, 0

lblCount(1).Caption = List1.Rows
cmdFixEng.value = True
End If

If booFixSMS = True Then
List1.Select 0, 0, List1.Rows - 1, 0
lblCount(1).Caption = List1.SelectedRows
cmdFixSms.value = True
End If
If booFixFix = True Then
List1.Select 0, 0, List1.Rows - 1, 0
lblCount(1).Caption = List1.SelectedRows
cmdFixFix.value = True
End If
If booFixSD = True Then
'For i = 0 To List1.Rows - 1
'If List1.IsSelected(i) = False Then
'List1.IsSelected(i) = True
'End If
'Next i
List1.Select 0, 0, List1.Rows - 1, 0
DoEvents

lblCount(1).Caption = List1.SelectedRows
cmdFixSD.value = True
End If
If booFixCon = True Then
'For i = 0 To List1.Rows - 1
'If List1.IsSelected(i) = False Then
'List1.IsSelected(i) = True
'End If
'Next i
List1.Select 0, 0, List1.Rows - 1, 0
DoEvents
lblCount(1).Caption = List1.SelectedRows
cmdFixCon.value = True
End If
If booFixBB = True Then
'For i = 0 To List1.Rows - 1
'If List1.IsSelected(i) = False Then
'List1.IsSelected(i) = True
'End If
'Next i
List1.Select 0, 0, List1.Rows - 1, 0
lblCount(1).Caption = List1.SelectedRows
cmdFixBB.value = True
End If

If booFixCVF = True Then
'For i = 0 To List1.Rows - 1
'If List1.IsSelected(i) = False Then
'List1.IsSelected(i) = True
'End If
'Next i
List1.Select 0, 0, List1.Rows - 1, 0
DoEvents
lblCount(1).Caption = List1.SelectedRows
cmdFixCVF.value = True
End If
MousePointer = 0
End Sub

Private Function WriteUniFile(CompleteFilePath As String, MyString As String) As String


Dim File_obj As Object, The_obj As Object
On Error GoTo Errtrap

Set File_obj = CreateObject("Scripting.FileSystemObject")

Set The_obj = File_obj.CreateTextFile(CompleteFilePath, True, True)
The_obj.Write (MyString)
The_obj.Close
Exit Function
Errtrap:
Call MsgBox(Err.Description & " occurred in WriteUniFile while processing" _
            & vbCrLf & "file " & CompleteFilePath _
            , vbExclamation, App.Title)


End Function



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
    'MsgBox chrw$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34) & Lang(402), vbInformation, Me.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
  mytristate = 0
'    MsgBox chrw$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34) & Lang(403), vbInformation, Me.Caption
'    Exit Function
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



' read MSTS file into byte array
' works on if GZip compressed too
Private Function readFile2(fName As String) As Long
    Dim i As Integer
    Dim bufSize As Long
    Dim bHead(7) As Byte
    Dim bdata() As Byte
    
    
    i = FreeFile
    On Error GoTo ErrEx
    
    Open fName For Binary Access Read As i
    Get #i, , bHead()
    If bHead(7) > 64 Then
        Close i
        readFile2 = 0
    Else
        ReDim bdata(lOf(i) - 17)
        Get #i, 17, bdata()
        Close i
        comp.CompressData bdata
        bufSize = comp.m_OriginalSize
        Kill fName
        Open fName For Binary Access Write As i
        bHead(7) = 70
        Put #i, , bHead
        Put #i, , bufSize
        Put #i, , "@@@@"
        Put #i, , bdata
        Close i
        readFile2 = comp.m_OriginalSize - comp.m_CompressedSize
    End If
    Exit Function
ErrEx:
    readFile2 = Err.Number
End Function


' read MSTS file into byte array
' works on if GZip compressed too
Public Function readFile(fName As String, ByRef bxdata() As Byte) As Long
    Dim i As Integer
    Dim bufSize As Long
    Dim bHead(7) As Byte
   
    
    i = FreeFile
    On Error GoTo ErrEx
    
    Open fName For Binary Access Read As i
    Get #i, , bHead()
    If bHead(7) > 64 Then
        Get #i, , bufSize
        ReDim bxdata(lOf(i) - 17)
        Get #i, 17, bxdata()
        Set comp = New CompressZIt
        readFile = comp.DecompressData(bxdata(), bufSize)
        Set comp = Nothing
    Else
        ReDim bxdata(lOf(i) - 17)
        Get #i, 17, bxdata()
    End If
    Close i
    Exit Function
ErrEx:
    readFile = Err.Number
End Function


Private Sub List1_Click()
lblCount(1).Caption = List1.SelectedRows
End Sub




Private Sub SP_FileMatch(Filename As String, Path As String)
Dim Entry As String, x As Long, j As Integer

Entry = Path & Filename

If Filename = "Default.wag" Then GoTo SkipThis

j = InStr(Path, "\Stored")
If j > 0 Then GoTo SkipThis
j = InStr(Path, "\BackUp")
If j > 0 Then GoTo SkipThis

If Right$(Path, 7) = "Stored\" Then GoTo SkipThis
        If Right$(Path, 9) = "Backup\" Then GoTo SkipThis

        If Right$(Path, 14) = "SpareConsists\" Then GoTo SkipThis
        If Right$(Path, 9) = "SpareCon\" Then GoTo SkipThis
        If Right$(Path, 7) = "Spares\" Then GoTo SkipThis
        If frmUtils.Check1.value = 0 Then
        x = InStrRev(Path, "\", Len(Path) - 1)
        If Mid$(Path, x + 1, 7) = "Cabview" Then
        GoTo SkipThis
        End If
        End If
        
            List1.AddItem Entry
            lblCount(0).Caption = Str(Val(lblCount(0).Caption) + 1)
            DoEvents
            If booAbort = True Then
            SearchFlag = False
            Exit Sub
            End If
SkipThis:
End Sub


