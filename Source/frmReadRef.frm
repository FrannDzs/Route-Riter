VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmReadRef 
   Caption         =   "RefEditor"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Add Unreferenced Shapes"
      Height          =   615
      Left            =   7560
      TabIndex        =   10
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Print Grid"
      Height          =   615
      Left            =   6480
      TabIndex        =   9
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Add to Class List"
      Height          =   615
      Left            =   11400
      TabIndex        =   8
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete Duplicates"
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Delete Row"
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add Row"
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Re-Index .Ref File"
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save new .Ref"
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit Without Change"
      Height          =   615
      Left            =   5400
      TabIndex        =   1
      Top             =   8520
      Width           =   1095
   End
   Begin VSFlex8LCtl.VSFlexGrid RefGrid 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      _cx             =   22463
      _cy             =   14420
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmReadRef.frx":0000
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
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
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
End
Attribute VB_Name = "frmReadRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim MasterFile() As String
Dim MasterIndex() As String
Dim intShapes As Integer
Dim strRefFile As String
Dim booNoChange As Boolean
Dim booEndFile As Boolean
Dim MasterClass() As Variant
Dim MasterClass2() As Variant
Dim numClass As Integer
Dim strClass As String
Dim MasterShape() As String
Dim intMasterShape As Integer

Const REF_CHUNK = 100
Private Sub FillGrid()
Dim i As Integer, x As Integer, strTemp As String, Y As Integer
Dim strGrid(0 To 9) As String, result As String, strTemp2 As String


RefGrid.AllowUserResizing = flexResizeBoth
    RefGrid.ExtendLastCol = True
    
   
   RefGrid.Cell(flexcpFontBold, 0, 0, 0, 9) = True
   
   RefGrid.ExplorerBar = flexExSort
   RefGrid.BackColor = vbWhite
RefGrid.Rows = 1

For i = 0 To intShapes
For x = 0 To 9: strGrid(x) = vbNullString: Next

strTemp = MasterFile(i)

x = InStr(strTemp, "(")
strTemp2 = Trim$(Left$(strTemp, x - 1))
For x = 1 To Len(strTemp2)
If Left$(strTemp2, 1) = vbCr Or Left$(strTemp2, 1) = vbLf Or Left$(strTemp2, 1) = vbTab Then
strTemp2 = Mid$(strTemp2, 2)
End If
Next x

If strTemp2 = "Platform" Or strTemp2 = "Siding" Or strTemp2 = "CarSpawner" Or strTemp2 = "Dyntrack" Then GoTo CarryON
strGrid(1) = strTemp2
strGrid(0) = Str(Y)
Y = Y + 1
x = InStr(strTemp, "Filename")
If x > 0 Then

Call GetData(Mid$(strTemp, x), result)
strGrid(2) = result
Else
strGrid(2) = vbNullString
End If
x = InStr(strTemp, "Shadow ")
If x > 0 Then
Call GetData(Mid$(strTemp, x), result)
strGrid(3) = result
Else
strGrid(3) = vbNullString
End If
x = InStr(strTemp, "Class ")
If x > 0 Then

Call GetData(Mid$(strTemp, x), result)
strGrid(4) = result
Else
x = InStr(strTemp, "Class(")
    If x > 0 Then
    strTemp = Left$(strTemp, x + 4) & " " & Mid$(strTemp, x + 5)
    Call GetData(Mid$(strTemp, x), result)
    strGrid(4) = result
    
Else
strGrid(4) = vbNullString
End If
End If
x = InStr(strTemp, "Align ")
If x > 0 Then
Call GetData(Mid$(strTemp, x), result)
strGrid(5) = result
Else
strGrid(5) = vbNullString
End If
x = InStr(strTemp, "Description")
If x > 0 Then
Call GetData(Mid$(strTemp, x), result)
strGrid(6) = result
Else
strGrid(6) = vbNullString
End If
x = InStr(strTemp, "StoreMatrix")
If x > 0 Then

strGrid(7) = "Yes"
Else
strGrid(7) = vbNullString
End If
x = InStr(strTemp, "PickupType")
If x > 0 Then
Call GetData(Mid$(strTemp, x), result)
strGrid(8) = result
Else
strGrid(8) = vbNullString
End If
x = InStr(strTemp, "TunnelEntrance")
If x > 0 Then
strGrid(9) = "Yes"
Else
strGrid(9) = vbNullString
End If

RefGrid.AddItem strGrid(0) & vbTab & strGrid(1) & vbTab & strGrid(2) & vbTab & strGrid(3) & vbTab & strGrid(4) & vbTab & strGrid(5) & vbTab & strGrid(6) & vbTab & strGrid(7) & vbTab & strGrid(8) & vbTab & strGrid(9)
CarryON:
Next i
End Sub

Private Sub GetData(strTemp As String, strNew As String)
Dim x As Integer, Y As Integer, xx As Integer, Z As Integer, zz As Integer



Y = InStr(strTemp, "(")
zz = InStr(strTemp, vbCr)
Z = InStr(strTemp, " )")
If Z = 0 Or Z > zz Then
Z = InStr(strTemp, ")")
End If
'yy = InStrRev(strTemp, ")", Z)
strTemp = Mid$(strTemp, Y + 1, (Z - Y) - 1)
strTemp = Trim$(strTemp)
x = InStr(strTemp, ChrW$(34))
If x > 0 Then

  xx = InStr(x + 1, strTemp, ChrW$(34))
  If xx = 0 Then
  strTemp = strTemp & ChrW$(34)
  xx = Len(strTemp)
  End If
   strNew = Mid$(strTemp, x + 1, xx - x - 1)
   Else
   strNew = strTemp
   End If
   If strNew = "None)" Then
   strNew = "NONE"
   End If
End Sub


Private Sub GetFirstToken(strRef As String, Min As Double, MinIndex As Long, x As Long)
Dim refToken(0 To 11) As Double, Max As Double, maxIndex As Long

refToken(0) = InStr(x, strRef, "Static (")
refToken(1) = InStr(x, strRef, "Forest (")
refToken(2) = InStr(x, strRef, "Hazard (")
refToken(3) = InStr(x, strRef, "LevelCr (")
refToken(4) = InStr(x, strRef, "Transfer (")
refToken(5) = InStr(x, strRef, "Pickup (")
refToken(6) = InStr(x, strRef, "Dyntrack (")
refToken(7) = InStr(x, strRef, "Platform (")
refToken(8) = InStr(x, strRef, "Siding (")
refToken(9) = InStr(x, strRef, "Carspawner (")
refToken(10) = InStr(x, strRef, "Skip(")
refToken(11) = InStr(x, strRef, "Skip (")

Call FindMinMax(refToken(), Min, Max, MinIndex, maxIndex)
DoEvents
If Min = 0 And Max > 0 Then
Min = Max
MinIndex = maxIndex
End If
End Sub

Private Sub GetNextToken(strRef As String, Min As Double, MinIndex As Long, x As Long)
Dim refToken(0 To 11) As Double, Max As Double, maxIndex As Long, i As Integer

Dim booToken As Boolean


refToken(0) = InStr(x, strRef, "Static (")
refToken(1) = InStr(x, strRef, "Forest (")
refToken(2) = InStr(x, strRef, "Hazard (")
refToken(3) = InStr(x, strRef, "LevelCr (")
refToken(4) = InStr(x, strRef, "Transfer (")
refToken(5) = InStr(x, strRef, "Pickup (")
refToken(6) = InStr(x, strRef, "Dyntrack (")
refToken(7) = InStr(x, strRef, "Platform (")
refToken(8) = InStr(x, strRef, "Siding (")
refToken(9) = InStr(x, strRef, "Carspawner (")
refToken(10) = InStr(x, strRef, "Skip(")
refToken(11) = InStr(x, strRef, "Skip (")
For i = 0 To 11
If refToken(i) > 0 Then
booToken = True
End If
Next i

Call FindMinMax(refToken(), Min, Max, MinIndex, maxIndex)
If Min = 0 And Max > 0 Then
Min = Max
MinIndex = maxIndex
ElseIf Min = 0 And Max = 0 Then  'end of file
Min = Len(strRef)
booEndFile = True
End If
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
Private Sub ReadRef2(strRefName As String)
Dim strNew As String, strTemp As String, x As Long, i As Long
Dim Y As Integer, yy As Integer, strTemp2 As String
Dim strRef As String, Min As Double, strNew2 As String
Dim MinIndex As Long, tokStart As Double, xx As Long, MinNext As Double

ReDim MasterIndex(0 To REF_CHUNK)
ReDim MasterFile(0 To REF_CHUNK)
ReDim MasterClass(0 To REF_CHUNK)
ReDim MasterShape(0 To REF_CHUNK)
intMasterShape = 0
'Open "e:\temp\reftest.txt" For Output As #7
strClass = "          "
strRef = ReadUniFile(strRefName)
strRef = Replace(strRef, vbTab, " ")
strRef = Replace(strRef, "        (", " (")
strRef = Replace(strRef, "       (", " (")
strRef = Replace(strRef, "      (", " (")
strRef = Replace(strRef, "     (", " (")
strRef = Replace(strRef, "    (", " (")
strRef = Replace(strRef, "   (", " (")
strRef = Replace(strRef, "  (", " (")
strRef = Replace(strRef, "         (", " (")
strRef = Replace(strRef, "          (", " (")
strRef = Replace(strRef, "Static(", "Static (")
strRef = Replace(strRef, ".s)", ".s )")
strRef = Replace(strRef, vbCr & vbLf & " )", vbCr & vbLf & ")")

x = 1

GetAnother:
Call GetFirstToken(strRef, Min, MinIndex, x)
             

tokStart = Min
x = Min + 1

Call GetNextToken(strRef, MinNext, MinIndex, x)


If MinNext = 0 Then GoTo Label1
strTemp2 = Mid$(strRef, tokStart, MinNext - tokStart)



              Rem***************
              xx = InStr(strTemp2, "FileName")
           
                     If xx > 0 Then
                    
                     Y = InStr(xx, strTemp2, "(")
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
                         
                          frmUtils.SB1.Panels(2).Text = strNew
                           If Right(strNew, 2) = ".s" Then
                                 MasterShape(intMasterShape) = strNew
                                 intMasterShape = intMasterShape + 1
                                 If intMasterShape > UBound(MasterShape) Then
                                ReDim Preserve MasterShape(0 To intMasterShape + REF_CHUNK)
                                End If
                                 
                            End If
                    End If
                    
              Rem****************
              xx = InStr(strTemp2, "Class")
             
                     If xx > 0 Then
                     Y = InStr(xx, strTemp2, "(")
                     yy = InStr(Y, strTemp2, ")")
                     strNew2 = Mid$(strTemp2, Y + 1, yy - (Y + 1))
                     strNew2 = Trim$(strNew2)
                          If Left$(strNew2, 1) = ChrW$(34) Then
                                strNew2 = Mid$(strNew2, 2)
                                Y = InStr(strNew2, ChrW$(34))
                                 If Y > 0 Then
                                 strNew2 = Left$(strNew2, Y - 1)
                                 End If
                          End If
                    End If
              Rem *********** fix temp2 ****************
              j = InStrRev(strTemp2, ")")
              strTemp2 = Left$(strTemp2, j)
              strTemp2 = strTemp2 & vbCrLf
              
                Rem*************
              MasterIndex(i) = strNew
              MasterFile(i) = strTemp2
              MasterClass(i) = strNew2
              strTemp = vbNullString
              strTemp2 = vbNullString
             ' strTemp3 = vbNullString
              i = i + 1
              
              If i > UBound(MasterIndex) Then
           ReDim Preserve MasterIndex(0 To i + REF_CHUNK)
           ReDim Preserve MasterFile(0 To i + REF_CHUNK)
           ReDim Preserve MasterClass(0 To i + REF_CHUNK)
           End If

TryAgain:
      
       x = MinNext - 1
       If booEndFile = False Then
       GoTo GetAnother
       End If
  

Label1:
 
    ReDim Preserve MasterClass(0 To i)
    QSort3 MasterClass(), 0, i
    DoEvents
    RemD2 MasterClass(), MasterClass2()
    DoEvents
    numClass = UBound(MasterClass2)
   For x = 0 To numClass
   strClass = strClass & "|" & MasterClass2(x)
   Next x
 
 ReDim Preserve MasterIndex(0 To i)
 ReDim Preserve MasterFile(0 To i)
 intShapes = i - 1
 Call FillGrid
 booEndFile = False
 
 'Close #7
End Sub

Public Sub FindMinMax(ByRef dArray() As Double, ByRef dLowVal As Double, ByRef dHighVal As Double, MinIndex As Long, maxIndex As Long)

    Dim lIndex As Long
    Dim dFirstValIdx As Double
    Dim dLastValIdx As Double
    Dim dActVal As Double
    dFirstValIdx = LBound(dArray)
    dLastValIdx = UBound(dArray)
    dLowVal = dArray(dFirstValIdx) 'start value
    dHighVal = dArray(dFirstValIdx) 'start value


    For lIndex = dFirstValIdx To dLastValIdx
        dActVal = dArray(lIndex)


        If dActVal > dHighVal Then
            dHighVal = dActVal
            maxIndex = lIndex
        Else 'if value smaller Then high value


            If dActVal < dLowVal And dActVal > 0 Then
                dLowVal = dActVal
                MinIndex = lIndex
            End If

        End If

    Next lIndex

End Sub

Private Sub SetLang()
Command4.Caption = Lang(621)
Command5.Caption = Lang(622)
Command3.Caption = Lang(623)
Command2.Caption = Lang(624)
Command1.Caption = Lang(625)

End Sub

Private Sub WriteRow(lRow As Long)
Dim i As Integer, strTemp As String
Dim strGrid(0 To 9) As String

For i = 0 To 9
RefGrid.Select lRow, i
strGrid(i) = RefGrid.TextMatrix(lRow, i)
Next i
strTemp = strGrid(1) & "   (" & vbCrLf
RefGrid.Select lRow, 2
strTemp = strTemp & "Filename        ( " & ChrW$(34) & strGrid(2) & ChrW$(34) & " )" & vbCrLf
RefGrid.Select lRow, 3
If strGrid(3) <> vbNullString Then
strTemp = strTemp & "Shadow          ( " & strGrid(3) & " )" & vbCrLf
End If
RefGrid.Select lRow, 4
If strGrid(4) <> vbNullString Then
strTemp = strTemp & "Class           ( " & ChrW$(34) & strGrid(4) & ChrW$(34) & " )" & vbCrLf
End If
RefGrid.Select lRow, 5
If strGrid(5) <> vbNullString Then
strTemp = strTemp & "Align           ( " & strGrid(5) & " )" & vbCrLf
End If
RefGrid.Select lRow, 6
If strGrid(6) <> vbNullString Then
strTemp = strTemp & "Description     ( " & ChrW$(34) & strGrid(6) & ChrW$(34) & " )" & vbCrLf
End If
RefGrid.Select lRow, 7
If strGrid(7) <> vbNullString Then
strTemp = strTemp & "StoreMatrix ( )" & vbCrLf
End If
RefGrid.Select lRow, 8
If strGrid(8) <> vbNullString Then
strTemp = strTemp & "PickupType      ( " & strGrid(8) & " )" & vbCrLf
End If
RefGrid.Select lRow, 9
If strGrid(9) <> vbNullString Then
strTemp = strTemp & "StoreMatrix ( )" & vbCrLf
End If

MasterFile(Val(strGrid(0))) = strTemp & vbCrLf & ")" & vbCrLf


End Sub

Private Sub Command1_Click()
booNoChange = True
Unload Me

End Sub

Private Sub Command2_Click()
Dim i As Long, NewFile As Integer
Dim flagway As Integer, Filpath1$

For i = 0 To RefGrid.Rows - 1
RefGrid.Select i, 0
Call WriteRow(i)
DoEvents
Next i

Filpath1$ = App.Path & "\setupfiles\"
If FileExists(Filpath1$ & "TempRef.ref") Then
Kill Filpath1$ & "TempRef.ref"
DoEvents
End If
FileCopy Filpath1$ & "reffilestart.txt", Filpath1$ & "TempRef.ref"

NewFile = FreeFile
   Open Filpath1$ & "TempRef.ref" For Append As #NewFile
   Print #NewFile, vbCrLf
For i = 0 To RefGrid.Rows - 1
   
               Print #NewFile, MasterFile(i)
               Next i
               Close NewFile
 
 
flagway = 1
Call ConvertIt(Filpath1$ & "TempRef.ref", flagway)
DoEvents
Kill strRefFile
DoEvents
FileCopy Filpath1$ & "TempRef.ref", strRefFile

Unload Me

End Sub

Private Sub Command3_Click()
Dim i As Integer
For i = 0 To RefGrid.Rows - 2

RefGrid.Select i + 1, 0
RefGrid.TextMatrix(i + 1, 0) = Str(i)
Next
End Sub



Private Sub Command4_Click()
Dim totRows As Long

totRows = RefGrid.Rows
RefGrid.Rows = totRows + 1
intShapes = intShapes + 1
intShapes = intShapes - x
For i = 0 To RefGrid.Rows - 2

RefGrid.Select i + 1, 0
RefGrid.TextMatrix(i + 1, 0) = Str(i)
Next
    If i > UBound(MasterIndex) Then
           ReDim Preserve MasterIndex(0 To i + REF_CHUNK)
           ReDim Preserve MasterFile(0 To i + REF_CHUNK)
           ReDim Preserve MasterClass(0 To i + REF_CHUNK)
           End If
End Sub

Private Sub Command5_Click()
Dim i As Long, x As Integer

x = (RefGrid.RowSel - RefGrid.row) + 1

For i = RefGrid.RowSel To RefGrid.row Step -1
RefGrid.RemoveItem
Next i
intShapes = intShapes - x
For i = 0 To RefGrid.Rows - 2

RefGrid.Select i + 1, 0
RefGrid.TextMatrix(i + 1, 0) = Str(i)
Next

End Sub

Private Sub Command6_Click()
Dim i As Long

On Error GoTo Errtrap

Select Case MsgBox("It is possible to have two items using the same .s file so take care with this option." _
                   & vbCrLf & "Click OK if you wish to continue" _
                   , vbOKCancel Or vbExclamation Or vbDefaultButton1, "Warning !!")

    Case vbOK

    Case vbCancel
Exit Sub
End Select
RefGrid.col = 2
RefGrid.Sort = flexSortGenericAscending

For i = RefGrid.Rows - 1 To 1 Step -1

If RefGrid.Cell(flexcpText, i, 2) = RefGrid.Cell(flexcpText, i - 1, 2) Then

RefGrid.row = i
RefGrid.RemoveItem
End If
Next i
DoEvents
For i = 0 To RefGrid.Rows - 2

RefGrid.Select i + 1, 0
RefGrid.TextMatrix(i + 1, 0) = Str(i)
Next
Exit Sub
Errtrap:

Resume Next

End Sub

Private Sub Command7_Click()
If Text1.Text <> vbNullString Then
strClass = strClass & "|" & Text1
RefGrid.ColComboList(4) = strClass
End If
Text1 = vbNullString
End Sub

Private Sub Command8_Click()
flagPrint = 18
fEZPrint.Show

End Sub

Private Sub Command9_Click()
Dim i As Integer, strTemp As String, j As Integer, booFound As Boolean
Dim nextrow As Integer
MousePointer = 11
frmUtils.Dir1(0).Path = RoutePath & "\Shapes"
frmUtils.Text1(0).Text = "*.s"
DoEvents
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i
For i = 0 To frmUtils.File1(cursouind).ListCount - 1

   If frmUtils.File1(cursouind).Selected(i) Then
     strTemp = frmUtils.File1(cursouind).List(i)
     For j = 0 To intMasterShape - 1
        If strTemp = MasterShape(j) Then
        booFound = True
        Exit For
        End If
     Next j

        If booFound = False Then
        nextrow = RefGrid.Rows + 1
        RefGrid.AddItem Str(nextrow) & vbTab & "Static" & vbTab & strTemp & vbTab & "" & vbTab & "Misc" & vbTab & "" & vbTab & Left(strTemp, Len(strTemp) - 2) & vbTab & "" & vbTab & "" & vbTab & ""
        Else
        booFound = False
        
        End If
     End If
     Next i
     ReDim Preserve MasterFile(0 To nextrow)
     Command3.value = True
MousePointer = 0
End Sub

Private Sub Form_Load()
Dim flagway As Integer

Call SetLang
MousePointer = 11
strRefFile = frmUtils.Label2(0).Caption
RefGrid.ColComboList(3) = "     |DYNAMIC|RECT|ROUND|Shadow"
RefGrid.ColComboList(1) = "CarSpawner|Dyntrack|Hazard|LevelCr|PickUp|Platform|Siding|Static|Transfer"


Me.Caption = Me.Caption & " - " & strRefFile
flagway = 0
'Call ConvertIt(strRefFile, flagway)


Call ReadRef2(strRefFile)
DoEvents
RefGrid.ExplorerBar = flexExSort
RefGrid.ColComboList(4) = strClass
DoEvents
MousePointer = 0
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
Public Function RemD2(ByRef rArray() As Variant, xArray() As Variant) As Variant

    'Declare variables
    Dim ii As Long, jj As Long
    
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
        xArray(jj) = rArray(ii)
        jj = jj + 1
End If
Next ii
If rArray(high) <> rArray(high - 1) Then
xArray(jj) = rArray(high)
jj = jj + 1
End If
        

  
ReDim Preserve xArray(0 To jj - 1)

End Function


Private Function ConvertIt(CompleteFilePath As String, flagway As Integer) As Boolean


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
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & " is too large to convert!", vbInformation, Me.Caption
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
    'MsgBox chrw$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34) & Lang(402), vbInformation, Me.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
   ' MsgBox chrw$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34) & Lang(403), vbInformation, Me.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll
The_obj.Close
fileflag = False
MyString = Replace(MyString, vbCr & vbLf & vbCr & vbLf, vbCrLf)

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






Private Sub Form_Resize()
On Error GoTo Errtrap
RefGrid.width = frmReadRef.width * 0.95
RefGrid.height = frmReadRef.height * 0.85
RefGrid.Left = 100
RefGrid.Top = 100
Command1.Top = RefGrid.Top + RefGrid.height + 100
Command2.Top = Command1.Top
Command3.Top = Command1.Top
Command4.Top = Command1.Top
Command5.Top = Command1.Top
Command6.Top = Command1.Top
Command6.Left = RefGrid.Left
Command5.Left = Command6.Left + Command6.width
Command4.Left = Command5.Left + Command5.width
Command3.Left = Command4.Left + Command4.width
Command2.Left = Command3.Left + Command3.width
Command1.Left = Command2.Left + Command2.width
Text1.Left = Command8.Left + Command8.width + 300
Command7.Left = Text1.Left + Text1.width + 50
Text1.Top = Command1.Top
Command7.Top = Command1.Top
Command8.Top = Command1.Top
Command9.Top = Command1.Top
Command8.Left = Command1.Left + Command1.width
Command9.Left = Command8.Left + Command8.width
Exit Sub
Errtrap:
If Err = 380 Then
Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If booNoChange = True Then
Dim flagway As Integer
flagway = 1
Call ConvertIt(strRefFile, flagway)
booNoChange = False
DoEvents
End If
End Sub




Private Sub RefGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim strTemp As String
If Button = 2 Then
strTemp = InputBox("Change Selected items to:", "Editor")
For i = RefGrid.row To RefGrid.RowSel
RefGrid.TextMatrix(i, RefGrid.col) = strTemp
Next i
End If
End Sub


