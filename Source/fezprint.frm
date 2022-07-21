VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "VSPrint8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fEZPrint 
   Caption         =   "Print Preview"
   ClientHeight    =   9045
   ClientLeft      =   2235
   ClientTop       =   1740
   ClientWidth     =   13260
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "fezprint.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9045
   ScaleWidth      =   13260
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   7680
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog Cdl1 
      Left            =   5520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load"
      Height          =   435
      Left            =   8840
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   435
      Left            =   7720
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   1200
      Left            =   120
      ScaleHeight     =   1072.941
      ScaleMode       =   0  'User
      ScaleWidth      =   1545
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter 
      Height          =   6855
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   9255
      _cx             =   16325
      _cy             =   12091
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   35.9180035650624
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Pre&view"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   9960
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11280
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   4920
      Top             =   7200
      Width           =   375
   End
End
Attribute VB_Name = "fEZPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Dim MyPage%         'Keep the output view to be printed
Dim OldOrientation  'Don't mess with my printer settings

Dim booCancel As Boolean






Private Sub DoGrid()
VSPrinter.Orientation = orLandscape
 VSPrinter.StartDoc
        VSPrinter.RenderControl = frmGrid.Grid1.hwnd
        VSPrinter.EndDoc

End Sub

Private Sub DoList()
Dim i As Integer, j As Integer, x As Integer, y As Integer

    With VSPrinter
.MarginLeft = 50
.MarginHeader = 250
.MarginTop = 750
.PenWidth = 1
       ' .TextAlign = taLeftMiddle
    .Header = "Produced using Route_Riter||Page %d"
        .FontSize = 10
        .TableBorder = 7
       
        .Paragraph = vbNullString
        j = 0
For i = 0 To intNumPix - 1

Set Picture1.Picture = LoadPicture(strPixPath & PixRealName(i))

If j = 0 Or (Int(j / 2) = j / 2) Then
.DrawPicture Picture1.Picture, 500, (j + 1) * 2000, "60%", "60%"
y = ((j + 1) * 2000) + 3600
x = 1500
.TextBox PixPath(i), x, y, 4500, 200
j = j + 1
Else
.DrawPicture Picture1.Picture, 6000, (j) * 2000, "60%", "60%"
y = ((j) * 2000) + 3600
x = 7000
.TextBox PixPath(i), x, y, 4500, 200
j = j + 1
End If

If j = 6 Then
j = 0
.NewPage
End If
    Next i
    End With
 ' all done
  VSPrinter.Action = 6 'End Document
  MousePointer = 0
End Sub

Private Sub DoLooseActCon()
VSPrinter.Columns = 2
VSPrinter.Orientation = orPortrait
VSPrinter.Header = "Consists||Page %d"
 VSPrinter.StartDoc
        VSPrinter.RenderControl = frmActGrid.Grid2.hwnd
        VSPrinter.EndDoc
End Sub

Private Sub DoLooseCon()
VSPrinter.Columns = 2
VSPrinter.Orientation = orPortrait
VSPrinter.Header = "Consists||Page %d"
 VSPrinter.StartDoc
        VSPrinter.RenderControl = frmGrid.Grid2.hwnd
        VSPrinter.EndDoc
End Sub
Private Sub doRefFiles()
'VSPrinter.Columns = 2
VSPrinter.Orientation = orLandscape
VSPrinter.Header = "Ref. File||Page %d"
 VSPrinter.StartDoc
        VSPrinter.RenderControl = frmReadRef.RefGrid.hwnd
        VSPrinter.EndDoc
End Sub

Private Sub DoTraffic()
VSPrinter.Orientation = orPortrait
VSPrinter.Columns = 2
VSPrinter.Header = "Unused Traffic||Page %d"
 VSPrinter.StartDoc
        VSPrinter.RenderControl = frmUnusedSrv.GridTfc.hwnd
        VSPrinter.EndDoc
End Sub

Private Sub DoUnusedCon()
VSPrinter.Columns = 3
VSPrinter.Orientation = orPortrait
VSPrinter.Header = "Unused Consists||Page %d"
 VSPrinter.StartDoc
        VSPrinter.RenderControl = frmUnusedSrv.GridCon.hwnd
        VSPrinter.EndDoc

End Sub

Private Sub DoUnusedPaths()
VSPrinter.Orientation = orPortrait
VSPrinter.Columns = 2
VSPrinter.Header = "Unused Paths||Page %d"
 VSPrinter.StartDoc
        VSPrinter.RenderControl = frmUnusedSrv.GridPaths.hwnd
        VSPrinter.EndDoc
End Sub

Private Sub DoUnusedSrv()
VSPrinter.Orientation = orPortrait
VSPrinter.Columns = 2
VSPrinter.Header = "Unused Services||Page %d"
 VSPrinter.StartDoc
        VSPrinter.RenderControl = frmUnusedSrv.GridUnused.hwnd
        VSPrinter.EndDoc

End Sub

Private Sub DoMissing()
VSPrinter.Orientation = orPortrait
VSPrinter.Header = "Missing Rolling-Stock||Page %d"
 VSPrinter.StartDoc
        VSPrinter.RenderControl = frmStock.Grid3.hwnd
        VSPrinter.EndDoc

End Sub

Private Sub DoStock2()
VSPrinter.Orientation = orLandscape
VSPrinter.Header = "||Page %d"
 VSPrinter.StartDoc
        VSPrinter.RenderControl = frmStock.GridUnused.hwnd
        VSPrinter.EndDoc

End Sub


Private Sub DoStock()
VSPrinter.Orientation = orLandscape
VSPrinter.Header = "||Page %d"
 VSPrinter.StartDoc
        VSPrinter.RenderControl = frmStock.GridStock.hwnd
        VSPrinter.EndDoc

End Sub
Private Sub DoUnused()
'vsPrinter.Orientation = orPortrait
'vsPrinter.Header = "Unused Rolling-Stock||Page %d"
' vsPrinter.StartDoc
'        vsPrinter.RenderControl = frmStock.Grid2.hwnd
'        vsPrinter.EndDoc

End Sub


Private Sub gettext()
    With VSPrinter
VSPrinter.Header = "Search Results:-||Page %d"
    .Orientation = orPortrait
        .PenWidth = 1
        .TextAlign = 0
        .FontSize = 8
        .TableBorder = 7
        .PageBorder = 2
 End With
If booListAce = True Then
booListAce = False
 VSPrinter.Columns = 2
 VSPrinter.ColumnSpacing = 10
 End If
 
    VSPrinter = objList ' render regular text and RTF
 
  

MousePointer = 0

End Sub







Private Sub Command1_Click(Index%)


  ' remember page for use with Print command
  MyPage = Index%

  ' we have a print job, so let's enable these guys
 ' cmbZoom.Enabled = True
  
  Command1(2).Visible = False
  If selFlag = 1 Then
  VSPrinter.Orientation = orLandscape
  Else
  VSPrinter.Orientation = orPortrait
  
  End If
  ' start the print preview job
  VSPrinter.Action = 3 ' StartDoc
  If VSPrinter.Error Then Beep: Exit Sub
 ' MousePointer = 11

  ' set default style
  VSPrinter.FontName = "Arial"
  VSPrinter.FontSize = 30
  VSPrinter.FontBold = False
  VSPrinter.FontItalic = False
  VSPrinter.TextAlign = 0  'Left
  VSPrinter.TableBorder = 7
  VSPrinter.PageBorder = 2
  VSPrinter.PenStyle = 0
  VSPrinter.BrushStyle = 0
  VSPrinter.PenWidth = 2
  VSPrinter.PenColor = 0
  VSPrinter.BrushColor = 0
  VSPrinter.TextColor = 0
  VSPrinter.Columns = 1
  If flagPrint = 1 Then
  Call DoTable
  ElseIf flagPrint = 2 Then
  Call DoTable2
  ElseIf flagPrint = 3 Then
  Call DoPrint
  ElseIf flagPrint = 4 Then
  Call DoGrid
  ElseIf flagPrint = 5 Then
  Call DoStock
  ElseIf flagPrint = 6 Then
  Call DoUnused
  ElseIf flagPrint = 7 Then
  Call DoMissing
  ElseIf flagPrint = 8 Then
  Call DoUnusedSrv
  ElseIf flagPrint = 9 Then
  Call DoUnusedCon
  ElseIf flagPrint = 10 Then
  Call DoLooseCon
  ElseIf flagPrint = 11 Then
  Call DoTraffic
  ElseIf flagPrint = 12 Then
  Call DoLooseActCon
  ElseIf flagPrint = 13 Then
  Call DoUnusedPaths
  ElseIf flagPrint = 14 Then
  Call DoList
  ElseIf flagPrint = 15 Then
  Show
  ElseIf flagPrint = 16 Then
  Call DoStock2
  ElseIf flagPrint = 18 Then
  Call doRefFiles
  ElseIf flagPrint = 19 Then
  Call DoPrint2
  Else
Call gettext
End If
  ' choose what to print based on button index
  
 VSPrinter.Action = 6 'End Document

End Sub

Private Sub Command2_Click()


Unload Me
End Sub




Private Sub Command3_Click()
Dim strList As String

MousePointer = 11
Command5.Visible = True

CDL1.Filter = "VSView Files (*.vsv)|*.vsv"
CDL1.DialogTitle = "Save Current Document"
CDL1.FilterIndex = 1
CDL1.Action = 2
strList = CDL1.FileName
VSPrinter.SaveDoc strList

DoEvents
MousePointer = 0
End Sub

Private Sub Command4_Click()
Dim strList As String

CDL1.Filter = "VSView Files (*.vsv)|*.vsv"
CDL1.DialogTitle = "Select Document to Load"
CDL1.FilterIndex = 1
CDL1.Action = 1
strList = CDL1.FileName
If strList = vbNullString Then Exit Sub
VSPrinter.LoadDoc strList
End Sub


Private Sub Command5_Click()
booCancel = True

End Sub



Private Sub Form_Load()

   
    
    '------------------------------------------------------
    ' save orientation to clean up later
    '------------------------------------------------------
    OldOrientation = VSPrinter.Orientation
    VSPrinter.Orientation = orPortrait
    
    MyPage = -1                         ' no current page
 
  Command1(2).Visible = True
   Command1(2).Caption = Lang(182)
   Command2.Caption = Lang(38)
    '------------------------------------------------------
    ' orientation (you cannot choose your own)
    '------------------------------------------------------
    
VSPrinter.Orientation = orPortrait

    '------------------------------------------------------
    ' ready, set default page to 0
    '------------------------------------------------------
    MyPage = 0
  
    With VSPrinter
        
        .Preview = True         ' Show preview to screen
        .PreviewPage = 1        ' default preview page to first page

        '------------------------------------------------------
        ' show available devices
        ' and honor Windows default selection
        '------------------------------------------------------
        
        
    End With

  VSPrinter.Action = 6 'End Document
  MousePointer = 0
Command1(2).Value = True
End Sub

Private Sub DoTable2()
    Dim s$, fmt$

    With VSPrinter
.MarginLeft = 150
.PenWidth = 1
        .TextAlign = 1
    .Header = Lang(522) & "||Page %d"
        .FontSize = 11
        .TableBorder = 7
        '--------------------------------------------------------
        ' create table format
        '--------------------------------------------------------
        fmt = "^5500|^5500;"                   '^Center > Right
    
        '--------------------------------------------------------
        ' create table string
        '--------------------------------------------------------
        s = fmt & Lang(523) & "|" & Lang(523) & ";"
        .Table = s
        'Set header
        fmt = "5500|5500;"
        s = fmt & strForPrint
        

        '--------------------------------------------------------
        ' print the table in three flavors
        '--------------------------------------------------------
        .PenWidth = 1
        .TextAlign = 1
    
        .FontSize = 11
        .TableBorder = 7
        .Table = s         ' flavor 1
        .Paragraph = vbNullString
    
        
    End With
 ' all done
  VSPrinter.Action = 6 'End Document
  MousePointer = 0

End Sub


Private Sub DoPrint2()
   

    With VSPrinter
.MarginLeft = 150
.PenWidth = 1
        .TextAlign = 0
    .Header = "||Page %d"
        .FontSize = 10
        .TableBorder = 7
       .Text = strForPrint
        .Paragraph = vbNullString
    
        
    End With
 ' all done
  VSPrinter.Action = 6 'End Document
  MousePointer = 0

End Sub

Private Sub DoPrint()
   

    With VSPrinter
.MarginLeft = 50
.PenWidth = 1
        .TextAlign = 1
    .Header = "||Page %d"
        .FontSize = 12
        .TableBorder = 7
       .Text = strForPrint
        .Paragraph = vbNullString
    
        
    End With
 ' all done
  VSPrinter.Action = 6 'End Document
  MousePointer = 0

End Sub

Private Sub DoTable()
    Dim s$, fmt$

    With VSPrinter
.MarginLeft = 50
        '--------------------------------------------------------
        ' print page title
        '--------------------------------------------------------
'        s = "You can now print reports that include paragraphs "
'        s = s & "and tables.  With VSView, printing a grid it is very easy."
'
'        .Paragraph = "Print by Table"
'        .FontSize = 18
'        .Paragraph = vbNullString
'        .Paragraph = s
'        .Paragraph = vbNullString
.PenWidth = 1
        .TextAlign = 1
    .Header = "||Page %d"
        .FontSize = 12
        .TableBorder = 7
        '--------------------------------------------------------
        ' create table format
        '--------------------------------------------------------
        fmt = "^6000|^6000;"                   '^Center > Right
    
        '--------------------------------------------------------
        ' create table string
        '--------------------------------------------------------
        s = fmt & Lang(65) & "|" & Lang(61) & ";"
        .Table = s
        'Set header
        fmt = "6000|6000;"
        s = fmt & strForPrint
        .Table = s
        .Paragraph = vbNullString
s = fmt & "Consist|Wagons;"
        .Table = s
        'Set header
        fmt = "6000|6000;"
        s = fmt & strForPrint2
        '--------------------------------------------------------
        ' print the table in three flavors
        '--------------------------------------------------------
        .PenWidth = 1
        .TextAlign = 1
    
        .FontSize = 12
        .TableBorder = 7
        .Table = s         ' flavor 1
        .Paragraph = vbNullString
    
        
    End With
 ' all done
  VSPrinter.Action = 6 'End Document
  MousePointer = 0

End Sub







Private Sub Form_Resize()
On Error GoTo Errtrap
VSPrinter.Top = Command2.Top + Command2.height + 100
VSPrinter.height = ScaleHeight - (Command2.Top + Command2.height + 100)
VSPrinter.Left = 150
VSPrinter.width = ScaleWidth - 300
'VSPrinter.Height = fEZPrint.Height - 1400
Exit Sub
Errtrap:
If Err = 380 Then
Exit Sub
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
  
  '--------------------------------------------------------
  ' restore printer orientation
  '--------------------------------------------------------
  VSPrinter.Orientation = OldOrientation

End Sub




Private Sub VSPrinter_SavingDoc(ByVal Page As Integer, ByVal Of As Integer, Cancel As Boolean)
Label1.Caption = "Saving Page " & Page & " Of " & Of
DoEvents
If booCancel = True Then
Cancel = True
Command5.Visible = False
End If
Refresh
End Sub


