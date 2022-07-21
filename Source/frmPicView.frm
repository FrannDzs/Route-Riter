VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPicView 
   Caption         =   "PicView"
   ClientHeight    =   5145
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   8790
   Icon            =   "frmPicView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      Height          =   5145
      Left            =   0
      ScaleHeight     =   5085
      ScaleWidth      =   7395
      TabIndex        =   0
      ToolTipText     =   "Click on the image to display full screen"
      Top             =   0
      Width           =   7455
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox imlToolbarIcons 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   1740
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   1350
      Width           =   1200
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu image 
      Caption         =   "Image"
      Begin VB.Menu grey 
         Caption         =   "Greyscale"
      End
      Begin VB.Menu flip 
         Caption         =   "Flip"
      End
      Begin VB.Menu rot 
         Caption         =   "Rotate 90 degrees"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuBrowser 
         Caption         =   "Browser"
      End
      Begin VB.Menu mnuSlide 
         Caption         =   "Slide_Show"
      End
   End
   Begin VB.Menu mnuAbort 
      Caption         =   "Abort"
   End
End
Attribute VB_Name = "frmPicView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strPicDir As String

Private Sub ShowPix()
Dim sFile As String
    Dim tFile As String
    Dim res As Long
    Dim p As Pic
    Dim aComment As String
    Dim i As Integer
    Dim z As Integer
    
        sFile = strPicView
        
    
    'ToDo: add code to process the opened file
    tFile = "c:\temp.bmp"
    
    'call the mwgfx.dll routine
    res = anytobmps(sFile, tFile, p, 0, 0)
    
    If p.width > 800 Then
    z = WinImageSize(tFile)
    End If
    Picture1.AutoSize = True
    Picture1.Picture = LoadPicture(tFile)
    Picture1.Align = 3
    Picture1.Align = 1
    'get the data from the Pic structure
    For i = 0 To 80 Step 1
        If p.comment(i) = 0 Then Exit For
        aComment = aComment + Chr$(p.comment(i))
    Next i
           
    Caption = "PicView - " + aComment + " (" + str(p.width) + "x" + str(p.height) + "x" + str(p.depth) + ")"
    

End Sub


Private Sub flip_Click()
    Dim tFile As String
    Dim p As Pic
        
    tFile = "c:\temp.bmp"
       
    'call the mwgfx.dll routine
    Call bmprocess(tFile, tFile, p, 106)
    Picture1.AutoSize = True
    Picture1.Picture = LoadPicture(tFile)
    Picture1.Align = 3
    Picture1.Align = 1

End Sub

Private Sub Form_Load()
Dim x As Integer

    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.width = GetSetting(App.Title, "Settings", "MainWidth", 12800)
    Me.height = GetSetting(App.Title, "Settings", "MainHeight", 8000)
    If strPicView <> vbNullString Then
    x = InStrRev(strPicView, "\")
    strPicDir = Left$(strPicView, x - 1)
    Call ShowPix
 
    End If
End Sub


Private Sub Form_Resize()
    Picture1.Align = 3
    Picture1.Align = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Dim i As Integer
'
'strPicView = vbNullString
'    'close all sub forms
'    For i = Forms.Count - 1 To 1 Step -1
'        Unload Forms(i)
'    Next
'    If Me.WindowState <> vbMinimized Then
'        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
'        SaveSetting App.Title, "Settings", "MainTop", Me.Top
'        SaveSetting App.Title, "Settings", "MainWidth", Me.width
'        SaveSetting App.Title, "Settings", "MainHeight", Me.height
'    End If
End Sub



Private Sub grey_Click()
    Dim tFile As String
    Dim p As Pic
    tFile = "c:\temp.bmp"
       
    'call the mwgfx.dll routine
    Call bmprocess(tFile, tFile, p, 103)
    Picture1.AutoSize = True
    Picture1.Picture = LoadPicture(tFile)
    Picture1.Align = 3
    Picture1.Align = 1
    
End Sub

Private Sub mnuAbort_Click()
booAbort = True
Unload Me
End Sub

Private Sub mnuBrowser_Click()
WinImageBrowse ("c:\temp.bmp")
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFilePrint_Click()
    'ToDo: Add 'mnuFilePrint_Click' code.
    WinImagePrint ("c:\temp.bmp")
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String
    Dim tFile As String
    Dim res As Long
    Dim p As Pic
    Dim aComment As String
    Dim i As Integer, x As Integer
    
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    x = InStrRev(sFile, "\")
    strPicDir = Left$(sFile, x - 1)
    
    'ToDo: add code to process the opened file
    tFile = "c:\temp.bmp"
    
    'call the mwgfx.dll routine
    res = anytobmps(sFile, tFile, p, 0, 0)
    
    Picture1.AutoSize = True
    Picture1.Picture = LoadPicture(tFile)
    Picture1.Align = 3
    Picture1.Align = 1
    'get the data from the Pic structure
    For i = 0 To 80 Step 1
        If p.comment(i) = 0 Then Exit For
        aComment = aComment + Chr$(p.comment(i))
    Next i
     
    Caption = "PicView - " + aComment + " (" + str(p.width) + "x" + str(p.height) + "x" + str(p.depth) + ")"
    

End Sub

Private Sub mnuSlide_Click()
Call WinSlideShow(strPicDir, 3, 0, 0, 0, 0)
End Sub

Private Sub Picture1_Click()
     Dim tFile As String
     tFile = "c:\temp.bmp"
     Call WinImageShow(tFile, 0)
End Sub

Private Sub rot_Click()
    Dim tFile As String
    Dim p As Pic
    tFile = "c:\temp.bmp"
       
    'call the mwgfx.dll routine
    Call bmprocess(tFile, tFile, p, 100)
    Picture1.AutoSize = True
    Picture1.Picture = LoadPicture(tFile)
    Picture1.Align = 3
    Picture1.Align = 1

End Sub
