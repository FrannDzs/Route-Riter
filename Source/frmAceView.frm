VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAceView 
   Caption         =   "AceView"
   ClientHeight    =   4185
   ClientLeft      =   3615
   ClientTop       =   2550
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   6270
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      Height          =   4185
      Left            =   0
      ScaleHeight     =   4125
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   0
      Width           =   6735
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
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu test1 
      Caption         =   "Image"
      Begin VB.Menu ShowAlpha 
         Caption         =   "Show Alpha"
      End
      Begin VB.Menu showmain 
         Caption         =   "Show Image"
      End
   End
End
Attribute VB_Name = "frmAceView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
End Sub


Private Sub Form_Resize()
    Picture1.Align = 3
    Picture1.Align = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.width
        SaveSetting App.Title, "Settings", "MainHeight", Me.height
    End If
End Sub



Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub


Private Sub mnuFileOpen_Click()
    Dim sFile As String
    Dim tFile As String
    Dim aFile As String
    Dim aComment As String
    Dim i As Integer
    Dim res As Long
    Dim p As Pic

    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
         .Filter = "ACE Files (*.ace)|*.ace"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    tFile = "temp.bmp"
    aFile = "tema.bmp"
    
    res = AceToBmps(sFile, tFile, aFile)
    Picture1.AutoSize = True
    Picture1.Picture = LoadPicture(tFile)
    Picture1.Align = 3
    Picture1.Align = 1
    res = CheckAce(sFile, p)
    For i = 0 To 80 Step 1
        If p.comment(i) = 0 Then Exit For
        aComment = aComment + Chr$(p.comment(i))
    Next i
    
        
    Caption = "AceView - " + aComment + " (" + str(p.width) + "x" + str(p.height) + "x" + str(p.depth) + ")"
    
End Sub

Private Sub ShowAlpha_Click()
    Dim tFile As String
    
    tFile = "tema.bmp"
    Picture1.AutoSize = True
    Picture1.Picture = LoadPicture(tFile)
    Picture1.Align = 3
    Picture1.Align = 1

End Sub

Private Sub showmain_Click()
    Dim tFile As String
    
    tFile = "temp.bmp"
    Picture1.AutoSize = True
    Picture1.Picture = LoadPicture(tFile)
    Picture1.Align = 3
    Picture1.Align = 1

End Sub
