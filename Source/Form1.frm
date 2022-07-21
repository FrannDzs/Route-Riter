VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C7083F68-7BF5-4755-9CF0-38D810EC405C}#2.1#0"; "trainlib.ocx"
Begin VB.Form frmLoad 
   Caption         =   "Shape Viewer"
   ClientHeight    =   7740
   ClientLeft      =   255
   ClientTop       =   630
   ClientWidth     =   10350
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   690
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optTab 
      Caption         =   "Animation"
      Height          =   315
      Index           =   2
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   630
      Width           =   1215
   End
   Begin VB.OptionButton optTab 
      Caption         =   "Lighting"
      Height          =   315
      Index           =   1
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   1215
   End
   Begin VB.OptionButton optTab 
      Caption         =   "Main"
      Height          =   315
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Value           =   -1  'True
      Width           =   1215
   End
   Begin trainlib.sfCanvas sfCanvas1 
      Height          =   6675
      Left            =   60
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1020
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11774
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      fogFactor       =   0
      skyTexture      =   ""
      screenshotLocation=   ""
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   10860
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FileName        =   "*.s"
      InitDir         =   "MSTS Shape Files|*.s"
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   0
      Left            =   1260
      ScaleHeight     =   945
      ScaleWidth      =   8985
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   9015
      Begin VB.OptionButton optTex 
         Caption         =   "Winter"
         Height          =   195
         Index           =   5
         Left            =   5940
         TabIndex        =   30
         Top             =   510
         Width           =   1155
      End
      Begin VB.OptionButton optTex 
         Caption         =   "Snow"
         Height          =   195
         Index           =   2
         Left            =   4980
         TabIndex        =   29
         Top             =   510
         Width           =   1155
      End
      Begin VB.OptionButton optTex 
         Caption         =   "Spring"
         Height          =   195
         Index           =   3
         Left            =   5940
         TabIndex        =   28
         Top             =   60
         Width           =   1155
      End
      Begin VB.OptionButton optTex 
         Caption         =   "Autumn"
         Height          =   195
         Index           =   4
         Left            =   5940
         TabIndex        =   27
         Top             =   285
         Width           =   1155
      End
      Begin VB.OptionButton optTex 
         Caption         =   "Night"
         Height          =   195
         Index           =   1
         Left            =   4980
         TabIndex        =   26
         Top             =   285
         Width           =   1155
      End
      Begin VB.OptionButton optTex 
         Caption         =   "Summer"
         Height          =   195
         Index           =   0
         Left            =   4980
         TabIndex        =   25
         Top             =   60
         Width           =   1155
      End
      Begin VB.ListBox lstLod 
         Height          =   840
         Left            =   3180
         TabIndex        =   7
         Top             =   60
         Width           =   1695
      End
      Begin VB.CommandButton cmdAce 
         Caption         =   "Load Shape"
         Height          =   375
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   10260
         Top             =   0
      End
      Begin VB.HScrollBar hDist 
         Height          =   255
         LargeChange     =   5
         Left            =   1200
         Max             =   700
         Min             =   2
         TabIndex        =   6
         Top             =   300
         Value           =   10
         Width           =   1935
      End
      Begin VB.CommandButton cmdShot 
         Caption         =   "Screenshot"
         Height          =   375
         Left            =   60
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lbl 
         Height          =   915
         Left            =   7440
         TabIndex        =   13
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Viewing Distance"
         Height          =   195
         Left            =   1140
         TabIndex        =   12
         Top             =   60
         Width           =   2055
      End
      Begin VB.Label lblDist 
         Alignment       =   1  'Right Justify
         Caption         =   "10.0m"
         Height          =   195
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   2
      Left            =   1260
      ScaleHeight     =   945
      ScaleWidth      =   8985
      TabIndex        =   21
      Top             =   0
      Width           =   9015
      Begin VB.CheckBox chkAni 
         Caption         =   "Animate"
         Height          =   315
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   60
         Width           =   1215
      End
      Begin VB.HScrollBar hAnim 
         Height          =   255
         LargeChange     =   2
         Left            =   2820
         Max             =   1
         Min             =   12
         TabIndex        =   22
         Top             =   360
         Value           =   2
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Animation Speed"
         Height          =   195
         Left            =   720
         TabIndex        =   23
         Top             =   420
         Width           =   2055
      End
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   1
      Left            =   1260
      ScaleHeight     =   945
      ScaleWidth      =   8985
      TabIndex        =   15
      Top             =   0
      Width           =   9015
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   60
         Left            =   5580
         Max             =   1080
         Min             =   360
         TabIndex        =   9
         Top             =   60
         Value           =   360
         Width           =   2055
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         LargeChange     =   10
         Left            =   5580
         Max             =   66
         Min             =   -66
         TabIndex        =   10
         Top             =   420
         Value           =   50
         Width           =   2055
      End
      Begin VB.HScrollBar hShade 
         Height          =   255
         LargeChange     =   5
         Left            =   2580
         Max             =   255
         TabIndex        =   8
         Top             =   60
         Value           =   128
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Time"
         Height          =   195
         Left            =   4920
         TabIndex        =   20
         Top             =   60
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Latitude"
         Height          =   195
         Left            =   4860
         TabIndex        =   19
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "09:00"
         Height          =   195
         Left            =   7800
         TabIndex        =   18
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "50N"
         Height          =   195
         Left            =   7800
         TabIndex        =   17
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Brightness"
         Height          =   195
         Left            =   420
         TabIndex        =   16
         Top             =   60
         Width           =   2115
      End
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mReload 
         Caption         =   "&Reload"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mSecond 
         Caption         =   "Open &Second Shape"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mPos 
         Caption         =   "&Position Second Shape"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu m1 
         Caption         =   "-"
      End
      Begin VB.Menu mQuit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mVis 
         Caption         =   "Parts &Visibility ..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mTextures 
         Caption         =   "&Textures ..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mTool 
      Caption         =   "&Tools"
      Begin VB.Menu mColor 
         Caption         =   "Background &Colour ..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu m4 
         Caption         =   "-"
      End
      Begin VB.Menu mAnim 
         Caption         =   "&Animate"
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu m2 
         Caption         =   "-"
      End
      Begin VB.Menu mBase 
         Caption         =   "&Display Base"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mSky 
         Caption         =   "&Sky"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mGrass 
         Caption         =   "G&rass"
      End
      Begin VB.Menu mGrid 
         Caption         =   "&Grid"
      End
      Begin VB.Menu m5 
         Caption         =   "-"
      End
      Begin VB.Menu mOptions 
         Caption         =   "&Options"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mSkyFile 
         Caption         =   "&Select Sky"
         Shortcut        =   ^S
      End
      Begin VB.Menu mShot 
         Caption         =   "Set Screenshot &Folder"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu mContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu m3 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "&About SView"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About sfCanvas"
      End
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modl As sfModel
Dim mod2 As sfModel
Dim tcbase As sfModel

Dim basemx As D3DMATRIX

Dim vAngl As Single, hAngl As Single, bmove As Boolean  ' view angles
Dim sx As Single, sy As Single      ' mouse shifts

Dim cx As Single, cy As Single, cz As Single, cr As Single    ' centre of viewpoint
Dim mv As D3DVECTOR

Dim rAng As Single, mTim As Single
Dim bLine As Boolean
Dim bColor As Long

Dim iAnim As Integer    ' animation speed

Const GRIDSIZE = 11
Const GRIDM1 = GRIDSIZE - 1

Dim gridVert(GRIDSIZE * GRIDSIZE) As D3DVERTEX
Dim gridVB As Direct3DVertexBuffer8
Dim gridI(GRIDSIZE * 2) As Integer
Dim gridIB(GRIDM1) As Direct3DIndexBuffer8
Dim gridTex As Direct3DTexture8

Dim gridState As Long
Dim FFname As String
Dim FFanim As String


Const LANG_ENGLISH = &H9
Const LANG_FRENCH = &HC
Const LANG_GERMAN = &H7
Const LANG_SPANISH = &HA
Dim lcid As Long
Private Declare Function SetThreadLocale Lib "KERNEL32" (ByVal Locale As Long) As Long
Private Declare Function GetThreadLocale Lib "KERNEL32" () As Long
Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

Const TORAD = 3.14159265358979 / 180   ' to radians

Private Sub chkAni_Click()
    mAnim_Click
End Sub

Private Sub cmdAce_Click()
    On Error Resume Next
    
    If sfCanvas1.isWindowed Then
        Timer1.Enabled = False
        
        mReload.Enabled = False
        mTextures.Enabled = False
        mVis.Enabled = False
        mSecond.Enabled = False
        mAnim.Enabled = False
        mPos.Enabled = False
            
        DoEvents
        
        cdlg.Flags = cdlOFNAllowMultiselect Or cdlOFNFileMustExist Or cdlOFNExplorer
        cdlg.Filter = "MSTS S file|*.s"
        cdlg.DialogTitle = "Select Shape File"
        On Error Resume Next
        cdlg.ShowOpen
        
        ' trap cancel errors
        If Err.Number = 0 Then
            Set modl = Nothing
            Set mod2 = Nothing
            
            FFanim = ""
            showFile cdlg.filename
            mv.x = 0#
            mv.y = 0#
            mv.z = 0#
        Else
            If modl Is Nothing Then
                Me.Caption = "Shape Viewer"
            Else
                Timer1.Enabled = True
            
                mReload.Enabled = True
                mTextures.Enabled = True
                mVis.Enabled = True
                mSecond.Enabled = True
                mAnim.Enabled = True
                mPos.Enabled = True
            End If
    
        End If
    End If

End Sub

Public Sub showFile(fName As String)
    Dim i As Long, j As Long
    
        #If DebugLog = 1 Then
        App.StartLogging App.path & "\sview.log", vbLogToFile
        App.LogEvent "Selected " & cdlg.filename
        #End If
        
    ' piece out multiple names ?
    i = InStr(fName, Chr(0))
    If i > 0 Then
        j = InStr(i + 1, fName, Chr(0))
        FFanim = Left(fName, i - 1) & "\" & Mid(fName, j + 1)
        FFanim = Replace(FFanim, """", "")
        fName = Replace(Left(fName, j - 1), Chr(0), "\")
    End If
    
    ' trigger loading of file
    fName = Replace(fName, """", "")
    FFname = fName
    
    Screen.MousePointer = vbHourglass
    
    #If DebugLog = 1 Then
    App.LogEvent "Setting name " & fName
    #End If
    setModel modl, fName
    
    If FFanim <> "" Then
        setModel mod2, FFanim
        'gmfv.x = 0#
        'gmfv.y = 0#
        'gmfv.z = 0#
    End If
    
    #If DebugLog = 1 Then
    App.LogEvent "sfModel finished model load"
    #End If
    Me.Caption = "Shape Viewer - " & modl.filename & IIf(FFanim <> "", " (" & FFanim & ")", "")
    
    lstLod.Clear
    If modl.numLODs = 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    For i = 1 To modl.numLODs
        lstLod.AddItem modl.lodRange(i) & "m " & modl.lodPolys(i) & "polys"
    Next
    lstLod.ListIndex = 0
    ' get model approx size
    modl.getSphere cx, cy, cz, cr
    ' sets the camera position
    
    If cr * 4 < 4 Then
        hDist.value = 4
    ElseIf cr * 4 > hDist.max Then
        hDist.value = hDist.max
    Else
        hDist.value = cr * 4
    End If
    
    ' error messages ?
    If modl.info <> "" Then lbl.Caption = modl.info
    
    #If DebugLog = 1 Then
    App.LogEvent "set controls complete"
    #End If
    ' defaults transforms and lighting
    
    
    mTim = Timer
    #If DebugLog = 1 Then
    App.LogEvent "init dx and timer"
    #End If
    
    mReload.Enabled = True
    mTextures.Enabled = True
    mVis.Enabled = True
    mAnim.Enabled = True
    mSecond.Enabled = True
    If Not mod2 Is Nothing Then mPos.Enabled = True
    
    hShade_scroll
    
    Screen.MousePointer = vbDefault
    Timer1.Enabled = True
End Sub

Private Sub render()
    
    With sfCanvas1
        ' clear the screenbuffer
        If .BeginScene Then
            If mGrid.Checked Or mGrass.Checked Then showGrid
            If mBase.Checked Then
                basemx.m41 = mv.x
                basemx.m42 = mv.y - 0.54
                basemx.m43 = mv.z
                tcbase.setMatrix basemx
                tcbase.showModel
            End If
            ' render the model - 5 second rotation time
            If mAnim.Checked Then
                modl.showModel lstLod.ListIndex + 1, (Timer * 100 Mod (iAnim * 100)) / iAnim
            Else
                modl.showModel lstLod.ListIndex + 1
            End If
            If Not mod2 Is Nothing Then mod2.showModel
            .EndScene
        Else
            modl.ResetState
            .Reset
        End If
    End With
End Sub


Private Sub cmdShot_Click()
    sfCanvas1.saveImage
End Sub

Private Sub Form_Activate()
    Timer1.Interval = 1
End Sub

Private Sub Form_Deactivate()
    Timer1.Interval = 100
End Sub

Private Sub Form_GotFocus()
    Timer1.Interval = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not modl Is Nothing Then
        Select Case KeyCode
        Case vbKeyDown
            If Shift = 0 Then
                vAngl = vAngl - TORAD * 2
            ElseIf Shift And vbCtrlMask Then
                mv.y = mv.y + 0.25
            End If
            hDist_scroll
            KeyCode = 0
        Case vbKeyUp
            If Shift = 0 Then
                vAngl = vAngl + TORAD * 2
            ElseIf Shift And vbCtrlMask Then
                mv.y = mv.y - 0.25
            End If
            hDist_scroll
            KeyCode = 0
        Case vbKeyLeft
            If Shift = 0 Then
                hAngl = hAngl + TORAD * 2
            ElseIf Shift And vbCtrlMask Then
                mv.z = mv.z + Cos(hAngl) / 3
                mv.x = mv.x - Sin(hAngl) / 3
            End If
            hDist_scroll
            KeyCode = 0
        Case vbKeyRight
            If Shift = 0 Then
                hAngl = hAngl - TORAD * 2
            ElseIf Shift And vbCtrlMask Then
                mv.z = mv.z - Cos(hAngl) / 3
                mv.x = mv.x + Sin(hAngl) / 3
            End If
            hDist_scroll
            KeyCode = 0
        Case vbKeyF12
            Timer1.Enabled = False
            ' remove model first
            gridState = 0
            Set modl = Nothing
            Set mod2 = Nothing
            Set tcbase = Nothing
            Set gridVB = Nothing
            If sfCanvas1.isWindowed Then
                pic1(0).Visible = False
                pic1(1).Visible = False
                pic1(2).Visible = False
                sfCanvas1.FullScreen 0, 0
            Else
                pic1(0).Visible = True
                pic1(1).Visible = True
                pic1(2).Visible = True
                sfCanvas1.Windowed
            End If
            
            sfCanvas1.initDefaults
            
            setModel modl, FFname
            
            If FFanim <> "" Then
                setModel mod2, FFanim
            End If
            
            initGrid
            hShade_scroll
            Timer1.Enabled = True
            
        Case vbKeyPrint, vbKeySnapshot, vbKeyF11
            cmdShot_Click
        Case vbKeySubtract
            If hDist.value < hDist.max Then hDist.value = hDist.value + 1
            KeyCode = 0
        Case vbKeyAdd
            If hDist.value > hDist.Min Then hDist.value = hDist.value - 1
            KeyCode = 0
        Case vbKeyPageUp
            If hAnim.value < hAnim.max Then hAnim.value = hAnim.value + 1
            KeyCode = 0
        Case vbKeyPageDown
            If hAnim.value > hAnim.Min Then hAnim.value = hAnim.value - 1
            KeyCode = 0
        Case vbKeyInsert
            If hShade.value < hShade.max Then hShade.value = hShade.value + 1
            KeyCode = 0
        Case vbKeyDelete
            If hShade.value > hShade.Min Then hShade.value = hShade.value - 1
            KeyCode = 0
        Case vbKeyHome
            mv.x = 0
            mv.y = 0
            mv.z = 0
        Case vbKeyBack
            bLine = Not bLine
            If bLine Then
                sfCanvas1.d3dd.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
            Else
                sfCanvas1.d3dd.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
            End If
        Case vbKeyG
            If mGrid.Checked Then
                mGrass_Click
            ElseIf mGrass.Checked Then
                mGrass_Click
            Else
                mGrid_Click
            End If
        Case vbKeyEscape
            Unload Me
            End
        Case vbKey0 To vbKey9
            
            If Shift And vbCtrlMask Then
                saveCam KeyCode - vbKey0
            Else
                restCam KeyCode - vbKey0
            End If
            KeyCode = 0
        End Select
    End If
    
End Sub


Private Sub Form_Load()
    loadLang
    
    ' make form visible before running startup method
    On Error Resume Next
    sfCanvas1.BackColor = CLng(GetSetting("Decapod", App.Title, "BackgroundCol", &HFFFFFF))
    bColor = sfCanvas1.BackColor
    hAnim.value = CLng(GetSetting("Decapod", App.Title, "AnimSpeed", 2))
    hShade.value = CLng(GetSetting("Decapod", App.Title, "BackgroundShade", 128))
    mGrid.Checked = CBool(GetSetting("Decapod", App.Title, "ShowGrid", "False"))
    mGrass.Checked = CBool(GetSetting("Decapod", App.Title, "ShowGrass", "False"))
    mBase.Checked = CBool(GetSetting("Decapod", App.Title, "ShowBase", "False"))
    
    sfCanvas1.showFPS = CBool(GetSetting("Decapod", App.Title, "ShowFPS", "False"))
    sfCanvas1.tripleBuffer = CBool(GetSetting("Decapod", App.Title, "tripleBuffer", "False"))
    sfCanvas1.AntiAlias = CBool(GetSetting("Decapod", App.Title, "AntiAlias", 0))
    sfCanvas1.fogFactor = Val(GetSetting("Decapod", App.Title, "fogFactor", 0.08))
    sfCanvas1.fog = CInt(GetSetting("Decapod", App.Title, "fog", 0))
    sfCanvas1.anisotropicFilter = CBool(GetSetting("Decapod", App.Title, "anisotropicFilter", "False"))
    sfCanvas1.sky = CBool(GetSetting("Decapod", App.Title, "Sky", 0))
    
    sfCanvas1.skyTexture = GetSetting("Decapod", App.Title, "skyTexture", App.path & "\sky.jpg")
    sfCanvas1.screenshotLocation = GetSetting("Decapod", App.Title, "screenshotLocation", App.path & "\SFCAN")
    cdlg.filename = GetSetting("Decapod", App.Title, "Last", App.path & "\*.s")
    sfCanvas1.time = CLng(GetSetting("Decapod", App.Title, "Time", 540))
    sfCanvas1.latitude = CLng(GetSetting("Decapod", App.Title, "Lat", 50))
    
    mSky.Checked = sfCanvas1.sky
    iAnim = hAnim.value
    HScroll1.value = sfCanvas1.time
    HScroll2.value = sfCanvas1.latitude
    On Error GoTo 0
    
    Show
    sfCanvas1.Startup
    sfCanvas1.initDefaults
            
    ' set up the grid
    initGrid
    
    If Command <> "" Then
        ' swap dir when double click
        On Error Resume Next
        cdlg.filename = Left(Command, InStrRev(Command, "\")) & "*.s"
        showFile Command
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Function LoadRString(code As Long) As String
   Dim s As String * 255
   LoadString App.hInstance, code, s, 255
   Dim posNull As Long
   posNull = InStr(s, Chr$(0))
    If posNull > 0 Then
        LoadRString = Left$(s, posNull - 1)
    Else
        LoadRString = s
    End If
End Function


Private Sub loadLang()
    Dim s As String, i As Integer
    On Error Resume Next
    
    lcid = GetThreadLocale
    Select Case lcid And &HFF
    Case LANG_GERMAN, LANG_FRENCH
        lcid = 1024 Or (lcid And &HFF)
        SetThreadLocale lcid
    Case LANG_SPANISH
        lcid = 3072 Or (lcid And &HFF)
        SetThreadLocale lcid
    End Select
    
    s = LoadRString(101)
    If s <> "" Then mColor.Caption = s
    s = LoadRString(102)
    If s <> "" Then mOpen.Caption = s
    s = LoadRString(103)
    If s <> "" Then mQuit.Caption = s
    s = LoadRString(104)
    If s <> "" Then mSky.Caption = s
    s = LoadRString(105)
    If s <> "" Then mGrid.Caption = s
    s = LoadRString(106)
    If s <> "" Then mReload.Caption = s
    s = LoadRString(107)
    If s <> "" Then cmdAce.Caption = s
    s = LoadRString(108)
    If s <> "" Then mHelp.Caption = s
    s = LoadRString(109)
    If s <> "" Then mTool.Caption = s
    s = LoadRString(110)
    If s <> "" Then mEdit.Caption = s
    s = LoadRString(111)
    If s <> "" Then
        mAbout.Caption = s & " SView"
        mnuAbout.Caption = s & " sfCanvas"
    End If
    s = LoadRString(112)
    If s <> "" Then mContents.Caption = s
    s = LoadRString(113)
    If s <> "" Then mOptions.Caption = s
    
    s = LoadRString(117)
    If s <> "" Then mTextures.Caption = s
    s = LoadRString(118)
    If s <> "" Then mGrass.Caption = s
    
    s = LoadRString(119)
    If s <> "" Then Label1.Caption = s
    s = LoadRString(120)
    If s <> "" Then Label2.Caption = s
    s = LoadRString(121)
    If s <> "" Then Label4.Caption = s
    s = LoadRString(122)
    If s <> "" Then mSecond.Caption = s
    s = LoadRString(123)
    If s <> "" Then mPos.Caption = s
    s = LoadRString(124)
    If s <> "" Then optTab(0).Caption = s
    s = LoadRString(125)
    If s <> "" Then optTab(2).Caption = s
    s = LoadRString(126)
    If s <> "" Then optTab(1).Caption = s
    s = LoadRString(127)
    If s <> "" Then Label7.Caption = s

    For i = 0 To 5
        s = LoadRString(128 + i)
        If s <> "" Then optTex(i).Caption = s
    Next
End Sub

Private Sub Form_LostFocus()
    Timer1.Interval = 100
End Sub

Private Sub Form_Resize()
    Dim b As Boolean
    
    b = Timer1.Enabled
    If b Then Timer1.Enabled = False
    
    If WindowState <> vbMinimized Then
        If height < 2000 Then
            height = 2000
            Exit Sub
        End If
        If width < 1000 Then
            width = 1000
            Exit Sub
        End If
        
        With sfCanvas1
            '.Visible = True
            .width = ScaleWidth - .Left * 2
            .height = ScaleHeight - .Top - .Left
            gridState = 0
            If Not modl Is Nothing Then modl.ResetState
            If Not mod2 Is Nothing Then mod2.ResetState
            If Not tcbase Is Nothing Then tcbase.ResetState
            sfCanvas1.Reset
        End With
        Timer1.Interval = 1
    Else
        Timer1.Interval = 1000
        'sfCanvas1.Visible = False
    End If
    
    Timer1.Enabled = b
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set modl = Nothing
    Set mod2 = Nothing
    Set tcbase = Nothing
    SaveSetting "Decapod", App.Title, "AnimSpeed", hAnim.value
    SaveSetting "Decapod", App.Title, "BackgroundShade", hShade.value
    SaveSetting "Decapod", App.Title, "BackgroundCol", bColor
    SaveSetting "Decapod", App.Title, "ShowGrid", mGrid.Checked
    SaveSetting "Decapod", App.Title, "ShowBase", mBase.Checked
    SaveSetting "Decapod", App.Title, "ShowGrass", mGrass.Checked
    SaveSetting "Decapod", App.Title, "ShowFPS", sfCanvas1.showFPS
    SaveSetting "Decapod", App.Title, "tripleBuffer", sfCanvas1.tripleBuffer
    SaveSetting "Decapod", App.Title, "fog", sfCanvas1.fog
    SaveSetting "Decapod", App.Title, "fogFactor", Replace(CStr(sfCanvas1.fogFactor), ",", ".")
    SaveSetting "Decapod", App.Title, "anisotropicFilter", sfCanvas1.anisotropicFilter
    SaveSetting "Decapod", App.Title, "AntiAlias", sfCanvas1.AntiAlias
    SaveSetting "Decapod", App.Title, "Sky", sfCanvas1.sky
    SaveSetting "Decapod", App.Title, "skyTexture", sfCanvas1.skyTexture
    SaveSetting "Decapod", App.Title, "screenshotLocation", sfCanvas1.screenshotLocation
    SaveSetting "Decapod", App.Title, "Last", cdlg.filename
    SaveSetting "Decapod", App.Title, "Time", sfCanvas1.time
    SaveSetting "Decapod", App.Title, "Lat", sfCanvas1.latitude

    End
End Sub


Private Sub hAnim_Change()
    hAnim_Scroll
End Sub

Private Sub hAnim_Scroll()
    iAnim = hAnim.value
End Sub

Private Sub hDist_Change()
    hDist_scroll
End Sub

Private Sub hDist_scroll()
    
    sfCanvas1.setCamera hDist.value * Cos(vAngl) * Cos(hAngl) / 2 + cx, hDist.value * Sin(vAngl) / 2 + cy, hDist.value * Sin(hAngl) * Cos(vAngl) / 2 + cz, cx, cy, cz, 0#, 1#, 0#
    
    lblDist.Caption = Format(hDist.value / 2, "#0.0") & "m"

End Sub

Private Sub hShade_Change()
    hShade_scroll
End Sub

Private Sub hShade_scroll()
    Dim r As Long, g As Long, b As Long
    'b = bColor Mod 256
    'g = (bColor \ 256) Mod 256
    'r = (bColor \ 65536) Mod 256
    'sfCanvas1.BackColor = RGB(b * hShade.Value / 255, g * hShade.Value / 255, r * hShade.Value / 255)
    b = hShade.value / hShade.max * 255
    If Not sfCanvas1.d3dd Is Nothing Then sfCanvas1.d3dd.SetRenderState D3DRS_AMBIENT, D3DColorXRGB(b, b, b)
End Sub

Private Sub lstLod_Click()
    Dim i As Long
    i = lstLod.ListIndex + 1
    lbl.Caption = Format(modl.lodPolys(i) * Sqr(modl.lodPrims(i) * 2) / 21803, "##0.00") & " kujus"
End Sub

Private Sub mAbout_Click()
    frmAbout1.Show vbModal
End Sub

Private Sub mAnim_Click()
    mAnim.Checked = Not mAnim.Checked
End Sub

Private Sub mBase_Click()
    mBase.Checked = Not mBase.Checked
End Sub

Private Sub mColor_Click()
    On Error Resume Next
    cdlg.Flags = cdlCCFullOpen
    cdlg.Color = bColor
    cdlg.DialogTitle = "Select Colour"
    cdlg.ShowColor
    If Err.Number = 0 Then
        bColor = cdlg.Color
        Dim r As Long, g As Long, b As Long
        b = bColor Mod 256
        g = (bColor \ 256) Mod 256
        r = (bColor \ 65536) Mod 256
        sfCanvas1.BackColor = RGB(b, g, r)
    End If
End Sub

Private Sub mContents_Click()
    HH_DISPLAY_Click Me.hwnd
End Sub

Private Sub mGrass_Click()
    mGrass.Checked = Not mGrass.Checked
    If mGrid.Checked Or mGrass.Checked Then
        Set gridTex = sfCanvas1.loadTex(App.path & "\grass.bmp")
        mGrid.Checked = False
    End If
End Sub

Private Sub mGrid_Click()
    mGrid.Checked = Not mGrid.Checked
    If mGrid.Checked Or mGrass.Checked Then
        Set gridTex = sfCanvas1.loadTex(App.path & "\grid.bmp")
        mGrass.Checked = False
    End If
End Sub



Private Sub mnuAbout_Click()
    sfCanvas1.About
End Sub

Private Sub mOpen_Click()
    cmdAce_Click
End Sub

Private Sub mOptions_Click()
    If sfCanvas1.isWindowed Then
        sfCanvas1.showOptions
        mSky.Checked = sfCanvas1.sky
    End If
End Sub

Private Sub mPos_Click()
    Dim f As frmPos
    
    If sfCanvas1.isWindowed Then
        Set f = New frmPos
        f.Move Left, Top
        f.Show vbModal
        Set f = Nothing
    End If
End Sub

Private Sub mQuit_Click()
    Unload Me
    End
End Sub

Private Sub mReload_Click()
    Timer1.Enabled = False
    showFile FFname
    If FFanim <> "" Then
        setModel mod2, FFanim
    End If
    Timer1.Enabled = True
End Sub

Private Sub mSecond_Click()
    On Error Resume Next
    
    cdlg.Flags = cdlOFNFileMustExist Or cdlOFNExplorer
    cdlg.Filter = "MSTS S file|*.s"
    cdlg.filename = "*.s"
    cdlg.DialogTitle = "Select Shape File"
    On Error Resume Next
    cdlg.ShowOpen
    
    ' trap cancel errors
    If Err.Number = 0 Then
        Set mod2 = Nothing
        FFanim = cdlg.filename
        setModel mod2, FFanim
        mPos.Enabled = True
        mPos_Click
    End If
    
End Sub

Private Sub mShot_Click()
    If sfCanvas1.isWindowed Then
        cdlg.DialogTitle = "Goto the folder and enter the prefix you require"
        cdlg.Flags = cdlOFNExplorer
        cdlg.Filter = "Enter the prefix|*.asdgfasd"
        cdlg.filename = sfCanvas1.screenshotLocation
        On Error Resume Next
        cdlg.ShowOpen
        
        ' trap cancel errors
        If Err.Number = 0 Then
            sfCanvas1.screenshotLocation = Left(cdlg.filename, InStrRev(cdlg.filename, ".") - 1)
        End If
    End If
End Sub

Private Sub mSky_Click()
    mSky.Checked = Not mSky.Checked
    sfCanvas1.sky = mSky.Checked
End Sub

Private Sub mSkyFile_Click()
    Dim s As String
    If sfCanvas1.isWindowed Then
        s = cdlg.filename
        cdlg.filename = App.path & "\*.jpg"
        cdlg.Flags = cdlOFNFileMustExist Or cdlOFNExplorer
        cdlg.DialogTitle = "Select Sky Texture file"
        cdlg.Filter = "Image file (bmp,jpg,tga)|*.jpg;*.bmp;*.tga;*.dds"
        cdlg.filename = sfCanvas1.skyTexture
        On Error Resume Next
        cdlg.ShowOpen
        
        ' trap cancel errors
        If Err.Number = 0 Then
            sfCanvas1.skyTexture = cdlg.filename
        End If
        cdlg.filename = s
    End If
End Sub

Private Sub mTextures_Click()
    Dim f As Dialog
    Set f = New Dialog
    Set f.modl = modl
    Set f.mod2 = mod2
    f.Show vbModal
    Unload f
    Set f = Nothing
End Sub


Private Sub optTab_Click(Index As Integer)
    pic1(Index).ZOrder
End Sub

Private Sub optTex_Click(Index As Integer)
    Dim s As String, sf As String, s1 As String
    Dim w As Long, h As Long, l As Boolean
    
    Dim i As Long
    
    Select Case Index
    Case 0
        s = "..\textures\"
    Case 1
        s = "..\textures\night\;..\textures\"
    Case 2
        s = "..\textures\snow\;..\textures\wintersnow\;..\textures\"
    Case 3
        s = "..\textures\spring\;..\textures\"
    Case 4
        s = "..\textures\autumn\;..\textures\"
    Case 5
        s = "..\textures\winter\;..\textures\"
    End Select
    
    If Not modl Is Nothing Then
        modl.textureSearchPath = s
        For i = 0 To modl.nTextures - 1
            modl.getTexInfo i, sf, s1, w, h, l
            
            modl.SetTexture i, modl.loadAce(sf)
        Next
    End If
    If Not mod2 Is Nothing Then
        mod2.textureSearchPath = s
        For i = 0 To mod2.nTextures - 1
            mod2.getTexInfo i, sf, s1, w, h, l
            mod2.SetTexture i, mod2.loadAce(sf)
        Next
    End If
End Sub

Private Sub sfCanvas1_KeyDown(KeyCode As Integer, Shift As Integer)
    ' arrows ?
    If KeyCode > 36 And KeyCode < 41 Then
        Form_KeyDown KeyCode, Shift
    End If
End Sub

Private Sub sfCanvas1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' get start position
    bmove = True
    sx = x
    sy = y
End Sub

Private Sub sfCanvas1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim dy As Single
    
    If Not bmove Then Exit Sub
    dy = y - sy
    If Button = vbLeftButton Then
        ' rotate
        vAngl = vAngl + dy / 2500
        hAngl = hAngl - (x - sx) / 2500
        hDist_scroll
    Else
        ' zoom
        dy = hDist.value + dy / 100
        If dy > hDist.max Then
            dy = hDist.max
        ElseIf dy < hDist.Min Then
            dy = hDist.Min
        End If
        hDist.value = dy
    End If
    sy = y
    sx = x
    
End Sub

Private Sub sfCanvas1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bmove = False
End Sub

Private Sub Timer1_Timer()
    Dim mw As D3DMATRIX, mt As D3DMATRIX
    
    D3DXMatrixIdentity mw
    
    DXCopyMemory mw.m41, mv, 12
    
    modl.setMatrix mw

    If Not mod2 Is Nothing Then
        mw.m41 = mw.m41 + gmfv.x
        mw.m42 = mw.m42 + gmfv.y
        mw.m43 = mw.m43 + gmfv.z
        If gflip Then
            D3DXMatrixRotationY mt, 180 * TORAD
            D3DXMatrixMultiply mw, mt, mw
        End If
        
        mod2.setMatrix mw
    End If
        
    ' show the model
    render
End Sub

Private Sub initGrid()
    Dim i As Long, j As Long
    
    ' build vertices
    For i = 0 To GRIDM1
        For j = 0 To GRIDM1
            setVert gridVert(i * GRIDSIZE + j), i * 400 / GRIDM1 - 200, 0, j * 400 / GRIDM1 - 200, 0, 1, 0, i * 100 / GRIDM1, j * 100 / GRIDM1
        Next
    Next
    ' build strips
    For i = 0 To GRIDSIZE - 2
        For j = 0 To GRIDM1
            gridI(j * 2) = (i + 1) * GRIDSIZE + j
            gridI(j * 2 + 1) = i * GRIDSIZE + j
        Next
        Set gridIB(i) = sfCanvas1.d3dd.CreateIndexBuffer(4 * GRIDSIZE, D3DUSAGE_WRITEONLY, D3DFMT_INDEX16, D3DPOOL_MANAGED)
        D3DIndexBuffer8SetData gridIB(i), 0, GRIDSIZE * 4, 0, gridI(0)
    Next
    Set gridVB = sfCanvas1.d3dd.CreateVertexBuffer(Len(gridVert(0)) * GRIDSIZE * GRIDSIZE, D3DUSAGE_WRITEONLY, D3DFVF_VERTEX, D3DPOOL_MANAGED)
    D3DVertexBuffer8SetData gridVB, 0, GRIDSIZE * GRIDSIZE * Len(gridVert(0)), 0, gridVert(0)
    
    On Error Resume Next
    
    If mGrass.Checked Then
        Set gridTex = sfCanvas1.loadTex(App.path & "\grass.bmp")
    Else
        Set gridTex = sfCanvas1.loadTex(App.path & "\grid.bmp")
    End If
    
    ' load display base
    Set tcbase = New sfModel
    Set tcbase.d3dd = sfCanvas1.d3dd
    tcbase.filename = App.path & "\TC_Displaybase.s"
    D3DXMatrixIdentity basemx
    basemx.m42 = -0.54
    tcbase.setMatrix basemx
    
End Sub

Private Sub showGrid()
    Dim i As Long
    Dim m As D3DMATRIX
    Dim mtrl As D3DMATERIAL8
    
    
    D3DXMatrixIdentity m
    DXCopyMemory m.m41, mv, 12
    
    With sfCanvas1.d3dd
        .SetTransform D3DTS_WORLD, m
        If gridState = 0 Then
            mtrl.diffuse.r = 1
            mtrl.diffuse.g = 1
            mtrl.diffuse.b = 1
            mtrl.diffuse.a = 1
            mtrl.Ambient = mtrl.diffuse
            .BeginStateBlock
            .SetMaterial mtrl
            .SetRenderState D3DRS_ALPHABLENDENABLE, 0
            .SetRenderState D3DRS_ZENABLE, 0
            .SetRenderState D3DRS_SPECULARENABLE, 0
            .SetVertexShader D3DFVF_VERTEX
            .SetStreamSource 0, gridVB, Len(gridVert(0))
            gridState = .EndStateBlock
        End If
        .ApplyStateBlock gridState
        .SetTexture 0, gridTex
        For i = 0 To GRIDM1 - 1
            .SetIndices gridIB(i), 0
            .DrawIndexedPrimitive D3DPT_TRIANGLESTRIP, i * GRIDSIZE, GRIDSIZE * 2, 0, GRIDM1 * 2
        Next
        .SetRenderState D3DRS_ZENABLE, 1
    End With
End Sub

Private Sub setVert(v As D3DVERTEX, x As Single, y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single)
    v.x = x
    v.y = y
    v.z = z
    v.nx = nx
    v.ny = ny
    v.nz = nz
    v.tu = tu
    v.tv = tv
End Sub

Private Sub setModel(modx As sfModel, s As String)
    Set modx = New sfModel
    ' set the drawing device
    Set modx.d3dd = sfCanvas1.d3dd
    modx.textureSearchPath = "..\textures\"
    ' load the S file
    modx.filename = s
    optTex(0).value = True
End Sub

Private Sub HScroll1_Change()
    Dim t As Single
    t = HScroll1.value
    Label5.Caption = t \ 60 & ":" & Format(t Mod 60, "00")
    sfCanvas1.time = HScroll1.value
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub HScroll2_Change()
    Label3.Caption = Abs(HScroll2.value) & IIf(HScroll2.value < 0, "S", "N")
    sfCanvas1.latitude = HScroll2.value
End Sub

Private Sub HScroll2_Scroll()
    HScroll2_Change
End Sub

Private Sub saveCam(k As Integer)
    SaveSetting "Decapod", App.Title, "kd" & k, hDist.value
    SaveSetting "Decapod", App.Title, "kv" & k, vAngl
    SaveSetting "Decapod", App.Title, "kh" & k, hAngl
    SaveSetting "Decapod", App.Title, "kcx" & k, cx
    SaveSetting "Decapod", App.Title, "kcy" & k, cy
    SaveSetting "Decapod", App.Title, "kcz" & k, cz
    lbl.Caption = "Camera position " & k & " saved."
End Sub
Private Sub restCam(k As Integer)
    vAngl = GetSetting("Decapod", App.Title, "kv" & k, 0)
    hAngl = GetSetting("Decapod", App.Title, "kh" & k, 0)
    cx = GetSetting("Decapod", App.Title, "kcx" & k, 0)
    cy = GetSetting("Decapod", App.Title, "kcy" & k, 2)
    cz = GetSetting("Decapod", App.Title, "kcz" & k, 0)
    hDist.value = GetSetting("Decapod", App.Title, "kd" & k, 44)
    hDist_Change
    lbl.Caption = ""
End Sub
