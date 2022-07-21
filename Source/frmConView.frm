VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C7083F68-7BF5-4755-9CF0-38D810EC405C}#2.20#0"; "trainlib.ocx"
Begin VB.Form frmConView 
   AutoRedraw      =   -1  'True
   Caption         =   "MSTS Consist Viewer"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   13590
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar Htime 
      Height          =   375
      LargeChange     =   60
      Left            =   8280
      Max             =   1080
      Min             =   260
      SmallChange     =   20
      TabIndex        =   14
      Top             =   8160
      Value           =   540
      Width           =   1455
   End
   Begin VB.HScrollBar hDist 
      Height          =   375
      LargeChange     =   5
      Left            =   9840
      Max             =   700
      Min             =   2
      TabIndex        =   11
      Top             =   8160
      Value           =   30
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   11280
      TabIndex        =   10
      Top             =   8160
      Width           =   1095
   End
   Begin trainlib.sfCanvas sfCanvas1 
      Height          =   7455
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   13150
      BackColor       =   14737632
      AutoRedraw      =   0   'False
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
   Begin VB.CommandButton cmdCamPrev 
      Caption         =   "cam x+"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCamNext 
      Caption         =   "cam x-"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCamReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCamDown 
      Caption         =   "cam v"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCamUp 
      Caption         =   "cam ^"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCamForward 
      Caption         =   "cam ~|"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCamBack 
      Caption         =   "cam L"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCamRight 
      Caption         =   "cam >"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCamLeft 
      Caption         =   "cam <"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   8040
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgShapeSelector 
      Left            =   7080
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Light:"
      Height          =   255
      Left            =   8760
      TabIndex        =   15
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Viewing Dist."
      Height          =   255
      Left            =   9840
      TabIndex        =   12
      Top             =   7800
      Width           =   1215
   End
End
Attribute VB_Name = "frmConView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text


Dim Trains As Collection
Dim Train As MovableObject
'Dim f As sfModel
Dim ReSize As clsRsize
Dim CamX As Single
Dim CamY As Single
Dim CamZ As Single
Dim LookAtX As Single
Dim LookAtY As Single
Dim LookAtZ As Single
Dim LookAt As Integer
Dim i As Integer
Dim gridState As Long

Const GRIDSIZE = 11
Const GRIDM1 = GRIDSIZE - 1

Dim gridVert(GRIDSIZE * GRIDSIZE) As D3DVERTEX
Dim gridVB As Direct3DVertexBuffer8
Dim gridI(GRIDSIZE * 2) As Integer
Dim gridIB(GRIDM1) As Direct3DIndexBuffer8
Dim gridTex As Direct3DTexture8



Dim vAngl As Single, hAngl As Single, bmove As Boolean  ' view angles
Dim sX As Single, sY As Single      ' mouse shifts









Private Sub ResetCamera()
    sfCanvas1.setCamera CamX, CamY, CamZ, LookAtX, LookAtY, LookAtZ, 0, 1, 0
    RefreshForm
End Sub

Private Sub cmdCamBack_Click()
    
    CamZ = CamZ + 4
    ResetCamera
End Sub

Private Sub cmdCamDown_Click()

    CamY = CamY - 2
  '  LookAtY = LookAtY - 2
    ResetCamera
End Sub

Private Sub cmdCamForward_Click()

    CamZ = CamZ - 4
    ResetCamera
End Sub

Private Sub cmdCamLeft_Click()

    CamZ = CamZ - 15
    LookAtZ = LookAtZ - 15
    ResetCamera
End Sub




Private Sub cmdCamNext_Click()
CamX = CamX + 2
LookAtX = LookAtX + 2
ResetCamera

End Sub

Private Sub cmdCamPrev_Click()
CamX = CamX - 2
LookAtX = LookAtX - 2
ResetCamera

End Sub


Private Sub cmdCamReset_Click()
CamZ = -6
    LookAtZ = -6
 CamY = 2
 LookAtY = 2
   CamX = 15
    LookAtX = 0
    ResetCamera
End Sub

Private Sub cmdCamRight_Click()
CamZ = CamZ + 15
    LookAtZ = LookAtZ + 15
    
    ResetCamera
End Sub

Private Sub cmdCamUp_Click()


    CamY = CamY + 2
  
    ResetCamera
End Sub













Sub LoadShape(Shapefile As String, ErrName As Boolean)
    
    
    ' create a new model object
    Set Train = New MovableObject
    'Set f = New sfModel
    ' the d3dd object must come from the sfCanvas control
    Set Train.Model.d3dd = sfCanvas1.d3dd
    
    ' by default, the program searches the current folder first.
    ' if not there, it uses the textureSearchPath (delimited by ;) to look for the texture files
    
    Train.Model.textureSearchPath = "..\textures"
    
    ' setting this property triggers loading of the S file and textures
    If Not Train.LoadShape(Shapefile) Then
        'MsgBox Train.LastError, vbCritical, "Error Loading Shape"
        Set Train = Nothing
        ErrName = True
    End If

    Trains.Add Train
    
   Exit Sub

End Sub

Private Sub Command1_Click()
Unload Me

End Sub



Private Sub Form_Load()
    Dim TrainLen As Single, ThisEngSize As Single, LastEngSize As Single
    Dim ErrCall As Boolean
    ' the form containing the sfCanvas control must be visible before running
    ' the startup method
   
    MousePointer = 11
    
    Set ReSize = New clsRsize
    ReSize.HandleForm = Me
    ReSize.Attach Me.sfCanvas1, True, True, True, True
    ReSize.Attach Me.cmdCamLeft, True, False, False, True
    ReSize.Attach Me.cmdCamRight, True, False, False, True
    ReSize.Attach Me.cmdCamUp, True, False, False, True
    ReSize.Attach Me.cmdCamDown, True, False, False, True
    ReSize.Attach Me.cmdCamForward, True, False, False, True
    ReSize.Attach Me.cmdCamBack, True, False, False, True
    ReSize.Attach Me.cmdCamReset, True, False, False, True

    ReSize.Attach Me.cmdCamNext, True, False, False, True
    ReSize.Attach Me.cmdCamPrev, True, False, False, True
ReSize.Attach Me.Command1, False, True, False, True
    
    ReSize.Attach hDist, False, True, False, True
    ReSize.Attach Htime, False, True, False, True
    ReSize.Attach Label1, False, True, False, True
    ReSize.Attach Label2, False, True, False, True
    ReSize.Attach Label3, False, True, False, True
    ReSize.Ready = True

    Set Trains = New Collection
    
   ' Show
    
    ' finalises the directX setup
    sfCanvas1.Startup
    
    ' this turns on the lights (ambient and directional) and sets up the
    ' camera viewing angle and a few rendering options
    sfCanvas1.initDefaults
    
    ' this moves the camera to x=10, y=2, z=0 to look at a point 2m above the origin
    CamX = 15#
    CamY = 2#
    CamZ = 0#
    LookAtX = 0#
    LookAtY = 2#
    LookAtZ = 0#
    
    ' Not looking at anything
    LookAt = 0
    
    sfCanvas1.setCamera CamX, CamY, CamZ, LookAtX, LookAtY, LookAtZ, 0, 1, 0
     
    sfCanvas1.showFPS = True
    initGrid
    
    DoEvents
   
    For i = 0 To conNumber - 1
    If Not FileExists(conItem(i)) Then
    Call MsgBox(conItem(i) & Lang(353), vbExclamation, App.Title)
    Exit Sub
    End If
   ThisCon = i
    Call LoadShape(conItem(i), ErrCall)
    If ErrCall = True Then
    Exit Sub
    End If
    If conFlip(i) = True Then
        Train.RotateBy -3.1415926    ' rotate clockwise by 180°
    RefreshForm
    ThisEngSize = EngSize / 2
    TrainLen = TrainLen + ThisEngSize + LastEngSize
    
    Train.MoveBy , , TrainLen + 2
    
    LastEngSize = ThisEngSize
    Else
    ThisEngSize = EngSize / 2
    TrainLen = TrainLen + ThisEngSize + LastEngSize
    
    Train.MoveBy , , -(TrainLen + 2)
    
    LastEngSize = ThisEngSize
    End If
    gridState = 0
    RefreshForm
'    Train.RotateBy -3.1415926    ' rotate clockwise by 180°
'    RefreshForm
    Next i
   ' sfCanvas1.d3dd.SetRenderState D3DRS_AMBIENT, D3DColorXRGB(255, 255, 255)
    sfCanvas1.d3dd.SetRenderState D3DRS_AMBIENT, D3DColorXRGB(255, 255, 255)
        Me.height = 10000
    Me.width = 12900
    DoEvents
    CamZ = -6
    LookAtZ = -6
    
    CamX = 15
    LookAtX = 0
    
    
    ResetCamera
    RefreshForm
    MousePointer = 0

End Sub

Sub RefreshForm()

    Dim objTrain As MovableObject
    'sfCanvas1.Width = ScaleWidth - sfCanvas1.Left * 2
    'sfCanvas1.Height = ScaleHeight - sfCanvas1.Top * 2
    
    DoEvents
    
    ' this method adjusts the viewport camera parameters to accomodate
    ' changes in aspect ratio
    
    sfCanvas1.Reset
    
    ' resetState should be called for each model that exists
    For Each objTrain In Trains
 objTrain.Model.ResetState
    Next
    
    ' show the new image (but only if d3dd has been initialised)
    If Not sfCanvas1.d3dd Is Nothing Then
        showModels
    End If
    DoEvents
End Sub


Private Sub Form_Resize()
'    ReSize.ReSize
'    RefreshForm
End Sub

Private Sub showModels()

    Dim objTrain As MovableObject

    ' signal start of 3D drawing mode
    If sfCanvas1.BeginScene Then
    
'    sfCanvas1.sky = True
'sfCanvas1.skyTexture = App.path & "\sky.jpg"
sfCanvas1.time = Htime.Value
Set gridTex = sfCanvas1.loadTex(App.Path & "\grass.bmp")
        ' show the model
        For Each objTrain In Trains
            objTrain.Model.showModel
        Next
        ' signal end and display the model in the scene
        sfCanvas1.EndScene
               
    End If

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
    
'    If mGrass.Checked Then
        Set gridTex = sfCanvas1.loadTex(App.Path & "\grass.bmp")
'    Else
'        Set gridTex = sfCanvas1.loadTex(App.path & "\grid.bmp")
'    End If
    
    ' load display base
'    Set tcbase = New sfModel
'    Set tcbase.d3dd = sfCanvas1.d3dd
'    tcbase.filename = App.path & "\TC_Displaybase.s"
'    D3DXMatrixIdentity basemx
'    basemx.m42 = -0.54
'    tcbase.setMatrix basemx
    
End Sub

Private Sub setVert(v As D3DVERTEX, X As Single, Y As Single, Z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single)
    v.X = X
    v.Y = Y
    v.Z = Z
    v.nx = nx
    v.ny = ny
    v.nz = nz
    v.tu = tu
    v.tv = tv
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set ReSize = Nothing
End Sub

Function GetMSTSFolder() As String
    GetMSTSFolder = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Microsoft Games\Train Simulator\1.0", "Path")
End Function





Private Sub hDist_Change()
hDist_scroll
End Sub

Private Sub hDist_scroll()
'sfCanvas1.setCamera HDist.value * Cos(vAngl) * Cos(hAngl) / 2 + cx, HDist.value * Sin(vAngl) / 2 + cy, HDist.value * Sin(hAngl) * Cos(vAngl) / 2 + cz, cx, cy, cz, 0#, 1#, 0#

sfCanvas1.setCamera hDist.Value * Cos(vAngl) * Cos(hAngl) / 2 + CamX, hDist.Value * Sin(vAngl) / 2 + CamY, hDist.Value * Sin(hAngl) * Cos(vAngl) / 2 + CamZ, LookAtX, LookAtY, LookAtZ, 0#, 1#, 0#
Label2.Caption = Format(hDist.Value / 2, "#0.0") & "m"
    RefreshForm
End Sub





Private Sub Htime_Change()
sfCanvas1.time = Htime.Value
ResetCamera

End Sub

Private Sub sfCanvas1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then
'DrawNow = True
'FirstX = x
'End If
'If Button = 2 Then
'DrawNow = True
'FirstY = y
'End If
 bmove = True
    sX = X
    sY = Y
End Sub

Private Sub sfCanvas1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Button = 1 And DrawNow = True Then
'If x > FirstX Then
'CamZ = CamZ - 2
'LookAtZ = LookAtZ - 2
'    ResetCamera
'    ElseIf x < FirstX Then
'    CamZ = CamZ + 2
'    LookAtZ = LookAtZ + 2
'    ResetCamera
'    End If
'    End If
'    If Button = 2 And DrawNow = True Then
'     If y > FirstY Then
'
'    CamY = CamY - 1
'    ResetCamera
'    Else
'    CamY = CamY + 1
'
'    ResetCamera
'    End If
'End If
  Dim DY As Single
    
    If Not bmove Then Exit Sub
    DY = Y - sY
    If Button = vbLeftButton Then
    
        ' rotate
        vAngl = vAngl + DY / 3500
        hAngl = hAngl - (X - sX) / 3500
        hDist_scroll
    Else
        ' zoom
        DY = hDist.Value + DY / 100
        If DY > hDist.Max Then
            DY = hDist.Max
        ElseIf DY < hDist.Min Then
            DY = hDist.Min
        End If
        hDist.Value = DY
    End If
    sY = Y
    sX = X
End Sub


Private Sub sfCanvas1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Or Button = 2 Then
'DrawNow = False
'End If
bmove = False
End Sub




