VERSION 5.00
Object = "{C7083F68-7BF5-4755-9CF0-38D810EC405C}#1.0#0"; "trainlib.ocx"
Begin VB.Form frmConsist 
   Caption         =   "Consist Viewer - "
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin trainlib.sfCanvas sfCanvas1 
      Height          =   1575
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2778
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
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   3840
      Max             =   -5
      Min             =   -25
      TabIndex        =   2
      Top             =   2040
      Value           =   -15
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   1
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2040
      Width           =   495
   End
   Begin trainlib.sfCanvas sfCanvas1 
      Height          =   1575
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2778
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
   Begin trainlib.sfCanvas sfCanvas1 
      Height          =   1575
      Index           =   2
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2778
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
   Begin trainlib.sfCanvas sfCanvas1 
      Height          =   1575
      Index           =   3
      Left            =   7080
      TabIndex        =   9
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2778
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
   Begin trainlib.sfCanvas sfCanvas1 
      Height          =   1575
      Index           =   4
      Left            =   9360
      TabIndex        =   10
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2778
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
   Begin VB.Label Label1 
      Caption         =   "Distance"
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "frmConsist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Text
Dim f(0 To 4) As sfModel
Dim FirstCon As Integer
Dim zz(0 To 4) As Boolean
Private Sub Command1_Click()
Dim i As Integer
FirstCon = FirstCon - 1
For i = 0 To 4
zz(i) = False
Next
If FirstCon < 0 Then
FirstCon = 0
Exit Sub
End If
For i = 0 To 4
    Select Case i
       Case 0
    f(i).filename = conItem(FirstCon)
    If conFlip(FirstCon) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 1
    f(i).filename = conItem(FirstCon + 1)
    If conFlip(FirstCon + 1) = True Then
    zz(i) = True = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 2
    f(i).filename = conItem(FirstCon + 2)
    If conFlip(FirstCon + 2) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
   Case 3
   f(i).filename = conItem(FirstCon + 3)
   If conFlip(FirstCon + 3) = True Then
   zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 4
    f(i).filename = conItem(FirstCon + 4)
    If conFlip(FirstCon + 4) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
   End Select
 
    DoEvents
    
   Call showModel(i)
    Next i
End Sub

Private Sub Command2_Click()
Dim i As Integer
FirstCon = FirstCon + 1
For i = 0 To 4
zz(i) = False
Next
If FirstCon > conNumber - 5 Or FirstCon < 0 Then
FirstCon = FirstCon - 1
Exit Sub
End If

For i = 0 To 4
    Select Case i
       Case 0
    f(i).filename = conItem(FirstCon)
    If conFlip(FirstCon) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 1
    f(i).filename = conItem(FirstCon + 1)
    If conFlip(FirstCon + 1) = True Then
    zz(i) = True = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 2
    f(i).filename = conItem(FirstCon + 2)
    If conFlip(FirstCon + 2) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
   Case 3
   f(i).filename = conItem(FirstCon + 3)
   If conFlip(FirstCon + 3) = True Then
   zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 4
    f(i).filename = conItem(FirstCon + 4)
    If conFlip(FirstCon + 4) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
   End Select
 
    DoEvents
    
   Call showModel(i)
    Next i
End Sub

Private Sub Command3_Click()
Dim i As Integer
FirstCon = 0
For i = 0 To 4
zz(i) = False
Next

For i = 0 To 4
    Select Case i
        Case 0
    f(i).filename = conItem(FirstCon)
    If conFlip(FirstCon) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 1
    f(i).filename = conItem(FirstCon + 1)
    If conFlip(FirstCon + 1) = True Then
    zz(i) = True = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 2
    f(i).filename = conItem(FirstCon + 2)
    If conFlip(FirstCon + 2) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
   Case 3
   f(i).filename = conItem(FirstCon + 3)
   If conFlip(FirstCon + 3) = True Then
   zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 4
    f(i).filename = conItem(FirstCon + 4)
    If conFlip(FirstCon + 4) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
   End Select
    DoEvents
    
   Call showModel(i)
    Next i
End Sub

Private Sub Command4_Click()
Dim i As Integer
FirstCon = conNumber - 5
For i = 0 To 4
zz(i) = False
Next


For i = 0 To 4
    Select Case i
        Case 0
    f(i).filename = conItem(FirstCon)
    If conFlip(FirstCon) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 1
    f(i).filename = conItem(FirstCon + 1)
    If conFlip(FirstCon + 1) = True Then
    zz(i) = True = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 2
    f(i).filename = conItem(FirstCon + 2)
    If conFlip(FirstCon + 2) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
   Case 3
   f(i).filename = conItem(FirstCon + 3)
   If conFlip(FirstCon + 3) = True Then
   zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 4
    f(i).filename = conItem(FirstCon + 4)
    If conFlip(FirstCon + 4) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
   End Select
    DoEvents
    
   Call showModel(i)
    Next i
End Sub

Private Sub Form_Load()
Dim i As Integer, mw As D3DMATRIX, mt As D3DMATRIX
Dim x As Single, y As Single, z As Single, radius As Single

On Error GoTo Errtrap
    ' the form containing the sfCanvas control must be visible before running
    ' the startup method
FirstCon = 0
    Show
    For i = 0 To 4
    ' finalises the directX setup
    sfCanvas1(i).Startup
    
    ' this turns on the lights (ambient and directional) and sets up the
    ' camera viewing angle and a few rendering options
    sfCanvas1(i).initDefaults
    
    ' this moves the camera to x=10, y=2, z=0 to look at a point 2m above the origin
    sfCanvas1(i).setCamera -14, 2, 0, 0, 2, 0, 0, 1, 0
    
    ' create a new model object
    Set f(i) = New sfModel
    ' the d3dd object must come from the sfCanvas control
    Set f(i).d3dd = sfCanvas1(i).d3dd
    
    ' by default, the program searches the current folder first.
    ' if not there, it uses the textureSearchPath (delimited by ;) to look for the texture files
    
    f(i).textureSearchPath = "..\textures"
    
'    If Not FileExists("D:\program files\microsoft games\train simulator\trains\trainset\scotsman\scotsman.s") Then
'    Debug.Print "File missing"
'    End If
    ' setting this property triggers loading of the S file and textures
  
    Select Case i
    Case 0
    f(i).filename = conItem(FirstCon)
    If conFlip(FirstCon) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 1
    f(i).filename = conItem(FirstCon + 1)
    If conFlip(FirstCon + 1) = True Then
    zz(i) = True = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 2
    f(i).filename = conItem(FirstCon + 2)
    If conFlip(FirstCon + 2) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
   Case 3
   f(i).filename = conItem(FirstCon + 3)
   If conFlip(FirstCon + 3) = True Then
   zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
    Case 4
    f(i).filename = conItem(FirstCon + 4)
    If conFlip(FirstCon + 4) = True Then
    zz(i) = True
    sfCanvas1(i).setCamera 14, 2, 0, 0, 2, 0, 0, 1, 0
    End If
   End Select
 
    DoEvents
'    D3DXMatrixIdentity mw
'    D3DXMatrixRotationY mt, 3.1415926
'    D3DXMatrixMultiply mw, mt, mw
'    f(i).setMatrix mw
'    f(i).getSphere x, y, z, radius
'    sfCanvas1(i).setCamera radius * 2, 2, 0, x, y, z, 0, 1, 0
   Call showModel(i)
    Next i
    Exit Sub
Errtrap:
    Stop
End Sub

Private Sub showModel(i As Integer)
    ' signal start of 3D drawing mode
    Dim mc As D3DCOLORVALUE, v As D3DVECTOR

  '  sfCanvas1(i).d3dd.SetRenderState D3DRS_AMBIENT, D3DColorXRGB(255, 255, 255)
    
    If sfCanvas1(i).BeginScene Then
    
     sfCanvas1(i).d3dd.SetRenderState D3DRS_AMBIENT, D3DColorXRGB(255, 255, 255)
    
        ' show the model
        f(i).showModel
        ' signal end and display the model in the scene
        sfCanvas1(i).EndScene
    End If

End Sub

Private Sub HScroll1_Change()
Dim i As Integer
For i = 0 To 4
If zz(i) = True Then
sfCanvas1(i).setCamera -(HScroll1.value), 2, 0, 0, 2, 0, 0, 1, 0
Else
sfCanvas1(i).setCamera HScroll1.value, 2, 0, 0, 2, 0, 0, 1, 0
End If
showModel (i)
Next
End Sub


