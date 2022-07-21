VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SaveTex"
   ClientHeight    =   1965
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   8850
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1356.278
   ScaleMode       =   0  'User
   ScaleWidth      =   8310.606
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3780
      TabIndex        =   0
      Top             =   1440
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   56.343
      X2              =   8169.749
      Y1              =   869.674
      Y2              =   869.674
   End
   Begin VB.Label lblTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   8505
   End
   Begin VB.Label lblVersion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Width           =   5265
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim comp As CompressZIt
Dim fso As New FileSystemObject
Dim saved As Long

' read MSTS file into byte array
' works on if GZip compressed too
Private Function readFile(fName As String) As Long
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
        readFile = 0
    Else
        ReDim bdata(LOF(i) - 17)
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
        readFile = comp.m_OriginalSize - comp.m_CompressedSize
    End If
    Exit Function
ErrEx:
    readFile = Err.Number
End Function

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
'    Dim s As String
    Dim f As String
    Dim fl As Folder
    
    
    cmdOK.Enabled = False
    Set comp = New CompressZIt

f = frmUtils.Dir1(0).path & "\"
        If frmUtils.Dir1(0).path <> vbNullString Then
            Show
            Refresh
           Set fl = fso.GetFolder(f)
            scandir fl
           
        End If
        f = vbNullString

        lblVersion.Caption = lblVersion.Caption & " Finished.."
  
    cmdOK.Enabled = True
End Sub

Private Sub scandir(fl As Folder)
    Dim sf As String
    Dim fil As File
    
    lblTitle.Caption = "Scanning " & fl.path
    
    For Each fil In fl.Files
        sf = fil.Name
        If UCase(Right$(sf, 4)) = ".ACE" And Not (fil.Attributes And ReadOnly) Then
            saved = saved + readFile(fil.path)
            lblVersion.Caption = "Saved " & saved \ 1024 & "Kb"
            lblVersion.Refresh
        End If
    Next



End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set comp = Nothing
End Sub

