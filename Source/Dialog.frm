VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Texture Files"
   ClientHeight    =   3330
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7755
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save as File"
      Height          =   375
      Left            =   6420
      TabIndex        =   4
      Top             =   2580
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2535
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   420
      Width           =   6255
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6420
      TabIndex        =   2
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6420
      TabIndex        =   1
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Uncheck a texture to remove it from the model."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   7275
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public modl As sfModel
Public mod2 As sfModel
Dim d3dx8 As New d3dx8

Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    Dim tx As Direct3DTexture8, s As String
    Dim p As PALETTEENTRY
    
    If List1.ListIndex < 0 Then Exit Sub
    If List1.ListIndex = modl.nTextures Then Exit Sub
    
    Set tx = modl.GetTexture(List1.ListIndex)
    s = List1.List(List1.ListIndex)
    s = Left$(s, InStr(s, vbTab) - 5) & ".bmp"
    d3dx8.SaveTextureToFile s, D3DXIFF_BMP, tx, p
End Sub

Private Sub Form_Load()
    Dim lh As Long
    Dim f As String, fmt As String, w As Long, H As Long, loaded As Boolean
    
    For lh = 1 To modl.nTextures
        modl.getTexInfo lh - 1, f, fmt, w, H, loaded
        List1.AddItem f & vbTab & Left$(w & " x " & H & "           ", 16) & vbTab & vbTab & fmt
        List1.Selected(lh - 1) = loaded
        If Not loaded Then List1.ItemData(lh - 1) = 1
    Next
    
    If Not mod2 Is Nothing Then
        List1.AddItem mod2.Filename & ">>>>"
        For lh = 1 To mod2.nTextures
            mod2.getTexInfo lh - 1, f, fmt, w, H, loaded
            List1.AddItem f & vbTab & Left$(w & " x " & H & "           ", 16) & vbTab & vbTab & fmt
            List1.Selected(List1.NewIndex) = loaded
            If Not loaded Then List1.ItemData(List1.NewIndex) = 1
        Next
    End If
End Sub

Private Sub OKButton_Click()
    Dim i As Integer, s As String
    Dim j As Integer
    j = modl.nTextures
    For i = 0 To List1.ListCount - 1
        If i < j Then
            If Not List1.Selected(i) Then
                modl.SetTexture i, Nothing
            ElseIf List1.ItemData(i) = 1 Then
                s = List1.List(i)
                s = Left$(s, InStr(s, vbTab) - 1)
                modl.SetTexture i, modl.loadAce(s)
            End If
        ElseIf i > j Then
            If Not List1.Selected(i) Then
                mod2.SetTexture i - j - 1, Nothing
            ElseIf List1.ItemData(i) = 1 Then
                s = List1.List(i)
                s = Left$(s, InStr(s, vbTab) - 1)
                mod2.SetTexture i - j - 1, mod2.loadAce(s)
            End If
        End If
    Next
    Me.Hide
End Sub

