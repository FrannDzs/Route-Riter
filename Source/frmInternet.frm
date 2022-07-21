VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmInternet 
   Caption         =   "PayPal Browser - NOTE: PAYPAL DO NOT PROCESS DONATIONS OF LESS THAN $1.00"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   14025
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8895
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   13455
      ExtentX         =   23733
      ExtentY         =   15690
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   9600
      Width           =   6375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FORWARD>>"
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   3
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HOME"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< BACK"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   12360
      TabIndex        =   0
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "URL "
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   9600
      Width           =   975
   End
End
Attribute VB_Name = "frmInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmInternet.Hide

'Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)
On Error GoTo Errtrap
Select Case Index
Case 0
WebBrowser1.GoBack
Case 1
WebBrowser1.GoHome
Case 2
WebBrowser1.GoForward
End Select
Exit Sub
Errtrap:
Exit Sub
End Sub


Private Sub Form_Load()
On Error GoTo Errtrap

WebBrowser1.Top = 120
WebBrowser1.width = Me.width - 200
WebBrowser1.height = Me.height - 400

If flagInternet = 1 Then
WebBrowser1.Navigate "https://www.paypal.com/xclick/business=agene%40optusnet.com.au&item_name=Route_Riter+Payment&no_shipping=0&no_note=1&tax=0&currency_code=USD&lc=AU"

ElseIf flagInternet = 2 Then
WebBrowser1.Navigate "http://www.rstools.info/route_riter.html#RRv7"
'
ElseIf flagInternet = 3 Then
WebBrowser1.Navigate "http://www.rstools.info/faq.html"
ElseIf flagInternet = 4 Then
WebBrowser1.Navigate "http://www.rstools.info/index.html"

End If
Exit Sub
Errtrap:


End Sub


Private Sub Form_Resize()
Dim i As Integer
On Error GoTo Errtrap
WebBrowser1.Left = 300
WebBrowser1.Top = 400
WebBrowser1.width = Me.width - 800
WebBrowser1.height = Me.height - 2000
Command1.Top = WebBrowser1.Top + WebBrowser1.height + 200
'Command1.Left = Me.width / 2 - Command1.width / 2
For i = 0 To 2
Command2(i).Top = Command1.Top
Next i
Text1.Top = Command1.Top
Label1.Top = Command1.Top
Exit Sub
Errtrap:
If Err = 380 Then Exit Sub
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
WebBrowser1.Navigate Text1.Text
DoEvents
End If
End Sub


