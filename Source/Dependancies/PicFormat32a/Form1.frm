VERSION 5.00
Object = "{E5CEE37F-8CF8-489E-BFA0-8201CBD6AEE8}#1.0#0"; "PICFORMAT32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PicFormat32 OCX/DLL"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "GifToBmp"
      TabPicture(0)   =   "Form1.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "BmpToGif"
      TabPicture(1)   =   "Form1.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Text4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label7(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "JpegToBmp"
      TabPicture(2)   =   "Form1.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Text5"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label7(3)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label7(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "BmpToJpeg"
      TabPicture(3)   =   "Form1.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label7(4)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label7(5)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label3"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Command3"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Text7"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Text8"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Text9"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "HScroll1"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Command9"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      Begin VB.CommandButton Command9 
         Caption         =   "..."
         Height          =   255
         Left            =   -70800
         TabIndex        =   29
         Top             =   600
         Width           =   495
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   315
         Left            =   -72840
         Max             =   130
         TabIndex        =   28
         Top             =   1200
         Value           =   65
         Width           =   2475
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -73320
         TabIndex        =   27
         Text            =   "65"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -73560
         TabIndex        =   26
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -73560
         TabIndex        =   25
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton Command8 
         Caption         =   "..."
         Height          =   255
         Left            =   -70800
         TabIndex        =   21
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -73560
         TabIndex        =   20
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -73560
         TabIndex        =   19
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton Command7 
         Caption         =   "..."
         Height          =   255
         Left            =   -70800
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -73560
         TabIndex        =   15
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -73560
         TabIndex        =   14
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "BmpToJpeg"
         Height          =   495
         Left            =   -74880
         TabIndex        =   11
         Top             =   1680
         Width           =   4575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save JpegToBmp"
         Height          =   495
         Left            =   -74880
         TabIndex        =   10
         Top             =   1440
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save BmpToGif"
         Height          =   495
         Left            =   -74880
         TabIndex        =   9
         Top             =   1440
         Width           =   4575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save GifToBmp"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Compression Ratio :"
         Height          =   195
         Left            =   -74880
         TabIndex        =   24
         Top             =   1200
         Width           =   1410
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Jpeg Filename :"
         Height          =   195
         Index           =   5
         Left            =   -74880
         TabIndex        =   23
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Bitmap Filename :"
         Height          =   195
         Index           =   4
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Jpeg Filename :"
         Height          =   195
         Index           =   3
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Bitmap Filename :"
         Height          =   195
         Index           =   2
         Left            =   -74880
         TabIndex        =   17
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gif Filename :"
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   13
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Bitmap Filename :"
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gif Filename :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Bitmap Filename :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1245
      End
   End
   Begin PicFormat32a.PicFormat32 PicFormat321 
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Befor you can use PicFormat32.dll make sure its
'in C:\WINDOWS\SYSTEM folder.


Private Sub Command1_Click()
PicFormat321.SaveBmpToGif Text3, Text4
End Sub

Private Sub Command2_Click()
PicFormat321.SaveGifToBmp Text1.Text, Text2.Text
End Sub

Private Sub Command3_Click()
PicFormat321.SaveBmpToJpeg Text7, Text8, Text9
End Sub

Private Sub Command4_Click()
PicFormat321.SaveJpegToBmp Text5, Text6
End Sub

Private Sub Command5_Click()
With CommonDialog1
.FileName = "*.gif"
.Filter = "gif"
.DialogTitle = "Open Gif to save as Bmp"
.ShowOpen
Text1.Text = .FileName
Text2.Text = Left(.FileName, Len(.FileName) - 3) + ".bmp"
End With
End Sub

Private Sub Command6_Click()
PicFormat321.About
End Sub

Private Sub Command7_Click()
With CommonDialog1
.FileName = "*.bmp"
.Filter = "bmp"
.DialogTitle = "Open Bmp to save as Gif"
.ShowOpen
Text3.Text = .FileName
Text4.Text = Left(.FileName, Len(.FileName) - 3) + ".gif"
End With
End Sub

Private Sub Command8_Click()
With CommonDialog1
.FileName = "*.jpg"
.Filter = "jpg"
.DialogTitle = "Open Jpg to save as Bmp"
.ShowOpen
Text5.Text = .FileName
Text6.Text = Left(.FileName, Len(.FileName) - 3) + ".bmp"
End With
End Sub

Private Sub Command9_Click()
With CommonDialog1
.FileName = "*.bmp"
.Filter = "bmp"
.DialogTitle = "Open Bmp to save as Jpg"
.ShowOpen
Text7.Text = .FileName
Text8.Text = Left(.FileName, Len(.FileName) - 3) + ".jpg"
End With
End Sub

Private Sub HScroll1_Change()
Text9.Text = HScroll1.Value
If HScroll1.Value < 20 Then
HScroll1.Value = 20
Text9.Text = 20
End If
End Sub
