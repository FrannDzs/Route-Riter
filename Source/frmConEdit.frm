VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmConEdit 
   Caption         =   "List of Filtered Files."
   ClientHeight    =   13365
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   21195
   LinkTopic       =   "Form1"
   ScaleHeight     =   13365
   ScaleWidth      =   21195
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   13365
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   21195
      _cx             =   37386
      _cy             =   23574
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483637
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CheckBox Check8 
         BackColor       =   &H8000000B&
         Caption         =   "Length"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   13560
         TabIndex        =   112
         Top             =   9000
         Width           =   1575
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H8000000B&
         Caption         =   "Mass"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   11880
         TabIndex        =   111
         Top             =   9000
         Width           =   1575
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H8000000B&
         Caption         =   "Type"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10200
         TabIndex        =   110
         Top             =   9000
         Width           =   1575
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H8000000B&
         Caption         =   "Coupling"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8520
         TabIndex        =   109
         Top             =   9000
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H8000000B&
         Caption         =   "Length"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   108
         Top             =   9000
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H8000000B&
         Caption         =   "Max Power"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   107
         Top             =   9000
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H8000000B&
         Caption         =   "Mass"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4560
         TabIndex        =   106
         Top             =   9000
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000B&
         Caption         =   "Coupling"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   105
         Top             =   9000
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Save Activity"
         Enabled         =   0   'False
         Height          =   390
         Left            =   16005
         TabIndex        =   102
         Top             =   12270
         Width           =   1350
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Save Loose Consist"
         Enabled         =   0   'False
         Height          =   390
         Left            =   17475
         TabIndex        =   101
         Top             =   12270
         Width           =   1740
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   420
         Left            =   15600
         TabIndex        =   84
         Top             =   12765
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Display Loose Consist #"
         Height          =   510
         Left            =   13995
         TabIndex        =   83
         Top             =   12720
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   13560
         TabIndex        =   82
         Top             =   8685
         Width           =   2205
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   6015
         TabIndex        =   80
         Top             =   8685
         Width           =   2070
      End
      Begin VB.PictureBox Picture1 
         Height          =   2595
         Left            =   16005
         ScaleHeight     =   2535
         ScaleWidth      =   4815
         TabIndex        =   69
         Top             =   240
         Width           =   4875
         Begin VB.Image Image2 
            Height          =   2550
            Left            =   0
            Top             =   0
            Width           =   4575
         End
      End
      Begin VB.PictureBox pixStart 
         Height          =   2595
         Left            =   16005
         ScaleHeight     =   2535
         ScaleWidth      =   4815
         TabIndex        =   68
         Top             =   240
         Visible         =   0   'False
         Width           =   4875
      End
      Begin VB.PictureBox Picture2 
         Height          =   1290
         Index           =   2
         Left            =   5775
         ScaleHeight     =   1230
         ScaleWidth      =   2385
         TabIndex        =   51
         Top             =   9795
         Width           =   2445
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   2
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   1290
         Index           =   3
         Left            =   8280
         ScaleHeight     =   1230
         ScaleWidth      =   2385
         TabIndex        =   50
         Top             =   9795
         Width           =   2445
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   3
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   1290
         Index           =   4
         Left            =   10770
         ScaleHeight     =   1230
         ScaleWidth      =   2400
         TabIndex        =   49
         Top             =   9795
         Width           =   2460
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   4
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   1290
         Index           =   5
         Left            =   13290
         ScaleHeight     =   1230
         ScaleWidth      =   2400
         TabIndex        =   48
         Top             =   9795
         Width           =   2460
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   5
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   1290
         Index           =   6
         Left            =   15795
         ScaleHeight     =   1230
         ScaleWidth      =   2370
         TabIndex        =   47
         Top             =   9795
         Width           =   2430
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   6
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   1290
         Index           =   7
         Left            =   18285
         ScaleHeight     =   1230
         ScaleWidth      =   2400
         TabIndex        =   46
         Top             =   9795
         Width           =   2460
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   7
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   1290
         Index           =   1
         Left            =   3285
         ScaleHeight     =   1230
         ScaleWidth      =   2385
         TabIndex        =   45
         Top             =   9795
         Width           =   2445
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   1
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   1290
         Index           =   0
         Left            =   765
         ScaleHeight     =   1230
         ScaleWidth      =   2400
         TabIndex        =   44
         Top             =   9795
         Width           =   2460
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   7080
         TabIndex        =   33
         Top             =   12270
         Width           =   780
      End
      Begin VB.CommandButton Command3 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   7710
         TabIndex        =   32
         Top             =   12270
         Width           =   795
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   8490
         TabIndex        =   31
         Top             =   12270
         Width           =   780
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   9270
         TabIndex        =   30
         Top             =   12270
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Height          =   390
         Index           =   1
         Left            =   11460
         TabIndex        =   29
         Top             =   12270
         Width           =   4080
      End
      Begin VB.TextBox Text1 
         Height          =   390
         Index           =   0
         Left            =   1800
         TabIndex        =   28
         Top             =   12270
         Width           =   4740
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save Consist"
         Height          =   390
         Left            =   19320
         TabIndex        =   27
         Top             =   12270
         Width           =   1425
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Exit"
         Height          =   390
         Left            =   19320
         TabIndex        =   22
         Top             =   12765
         Width           =   1425
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Show in 3D"
         Height          =   390
         Left            =   17910
         TabIndex        =   21
         Top             =   2805
         Width           =   1305
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear All"
         Height          =   390
         Left            =   17850
         TabIndex        =   20
         Top             =   12765
         Width           =   1425
      End
      Begin RichTextLib.RichTextBox RTB1 
         Height          =   4545
         Left            =   16005
         TabIndex        =   8
         Top             =   3180
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   8017
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmConEdit.frx":0000
      End
      Begin VSFlex8LCtl.VSFlexGrid List2 
         Height          =   8340
         Left            =   8280
         TabIndex        =   23
         Top             =   240
         Width           =   7575
         _cx             =   13361
         _cy             =   14711
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConEdit.frx":0082
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox Picture_src 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   1095
            Left            =   19920
            ScaleHeight     =   1035
            ScaleWidth      =   2115
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   5320
            Width           =   2175
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid List1 
         Height          =   8340
         Left            =   600
         TabIndex        =   25
         Top             =   240
         Width           =   7560
         _cx             =   13335
         _cy             =   14711
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConEdit.frx":014D
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox PicStart 
            AutoSize        =   -1  'True
            Height          =   1095
            Index           =   0
            Left            =   1320
            ScaleHeight     =   1035
            ScaleWidth      =   1995
            TabIndex        =   26
            Top             =   6240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "Selected Loco or Wagon File Path"
         Height          =   390
         Left            =   240
         TabIndex        =   104
         Top             =   12840
         Width           =   1500
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   1800
         TabIndex        =   103
         Top             =   12720
         Width           =   12150
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   15
         Left            =   20400
         TabIndex        =   100
         Top             =   9420
         Width           =   270
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   14
         Left            =   17850
         TabIndex        =   99
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   13
         Left            =   15465
         TabIndex        =   98
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   12930
         TabIndex        =   97
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   11
         Left            =   10410
         TabIndex        =   96
         Top             =   9420
         Width           =   270
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   10
         Left            =   7860
         TabIndex        =   95
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   9
         Left            =   5460
         TabIndex        =   94
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   2940
         TabIndex        =   93
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   18255
         TabIndex        =   92
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   15870
         TabIndex        =   91
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   13335
         TabIndex        =   90
         Top             =   9420
         Width           =   270
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   10785
         TabIndex        =   89
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   8400
         TabIndex        =   88
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   5865
         TabIndex        =   87
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   3345
         TabIndex        =   86
         Top             =   9420
         Width           =   270
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   795
         TabIndex        =   85
         Top             =   9420
         Width           =   285
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000B&
         Caption         =   "Find"
         Height          =   270
         Index           =   1
         Left            =   13185
         TabIndex        =   81
         Top             =   8685
         Width           =   525
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000B&
         Caption         =   "Find"
         Height          =   270
         Index           =   0
         Left            =   5625
         TabIndex        =   79
         Top             =   8685
         Width           =   405
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   630
         Index           =   15
         Left            =   18420
         TabIndex        =   78
         Top             =   11535
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   630
         Index           =   14
         Left            =   16005
         TabIndex        =   77
         Top             =   11535
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   630
         Index           =   13
         Left            =   13425
         TabIndex        =   76
         Top             =   11535
         Width           =   2340
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   630
         Index           =   12
         Left            =   10875
         TabIndex        =   75
         Top             =   11535
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   630
         Index           =   11
         Left            =   8310
         TabIndex        =   74
         Top             =   11535
         Width           =   2310
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   630
         Index           =   10
         Left            =   5880
         TabIndex        =   73
         Top             =   11535
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   630
         Index           =   9
         Left            =   3345
         TabIndex        =   72
         Top             =   11535
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   630
         Index           =   8
         Left            =   765
         TabIndex        =   71
         Top             =   11535
         Width           =   2325
      End
      Begin VB.Label Label6 
         Caption         =   "Consist Length:"
         Height          =   375
         Index           =   3
         Left            =   16005
         TabIndex        =   70
         Top             =   8940
         Width           =   4875
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   0
         Left            =   1470
         TabIndex        =   67
         Top             =   9420
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   1
         Left            =   3855
         TabIndex        =   66
         Top             =   9420
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   2
         Left            =   6390
         TabIndex        =   65
         Top             =   9420
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   3
         Left            =   8940
         TabIndex        =   64
         Top             =   9420
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   4
         Left            =   11460
         TabIndex        =   63
         Top             =   9420
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   5
         Left            =   13860
         TabIndex        =   62
         Top             =   9420
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   6
         Left            =   16395
         TabIndex        =   61
         Top             =   9420
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   7
         Left            =   18810
         TabIndex        =   60
         Top             =   9420
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   15
         Left            =   19590
         TabIndex        =   59
         Top             =   9420
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   14
         Left            =   17145
         TabIndex        =   58
         Top             =   9420
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   13
         Left            =   14580
         TabIndex        =   57
         Top             =   9420
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   12120
         TabIndex        =   56
         Top             =   9420
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   11
         Left            =   9735
         TabIndex        =   55
         Top             =   9420
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   10
         Left            =   7065
         TabIndex        =   54
         Top             =   9420
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   9
         Left            =   4590
         TabIndex        =   53
         Top             =   9420
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Napa Heavy SF"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   2130
         TabIndex        =   52
         Top             =   9420
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   765
         TabIndex        =   43
         Top             =   11175
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   3345
         TabIndex        =   42
         Top             =   11175
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   5880
         TabIndex        =   41
         Top             =   11175
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   8310
         TabIndex        =   40
         Top             =   11175
         Width           =   2310
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   4
         Left            =   10875
         TabIndex        =   39
         Top             =   11175
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   5
         Left            =   13425
         TabIndex        =   38
         Top             =   11175
         Width           =   2340
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   6
         Left            =   16005
         TabIndex        =   37
         Top             =   11175
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   7
         Left            =   18420
         TabIndex        =   36
         Top             =   11175
         Width           =   2325
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "Consist File Name"
         Height          =   390
         Index           =   0
         Left            =   765
         TabIndex        =   35
         Top             =   12270
         Width           =   1020
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000B&
         Caption         =   "Description:"
         Height          =   390
         Index           =   1
         Left            =   10410
         TabIndex        =   34
         Top             =   12270
         Width           =   885
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   0
         Left            =   2190
         TabIndex        =   19
         Top             =   8685
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "Total Locos"
         Height          =   270
         Index           =   0
         Left            =   765
         TabIndex        =   18
         Top             =   8685
         Width           =   1305
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "Selected Locos"
         Height          =   510
         Index           =   0
         Left            =   3210
         TabIndex        =   17
         Top             =   8685
         Width           =   1290
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   1
         Left            =   4590
         TabIndex        =   16
         Top             =   8685
         Width           =   915
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   2
         Left            =   12165
         TabIndex        =   15
         Top             =   8685
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "Selected Wagons"
         Height          =   510
         Index           =   1
         Left            =   10605
         TabIndex        =   14
         Top             =   8685
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "Total Wagons"
         Height          =   510
         Index           =   1
         Left            =   8310
         TabIndex        =   13
         Top             =   8685
         Width           =   1305
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   3
         Left            =   9600
         TabIndex        =   12
         Top             =   8685
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "Consist Details - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   16005
         TabIndex        =   11
         Top             =   7830
         Width           =   4875
      End
      Begin VB.Label Label6 
         Caption         =   "Engine Mass:"
         Height          =   375
         Index           =   1
         Left            =   16005
         TabIndex        =   10
         Top             =   8205
         Width           =   4875
      End
      Begin VB.Label Label6 
         Caption         =   "Wagon Mass:"
         Height          =   390
         Index           =   2
         Left            =   16005
         TabIndex        =   9
         Top             =   8565
         Width           =   4875
      End
   End
   Begin VB.PictureBox PicStart 
      AutoSize        =   -1  'True
      Height          =   1095
      Index           =   7
      Left            =   360
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox PicStart 
      AutoSize        =   -1  'True
      Height          =   1095
      Index           =   6
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox PicStart 
      AutoSize        =   -1  'True
      Height          =   1095
      Index           =   5
      Left            =   1440
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox PicStart 
      AutoSize        =   -1  'True
      Height          =   1095
      Index           =   4
      Left            =   1920
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox PicStart 
      AutoSize        =   -1  'True
      Height          =   1095
      Index           =   3
      Left            =   2520
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox PicStart 
      AutoSize        =   -1  'True
      Height          =   1095
      Index           =   2
      Left            =   3000
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox PicStart 
      AutoSize        =   -1  'True
      Height          =   1095
      Index           =   1
      Left            =   2400
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   21960
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "Files"
      Begin VB.Menu mnuOpenAct 
         Caption         =   "Open .ACT file"
      End
      Begin VB.Menu mnuOCF 
         Caption         =   "Open .CON file"
      End
   End
   Begin VB.Menu mnuPop1 
      Caption         =   "PopupMenu1"
      Begin VB.Menu mnuDel 
         Caption         =   "Delete Item"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert Here"
      End
      Begin VB.Menu mnuFlip 
         Caption         =   "Flip Item"
      End
      Begin VB.Menu mnuSwap 
         Caption         =   "Swap with Following Item"
      End
      Begin VB.Menu mnuMulti 
         Caption         =   "Add Multiple Items"
      End
   End
   Begin VB.Menu mnupop2 
      Caption         =   "PopupMenu2"
      Visible         =   0   'False
      Begin VB.Menu mnuActDel 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuActIns 
         Caption         =   "Insert"
      End
   End
End
Attribute VB_Name = "frmConEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Text


Const Wag_CHUNK = 500
Dim strActFile As String, strCFile As String
Dim booAbort As Boolean, booIgnore As Boolean
Dim ActNum As Integer, booIsAct As Boolean, booActIns As Boolean
Dim intConSerial As Integer, origConLen As Single
Dim strActStart As String, strActEnd As String
Private WithEvents SP As cScanPath
Attribute SP.VB_VarHelpID = -1
Public strFilter As String
Dim booConEng As Boolean
Dim booConWag As Boolean
Dim strConEng() As String
Dim strConWag() As String
Dim strConEngPath() As String
Dim strConWagPath() As String
Dim strConEngCou() As String
Dim strConWagCou() As String
Dim numEng As Integer
Dim numWag As Integer
Dim strLoose() As String
Dim numStock As Integer
Dim strConPic(1 To 500) As String
Dim strConListName(1 To 500) As String
Dim PicNum As Integer
Dim numIndex As Integer
Dim strConItem(1 To 500) As String
Dim strConFolder(1 To 500) As String
Dim intConType(1 To 500) As Integer
Dim strConUid(1 To 500) As String
Dim strMainPix As String
Dim intIndex As Integer
Dim booConMulti As Boolean
Dim engLoad As Single, wagLoad As Single
Dim conLen As Single
Dim strConCou(1 To 500) As String
Dim numItem As Integer
Private Sub ClearAll()
Dim i As Integer

booIgnore = False
For i = 0 To 7
Set Picture2(i) = LoadPicture("")
Set Image1(i) = LoadPicture("")
Set PicStart(i) = LoadPicture("")
Label3(i).Caption = ""
Label5(i).Caption = ""
Label8(i).Caption = ""
Next i
For i = 8 To 15
Label3(i).Caption = ""
Label5(i).Caption = ""
Label8(i).Caption = ""
Next i
DoEvents
Text1(0) = ""
Text1(1) = ""
numStock = 0
intConSerial = 0
Label6(1).Caption = "Engine Mass: "
Label6(2).Caption = "Wagon Mass: "
Label6(3).Caption = "Consist Length: "
engLoad = 0
conLen = 0
wagLoad = 0
For i = 1 To 500
strConPic(i) = ""
strConListName(i) = ""
strConItem(i) = ""
strConFolder(i) = ""
strConUid(i) = ""
intConType(i) = 0
Next i
End Sub

Private Sub GetEngDet(strEng As String, strFind As String, strTemp As String)
Dim x As Long, xx As Long

x = InStr(strEng, strFind)
If x = 0 Then
strTemp = ""
Exit Sub
End If
x = InStr(x, strEng, "(")
xx = InStr(x, strEng, ")" & vbCrLf)
If xx = 0 Then
strTemp = ""
Else
strTemp = Mid(strEng, x + 1, xx - (x + 1))
strTemp = Trim(strTemp)
End If
End Sub

Private Sub GetWagPaths(strPix As String, strEP As String, strEN As String, sMass As Single)
Dim strTemp As String, i As Integer, ii As Integer, x As Long



strTemp = Left(strPix, Len(strPix) - 4)
If strTemp = App.Path & "\Thomas" Then Exit Sub
    i = InStrRev(strTemp, "\", Len(strTemp) - 1)
    strEN = Mid(strTemp, i + 1)
    ii = InStrRev(strTemp, "\", i - 1)
    strEP = Mid(strTemp, ii + 1, i - ii - 1)
    
    
 MyString = ReadUniFile(strTemp & ".wag")
x = InStr(MyString, "Mass ")
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp2 = Mid(MyString, x + 1, xx - (x + 1))
strTemp2 = Trim(strTemp2)
If Left(strTemp2, 1) = ChrW$(34) Then
strTemp2 = Mid(strTemp2, 2)
End If
If Right(strTemp2, 1) = ChrW$(34) Then
strTemp2 = Left(strTemp2, Len(strTemp2) - 1)
End If
sMass = Val(strTemp2)

End Sub

Private Sub GetEngPaths(strPix As String, strEP As String, strEN As String, sMass As Single, sMax As Single, sMaxVel As Single, intFlag As Integer)
Dim strTemp As String, MyString As String, x As Long, xx As Long, strTemp2 As String, strUnit As String
Dim i As Integer, ii As Integer
On Error GoTo Errtrap

strTemp = Left(strPix, Len(strPix) - 4)
    i = InStrRev(strTemp, "\", Len(strTemp) - 1)
    
    strEN = Mid(strTemp, i + 1)
    ii = InStrRev(strTemp, "\", i - 1)
    
    strEP = Mid(strTemp, ii + 1, i - ii - 1)
   
MyString = ReadUniFile(strTemp & ".eng")
x = InStr(MyString, "Mass ")
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp2 = Mid(MyString, x + 1, xx - (x + 1))
strTemp2 = Trim(strTemp2)
If Left(strTemp2, 1) = ChrW$(34) Then
strTemp2 = Mid(strTemp2, 2)
End If
If Right(strTemp2, 1) = ChrW$(34) Then
strTemp2 = Left(strTemp2, Len(strTemp2) - 1)
End If

sMass = Val(strTemp2)
x = InStr(MyString, "MaxForce ")
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp2 = Mid(MyString, x + 1, xx - (x + 1))
strTemp2 = Trim(strTemp2)
If Left(strTemp2, 1) = ChrW$(34) Then
strTemp2 = Mid(strTemp2, 2)
End If
If Right(strTemp2, 1) = ChrW$(34) Then
strTemp2 = Left(strTemp2, Len(strTemp2) - 1)
End If
sMax = Val(strTemp2)
x = InStr(MyString, "MaxVelocity ")
If x = 0 Then Exit Sub
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp2 = Mid(MyString, x + 1, xx - (x + 1))
strTemp2 = Trim(strTemp2)
If Left(strTemp2, 1) = ChrW$(34) Then
strTemp2 = Mid(strTemp2, 2)
End If
If Right(strTemp2, 1) = ChrW$(34) Then
strTemp2 = Left(strTemp2, Len(strTemp2) - 1)
End If
For i = 1 To Len(strTemp2)
If Not IsNumeric(Mid(strTemp, i, 1)) Then
strUnit = strUnit & Mid(strTemp, i, 1)
End If
Next i
If strUnit = "mph" Then
intFlag = 1
ElseIf strUnit = "kph" Then
intFlag = 2
End If
sMaxVel = Val(strTemp2)
Exit Sub
Errtrap:
'
Resume Next

End Sub

Private Sub GetWagDet(strWag As String, strFind As String, strTemp As String)
Dim x As Long, xx As Long

x = InStr(strWag, strFind)
If x = 0 Then
strTemp = ""
Exit Sub
End If
x = InStr(x, strWag, "(")
xx = InStr(x, strWag, ")")
strTemp = Mid(strWag, x + 1, xx - (x + 1))
strTemp = Trim(strTemp)

End Sub
Private Sub GetWagDetail(strEng As String, strCoupler As String, strWagType As String, strMass As String, strLen As String, strName As String)
Dim MyString As String, x As Long, xx As Long, strCouple As String, strTemp As String

MyString = ReadUniFile(strEng)
MyString = Replace(MyString, "   ", " ")
DoEvents
MyString = Replace(MyString, "  ", " ")
DoEvents
x = InStr(MyString, "Mass (")
If x > 0 Then
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strMass = Mid(MyString, x + 1, xx - (x + 1))
strMass = Trim(strMass)
If Left(strMass, 1) = ChrW$(34) Then
strMass = Mid(strMass, 2)
End If
If Right(strMass, 1) = ChrW$(34) Then
strMass = Left(strMass, Len(strMass) - 1)
End If
End If
x = InStr(MyString, "Size ")
If x > 0 Then
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp = Mid(MyString, x + 1, xx - (x + 1))
strTemp = Trim(strTemp)
x = InStr(strTemp, " ")
xx = InStr(x + 2, strTemp, " ")
strLen = Mid(strTemp, xx + 1)
strLen = Trim(strLen)
If Right(strLen, 1) = "m" Then
strLen = Left(strLen, Len(strLen) - 1)
End If
End If
    j = InStr(MyString, "Name ")
    If j > 0 Then
    j = InStr(j, MyString, "(")
    jj = InStr(j, MyString, vbCrLf)
    strName = Mid(MyString, j + 1, jj - (j + 1) - 1)
    strName = Trim(strName)
    End If
x = InStr(MyString, "Coupling (")
If x = 0 Then
strCoupler = "NONE"
GoTo CarryON
End If
x = InStr(x, MyString, "Type")
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strCoupler = Mid(MyString, x + 1, xx - (x + 1))
strCoupler = Trim(strCoupler)
Select Case strCoupler
Case "Chain"
strCoupler = "C"
Case "Bar"
strCoupler = "B"
Case "Automatic"
strCoupler = "A"
Case Else
'Call MsgBox("Unknown coupler in " & strEng, vbExclamation, App.Title)
strCoupler = "X"
End Select
x = InStr(xx, MyString, "Coupling (")
If x = 0 Or (x - xx) > 500 Then
strCoupler = strCoupler & strCoupler
GoTo CarryON
End If
x = InStr(x, MyString, "Type")
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strCouple = Mid(MyString, x + 1, xx - (x + 1))
strCouple = Trim(strCouple)

Select Case strCouple
Case "Chain"
strCoupler = "C" & strCoupler
Case "Bar"
strCoupler = "B" & strCoupler
Case "Automatic"
strCoupler = "A" & strCoupler
Case Else
'Call MsgBox("Unknown coupler in " & strEng, vbExclamation, App.Title)
strCoupler = strCoupler & "X"
End Select
CarryON:
x = InStr(MyString, "Type")
If x = 0 Then
strWagType = "Unknown"
Exit Sub
End If
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strWagType = Mid(MyString, x + 1, xx - (x + 1))
strWagType = Trim(strWagType)

End Sub

Private Sub GetEngDetail(strEng As String, strCoupler As String, strPower As String, strMass As String, strLen As String, strName As String)
Dim MyString As String, x As Long, xx As Long, strCouple

MyString = ReadUniFile(strEng)
MyString = Replace(MyString, vbTab, "  ")
x = InStr(MyString, "Mass ")
If x > 0 Then
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strMass = Mid(MyString, x + 1, xx - (x + 1))
strMass = Trim(strMass)
If Left(strMass, 1) = ChrW$(34) Then
strMass = Mid(strMass, 2)
End If
If Right(strMass, 1) = ChrW$(34) Then
strMass = Left(strMass, Len(strMass) - 1)
End If
End If
x = InStr(MyString, "Size ")
If x > 0 Then
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp = Mid(MyString, x + 1, xx - (x + 1))
strTemp = Trim(strTemp)
x = InStr(strTemp, " ")
xx = InStr(x + 2, strTemp, " ")
strLen = Mid(strTemp, xx + 1)
strLen = Trim(strLen)
If Right(strLen, 1) = "m" Then
strLen = Left(strLen, Len(strLen) - 1)
End If

End If

    j = InStr(MyString, "Name ")
    If j > 0 Then
    j = InStr(j, MyString, "(")
    jj = InStr(j, MyString, vbCrLf)
    strName = Mid(MyString, j + 1, jj - (j + 1) - 1)
    strName = Trim(strName)
    End If
 
x = InStr(MyString, "Coupling (")
If x = 0 Then
x = InStr(MyString, "Coupling" & vbCrLf)
End If
If x = 0 Then
strCoupler = "X"
GoTo Another
End If
x = InStr(x, MyString, "Type")
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strCoupler = Mid(MyString, x + 1, xx - (x + 1))
strCoupler = Trim(strCoupler)
Another:
Select Case strCoupler
Case "Chain"
strCoupler = "C"
Case "Bar"
strCoupler = "B"
Case "Automatic"
strCoupler = "A"
Case Else

strCoupler = "X"
End Select
x = InStr(xx, MyString, "Coupling (")
If x = 0 Then
x = InStr(MyString, "Coupling" & vbCrLf)
If x = 0 Then
strCoupler = strCoupler & strCoupler
End If
GoTo CarryON
End If

If x - xx > 300 Then
strCoupler = strCoupler & strCoupler
GoTo CarryON
End If
x = InStr(x, MyString, "Type")
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strCouple = Mid(MyString, x + 1, xx - (x + 1))
strCouple = Trim(strCouple)
Select Case strCouple
Case "Chain"
strCoupler = "C" & strCoupler
Case "Bar"
strCoupler = "B" & strCoupler
Case "Automatic"
strCoupler = "A" & strCoupler
Case Else
'Call MsgBox("Unknown coupler in " & strEng, vbExclamation, App.Title)
strCoupler = strCoupler & "X"
End Select
CarryON:
x = InStr(MyString, "MaxPower")
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strPower = Mid(MyString, x + 1, xx - (x + 1))
strPower = Trim(strPower)

End Sub


Private Sub GetEngInfo(strEng As String)
Dim strTemp As String, strEngInf As String, strCou As String, intInd As Integer, strInd As String
Dim strCo As String, strMP As String, strMs As String


Call GetEngDet(strEng, "Description (", strTemp)

strTemp = Replace(strTemp, ChrW$(34), "")
strTemp = Replace(strTemp, "+", "")
strTemp = Replace(strTemp, "\n\n", "")

If strTemp <> "" Then
strEngInf = "Description : " & strTemp & vbCrLf
strEngInf = strEngInf & "**************************" & vbCrLf
End If

Call GetEngDet(strEng, "Type (", strTemp)
If strTemp <> "" Then
strEngInf = strEngInf & "Type : " & strTemp & vbCrLf
End If
Call GetEngDet(strEng, "WagonShape (", strTemp)
strWagonShape = strTemp
Call GetEngDet(strEng, "Size (", strTemp)
If strTemp <> "" Then
strEngInf = strEngInf & "Size : " & strTemp & vbCrLf
End If
Call GetEngDet(strEng, "Mass (", strTemp)
If strTemp <> "" Then
strEngInf = strEngInf & "Mass : " & strTemp & vbCrLf
End If
Rem **************
List1.Select List1.row, 1
strCo = List1.Cell(flexcpText)
List1.Select List1.row, 2
strMP = List1.Cell(flexcpText)
List1.Select List1.row, 4
strMs = List1.Cell(flexcpText)
Rem *********
List1.Select List1.row, 3
strInd = List1.Cell(flexcpText)
intInd = Val(strInd)
strCou = strConEngCou(intInd)
Select Case strCou
Case "AA"
strCou = "Automatic-Automatic"
Case "AB"
strCou = "Automatic-Bar"
Case "AC"
strCou = "Automatic-Chain"
Case "BB"
strCou = "Bar-Bar"
Case "BA"
strCou = "Bar-Automatic"
Case "BC"
strCou = "Bar-Chain"
Case "CC"
strCou = "Chain-Chain"
Case "CB"
strCou = "Chain-Bar"
Case "CA"
strCou = "Chain-Automatic"
Case Else
strCou = "Unknown"
End Select

strEngInf = strEngInf & "Coupler(s): " & strCou & vbCrLf
strEngInf = strEngInf & "Max Power: " & strMP
RTB1.Text = strEngInf


End Sub

Private Sub GetWagInfo(strWag As String)
Dim strTemp As String, strWagInf As String, strCou As String, intInd As Integer, strInd As String

Call GetWagDet(strWag, "Type (", strTemp)
If strTemp <> "" Then
strWagInf = "Type : " & strTemp & vbCrLf
End If
Call GetWagDet(strWag, "WagonShape (", strTemp)
strWagonShape = strTemp
Call GetWagDet(strWag, "Size (", strTemp)
If strTemp <> "" Then
strWagInf = strWagInf & "Size : " & strTemp & vbCrLf
End If
Call GetWagDet(strWag, "Mass (", strTemp)
If strTemp <> "" Then
strWagInf = strWagInf & "Mass : " & strTemp & vbCrLf
End If
List2.Select List2.row, 3
strInd = List2.Cell(flexcpText)
intInd = Val(strInd)
strCou = strConWagCou(intInd)
Select Case strCou
Case "AA"
strCou = "Automatic-Automatic"
Case "AB"
strCou = "Automatic-Bar"
Case "AC"
strCou = "Automatic-Chain"
Case "BB"
strCou = "Bar-Bar"
Case "BA"
strCou = "Bar-Automatic"
Case "BC"
strCou = "Bar-Chain"
Case "CC"
strCou = "Chain-Chain"
Case "CB"
strCou = "Chain-Bar"
Case "CA"
strCou = "Chain-Automatic"
Case Else
strCou = "Unknown"
End Select

strWagInf = strWagInf & "Coupler(s): " & strCou

RTB1.Text = strWagInf


End Sub



Private Sub MakeThumbNail(strModel2 As String, strEW As String)
Dim i As Integer, j As Integer, x As Integer
Dim strPix As String, strBatText As String, strTemp As String, MyString As String
Dim xx As Integer, strShape As String, strWagName As String, strAnim As String
Dim strStart As String, strTemp2 As String, PixPic As String, PixN As String, pixRN As String, PixP As String
Dim strModel As String

On Error GoTo Errtrap

booAbort = False
strPicView = vbNullString
strForPrint = vbNullString
strReport = vbNullString
flagThumb = 0
If Not DirExists(App.Path & "\TempFiles") Then

MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents

strSavePix = App.Path & "\Tempfiles"
strPixPath = strSavePix & "\"
SaveSetting "Decapod", "MSTS Shape Viewer", "screenshotLocation", strPixPath & "PIX"
DoEvents

strModel = Left(strModel2, Len(strModel2) - 3) & strEW

strTemp2 = strModel
strAnim = vbNullString


strStart = Left$(strTemp2, Len(strTemp2) - 3)
strStart = strStart & "jpg"

x = InStr(strTemp2, "Common.")
If x > 0 Then GoTo CarryON
x = InStr(strTemp2, "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(strTemp2, "Invisocar")
If x > 0 Then GoTo CarryON

MyString = ReadUniFile(strTemp2)
x = InStr(MyString, "wagonshape")
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strShape = Mid$(MyString, x + 1, xx - (x + 1))
strShape = Trim$(strShape)
If Left$(strShape, 1) = ChrW$(34) Then
strShape = Mid$(strShape, 2)
End If
If Right$(strShape, 1) = ChrW$(34) Then
strShape = Left$(strShape, Len(strShape) - 1)
End If
If Right$(strShape, 2) <> ".s" Then
strReport = strReport & "File " & strTemp2 & " has an invalid WagonShape entry so could not be processed" & vbCrLf

GoTo CarryON
End If

x = InStrRev(strTemp2, "\")
strShapePath = Left$(strTemp2, x)
strWagName = Mid$(strTemp2, x + 1)
Rem ************* Look for freightanim
x = InStr(MyString, "FreightAnim")
If x > 0 Then
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ".s")
strAnim = Mid$(MyString, x + 1, (xx + 2) - (x + 1))
strAnim = Trim$(strAnim)
If Left$(strAnim, 1) = ChrW$(34) Then
strAnim = Mid$(strAnim, 2)
End If
If Right$(strAnim, 1) = ChrW$(34) Then
strAnim = Left$(strAnim, Len(strAnim) - 1)
End If
If Left$(strAnim, 2) = ".." Then strAnim = vbNullString
End If

Rem *********************** Continue from here.
PixPic = strShapePath & strWagName
x = InStr(PixPic, "Common.")
If x > 0 Then GoTo CarryON
pixRN = PixPic
pixRN = Left$(pixRN, Len(pixRN) - 4) & ".jpg"
x = InStrRev(PixPic, "\", x - 1)
PixP = Mid$(PixPic, x + 1)
intNextPix = CInt(GetSetting("Decapod", "3D Train Control", "Lastscreenshot", 0))
If intNextPix > 999 Then
SaveSetting "Decapod", "3D Train Control", "Lastscreenshot", 0
intNextPix = 0
End If
strPix = Trim$(Str(intNextPix))
If Len(strPix) < 3 Then
strPix = String(3 - Len(strPix), "0") & strPix
End If
PixN = "Pix" & strPix & ".jpg"

strPicView = strShapePath & strShape
If strAnim <> vbNullString Then
strPicView = strPicView & ";" & strShapePath & strAnim
End If

strBatText = ChrW$(34) & App.Path & "\sviewRR4.exe" & ChrW$(34) & " " & ChrW$(34) & strPicView & ChrW$(34)
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)


TryAgain:
If Not FileExists(pixRN) Then
 FileCopy strPixPath & PixN, pixRN
 DoEvents
 Kill strPixPath & PixN
 DoEvents
 ElseIf FileExists(pixRN) Then
 x = InStrRev(pixRN, "\")
 strTemp = Mid$(pixRN, x + 1)

 If flagThumb = 0 Or flagThumb = 2 Then
 frmThumb.Show 1
 End If

If flagThumb = 0 Or flagThumb = 1 Then
 Kill pixRN
 DoEvents
 FileCopy strPixPath & PixN, pixRN
 DoEvents
 Kill strPixPath & PixN
 DoEvents
 ElseIf flagThumb = 2 Or flagThumb = 3 Then
 Kill strPixPath & PixN
 DoEvents
 End If
 End If



 j = j + 1
 If j > 999 Then
 j = 0
 End If


DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.Rows - 1
End If
DoEvents


MousePointer = 0
   If strReport <> vbNullString Then
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   End If
 
Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If

End Sub

Public Function RemD2(ByRef rArray() As Variant, xArray() As Variant) As Variant

    'Declare variables
    Dim ii As Long, jj As Long
 On Error GoTo Errtrap
    'Initialize variables
    count3 = 1
    high = UBound(rArray)
    'Declare temp array
    
    ReDim xArray(0 To high)
    
    'Start duplicates removal code

xArray(0) = rArray(0)
jj = 1
    For ii = 1 To high
        If rArray(ii) <> rArray(ii - 1) Then
        If rArray(ii) <> vbNullString Then
        xArray(jj) = rArray(ii)
        jj = jj + 1
        End If
End If
Next ii

If high > 0 Then
If rArray(high) <> rArray(high - 1) Then
xArray(high) = rArray(high)
End If
 End If
CarryON:

ReDim Preserve xArray(0 To jj - 1)

Exit Function
Errtrap:

Call MsgBox("An error " & Err & " occurred in subroutine 'RemD2' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Function

Private Function ReadUniFile(CompleteFilePath As String) As String

Dim length As Long, mytristate As Integer
Dim MyString As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean


Set File_obj = CreateObject("Scripting.FileSystemObject")
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, Me.Caption
  Exit Function
End If
mytristate = -1
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll
The_obj.Close
fileflag = False
ReadUniFile = MyString
End Function











Public Function QSort3(strList() As Variant, lLbound As Long, lUbound As Long)
    
    Dim strTemp As String
    Dim strBuffer As String
    Dim lngCurLow As Long
    Dim lngCurHigh As Long
    Dim lngCurMidpoint As Long
    
    lngCurLow = lLbound ' Start current low and high at actual low/high
    lngCurHigh = lUbound
    
    If lUbound <= lLbound Then Exit Function ' Error!
    lngCurMidpoint = (lLbound + lUbound) \ 2 ' Find the approx midpoint of the array
    
    strTemp = strList(lngCurMidpoint) ' Pick as a starting point (we are making
    ' an assumption that the data *might* be
    '
    ' in semi-sorted order already!
    


    Do While (lngCurLow <= lngCurHigh)


        Do While strList(lngCurLow) < strTemp
            lngCurLow = lngCurLow + 1
            If lngCurLow = lUbound Then Exit Do
        Loop
        


        Do While strTemp < strList(lngCurHigh)
            lngCurHigh = lngCurHigh - 1
            If lngCurHigh = lLbound Then Exit Do
        Loop


        If (lngCurLow <= lngCurHigh) Then ' if low is <= high then swap
            strBuffer = strList(lngCurLow)
            strList(lngCurLow) = strList(lngCurHigh)
            strList(lngCurHigh) = strBuffer
            '
            lngCurLow = lngCurLow + 1 ' CurLow++
            lngCurHigh = lngCurHigh - 1 ' CurLow--
        End If
        
    Loop
    


    If lLbound < lngCurHigh Then ' Recurse if necessary
        QSort3 strList(), lLbound, lngCurHigh
    End If
    


    If lngCurLow < lUbound Then ' Recurse if necessary
        QSort3 strList(), lngCurLow, lUbound
    End If
   
End Function


Private Sub Check1_Click()

If Check1.value = 0 Then
List1.ColHidden(1) = True
ElseIf Check1.value = 1 Then
List1.ColHidden(1) = False
End If

List1.Refresh
DoEvents
End Sub

Private Sub Check2_Click()
If Check2.value = 0 Then
List1.ColHidden(4) = True
ElseIf Check2.value = 1 Then
List1.ColHidden(4) = False
End If

List1.Refresh
DoEvents
End Sub

Private Sub Check3_Click()
If Check3.value = 0 Then
List1.ColHidden(2) = True
ElseIf Check3.value = 1 Then
List1.ColHidden(2) = False
End If

List1.Refresh
DoEvents
End Sub


Private Sub Check4_Click()
If Check4.value = 0 Then
List1.ColHidden(5) = True
ElseIf Check4.value = 1 Then
List1.ColHidden(5) = False
End If

List1.Refresh
DoEvents
End Sub

Private Sub Check5_Click()
If Check5.value = 0 Then
List2.ColHidden(1) = True
ElseIf Check5.value = 1 Then
List2.ColHidden(1) = False
End If

List2.Refresh
DoEvents
End Sub

Private Sub Check6_Click()
If Check6.value = 0 Then
List2.ColHidden(2) = True
ElseIf Check6.value = 1 Then
List2.ColHidden(2) = False
End If

List2.Refresh
DoEvents
End Sub


Private Sub Check7_Click()
If Check7.value = 0 Then
List2.ColHidden(4) = True
ElseIf Check7.value = 1 Then
List2.ColHidden(4) = False
End If

List2.Refresh
DoEvents
End Sub


Private Sub Check8_Click()
If Check8.value = 0 Then
List2.ColHidden(5) = True
ElseIf Check8.value = 1 Then
List2.ColHidden(5) = False
End If

List2.Refresh
DoEvents
End Sub


Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Command2_Click()
Dim i As Integer, strConsist As String, strEngpath As String, strEngname As String
Dim strWagPath As String, strWagName As String, strConHead As String, strConName As String
Dim strConDisp As String, sTon As Single, sForce As Single, sTotW As Single, sTotForce As Single
Dim sMaxVel As Single, flagVel As Integer, sAcc As Single, strMaxVel As String, strAcc As String
Dim strTrainCfg As String, intNextUiD As Integer, j As Integer, strConIt As String, strConFld As String

booIgnore = False
If Text1(0) = "" Then
Call MsgBox("You have not provided a File Name for your Consist?", vbExclamation Or vbDefaultButton1, App.Title)
Exit Sub
End If
strConName = Text1(0)
If Right(strConName, 4) <> ".con" Then
strConName = strConName & ".con"
End If
'End If
If Text1(1) <> "" Then
strConDisp = Text1(1)
Else
strConDisp = ""
End If

For i = 1 To numStock
Select Case intConType(i)
Case 1
Call GetEngPaths(strConPic(i), strEngpath, strEngname, sTon, sForce, sMaxVel, flagVel)
strConItem(i) = strEngname
strConFolder(i) = strEngpath
strConUid(i) = "UiD ( " & Trim(Str(i - 1)) & " )"
strConsist = strConsist & vbTab & vbTab & "Engine (" & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & strConUid(i) & vbCrLf
strConIt = strConItem(i)
strConFld = strConFolder(i)
j = InStr(strConIt, " ")
j = InStr(strConFld, " ")
If j > 0 Then
strConIt = ChrW$(34) & strConIt & ChrW$(34)
End If
If j > 0 Then
strConFld = ChrW$(34) & strConFld & ChrW$(34)
End If
strConsist = strConsist & vbTab & vbTab & vbTab & "EngineData ( " & strConIt & " " & strConFld & " )" & vbCrLf
strConsist = strConsist & vbTab & vbTab & ")" & vbCrLf

sTotW = sTotW + sTon
sTotForce = sTotForce + sForce
Case 2
Call GetWagPaths(strConPic(i), strWagPath, strWagName, sTon)
strConItem(i) = strWagName
strConFolder(i) = strWagPath
strConUid(i) = "UiD ( " & Trim(Str(i - 1)) & " )"
strConsist = strConsist & vbTab & vbTab & "Wagon (" & vbCrLf

strConIt = strConItem(i)
strConFld = strConFolder(i)
j = InStr(strConIt, " ")
j = InStr(strConFld, " ")
If j > 0 Then
strConIt = ChrW$(34) & strConIt & ChrW$(34)
End If
If j > 0 Then
strConFld = ChrW$(34) & strConFld & ChrW$(34)
End If
strConsist = strConsist & vbTab & vbTab & vbTab & "WagonData ( " & strConIt & " " & strConFld & " )" & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & strConUid(i) & vbCrLf
strConsist = strConsist & vbTab & vbTab & ")" & vbCrLf

sTotW = sTotW + sTon
Case 3
Call GetEngPaths(strConPic(i), strEngpath, strEngname, sTon, sForce, sMaxVel, flagVel)
strConItem(i) = strEngname
strConFolder(i) = strEngpath
strConUid(i) = "UiD ( " & Trim(Str(i - 1)) & " )"
strConsist = strConsist & vbTab & vbTab & "Engine (" & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & strConUid(i) & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & "Flip ( )" & vbCrLf
strConIt = strConItem(i)
strConFld = strConFolder(i)
j = InStr(strConIt, " ")
j = InStr(strConFld, " ")
If j > 0 Then
strConIt = ChrW$(34) & strConIt & ChrW$(34)
End If
If j > 0 Then
strConFld = ChrW$(34) & strConFld & ChrW$(34)
End If
strConsist = strConsist & vbTab & vbTab & vbTab & "EngineData ( " & strConIt & " " & strConFld & " )" & vbCrLf
strConsist = strConsist & vbTab & vbTab & ")" & vbCrLf
sTotW = sTotW + sTon
sTotForce = sTotForce + sForce
Case 4
Call GetWagPaths(strConPic(i), strWagPath, strWagName, sTon)
strConItem(i) = strWagName
strConFolder(i) = strWagPath
strConUid(i) = "UiD ( " & Trim(Str(i - 1)) & " )"
strConsist = strConsist & vbTab & vbTab & "Wagon (" & vbCrLf

strConIt = strConItem(i)
strConFld = strConFolder(i)
j = InStr(strConIt, " ")
j = InStr(strConFld, " ")
If j > 0 Then
strConIt = ChrW$(34) & strConIt & ChrW$(34)
End If
If j > 0 Then
strConFld = ChrW$(34) & strConFld & ChrW$(34)
End If
strConsist = strConsist & vbTab & vbTab & vbTab & "WagonData ( " & strConIt & " " & strConFld & " )" & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & strConUid(i) & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & "Flip ( )" & vbCrLf
strConsist = strConsist & vbTab & vbTab & ")" & vbCrLf
sTotW = sTotW + sTon
End Select

Next i
intNextUiD = i

strConsist = strConsist & vbTab & ")" & vbCrLf
strConsist = strConsist & ")" & vbCrLf
If flagVel = 0 Then flagVel = 1
If flagVel = 1 Then
sMaxVel = sMaxVel / 2.236
ElseIf flagVel = 2 Then
sMaxVel = sMaxVel / 1.6
End If
sAcc = (0.8 * sTotForce) / sTotW
strMaxVel = Str(sMaxVel)
strAcc = Str(sAcc)
strAcc = Format(strAcc, "0.0000")
strMaxVel = Format(strMaxVel, "0.0000")
strTrainCfg = Left(strConName, Len(strConName) - 4)
i = InStr(strTrainCfg, " ")
If i > 0 Then
strTrainCfg = ChrW$(34) & strTrainCfg & ChrW$(34)
End If
strConHead = "SIMISA@@@@@@@@@@JINX0D0t______" & vbCrLf & vbCrLf
strConHead = strConHead & "Train (" & vbCrLf
strConHead = strConHead & "  TrainCfg ( " & strTrainCfg & vbCrLf
If strConDisp <> "" Then
i = InStr(strConDisp, " ")
If i > 0 Then
strConHead = strConHead & vbTab & "Name ( " & ChrW$(34) & strConDisp & ChrW$(34) & " )" & vbCrLf
Else
strConHead = strConHead & vbTab & "Name ( " & strConDisp & " )" & vbCrLf
End If
End If
strConHead = strConHead & vbTab & vbTab & "Serial ( " & Trim(Str(intConSerial + 1)) & " )" & vbCrLf
strConHead = strConHead & vbTab & vbTab & "MaxVelocity ( " & strMaxVel & " " & strAcc & " )" & vbCrLf
strConHead = strConHead & vbTab & vbTab & "NextWagonUID ( " & Trim(Str(intNextUiD - 1)) & " )" & vbCrLf
strConHead = strConHead & vbTab & vbTab & "Durability ( 1.00 )" & vbCrLf

strConsist = strConHead & strConsist

DoEvents
Call WriteUniFile(MSTSPath & "\Trains\Consists\" & strConName, strConsist)
End Sub

Private Sub Command3_Click(Index As Integer)
Dim x As Long, xx As Long, i As Integer
Dim Y As Long, yy As Long, q As Integer, j As Integer, strTemp As String, maxPic As Integer


On Error GoTo Errtrap


If numStock < 8 Then
maxPic = numStock
Else
maxPic = 7
End If
Select Case Index
Case 0
If numStock > 8 Then
    For PicNum = 0 To maxPic
    Image1(PicNum).Visible = False
    PicStart(PicNum).Picture = LoadPicture(strConPic(PicNum + 1))
    Label3(PicNum).Caption = PicNum + 1
        If intConType(PicNum + 1) > 2 Then
        Label3(PicNum + 8).Caption = ">>>>"
        Else
        Label3(PicNum + 8).Caption = ""
    End If
    j = InStrRev(strConPic(PicNum + 1), "\")
    strTemp = Mid(strConPic(PicNum + 1), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
    Label5(PicNum + 8).Caption = strConListName(PicNum + 1)
    Label8(PicNum).Caption = Left(strConCou(PicNum + 1), 1)
Label8(PicNum + 8).Caption = Mid(strConCou(PicNum + 1), 2)
    x = PicStart(PicNum).width
    Y = PicStart(PicNum).height
    xx = x
    yy = Y
    If x > Picture2(PicNum).width Then
        Do Until x = Picture2(PicNum).width
            x = x - 1
            Y = Y - 1
        Loop
    End If
    Y = (x / xx) * yy
    Image1(PicNum).width = x
    Image1(PicNum).height = Y
    Image1(PicNum).Top = (Picture2(PicNum).height \ 2) - (Image1(PicNum).height \ 2)
    Image1(PicNum).Left = (Picture2(PicNum).width \ 2) - (Image1(PicNum).width \ 2)
    Image1(PicNum).Picture = PicStart(PicNum).Picture
    Image1(PicNum).Visible = True
    'q = q - 1
    Next PicNum
    End If
Case 4

If numStock > 8 Then

q = Val(Label3(1).Caption)
If q > numStock Then Exit Sub
maxPic = numStock - q
If maxPic > 7 Then maxPic = 7
For PicNum = 0 To maxPic

    Image1(PicNum).Visible = False
    PicStart(PicNum).Picture = LoadPicture(strConPic(q))
    Label3(PicNum).Caption = q
        If intConType(q) > 2 Then
        Label3(PicNum + 8).Caption = ">>>>"
        Else
        Label3(PicNum + 8).Caption = ""
    End If
    j = InStrRev(strConPic(q), "\")
    strTemp = Mid(strConPic(q), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
     Label5(PicNum + 8).Caption = strConListName(q)
     Label8(PicNum).Caption = Left(strConCou(q), 1)
Label8(PicNum + 8).Caption = Mid(strConCou(q), 2)
    x = PicStart(PicNum).width
    Y = PicStart(PicNum).height
    xx = x
    yy = Y
    If x > Picture2(PicNum).width Then
        Do Until x = Picture2(PicNum).width
            x = x - 1
            Y = Y - 1
        Loop
    End If
    Y = (x / xx) * yy
    Image1(PicNum).width = x
    Image1(PicNum).height = Y
    Image1(PicNum).Top = (Picture2(PicNum).height \ 2) - (Image1(PicNum).height \ 2)
    Image1(PicNum).Left = (Picture2(PicNum).width \ 2) - (Image1(PicNum).width \ 2)
    Image1(PicNum).Picture = PicStart(PicNum).Picture
    Image1(PicNum).Visible = True
   q = q + 1
    Next PicNum
    End If
    If maxPic < 7 Then
    For i = maxPic + 1 To 7
Set Picture2(i) = LoadPicture("")
Set Image1(i) = LoadPicture("")
Set PicStart(i) = LoadPicture("")
Label3(i).Caption = ""
Label5(i).Caption = ""
Label8(i).Caption = ""
Next i
For i = maxPic + 8 To 15
Label3(i).Caption = ""
Label5(i).Caption = ""
Label8(i).Caption = ""
Next i
End If
Case 2

Case 3

Case 1
If numStock > 8 Then
q = Val(Label3(0).Caption) - 1
If q = 0 Then Exit Sub
For PicNum = 0 To maxPic

    Image1(PicNum).Visible = False
    PicStart(PicNum).Picture = LoadPicture(strConPic(q))
    Label3(PicNum).Caption = q
        If intConType(q) > 2 Then
        Label3(PicNum + 8).Caption = ">>>>"
        Else
        Label3(PicNum + 8).Caption = ""
    End If
    j = InStrRev(strConPic(q), "\")
    strTemp = Mid(strConPic(q), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
     Label5(PicNum + 8).Caption = strConListName(q)
     Label8(PicNum).Caption = Left(strConCou(q), 1)
Label8(PicNum + 8).Caption = Mid(strConCou(q), 2)
    x = PicStart(PicNum).width
    Y = PicStart(PicNum).height
    xx = x
    yy = Y
    If x > Picture2(PicNum).width Then
        Do Until x = Picture2(PicNum).width
            x = x - 1
            Y = Y - 1
        Loop
    End If
    Y = (x / xx) * yy
    Image1(PicNum).width = x
    Image1(PicNum).height = Y
    Image1(PicNum).Top = (Picture2(PicNum).height \ 2) - (Image1(PicNum).height \ 2)
    Image1(PicNum).Left = (Picture2(PicNum).width \ 2) - (Image1(PicNum).width \ 2)
    Image1(PicNum).Picture = PicStart(PicNum).Picture
    Image1(PicNum).Visible = True
   q = q + 1
    Next PicNum
    End If
Case 5

If numStock > 8 Then
q = 7
For PicNum = 0 To maxPic

    Image1(PicNum).Visible = False
    PicStart(PicNum).Picture = LoadPicture(strConPic(numStock - q))
    Label3(PicNum).Caption = numStock - q
        If intConType(numStock - q) > 2 Then
        Label3(PicNum + 8).Caption = ">>>>"
        Else
        Label3(PicNum + 8).Caption = ""
    End If
    j = InStrRev(strConPic(numStock - q), "\")
    strTemp = Mid(strConPic(numStock - q), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
     Label5(PicNum + 8).Caption = strConListName(numStock - q)
     Label8(PicNum).Caption = Left(strConCou(numStock - q), 1)
Label8(PicNum + 8).Caption = Mid(strConCou(numStock - q), 2)
    x = PicStart(PicNum).width
    Y = PicStart(PicNum).height
    xx = x
    yy = Y
    If x > Picture2(PicNum).width Then
        Do Until x = Picture2(PicNum).width
            x = x - 1
            Y = Y - 1
        Loop
    End If
    Y = (x / xx) * yy
    Image1(PicNum).width = x
    Image1(PicNum).height = Y
    Image1(PicNum).Top = (Picture2(PicNum).height \ 2) - (Image1(PicNum).height \ 2)
    Image1(PicNum).Left = (Picture2(PicNum).width \ 2) - (Image1(PicNum).width \ 2)
    Image1(PicNum).Picture = PicStart(PicNum).Picture
    Image1(PicNum).Visible = True
    DoEvents
   q = q - 1
    Next PicNum
    DoEvents
    Else
    q = 1
For PicNum = 0 To maxPic - 1

    Image1(PicNum).Visible = False
    PicStart(PicNum).Picture = LoadPicture(strConPic(q))
    Label3(PicNum).Caption = q
        If intConType(q) > 2 Then
        Label3(PicNum + 8).Caption = ">>>>"
        Else
        Label3(PicNum + 8).Caption = ""
    End If
    j = InStrRev(strConPic(q), "\")
    strTemp = Mid(strConPic(q), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
     Label5(PicNum + 8).Caption = strConListName(q)
     Label8(PicNum).Caption = Left(strConCou(q), 1)
Label8(PicNum + 8).Caption = Mid(strConCou(q), 2)
    x = PicStart(PicNum).width
    Y = PicStart(PicNum).height
    xx = x
    yy = Y
    If x > Picture2(PicNum).width Then
        Do Until x = Picture2(PicNum).width
            x = x - 1
            Y = Y - 1
        Loop
    End If
    Y = (x / xx) * yy
    Image1(PicNum).width = x
    Image1(PicNum).height = Y
    Image1(PicNum).Top = (Picture2(PicNum).height \ 2) - (Image1(PicNum).height \ 2)
    Image1(PicNum).Left = (Picture2(PicNum).width \ 2) - (Image1(PicNum).width \ 2)
    Image1(PicNum).Picture = PicStart(PicNum).Picture
    Image1(PicNum).Visible = True
    DoEvents
   q = q + 1
    Next PicNum
    DoEvents
    End If
End Select
Exit Sub
Errtrap:
End Sub



Private Sub Command4_Click()
Dim strBatText As String, NewFile As Integer, i As Integer



strPicView = strMainPix
i = InStrRev(strPicView, "\")
strPicView = Left(strPicView, i) & strWagonShape
     NewFile = FreeFile
    strBatText = ChrW$(34) & App.Path & "\sviewRR.exe" & ChrW$(34) & " " & ChrW$(34) & strPicView & ChrW$(34) & ";"

   Open App.Path & "\TempFiles2\doBat.bat" For Output As #NewFile
   Print #NewFile, strBatText
   Close #NewFile
     ChDrive Left(App.Path, 1)
   ChDir App.Path & "\TempFiles2"
  DoEvents
    Call ShellAndWait("doBat.bat", True, vbHide)
        DoEvents
   
  
End Sub

Private Sub Command5_Click()
Call ClearAll
End Sub

Private Sub Command7_Click()
Dim MyString As String, i As Integer, x As Long, xx As Long, Y As Long, yy As Long
Dim strTemp As String, strCN As String, strCF As String, j As Integer
Dim Entry As String, strCoupler As String, strPower As String, strMass As String, strLen As String
Dim strWagType As String, sLen As Single, sLoad As Single, strName As String, strFLoad As String, strFLen As String
Dim strStock As String, ii As Integer

On Error GoTo Errtrap
Call ClearAll

Text1(0) = strActFile
i = Val(Text4)

If i > ActNum Then Exit Sub
MyString = strLoose(i)

Text1(1).Text = "Loose Consist #" & Str(i)
x = InStr(MyString, "Serial (")
If x > 0 Then
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp = Mid(MyString, x + 1, xx - (x + 1))
strTemp = Trim(strTemp)
intConSerial = Val(strTemp)
End If

numStock = 0
x = 1
Do
x = InStr(x + 1, MyString, "EngineData (")
If x > 0 Then
numStock = numStock + 1
End If
Loop While x > 0
x = 1
Do
x = InStr(x + 1, MyString, "WagonData (")
If x > 0 Then
numStock = numStock + 1
End If
Loop While x > 0

x = 1
For ii = 1 To numStock

x = InStr(x + 1, MyString, "Data (")
strTemp = Mid(MyString, x - 5, 5)
If strTemp = "ngine" Then

intConType(ii) = 1
Y = InStrRev(MyString, "Flip", x)
If Y > 0 And (x - Y) < 55 Then intConType(ii) = 3
ElseIf strTemp = "Wagon" Then
intConType(ii) = 2
Y = InStr(x, MyString, ")")
yy = InStr(Y, MyString, "Flip")
If yy > 0 And (yy - Y) < 55 Then intConType(ii) = 4

End If

x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp = Mid(MyString, x + 1, xx - (x + 1))
strTemp = Trim(strTemp)
j = InStr(strTemp, ChrW$(34))
If j = 0 Then   'No quotes
j = InStr(strTemp, " ")
strCN = Left(strTemp, j - 1)
strCF = Mid(strTemp, j + 1)
ElseIf j = 1 Then 'quotes around stock name
j = InStr(j + 1, strTemp, ChrW$(34))
strCN = Left(strTemp, j - 1)
strCF = Mid(strTemp, j + 2)
strCN = Trim(strCN)
strCF = Trim(strCF)
Else
strCN = Left(strTemp, j - 1)
strCF = Mid(strTemp, j + 1)
strCN = Trim(strCN)
strCF = Trim(strCF)
End If

If Left(strCN, 1) = ChrW$(34) Then
strCN = Mid(strCN, 2)
End If
If Right(strCN, 1) = ChrW$(34) Then
strCN = Left(strCN, Len(strCN) - 1)
End If
If Left(strCF, 1) = ChrW$(34) Then
strCF = Mid(strCF, 2)
End If
If Right(strCF, 1) = ChrW$(34) Then
strCF = Left(strCF, Len(strCF) - 1)
End If
If intConType(ii) = 1 Or intConType(ii) = 3 Then
strTemp = ".eng"
Else
strTemp = ".wag"
End If
strStock = MSTSPath & "\Trains\Trainset\" & strCF & "\" & strCN & strTemp
If Not FileExists(strStock) Then
Call MsgBox(strCF & "\" & strCN & strTemp & " listed in this consist appears to be missing?" _
    & vbCrLf & "Consist could not be loaded." _
    , vbExclamation, App.Title)
                                             

Exit Sub
End If

strConPic(ii) = MSTSPath & "\Trains\Trainset\" & strCF & "\" & strCN & ".jpg"
If Not FileExists(MSTSPath & "\Trains\Trainset\" & strCF & "\" & strCN & ".jpg") Then
strConPic(ii) = App.Path & "\Thomas.jpg"
End If
If intConType(ii) = 1 Then
Entry = MSTSPath & "\Trains\Trainset\" & strCF & "\" & strCN & ".eng"
strCoupler = ""
strPower = ""
strMass = ""
strLen = ""
strName = ""
Call GetEngDetail(Entry, strCoupler, strPower, strMass, strLen, strName)
strConListName(ii) = strName
 sLen = Val(strLen)
   conLen = conLen + sLen
   strFLen = Str(conLen * 3.28)
   strFLen = Format(strFLen, "####.000")
   Label6(3).Caption = "Consist Length: " & Str(conLen) & " m = " & strFLen & " ft"
  
    sLoad = Val(strMass)
   engLoad = engLoad + sLoad
    strFLoad = Str(engLoad * (2205 / 2000))
   strFLoad = Format(strFLoad, "####.000")
   Label6(1).Caption = "Engine Mass: " & Str(engLoad) & " t = " & strFLoad & " U.S. or Short tons"
   
  
ElseIf intConType(ii) = 2 Then
Entry = MSTSPath & "\Trains\Trainset\" & strCF & "\" & strCN & ".wag"
strCoupler = ""
strPower = ""
strMass = ""
strLen = ""
strName = ""

Call GetWagDetail(Entry, strCoupler, strWagType, strMass, strLen, strName)
strConListName(ii) = strName
 sLen = Val(strLen)
   conLen = conLen + sLen
   strFLen = Str(conLen * 3.28)
   strFLen = Format(strFLen, "####.000")
   Label6(3).Caption = "Consist Length: " & Str(conLen) & " m = " & strFLen & " ft"
   
  
    sLoad = Val(strMass)
   wagLoad = wagLoad + sLoad
   strFLoad = Str(wagLoad * (2205 / 2000))
   strFLoad = Format(strFLoad, "####.000")
   Label6(2).Caption = "Wagon Mass: " & Str(wagLoad) & " t = " & strFLoad & " U.S. or Short tons"
  
End If

Next ii
DoEvents
origConLen = conLen
Call LoadImage(numStock)
DoEvents
Call Command3_Click(5)
DoEvents
Call Command3_Click(0)
DoEvents
If i < ActNum Then
Text4 = Str(i + 1)
End If
Exit Sub
Errtrap:
Resume Next
End Sub

Private Sub Command8_Click()
Dim i As Integer, strConsist As String, strEngpath As String, strEngname As String
Dim strWagPath As String, strWagName As String
Dim sTon As Single, sForce As Single, sTotW As Single, sTotForce As Single
Dim sMaxVel As Single, flagVel As Integer
Dim intNextUiD As Integer, j As Integer, strConIt As String, strConFld As String
Dim intCon As Integer, strThisCon As String, strThisConStart As String, strThisConEnd As String
Dim strTemp As String, strTemp2 As String

booIgnore = False
If conLen > origConLen Then
Select Case MsgBox("During editing you have increased the length of this consist." _
                   & vbCrLf & "It may no longer fit in the position where it stands which would cause" _
                   & vbCrLf & "an error in MSTS. Do you still wish to proceed?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

    Case vbNo
Exit Sub
End Select



End If
intCon = Val(Text4) - 1
strThisCon = strLoose(intCon)
i = InStr(strThisCon, "Durability")
i = InStr(i, strThisCon, ")")
strThisConStart = Left(strThisCon, i)
i = InStr(strThisCon, "Direction")
i = InStrRev(strThisCon, ")", i)

strThisConEnd = Mid(strThisCon, i + 1)


For i = 1 To numStock
Select Case intConType(i)
Case 1
Call GetEngPaths(strConPic(i), strEngpath, strEngname, sTon, sForce, sMaxVel, flagVel)
strConItem(i) = strEngname
strConFolder(i) = strEngpath
strConUid(i) = "UiD ( " & Trim(Str(i - 1)) & " )"
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & "Engine (" & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & strConUid(i) & vbCrLf
strConIt = strConItem(i)
strConFld = strConFolder(i)
j = InStr(strConIt, " ")
j = InStr(strConFld, " ")
If j > 0 Then
strConIt = ChrW$(34) & strConIt & ChrW$(34)
End If
If j > 0 Then
strConFld = ChrW$(34) & strConFld & ChrW$(34)
End If
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "EngineData ( " & strConIt & " " & strConFld & " )" & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & ")" & vbCrLf

sTotW = sTotW + sTon
sTotForce = sTotForce + sForce
Case 2
Call GetWagPaths(strConPic(i), strWagPath, strWagName, sTon)
strConItem(i) = strWagName
strConFolder(i) = strWagPath
strConUid(i) = "UiD ( " & Trim(Str(i - 1)) & " )"
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & "Wagon (" & vbCrLf

strConIt = strConItem(i)
strConFld = strConFolder(i)
j = InStr(strConIt, " ")
j = InStr(strConFld, " ")
If j > 0 Then
strConIt = ChrW$(34) & strConIt & ChrW$(34)
End If
If j > 0 Then
strConFld = ChrW$(34) & strConFld & ChrW$(34)
End If
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "WagonData ( " & strConIt & " " & strConFld & " )" & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & strConUid(i) & vbCrLf
strConsist = strConsist & vbTab & vbTab & ")" & vbCrLf

sTotW = sTotW + sTon
Case 3
Call GetEngPaths(strConPic(i), strEngpath, strEngname, sTon, sForce, sMaxVel, flagVel)
strConItem(i) = strEngname
strConFolder(i) = strEngpath
strConUid(i) = "UiD ( " & Trim(Str(i - 1)) & " )"
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & "Engine (" & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & strConUid(i) & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Flip ( )" & vbCrLf
strConIt = strConItem(i)
strConFld = strConFolder(i)
j = InStr(strConIt, " ")
j = InStr(strConFld, " ")
If j > 0 Then
strConIt = ChrW$(34) & strConIt & ChrW$(34)
End If
If j > 0 Then
strConFld = ChrW$(34) & strConFld & ChrW$(34)
End If
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "EngineData ( " & strConIt & " " & strConFld & " )" & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & ")" & vbCrLf
sTotW = sTotW + sTon
sTotForce = sTotForce + sForce
Case 4
Call GetWagPaths(strConPic(i), strWagPath, strWagName, sTon)
strConItem(i) = strWagName
strConFolder(i) = strWagPath
strConUid(i) = "UiD ( " & Trim(Str(i - 1)) & " )"
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & "Wagon (" & vbCrLf

strConIt = strConItem(i)
strConFld = strConFolder(i)
j = InStr(strConIt, " ")
j = InStr(strConFld, " ")
If j > 0 Then
strConIt = ChrW$(34) & strConIt & ChrW$(34)
End If
If j > 0 Then
strConFld = ChrW$(34) & strConFld & ChrW$(34)
End If
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "WagonData ( " & strConIt & " " & strConFld & " )" & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & strConUid(i) & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Flip ( )" & vbCrLf
strConsist = strConsist & vbTab & vbTab & vbTab & vbTab & vbTab & ")" & vbCrLf
sTotW = sTotW + sTon
End Select

Next i
intNextUiD = i - 1

strConsist = strConsist & vbTab & ")" & vbCrLf
strConsist = strConsist & ")" & vbCrLf

strConsist = strThisConStart & vbCrLf & strConsist & strThisConEnd
j = InStr(strConsist, "NextWagonUID")
j = InStr(j, strConsist, "(")
i = InStr(j, strConsist, ")")
strTemp = Left(strConsist, j)
strTemp2 = Mid(strConsist, i)
strConsist = strTemp & " " & Trim(Str(intNextUiD)) & " " & strTemp2


strLoose(intCon) = strConsist
End Sub

Private Sub Command9_Click()
Dim i As Integer, strTemp As String

For i = 1 To ActNum
strTemp = strTemp & strLoose(i)
Next i
strTemp = strActStart & strTemp & strActEnd

If Not FileExists(strCFile & ".bak") Then
Name strCFile As strCFile & ".bak"
Else
For i = 1 To 50
    If Not FileExists(strCFile & ".bak" & Trim(Str(i))) Then
    Name strCFile As strCFile & ".bak" & Trim(Str(i))
    Exit For
    End If
Next i
End If

DoEvents


Call WriteUniFile(strCFile, strTemp)

booIsAct = False
End Sub

Private Sub Form_Load()
Dim FirstPath As String, DirCount As Integer

On Error GoTo Errtrap
ReDim strConEng(1 To Wag_CHUNK)
ReDim strConWag(1 To Wag_CHUNK)
ReDim strConEngPath(1 To Wag_CHUNK)
ReDim strConWagPath(1 To Wag_CHUNK)
ReDim strConEngCou(1 To Wag_CHUNK)
ReDim strConWagCou(1 To Wag_CHUNK)
List1.ColSort(0) = flexSortGenericAscending
List2.ColSort(0) = flexSortGenericAscending

If booLink = True Then Exit Sub

If frmUtils.Dir1(cursouind).Path <> frmUtils.Dir1(cursouind).List(frmUtils.Dir1(cursouind).ListIndex) Then
        frmUtils.Dir1(cursouind).Path = frmUtils.Dir1(cursouind).List(frmUtils.Dir1(cursouind).ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If


    frmUtils.File1(cursouind).Pattern = frmUtils.Text1(cursouind).Text
    FirstPath = frmUtils.Dir1(cursouind).Path
    DirCount = frmUtils.Dir1(cursouind).ListCount
    Rem ********************************************************************************
booConEng = True
booConWag = False
numEng = 1
Set SP = New cScanPath

strFilter = "*.eng"
With SP
            .Archive = True
            .Compressed = True
            .Hidden = False
            .Normal = True
            .ReadOnly = False
            .System = False
            .Filter = strFilter
            .StartScan FirstPath, True, False And False, True, False
        End With
        DoEvents
 Rem ****************** Now Wagons ***********************************
 booConWag = True
 booConEng = False
 Set SP = New cScanPath
numWag = 1
strFilter = "*.wag"
With SP
            .Archive = True
            .Compressed = True
            .Hidden = False
            .Normal = True
            .ReadOnly = False
            .System = False
            .Filter = strFilter
            .StartScan FirstPath, True, False And False, True, False
        End With
DoEvents
    
    
Rem *********************************************************************
    ' Start recursive direcory search.
                        ' Reset found files indicator.
  '  result = DirDiver(FirstPath, DirCount, "")
    If booAbort = True Then
'Unload frmConEdit
Exit Sub
End If
DoEvents
List1.Sort = 1
List2.Sort = 1
DoEvents
frmConEdit.Show
DoEvents
Exit Sub
Errtrap:
Call MsgBox(Err.Description & " occurred in frmConEdit Load", vbExclamation, App.Title)


End Sub

Private Function WriteUniFile(CompleteFilePath As String, MyString As String) As String


Dim File_obj As Object, The_obj As Object
On Error GoTo Errtrap

Set File_obj = CreateObject("Scripting.FileSystemObject")

Set The_obj = File_obj.CreateTextFile(CompleteFilePath, True, True)
The_obj.Write (MyString)
The_obj.Close
Exit Function
Errtrap:
Call MsgBox(Err.Description & " occurred in WriteUniFile while processing" _
            & vbCrLf & "file " & CompleteFilePath _
            , vbExclamation, App.Title)


End Function



Private Sub LoadImage2(strPic As String)
    Dim x As Long, xx As Long
    Dim Y As Long, yy As Long

    
    
    Image2.Visible = False
    If FileExists(strPic) Then
    pixStart.Picture = LoadPicture(strPic)
    
    Else
    pixStart.Picture = LoadPicture(App.Path & "\Thomas.jpg")
    
    End If
    x = pixStart.width
    Y = pixStart.height
    xx = x
    yy = Y
    If x > Picture1.width Then
        Do Until x = Picture1.width
            x = x - 1
            Y = Y - 1
        Loop
       
    End If
    Y = (x / xx) * yy
    Image2.width = x
    Image2.height = Y
    Image2.Top = (Picture1.height \ 2) - (Image2.height \ 2)
    Image2.Left = (Picture1.width \ 2) - (Image2.width \ 2)
    Image2.Picture = pixStart.Picture
    Image2.Visible = True
    
End Sub


Private Sub LoadImage(numStock As Integer)
    Dim x As Long, xx As Long
    Dim Y As Long, yy As Long, q As Integer, j As Integer, strTemp As String, strEWX As String

    If booInsert = False Then
    If numStock < 9 Then
    PicNum = numStock - 1
    
    Image1(PicNum).Visible = False
    If Not FileExists(strConPic(PicNum + 1)) Then
    If intConType(PicNum + 1) = 1 Or intConType(PicNum + 1) = 3 Then
    strEWX = "eng"
    Else
    strEWX = "wag"
    End If
    
    Call MakeThumbNail(strConPic(PicNum + 1), strEWX)
    DoEvents
    End If
    
    
    
    If FileExists(strConPic(PicNum + 1)) Then
    PicStart(PicNum).Picture = LoadPicture(strConPic(PicNum + 1))
    Label3(PicNum).Caption = PicNum + 1
    If intConType(PicNum + 1) > 2 Then
        Label3(PicNum + 8).Caption = ">>>>"
        Else
        Label3(PicNum + 8).Caption = ""
    End If
    j = InStrRev(strConPic(PicNum + 1), "\")
    strTemp = Mid(strConPic(PicNum + 1), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
    Label5(PicNum + 8).Caption = strConListName(PicNum + 1)
Label8(PicNum).Caption = Left(strConCou(PicNum + 1), 1)
Label8(PicNum + 8).Caption = Mid(strConCou(PicNum + 1), 2)
    Else
    PicStart(PicNum).Picture = LoadPicture(App.Path & "\Thomas.jpg")
    Label3(PicNum).Caption = PicNum + 1
    j = InStrRev(strConPic(PicNum + 1), "\")
    strTemp = Mid(strConPic(PicNum + 1), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
    Label5(PicNum + 8).Caption = strConListName(PicNum + 1)
    Label8(PicNum).Caption = Left(strConCou(PicNum + 1), 1)
Label8(PicNum + 8).Caption = Mid(strConCou(PicNum + 1), 2)
    End If
    x = PicStart(PicNum).width
    Y = PicStart(PicNum).height
    xx = x
    yy = Y
    If x > Picture2(PicNum).width Then
        Do Until x = Picture2(PicNum).width
            x = x - 1
            Y = Y - 1
        Loop
  
    End If
    Y = (x / xx) * yy
    Image1(PicNum).width = x
    Image1(PicNum).height = Y
    Image1(PicNum).Top = (Picture2(PicNum).height \ 2) - (Image1(PicNum).height \ 2)
    Image1(PicNum).Left = (Picture2(PicNum).width \ 2) - (Image1(PicNum).width \ 2)
    Image1(PicNum).Picture = PicStart(PicNum).Picture
    Image1(PicNum).Visible = True
    Else
    q = 7
    For PicNum = 0 To 7
    
    Image1(PicNum).Visible = False
    PicStart(PicNum).Picture = LoadPicture(strConPic(numStock - q))
    Label3(PicNum).Caption = numStock - q
        If intConType(numStock - q) > 2 Then
        Label3(PicNum + 8).Caption = ">>>>"
        Else
        Label3(PicNum + 8).Caption = ""
    End If
    j = InStrRev(strConPic(numStock - q), "\")
    strTemp = Mid(strConPic(numStock - q), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
    Label5(PicNum + 8).Caption = strConListName(numStock - q)
    Label8(PicNum).Caption = Left(strConCou(numStock - q), 1)
Label8(PicNum + 8).Caption = Mid(strConCou(numStock - q), 2)
    x = PicStart(PicNum).width
    Y = PicStart(PicNum).height
    xx = x
    yy = Y
    If x > Picture2(PicNum).width Then
        Do Until x = Picture2(PicNum).width
            x = x - 1
            Y = Y - 1
        Loop
    End If
    Y = (x / xx) * yy
    Image1(PicNum).width = x
    Image1(PicNum).height = Y
    Image1(PicNum).Top = (Picture2(PicNum).height \ 2) - (Image1(PicNum).height \ 2)
    Image1(PicNum).Left = (Picture2(PicNum).width \ 2) - (Image1(PicNum).width \ 2)
    Image1(PicNum).Picture = PicStart(PicNum).Picture
    Image1(PicNum).Visible = True
    q = q - 1
    Next PicNum
   End If
   ElseIf booInsert = True Then
   booInsert = False
   q = Val(Label3(0).Caption)
    For PicNum = 0 To 7
    Image1(PicNum).Visible = False
    If Not FileExists(strConPic(PicNum + 1)) Then
    
    If intConType(PicNum + 1) = 1 Or intConType(PicNum + 1) = 3 Then
    strEWX = "eng"
    Else
    strEWX = "wag"
    End If
    
    Call MakeThumbNail(strConPic(PicNum + 1), strEWX)
    DoEvents
    End If
    If FileExists(strConPic(PicNum + 1)) Then
    PicStart(PicNum).Picture = LoadPicture(strConPic(q))
    Label3(PicNum).Caption = q
        If intConType(q) > 2 Then
        Label3(PicNum + 8).Caption = ">>>>"
        Else
        Label3(PicNum + 8).Caption = ""
    End If
    j = InStrRev(strConPic(q), "\")
    strTemp = Mid(strConPic(q), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
    Label5(PicNum + 8).Caption = strConListName(q)

    Else
    PicStart(PicNum).Picture = LoadPicture(App.Path & "\Thomas.jpg")
    Label3(PicNum).Caption = q
    j = InStrRev(strConPic(q), "\")
    strTemp = Mid(strConPic(q), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
    Label5(PicNum + 8).Caption = strConListName(q)

    End If
    
  
    Label3(PicNum).Caption = q
    x = PicStart(PicNum).width
    Y = PicStart(PicNum).height
    xx = x
    yy = Y
    If x > Picture2(PicNum).width Then
        Do Until x = Picture2(PicNum).width
            x = x - 1
            Y = Y - 1
        Loop
    End If
    Y = (x / xx) * yy
    Image1(PicNum).width = x
    Image1(PicNum).height = Y
    Image1(PicNum).Top = (Picture2(PicNum).height \ 2) - (Image1(PicNum).height \ 2)
    Image1(PicNum).Left = (Picture2(PicNum).width \ 2) - (Image1(PicNum).width \ 2)
    Image1(PicNum).Picture = PicStart(PicNum).Picture
    Image1(PicNum).Visible = True
    q = q + 1
    Next PicNum
   End If
  
   If booIgnore = False Then
   For i = 0 To 6
   If Label8(i + 1).Caption = "" Then GoTo CarryON
   If Label8(i + 8).Caption = "" Then GoTo CarryON
 
   If Label8(i + 8).Caption <> Label8(i + 1).Caption Then
       ' If Label8(i + 1).Caption <> "" Or Label8(i + 8).Caption <> "" Then
        
            Select Case MsgBox("There are non-matching couplers in this consist." _
                           & vbCrLf & "Do you wish to continue?" _
                           , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)
        
            Case vbYes
        booIgnore = True
        Exit For
            Case vbNo
        Call Command5_Click
        Exit Sub
        End Select
        
        
   End If
CarryON:
   Next i
   End If
   If booActIns = True Then
   booActIns = False
   Call Command8_Click
   DoEvents
   End If
End Sub






Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If booIsAct = False Then
numIndex = Index
numItem = Val(Label3(numIndex).Caption)
PopupMenu mnuPop1
ElseIf booIsAct = True Then
numIndex = Index
numItem = Val(Label3(numIndex).Caption)
PopupMenu mnupop2
End If

End Sub


Private Sub List1_Click()

Dim strEP As String, MyString As String, strPix As String, intInd As Integer, strInd As String
Dim j As Long, jj As Long, strName As String

lblCount(1).Caption = List1.SelectedRows
List1.Select List1.row, 3
strInd = List1.Cell(flexcpText)
intInd = Val(strInd)
'If Button = 2 And intInd <> 0 Then
If intInd <> 0 Then
intPix = 1
strEP = strConEngPath(intInd) & strConEng(intInd)
If Right(strEP, 3) = "eng" Then
strPix = Left(strEP, Len(strEP) - 3) & "jpg"
If Not FileExists(strPix) Then
Call MakeThumbNail(strPix, "eng")
DoEvents
Call LoadImage2(strPix)

Else
Call LoadImage2(strPix)

End If
strMainPix = strPix
DoEvents

MyString = ReadUniFile(strEP)
Call GetEngInfo(MyString)
    j = InStr(MyString, "Name ")
    If j > 0 Then
    j = InStr(j, MyString, "(")
    jj = InStr(j, MyString, vbCrLf)
    strName = Mid(MyString, j + 1, jj - (j + 1) - 1)
    strName = Trim(strName)
    
    End If
    
    Label10.Caption = strEP
End If









End If

End Sub




Private Sub List1_DblClick()
Dim strEP As String, strPix As String, intInd As Integer, strInd As String, sLoad As Single, strLoad As String
Dim strLen As String, sLen As Integer, strFLoad As String, strFLen As String


On Error GoTo Errtrap
List1.Select List1.row, 3
strInd = List1.Cell(flexcpText)
intInd = Val(strInd)

If booInsert = False Then


strEP = strConEngPath(intInd) & strConEng(intInd)
    If Right(strEP, 3) = "eng" Then
    strPix = Left(strEP, Len(strEP) - 3) & "jpg"
    
    numStock = numStock + 1
    strConPic(numStock) = strPix
    
   intConType(numStock) = 1
   List1.Select List1.row, 4
   strLoad = List1.Cell(flexcpText)
   
   sLoad = Val(strLoad)
   engLoad = engLoad + sLoad
   strFLoad = Str(engLoad * (2205 / 2000))
   strFLoad = Format(strFLoad, "####.000")
   Label6(1).Caption = "Engine Mass: " & Str(engLoad) & " t = " & strFLoad & " U.S. or Short tons"
   
    List1.Select List1.row, 5
   strLen = List1.Cell(flexcpText)
    
   sLen = Val(strLen)
   conLen = conLen + sLen
   strFLen = Str(conLen * 3.28)
   strFLen = Format(strFLen, "####.000")
   Label6(3).Caption = "Consist Length: " & Str(conLen) & " m = " & strFLen & " ft"
     List1.Select List1.row, 6
   strConListName(numStock) = List1.Cell(flexcpText)
   List1.Select List1.row, 1
   strConCou(numStock) = List1.Cell(flexcpText)
   
    Call LoadImage(numStock)
    DoEvents
    End If
ElseIf booInsert = True Then
strEP = strConEngPath(intInd) & strConEng(intInd)
If Right(strEP, 3) = "eng" Then
strPix = Left(strEP, Len(strEP) - 3) & "jpg"
strConPic(intIndex) = strPix
intConType(intIndex) = 1
Rem *************************************************
List1.Select List1.row, 4
   strLoad = List1.Cell(flexcpText)
   
   sLoad = Val(strLoad)
   engLoad = engLoad + sLoad
   strFLoad = Str(engLoad * (2205 / 2000))
   strFLoad = Format(strFLoad, "####.000")
   Label6(1).Caption = "Engine Mass: " & Str(engLoad) & " t = " & strFLoad & " U.S. or Short tons"
   
    List1.Select List1.row, 5
   strLen = List1.Cell(flexcpText)
    
   sLen = Val(strLen)
   conLen = conLen + sLen
   strFLen = Str(conLen * 3.28)
   strFLen = Format(strFLen, "####.000")
   Label6(3).Caption = "Consist Length: " & Str(conLen) & " m = " & strFLen & " ft"

Rem ************************************************
Call LoadImage(intIndex)
DoEvents

End If
End If
Exit Sub
Errtrap:

End Sub


Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim strEP As String, MyString As String, strPix As String, intInd As Integer, strInd As String


List1.Select List1.row, 3
strInd = List1.Cell(flexcpText)
intInd = Val(strInd)
If Button = 2 And intInd <> 0 Then
intPix = 1
strEP = strConEngPath(intInd) & strConEng(intInd)
If Right(strEP, 3) = "eng" Then
strPix = Left(strEP, Len(strEP) - 3) & "jpg"
If Not FileExists(strPix) Then
Call MakeThumbNail(strPix, "eng")
DoEvents
Call LoadImage2(strPix)

Else
Call LoadImage2(strPix)

End If
strMainPix = strPix
DoEvents

MyString = ReadUniFile(strEP)
Call GetEngInfo(MyString)

End If









End If

End Sub


Private Sub List2_Click()

Dim strEP As String, MyString As String, strPix As String, intInd As Integer, strInd As String
Dim j As Long, jj As Long, strName As String

lblCount(2).Caption = List2.SelectedRows
List2.Select List2.row, 3
strInd = List2.Cell(flexcpText)
intInd = Val(strInd)

If intInd <> 0 Then
intPix = 2
strEP = strConWagPath(intInd) & strConWag(intInd)
If strConWag(intInd) = "Default.wag" Then Exit Sub
If Right(strEP, 3) = "wag" Then
strPix = Left(strEP, Len(strEP) - 3) & "jpg"
If Not FileExists(strPix) Then
Call MakeThumbNail(strPix, "wag")
DoEvents
Call LoadImage2(strPix)

Else

Call LoadImage2(strPix)
End If
strMainPix = strPix
DoEvents


MyString = ReadUniFile(strEP)
Call GetWagInfo(MyString)
   j = InStr(MyString, "Name ")
    If j > 0 Then
    j = InStr(j, MyString, "(")
    jj = InStr(j, MyString, vbCrLf)
    strName = Mid(MyString, j + 1, jj - (j + 1) - 1)
    strName = Trim(strName)
        End If
      
    Label10.Caption = strEP
End If









End If
End Sub


Private Sub List2_DblClick()
Dim strEP As String, strPix As String, intInd As Integer, strInd As String, sLoad As Single, strLoad As String
Dim strLen As String, sLen As Single, strFLoad As String, strFLen As String


On Error GoTo Errtrap
List2.Select List2.row, 3
strInd = List2.Cell(flexcpText)
intInd = Val(strInd)


If booMulti = True Then
Exit Sub
End If
If booInsert = False Then
strEP = strConWagPath(intInd) & strConWag(intInd)
    If Right(strEP, 3) = "Wag" Then
    strPix = Left(strEP, Len(strEP) - 3) & "jpg"
    numStock = numStock + 1
    strConPic(numStock) = strPix
  
    intConType(numStock) = 2
    List2.Select List2.row, 4
   strLoad = List2.Cell(flexcpText)
   sLoad = Val(strLoad)
   wagLoad = wagLoad + sLoad
    strFLoad = Str(wagLoad * (2205 / 2000))
   strFLoad = Format(strFLoad, "####.000")
   Label6(2).Caption = "Wagon Mass: " & Str(wagLoad) & " t = " & strFLoad & " U.S. or Short tons"
    List2.Select List2.row, 5
   
   strLen = List2.Cell(flexcpText)
   sLen = Val(strLen)
   conLen = conLen + sLen
   strFLen = Str(conLen * 3.28)
   strFLen = Format(strFLen, "####.000")
   Label6(3).Caption = "Consist Length: " & Str(conLen) & " m = " & strFLen & " ft"
      List2.Select List2.row, 6
   strConListName(numStock) = List2.Cell(flexcpText)
   List2.Select List2.row, 1
   strConCou(numStock) = List2.Cell(flexcpText)
   
    Call LoadImage(numStock)
    DoEvents
    End If
ElseIf booInsert = True Then
strEP = strConWagPath(intInd) & strConWag(intInd)
If Right(strEP, 3) = "Wag" Then
strPix = Left(strEP, Len(strEP) - 3) & "jpg"
strConPic(intIndex) = strPix
intConType(intIndex) = 2
Rem *****************************
    List2.Select List2.row, 4
   strLoad = List2.Cell(flexcpText)
   sLoad = Val(strLoad)
   wagLoad = wagLoad + sLoad
    strFLoad = Str(wagLoad * (2205 / 2000))
   strFLoad = Format(strFLoad, "####.000")
   Label6(2).Caption = "Wagon Mass: " & Str(wagLoad) & " t = " & strFLoad & " U.S. or Short tons"
    List2.Select List2.row, 5
   
   strLen = List2.Cell(flexcpText)
   sLen = Val(strLen)
   conLen = conLen + sLen
   strFLen = Str(conLen * 3.28)
   strFLen = Format(strFLen, "####.000")
   Label6(3).Caption = "Consist Length: " & Str(conLen) & " m = " & strFLen & " ft"

Rem *********************************
Call LoadImage(intIndex)
DoEvents

End If
End If
Exit Sub
Errtrap:

End Sub


Private Sub List2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim strEP As String, MyString As String, strPix As String, intInd As Integer, strInd As String

List2.Select List2.row, 3
strInd = List2.Cell(flexcpText)
intInd = Val(strInd)

If Button = 2 And intInd <> 0 Then
intPix = 2
strEP = strConWagPath(intInd) & strConWag(intInd)
If strConWag(intInd) = "Default.wag" Then Exit Sub
If Right(strEP, 3) = "wag" Then
strPix = Left(strEP, Len(strEP) - 3) & "jpg"
If Not FileExists(strPix) Then
Call MakeThumbNail(strPix, "wag")
DoEvents
Call LoadImage2(strPix)
'Call LoadImage2(App.Path & "\" & "thomas.jpg")
'Picture1.Picture = LoadPicture(App.Path & "\" & "thomas.jpg")
Else
'Picture1.Picture = LoadPicture(strPix)
Call LoadImage2(strPix)
End If
strMainPix = strPix
DoEvents


MyString = ReadUniFile(strEP)
Call GetWagInfo(MyString)

End If









End If
End Sub

Private Sub mnuActDel_Click()
Dim i As Integer, PicNum As Integer, ii As Integer, j As Integer

Set Picture2(numIndex) = LoadPicture("")
Set Image1(numIndex) = LoadPicture("")
Set PicStart(numIndex) = LoadPicture("")
Label5(numIndex).Caption = ""
Label5(numIndex + 8).Caption = ""
Label8(numIndex).Caption = ""
Label8(numIndex + 8).Caption = ""
numStock = numStock - 1

If numIndex < numStock Then
For i = numIndex + 1 To numStock
strConPic(i) = strConPic(i + 1)
intConType(i) = intConType(i + 1)
strConListName(i) = strConListName(i + 1)
Next i

strConPic(i + 1) = ""
strConListName(i + 1) = ""
strConItem(i + 1) = ""
strConFolder(i + 1) = ""
strConUid(i + 1) = ""
intConType(i + 1) = 0
End If
ii = Val(Label3(numIndex).Caption)
i = numStock - ii

If numIndex + i < 7 Then
i = numIndex + i
Else
i = 7
End If
For PicNum = numIndex To i
Image1(PicNum).Visible = False
    PicStart(PicNum).Picture = LoadPicture(strConPic(ii))
    j = InStrRev(strConPic(ii), "\")
   
    strTemp = Mid(strConPic(ii), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
     Label5(PicNum + 8).Caption = strConListName(ii)
     Label8(PicNum).Caption = Left(strConCou(ii), 1)
Label8(PicNum + 8).Caption = Mid(strConCou(ii), 2)
     If PicNum < 7 Then
     Label3(PicNum + 8).Caption = Label3(PicNum + 8 + 1).Caption
     End If
    x = PicStart(PicNum).width
    Y = PicStart(PicNum).height
    xx = x
    yy = Y
    If x > Picture2(PicNum).width Then
        Do Until x = Picture2(PicNum).width
            x = x - 1
            Y = Y - 1
        Loop
    End If
    Y = (x / xx) * yy
    Image1(PicNum).width = x
    Image1(PicNum).height = Y
    Image1(PicNum).Top = (Picture2(PicNum).height \ 2) - (Image1(PicNum).height \ 2)
    Image1(PicNum).Left = (Picture2(PicNum).width \ 2) - (Image1(PicNum).width \ 2)
    Image1(PicNum).Picture = PicStart(PicNum).Picture
    Image1(PicNum).Visible = True
    ii = ii + 1
    Next PicNum
For j = i + 1 To 7
Set Picture2(j) = LoadPicture("")
Set Image1(j) = LoadPicture("")
Set PicStart(j) = LoadPicture("")
Label5(j).Caption = ""
Label5(j + 8).Caption = ""
Label8(j).Caption = ""
Label8(j + 8).Caption = ""
Label3(j + 8).Caption = ""
Next j

Call Command8_Click
DoEvents

End Sub

Private Sub mnuActIns_Click()
Dim i As Integer

On Error GoTo Errtrap
For i = numStock + 1 To numIndex + 1 Step -1
If i - 1 = 0 Then Exit For
strConPic(i) = strConPic(i - 1)
intConType(i) = intConType(i - 1)
strConListName(i) = strConListName(i - 1)
strConCou(i) = strConCou(i - 1)
Next i
Set Picture2(numIndex) = LoadPicture("")
Set Image1(numIndex) = LoadPicture("")
Set PicStart(numIndex) = LoadPicture("")
Label5(numIndex).Caption = ""
Label5(numIndex + 8).Caption = ""
Label8(PicNum).Caption = ""
Label8(PicNum + 8).Caption = ""
numStock = numStock + 1
booInsert = True
intIndex = numIndex + 1
booActIns = True
Call MsgBox("Please double-click on the stock item you wish to insert", vbInformation, App.Title)
Exit Sub
Errtrap:
Resume Next
End Sub

Private Sub mnuDel_Click()
Dim i As Integer, PicNum As Integer, ii As Integer, j As Integer

Set Picture2(numIndex) = LoadPicture("")
Set Image1(numIndex) = LoadPicture("")
Set PicStart(numIndex) = LoadPicture("")
Label5(numIndex).Caption = ""
Label5(numIndex + 8).Caption = ""
Label8(numIndex).Caption = ""
Label8(numIndex + 8).Caption = ""
numStock = numStock - 1

If numItem < numStock Then
For i = numItem To numStock
strConPic(i) = strConPic(i + 1)
intConType(i) = intConType(i + 1)
strConListName(i) = strConListName(i + 1)
strConCou(i) = strConCou(i + 1)
Next i

strConPic(i) = ""
strConListName(i) = ""
strConItem(i) = ""
strConFolder(i) = ""
strConUid(i) = ""
intConType(i) = 0
End If
ii = numItem  'Val(Label3(numIndex).Caption)
i = numStock - ii

If numIndex + i < 7 Then
i = numIndex + i
Else
i = 7
End If
For PicNum = numIndex To i
Image1(PicNum).Visible = False
    PicStart(PicNum).Picture = LoadPicture(strConPic(ii))
    j = InStrRev(strConPic(ii), "\")
   
    strTemp = Mid(strConPic(ii), j + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
     Label5(PicNum + 8).Caption = strConListName(ii)
     Label8(PicNum).Caption = Left(strConCou(ii), 1)
Label8(PicNum + 8).Caption = Mid(strConCou(ii), 2)
     If PicNum < 7 Then
     Label3(PicNum + 8).Caption = Label3(PicNum + 8 + 1).Caption
     End If
    x = PicStart(PicNum).width
    Y = PicStart(PicNum).height
    xx = x
    yy = Y
    If x > Picture2(PicNum).width Then
        Do Until x = Picture2(PicNum).width
            x = x - 1
            Y = Y - 1
        Loop
    End If
    Y = (x / xx) * yy
    Image1(PicNum).width = x
    Image1(PicNum).height = Y
    Image1(PicNum).Top = (Picture2(PicNum).height \ 2) - (Image1(PicNum).height \ 2)
    Image1(PicNum).Left = (Picture2(PicNum).width \ 2) - (Image1(PicNum).width \ 2)
    Image1(PicNum).Picture = PicStart(PicNum).Picture
    Image1(PicNum).Visible = True
    ii = ii + 1
    Next PicNum
For j = i + 1 To 7
Set Picture2(j) = LoadPicture("")
Set Image1(j) = LoadPicture("")
Set PicStart(j) = LoadPicture("")
Label5(j).Caption = ""
Label5(j + 8).Caption = ""
Label8(j).Caption = ""
Label8(j + 8).Caption = ""
Label3(j + 8).Caption = ""
Next j




End Sub


Private Sub mnuFlip_Click()


Select Case MsgBox("Flip is not recommended (Unless you are using MSTSBin) as bogies on flipped" _
                   & vbCrLf & "stock may rotate in the wrong direction. Preferably use models" _
                   & vbCrLf & "built to run backwards. Only use this option to finalise your consist." _
                   , vbOKCancel Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbOK

intConType(numIndex + 1) = intConType(numIndex + 1) + 2
Label3(numIndex + 8) = ">>>>"
    Case vbCancel
Exit Sub
End Select

End Sub


Private Sub mnuInsert_Click()
Dim i As Integer

On Error GoTo Errtrap

For i = numStock + 1 To numIndex + 1 Step -1
If i - 1 = 0 Then Exit For
strConPic(i) = strConPic(i - 1)
intConType(i) = intConType(i - 1)
strConListName(i) = strConListName(i - 1)
strConCou(i) = strConCou(i - 1)
Next i

Set Picture2(numIndex) = LoadPicture("")
Set Image1(numIndex) = LoadPicture("")
Set PicStart(numIndex) = LoadPicture("")
Label5(numIndex).Caption = ""
Label5(numIndex + 8).Caption = ""
Label8(PicNum).Caption = ""
Label8(PicNum + 8).Caption = ""
numStock = numStock + 1
booInsert = True

intIndex = Val(Label3(numIndex).Caption)
Call MsgBox("Please double-click on the stock item you wish to insert", vbInformation, App.Title)
Exit Sub
Errtrap:
Resume Next

End Sub


Private Sub mnuMulti_Click()
Dim result As Integer, strTemp As String, strEP As String, j As Integer, i As Integer, intInd As Integer, strInd As String
Dim strCoupler As String, strWagType As String, strMass As String, strLen As String, strName As String
Dim sLen As Single, sLoad As Single, strFLen As String, strFLoad As String

On Error GoTo Errtrap
List2.Select List2.row, 3
strInd = List2.Cell(flexcpText)
intInd = Val(strInd)
booConMulti = True
strEP = strConWagPath(intInd) & strConWag(intInd)
DoEvents
strTemp = InputBox("How many of " & strConWag(intInd) & " do you wish to add?", "Enter Number of required item")
result = Val(strTemp)
DoEvents
Call GetWagDetail(strEP, strCoupler, strWagType, strMass, strLen, strName)

    If Right(strEP, 3) = "Wag" Then
    strPix = Left(strEP, Len(strEP) - 3) & "jpg"
    j = numStock + 1
    numStock = numStock + result

    For i = j To numStock
    strConPic(i) = strPix
    intConType(i) = 2
    strConListName(i) = strName
    sLen = Val(strLen)
   conLen = conLen + sLen
   strFLen = Str(conLen * 3.28)
   strFLen = Format(strFLen, "####.000")
   Label6(3).Caption = "Consist Length: " & Str(conLen) & " m = " & strFLen & " ft"
    sLoad = Val(strMass)
   wagLoad = wagLoad + sLoad
    strFLoad = Str(wagLoad * (2205 / 2000))
   strFLoad = Format(strFLoad, "####.000")
   Label6(2).Caption = "Wagon Mass: " & Str(wagLoad) & " t = " & strFLoad & " U.S. or Short tons"
    
    
Next i
DoEvents
intIndex = numIndex + 1
 Call LoadImage(j)
    DoEvents
    Call Command3_Click(5)
    
End If
Exit Sub
Errtrap:
If Err = 9 Then
Call MsgBox("You have not chosen a wagon type to include? Select one from the Wagons list.", vbExclamation, App.Title)

Exit Sub
End If

Resume Next
End Sub


Private Sub mnuOCF_Click()
Dim strCFile As String, MyString As String, i As Integer, x As Long, xx As Long, Y As Long, yy As Long
Dim strTemp As String, strCN As String, strCF As String, j As Integer
Dim Entry As String, strCoupler As String, strPower As String, strMass As String, strLen As String
Dim strWagType As String, sLen As Single, sLoad As Single, strName As String, strFLoad As String, strFLen As String
Dim strStock As String

Call ClearAll

CDL1.Filter = "Consist Files (*.con|*.con"
CDL1.DialogTitle = "Select Document to Load"
CDL1.InitDir = MSTSPath & "\Trains\Consists\"
CDL1.FilterIndex = 1
CDL1.Action = 1
strCFile = CDL1.Filename
booIsAct = False
If strCFile = vbNullString Then Exit Sub
Command2.Enabled = True
Command8.Enabled = False
Command9.Enabled = False
MyString = ReadUniFile(strCFile)
MyString = Replace(MyString, "  ", " ")

x = InStr(MyString, "TrainCfg")
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, vbCr)
strTemp = Mid(MyString, x + 1, xx - (x + 1))
strTemp = Trim(strTemp)
If Left(strTemp, 1) = ChrW$(34) Then
strTemp = Mid(strTemp, 2)
End If
If Right(strTemp, 1) = ChrW$(34) Then
strTemp = Left(strTemp, Len(strTemp) - 1)
End If
Text1(0) = strTemp
x = InStr(MyString, "Name (")
If x > 0 Then
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp = Mid(MyString, x + 1, xx - (x + 1))
strTemp = Trim(strTemp)
If Left(strTemp, 1) = ChrW$(34) Then
strTemp = Mid(strTemp, 2)
End If
If Right(strTemp, 1) = ChrW$(34) Then
strTemp = Left(strTemp, Len(strTemp) - 1)
End If
Text1(1) = strTemp
End If
x = InStr(MyString, "Serial (")
If x > 0 Then
x = InStr(x, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp = Mid(MyString, x + 1, xx - (x + 1))
strTemp = Trim(strTemp)
intConSerial = Val(strTemp)
End If

numStock = 0
x = 1
Do
x = InStr(x + 1, MyString, "Data (")
If x > 0 Then
numStock = numStock + 1
End If
Loop While x > 0

x = 1
For i = 1 To numStock
If i > numStock Then GoTo CarryON
x = InStr(x + 1, MyString, "Data (")
strTemp = Mid(MyString, x - 5, 5)
If strTemp = "ngine" Then
intConType(i) = 1
Y = InStrRev(MyString, "Flip", x)
If Y > 0 And (x - Y) < 25 Then intConType(i) = 3
ElseIf strTemp = "Wagon" Then
intConType(i) = 2
Y = InStr(x, MyString, ")")
yy = InStr(Y, MyString, "Flip")
If yy > 0 And (yy - Y) < 25 Then intConType(i) = 4

End If

x = InStr(x, MyString, "(")
'xx = InStr(X, MyString, ")")
xx = InStr(x, MyString, vbCr)
strTemp = Mid(MyString, x + 1, (xx - 1) - (x + 1))
strTemp = Trim(strTemp)
j = InStr(strTemp, ChrW$(34))
If j = 0 Then   'No quotes
j = InStr(strTemp, " ")
strCN = Left(strTemp, j - 1)
strCF = Mid(strTemp, j + 1)
ElseIf j = 1 Then 'quotes around stock name
j = InStr(j + 1, strTemp, ChrW$(34))
strCN = Left(strTemp, j - 1)
strCF = Mid(strTemp, j + 2)
strCN = Trim(strCN)
strCF = Trim(strCF)
Else
strCN = Left(strTemp, j - 1)
strCF = Mid(strTemp, j + 1)
strCN = Trim(strCN)
strCF = Trim(strCF)
End If

If Left(strCN, 1) = ChrW$(34) Then
strCN = Mid(strCN, 2)
End If
If Right(strCN, 1) = ChrW$(34) Then
strCN = Left(strCN, Len(strCN) - 1)
End If
If Left(strCF, 1) = ChrW$(34) Then
strCF = Mid(strCF, 2)
End If
If Right(strCF, 1) = ChrW$(34) Then
strCF = Left(strCF, Len(strCF) - 1)
End If
If intConType(i) = 1 Or intConType(i) = 3 Then
strTemp = ".eng"
Else
strTemp = ".wag"
End If
strStock = MSTSPath & "\Trains\Trainset\" & strCF & "\" & strCN & strTemp
If Not FileExists(strStock) Then

Call MsgBox(strCF & "\" & strCN & strTemp & " listed in this consist appears to be missing?" _
    & vbCrLf & "Missing item was in position " & Str(i) & ". You can resave the " _
    & vbCrLf & "consist without the missing items, or add replacements as required." _
    , vbExclamation, App.Title)
i = i - 1
numStock = numStock - 1

GoTo GetAnother
End If

strConPic(i) = MSTSPath & "\Trains\Trainset\" & strCF & "\" & strCN & ".jpg"
If Not FileExists(MSTSPath & "\Trains\Trainset\" & strCF & "\" & strCN & ".jpg") Then
strConPic(i) = App.Path & "\Thomas.jpg"
End If
If intConType(i) = 1 Then
Entry = MSTSPath & "\Trains\Trainset\" & strCF & "\" & strCN & ".eng"
strCoupler = ""
strPower = ""
strMass = ""
strLen = ""
strName = ""
Call GetEngDetail(Entry, strCoupler, strPower, strMass, strLen, strName)
strConListName(i) = strName
strConCou(i) = strCoupler
 sLen = Val(strLen)
   conLen = conLen + sLen
   strFLen = Str(conLen * 3.28)
   strFLen = Format(strFLen, "####.000")
   Label6(3).Caption = "Consist Length: " & Str(conLen) & " m = " & strFLen & " ft"
  ' Label6(3).Caption = "Consist Length: " & Str(conLen) & " m"
    sLoad = Val(strMass)
   engLoad = engLoad + sLoad
    strFLoad = Str(engLoad * (2205 / 2000))
   strFLoad = Format(strFLoad, "####.000")
   Label6(1).Caption = "Engine Mass: " & Str(engLoad) & " t = " & strFLoad & " U.S. or Short tons"
   
  ' Label6(1).Caption = "Engine Mass: " & Str(engLoad) & " t"
ElseIf intConType(i) = 2 Then
Entry = MSTSPath & "\Trains\Trainset\" & strCF & "\" & strCN & ".wag"
strCoupler = ""
strPower = ""
strMass = ""
strLen = ""
strName = ""
Call GetWagDetail(Entry, strCoupler, strWagType, strMass, strLen, strName)
strConListName(i) = strName
strConCou(i) = strCoupler
 sLen = Val(strLen)
   conLen = conLen + sLen
   strFLen = Str(conLen * 3.28)
   strFLen = Format(strFLen, "####.000")
   Label6(3).Caption = "Consist Length: " & Str(conLen) & " m = " & strFLen & " ft"
   
  ' Label6(3).Caption = "Consist Length: " & Str(conLen) & " m"
    sLoad = Val(strMass)
   wagLoad = wagLoad + sLoad
   strFLoad = Str(wagLoad * (2205 / 2000))
   strFLoad = Format(strFLoad, "####.000")
   Label6(2).Caption = "Wagon Mass: " & Str(wagLoad) & " t = " & strFLoad & " U.S. or Short tons"
  ' Label6(2).Caption = "Wagon Mass: " & Str(wagLoad) & " t"
End If
GetAnother:
Next i
DoEvents
CarryON:
Call LoadImage(numStock)
DoEvents
Call Command3_Click(5)
DoEvents

End Sub

Private Sub mnuOpenAct_Click()
Dim MyString As String, i As Integer, x As Long
Dim strTemp As String

On Error GoTo Errtrap
Call ClearAll
ActNum = 0
CDL1.Filter = "Activity Files (*.act|*.act"
CDL1.DialogTitle = "Select Document to Load"
CDL1.InitDir = MSTSPath & "\Routes\"
CDL1.FilterIndex = 1
CDL1.Action = 1
strCFile = CDL1.Filename
If strCFile = vbNullString Then Exit Sub
MyString = ReadUniFile(strCFile)
Command2.Enabled = False
Command8.Enabled = True
Command9.Enabled = True
i = InStrRev(strCFile, "\")
strActFile = Mid(strCFile, i + 1)
Text1(0) = strActFile
x = 1
Do
x = InStr(x, MyString, "TrainCfg")
If x = 0 Then Exit Do
ActNum = ActNum + 1
x = x + 1
Loop
Call MsgBox("This activity includes " & Str(ActNum) & " loose consists.", vbExclamation, App.Title)
If ActNum = 0 Then Exit Sub
ReDim strLoose(1 To ActNum)
x = InStr(MyString, "ActivityObject (")
strActStart = Left(MyString, x - 1)

strTemp = Mid(MyString, x)

For i = 1 To ActNum
x = InStr(2, strTemp, "ActivityObject (")
If x = 0 Then Exit For
strLoose(i) = Left(strTemp, x - 1)
strTemp = Mid(strTemp, x)
Next i
strLoose(i) = strTemp
x = InStr(strLoose(i), "Tile")
x = InStr(x, strLoose(i), ")")
x = InStr(x + 1, strLoose(i), ")")
strLoose(i) = Left(strLoose(i), x + 2)
strActEnd = Mid(strTemp, x + 3)

Command7.Visible = True
Text4.Visible = True
Text4.Text = "1"
booIsAct = True
Exit Sub
Errtrap:

Resume Next
End Sub


Private Sub mnuSwap_Click()
Dim i As Integer, strTemp As String, strTemp2 As String, j As Integer, jj As Integer, jq As Integer

strTemp = strConPic(numIndex + 1)
strTemp2 = strConPic(numIndex + 2)
strConPic(numIndex + 1) = strTemp2
strConPic(numIndex + 2) = strTemp
j = intConType(numIndex + 1)
jj = intConType(numIndex + 2)
intConType(numIndex + 1) = jj
intConType(numIndex + 2) = j
strTemp = strConListName(numIndex + 1)
strTemp2 = strConListName(numIndex + 2)
strConListName(numIndex + 1) = strTemp2
strConListName(numIndex + 2) = strTemp


ii = Val(Label3(0).Caption)
i = numStock - ii

If numIndex + i < 7 Then
i = numIndex + i
Else
i = 7
End If
For PicNum = 0 To i
Image1(PicNum).Visible = False
    PicStart(PicNum).Picture = LoadPicture(strConPic(ii))
    jq = InStrRev(strConPic(ii), "\")
    strTemp = Mid(strConPic(ii), jq + 1)
    strTemp = Left(strTemp, Len(strTemp) - 4)
    Label5(PicNum).Caption = strTemp
     Label5(PicNum + 8).Caption = strConListName(ii)
     Label8(PicNum).Caption = Left(strConCou(ii), 1)
Label8(PicNum + 8).Caption = Mid(strConCou(ii), 2)
    x = PicStart(PicNum).width
    Y = PicStart(PicNum).height
    xx = x
    yy = Y
    If x > Picture2(PicNum).width Then
        Do Until x = Picture2(PicNum).width
            x = x - 1
            Y = Y - 1
        Loop
    End If
    Y = (x / xx) * yy
    Image1(PicNum).width = x
    Image1(PicNum).height = Y
    Image1(PicNum).Top = (Picture2(PicNum).height \ 2) - (Image1(PicNum).height \ 2)
    Image1(PicNum).Left = (Picture2(PicNum).width \ 2) - (Image1(PicNum).width \ 2)
    Image1(PicNum).Picture = PicStart(PicNum).Picture
    Image1(PicNum).Visible = True
    ii = ii + 1
    Next PicNum

End Sub

Private Sub SP_FileMatch(Filename As String, Path As String)
Dim Entry As String, x As Long, j As Integer, strType As String, strType2 As String, strName As String
Dim strCoupler As String, strPower As String, strWagType As String, strMass As String, strLen As String
Dim sMass As Single, sLen As Single, i As Integer

On Error GoTo Errtrap
Entry = Path & Filename
If Filename = "Default.wag" Then Exit Sub
i = InStr(Entry, "\cabview")
If i > 0 Then Exit Sub
i = InStr(Entry, "\sound")
If i > 0 Then Exit Sub
i = InStr(Entry, "\backup")
If i > 0 Then Exit Sub
i = InStr(Entry, "\Common.")
If i > 0 Then Exit Sub
DoEvents
j = InStrRev(strFilter, ".")
If j > 0 Then
strType = Mid(strFilter, j + 1)
If strType = "*" Then GoTo CarryON
j = InStrRev(Filename, ".")
strType2 = Mid(Filename, j + 1)
If strType <> strType2 Then GoTo SkipThis
CarryON:
End If

If booConEng = True Then
j = InStr(Path, "\Stored")
If j > 0 Then GoTo SkipThis
j = InStr(Path, "\BackUp")
If j > 0 Then GoTo SkipThis

If Right$(Path, 7) = "Stored\" Then GoTo SkipThis
        If Right$(Path, 9) = "Backup\" Then GoTo SkipThis

        If Right$(Path, 14) = "SpareConsists\" Then GoTo SkipThis
        If Right$(Path, 9) = "SpareCon\" Then GoTo SkipThis
        If Right$(Path, 7) = "Spares\" Then GoTo SkipThis
        If frmUtils.Check1.value = 0 Then
        x = InStrRev(Path, "\", Len(Path) - 1)
        If Mid$(Path, x + 1, 7) = "Cabview" Then
        GoTo SkipThis
        End If
        End If
        Call GetEngDetail(Entry, strCoupler, strPower, strMass, strLen, strName)
        j = InStr(strMass, "#")
        If j > 0 Then
        strMass = Left(strMass, j - 1)
        strMass = Trim(strMass)
        End If
        If Left(strMass, 1) = ChrW$(34) Then
        strMass = Mid(strMass, 2)
        End If
        If Right(strMass, 1) = ChrW$(34) Then
        strMass = Left(strMass, Len(strMass) - 1)
        End If
        If IsNumeric(Right(strMass, 1)) Then
        strMass = strMass & "t"
        End If
        If Right(strMass, 1) <> "t" Then
        If Right(strMass, 2) = "lb" Then
        sMass = Val(strMass) / 2205
        strMass = Str(sMass) & "t"
        End If
        End If
        j = InStr(strLen, "#")
        If j > 0 Then
        strLen = Left(strLen, j - 1)
        strLen = Trim(strLen)
        End If
        If Left(strLen, 1) = ChrW$(34) Then
        strLen = Mid(strLen, 2)
        End If
        If Right(strLen, 1) = ChrW$(34) Then
        strLen = Left(strLen, Len(strLen) - 1)
        End If
        
        If IsNumeric(Right(strLen, 1)) Then
        strLen = strLen & "m"
        End If
        If Right(strLen, 1) <> "m" Then
        If Right(strLen, 2) = "ft" Then
        sLen = Val(strLen) / 3.28
        strLen = Str(sLen) & "m"
        End If
        End If
        
        
            List1.AddItem Filename & vbTab & strCoupler & vbTab & strPower & vbTab & numEng & vbTab & strMass & vbTab & strLen & vbTab & strName
            strConEng(numEng) = Filename
            strConEngPath(numEng) = Path
            strConEngCou(numEng) = strCoupler
            numEng = numEng + 1
            If numEng > UBound(strConEng) Then
           ReDim Preserve strConEng(1 To numEng + Wag_CHUNK)
           ReDim Preserve strConEngPath(1 To numEng + Wag_CHUNK)
           ReDim Preserve strConEngCou(1 To numEng + Wag_CHUNK)
            End If
            lblCount(0).Caption = Str(Val(lblCount(0).Caption) + 1)
            DoEvents
            If booAbort = True Then
            SearchFlag = False
            Exit Sub
            End If
            
ElseIf booConWag = True Then
j = InStr(Path, "\Stored")
If j > 0 Then GoTo SkipThis
j = InStr(Path, "\BackUp")
If j > 0 Then GoTo SkipThis

If Right$(Path, 7) = "Stored\" Then GoTo SkipThis
        If Right$(Path, 9) = "Backup\" Then GoTo SkipThis

        If Right$(Path, 14) = "SpareConsists\" Then GoTo SkipThis
        If Right$(Path, 9) = "SpareCon\" Then GoTo SkipThis
        If Right$(Path, 7) = "Spares\" Then GoTo SkipThis
        If frmUtils.Check1.value = 0 Then
        x = InStrRev(Path, "\", Len(Path) - 1)
        If Mid$(Path, x + 1, 7) = "Cabview" Then
        GoTo SkipThis
        End If
        End If
        Call GetWagDetail(Entry, strCoupler, strWagType, strMass, strLen, strName)
                j = InStr(strMass, "#")
        If j > 0 Then
        strMass = Left(strMass, j - 1)
        strMass = Trim(strMass)
        End If
        If Left(strMass, 1) = ChrW$(34) Then
        strMass = Mid(strMass, 2)
        End If
        If Right(strMass, 1) = ChrW$(34) Then
        strMass = Left(strMass, Len(strMass) - 1)
        End If
        If IsNumeric(Right(strMass, 1)) Then
        strMass = strMass & "t"
        End If
        If Right(strMass, 1) <> "t" Then
        If Right(strMass, 2) = "lb" Then
        sMass = Val(strMass) / 2205
        strMass = Str(sMass) & "t"
        End If
        End If
        j = InStr(strLen, "#")
        If j > 0 Then
        strLen = Left(strLen, j - 1)
        strLen = Trim(strLen)
        End If
        If Left(strLen, 1) = ChrW$(34) Then
        strLen = Mid(strLen, 2)
        End If
        If Right(strLen, 1) = ChrW$(34) Then
        strLen = Left(strLen, Len(strLen) - 1)
        End If
       ' If Not IsNumeric(Right(strLen, 1)) And Right(strLen, 1) <> "m" Then
            List2.AddItem Filename & vbTab & strCoupler & vbTab & strWagType & vbTab & numWag & vbTab & strMass & vbTab & strLen & vbTab & strName
            strConWag(numWag) = Filename
            strConWagPath(numWag) = Path
            strConWagCou(numWag) = strCoupler
            numWag = numWag + 1
            If numWag > UBound(strConWag) Then
           ReDim Preserve strConWag(1 To numWag + Wag_CHUNK)
           ReDim Preserve strConWagPath(1 To numWag + Wag_CHUNK)
           ReDim Preserve strConWagCou(1 To numWag + Wag_CHUNK)
            End If
            lblCount(3).Caption = Str(Val(lblCount(3).Caption) + 1)
            DoEvents
            If booAbort = True Then
            SearchFlag = False
            Exit Sub
            End If
End If
SkipThis:
Exit Sub
Errtrap:
Call MsgBox("An error occurred in SP_Filematch while processing " & Filename, vbExclamation, App.Title)


End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim MyRow As Long

If KeyCode = 13 Then
MyRow = List1.FindRow(Text2, 1, 0, False, False)
If MyRow = -1 Then
Call MsgBox("Search did not find that item.", vbExclamation, App.Title)

Exit Sub
End If
List1.ShowCell MyRow, 0
DoEvents
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim MyRow As Long

If KeyCode = 13 Then
MyRow = List2.FindRow(Text3, 1, 0, False, False)
If MyRow = -1 Then
Call MsgBox("Search did not find that item.", vbExclamation, App.Title)

Exit Sub
End If
List2.ShowCell MyRow, 0
DoEvents
End If
End Sub


Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call Command7_Click
End If
End Sub


