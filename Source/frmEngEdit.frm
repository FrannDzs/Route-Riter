VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmEngEdit 
   Caption         =   "Engine/Wagon Editor"
   ClientHeight    =   11250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11250
   ScaleWidth      =   15210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command30 
      Caption         =   "Add Shunting to selected"
      Height          =   495
      Left            =   12480
      TabIndex        =   88
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Disable Shunting"
      Height          =   495
      Left            =   12480
      TabIndex        =   87
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Enable Shunting"
      Height          =   495
      Left            =   12480
      TabIndex        =   86
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Correct COG"
      Height          =   495
      Left            =   5040
      TabIndex        =   85
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Use Formulae.txt"
      Height          =   495
      Left            =   4800
      TabIndex        =   84
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   82
      Top             =   3840
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Update Friction"
      Height          =   495
      Left            =   11040
      TabIndex        =   81
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Pro Throttle / Brake %"
      Height          =   495
      Left            =   11040
      TabIndex        =   80
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Use 24RL Brake System"
      Height          =   495
      Left            =   7920
      TabIndex        =   79
      Top             =   7680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Correct Brakes Pro-Version"
      Height          =   495
      Left            =   9480
      TabIndex        =   78
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Remove certain files from list"
      Height          =   495
      Left            =   6360
      TabIndex        =   77
      ToolTipText     =   "Removes items containing certain words from list"
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Delete Selected"
      Height          =   495
      Left            =   11040
      TabIndex        =   76
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Backups"
      Height          =   375
      Index           =   7
      Left            =   10920
      TabIndex        =   75
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "All Rolling-stock"
      Height          =   375
      Index           =   6
      Left            =   1800
      TabIndex        =   74
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Correct Couplers Automatically"
      Height          =   495
      Left            =   7920
      TabIndex        =   73
      ToolTipText     =   "Corrects Couplers Automatically based on Size using TurboBills settings"
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Remove all Lights from selected"
      Height          =   495
      Left            =   9480
      TabIndex        =   72
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Correct Brakes Standard"
      Height          =   495
      Left            =   9480
      TabIndex        =   71
      Top             =   7080
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   4800
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Correct Couplers Manually"
      Height          =   495
      Left            =   6360
      TabIndex        =   70
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Correct Mass and Couplings            Only"
      Height          =   615
      Left            =   4800
      TabIndex        =   67
      Top             =   6360
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Tweak Settings"
      Height          =   495
      Left            =   7920
      TabIndex        =   46
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Settings of Currently Selected Item"
      Height          =   2415
      Left            =   240
      TabIndex        =   33
      Top             =   8280
      Width           =   13095
      Begin VB.CommandButton Command26 
         Caption         =   "Update using Formula"
         Height          =   495
         Left            =   240
         TabIndex        =   83
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Equalise BBox"
         Height          =   495
         Left            =   1560
         TabIndex        =   69
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdApplyAll 
         Caption         =   "Apply to all selected files"
         Height          =   495
         Left            =   7920
         TabIndex        =   68
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton Command14 
         Caption         =   "View BB"
         Height          =   495
         Left            =   6600
         TabIndex        =   66
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   5400
         TabIndex        =   65
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command12 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   63
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   62
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   5400
         TabIndex        =   60
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Update"
         Height          =   495
         Left            =   5160
         TabIndex        =   59
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   12360
         TabIndex        =   58
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   12600
         TabIndex        =   57
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   11040
         TabIndex        =   56
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   10800
         TabIndex        =   55
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   5400
         TabIndex        =   54
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   5400
         TabIndex        =   53
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   12000
         TabIndex        =   52
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   11160
         TabIndex        =   51
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   10440
         TabIndex        =   50
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   9720
         TabIndex        =   49
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   7920
         TabIndex        =   44
         Top             =   1440
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   7920
         TabIndex        =   43
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7920
         TabIndex        =   42
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   41
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   40
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   39
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label10 
         Caption         =   "Change Z"
         Height          =   255
         Left            =   3360
         TabIndex        =   64
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Selected item"
         Height          =   375
         Left            =   4200
         TabIndex        =   61
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Change BBox +Z"
         Height          =   375
         Index           =   1
         Left            =   11520
         TabIndex        =   48
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Change BBox -Z"
         Height          =   375
         Index           =   0
         Left            =   9960
         TabIndex        =   47
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "ESD_Bounding_Box"
         Height          =   255
         Index           =   5
         Left            =   6360
         TabIndex        =   45
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "DerailBufferForce"
         Height          =   255
         Index           =   4
         Left            =   6360
         TabIndex        =   38
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "DerailRailForce"
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   37
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "InertiaTensor"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Mass"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Size"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Correct Settings"
      Height          =   495
      Left            =   6360
      TabIndex        =   32
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Make Backups of Altered Files"
      Height          =   615
      Left            =   4800
      TabIndex        =   31
      Top             =   5640
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Remove AI/Dead Eng from List"
      Height          =   495
      Left            =   9480
      TabIndex        =   30
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "If Key appears multiple times:-"
      Height          =   1935
      Left            =   600
      TabIndex        =   25
      Top             =   6120
      Width           =   3975
      Begin VB.CheckBox Check1 
         Caption         =   "Change or Delete in item shown above"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   3495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Change or Delete All Instances"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Change or Delete 2nd Instance"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Change or Delete 1st instance only"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Value           =   -1  'True
         Width           =   3375
      End
   End
   Begin VB.ComboBox cbInsert 
      Height          =   1350
      Left            =   600
      Style           =   1  'Simple Combo
      TabIndex        =   23
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   600
      TabIndex        =   21
      Top             =   3960
      Width           =   3855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      Caption         =   "Clear List"
      Height          =   375
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Unselect All"
      Height          =   495
      Left            =   7920
      TabIndex        =   19
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Select All"
      Height          =   495
      Left            =   6360
      TabIndex        =   18
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Insert Key"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      ToolTipText     =   "Insert this key BEFORE  the following Key in all selected files"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Change Value"
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      ToolTipText     =   "Change the parameter value for this key in all selected files"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Key"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      ToolTipText     =   "Delete the selected Key from all selected files"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ComboBox cbEng 
      Height          =   2130
      Left            =   480
      Style           =   1  'Simple Combo
      TabIndex        =   13
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tender"
      Height          =   375
      Index           =   5
      Left            =   9720
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Carriage"
      Height          =   375
      Index           =   4
      Left            =   8400
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Freight"
      Height          =   375
      Index           =   3
      Left            =   7080
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Electric Loco"
      Height          =   375
      Index           =   2
      Left            =   5760
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Diesel Loco"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   4920
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   8295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Steam Loco"
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Insert Before Line:-"
      Height          =   255
      Left            =   1320
      TabIndex        =   24
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "New Value"
      Height          =   255
      Left            =   1680
      TabIndex        =   22
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Select Key"
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   13800
      TabIndex        =   11
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Total Number of Files in List."
      Height          =   495
      Left            =   13560
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Number of Selected Files"
      Height          =   495
      Left            =   13680
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   13800
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Select Type of Rolling Stock"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmEngEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Dim EngIgnore() As String

Dim intIgnore As Integer
Dim strSize As String
Dim strInertia As String
Dim strSD As String
Dim MyString As String
Dim strMinusZ As String, strPlusZ As String, strBound As String
Dim strSDFile As String
Dim booEdited As Boolean
Dim fullpath$
Dim strNewSize As String
Dim flagType As Integer
Dim strParam(1 To 20) As String
Dim strFormula(1 To 20, 1 To 6) As String
Dim numFormula As Integer
Dim numItems As Integer
Private Function CheckDefaultWag()
Dim strWag As String, MyString As String, x As Integer, strTemp As String

strWag = MSTSPath & "\Trains\Trainset\Default\Default.wag"
If FileExists(strWag) Then
MyString = ReadUniFile(strWag)
x = InStr(MyString, "Collision threshold velocity")
If x > 0 Then
x = InStrRev(MyString, "m/s", x)
strTemp = Mid$(MyString, x - 3, 3)
If strTemp <> "0.1" Then
Select Case MsgBox("The Current Instance of MSTS Trainset does not appear to include the modified DEFAULT.wag file" _
                   & vbCrLf & "required for use with these Couplers - Do you wish to install it?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
FileCopy App.Path & "/Default/Default.wag", strWag
    Case vbNo
Exit Function
End Select
End If
End If
Else
Select Case MsgBox("The Current Instance of MSTS Trainset does not appear to include the modified DEFAULT.wag file" _
                   & vbCrLf & "required for use with these Couplers - Do you wish to install it?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
FileCopy App.Path & "/Default/Default.wag", strWag
    Case vbNo
Exit Function
End Select
End If

End Function

Private Function ReplaceAirBrakes(MyString As String, strBrakeFile As String, strFP As String) As Boolean



On Error GoTo ERRHANDLER
Dim x As Integer
Dim xx As Integer, strStart As String, strEnd As String
Dim strBrake As String


strBrake = ReadUniFile(strBrakeFile)
x = InStr(MyString, "AirBrakesAirCompressorPowerRating")
If x > 0 Then
xx = InStr(x + 1, MyString, "Brake_Dynamic")
End If
If x = 0 Or xx = 0 Then
strReport = strReport & strFP & " could not be modified automatically due to format discrepancies" & vbCrLf
Exit Function
End If


        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx - 1)
        MyString = strStart & strBrake & strEnd
   ' End If
    'End If
'End If



Exit Function
ERRHANDLER:
  Call MsgBox("An error " & Err.Description & " occurred in the function ReplaceBrakess ", vbExclamation, App.Title)

End Function

Private Function ReplaceBrakes(MyString As String, strBrakeFile As String, strFP As String) As Boolean



On Error GoTo ERRHANDLER
Dim x As Integer
Dim xx As Integer, Y As Integer, strStart As String, strEnd As String
Dim strBrake As String


strBrake = ReadUniFile(strBrakeFile)
x = InStr(MyString, "BrakeEquipmentType")
If x > 0 Then
xx = InStr(x + 1, MyString, "BrakeCylinderPressureForMaxBrakeBrakeForce")
End If
If x = 0 Or xx = 0 Then
strReport = strReport & strFP & " could not be modified automatically due to format discrepancies" & vbCrLf
Exit Function
End If
Rem ********* if xx>0 then two couplers else only 1 ************
Y = InStr(xx, MyString, ")")

        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & strBrake & strEnd
   ' End If
    'End If
'End If



Exit Function
ERRHANDLER:
  Call MsgBox("An error " & Err.Description & " occurred in the function ReplaceBrakess ", vbExclamation, App.Title)

End Function


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
Sub EvalExpression(E As String)
Dim vlist As String, pp As Integer, ppc As Integer, ppop As Integer, pp2 As Integer, ee As String, tp As String
Dim ppe As Integer, s As String, l1 As Double, l2 As Double, v1 As Double, v2 As Double, R As Double


   On Error GoTo matherr
  
   ' import all variables and their values
   vlist = "|"
'   lc = List1.ListCount - 1
'   For T = 0 To lc
'      vlist = vlist & UCase$(List1.List(T)) & "|"
'   Next
   
  ' e = Text1

   ' remove all spaces to avoid problems
   While InStr(E, " ") > 0
      pp = InStr(E, " ")
      E = Left$(E, pp - 1) & Right$(E, Len(E) - pp)
   Wend

   Do
      ' locate next ( ) section
      pp = InStr(E, ")")
      ppc = InStr(E, "(")
      If (ppc > 0 And pp = 0) Or (pp > 0 And ppc = 0) Then
         MsgBox "Parenthesis missmatch - check your opens and closes!", 16, "Stein Seal Advisor Message"
         Exit Sub ' error parsing - open/close parenthesis missmatch
      End If
      ppop = InStr(E, "/") + InStr(E, "*") + InStr(E, "^") + InStr(E, "+") + InStr(E, "-")
      If pp = 0 And ppop = 0 Then Exit Do
      pp2 = 0
      If pp > 0 Then pp2 = InStrRev(E, "(", pp)
      If pp > 0 And pp2 > 0 Then ee = Mid$(E, pp2 + 1, (pp - pp2 - 1)) Else ee = E
      ' evaluate expression
      Do
         If InStr(ee, "/") + InStr(ee, "*") + InStr(ee, "^") + InStr(ee, "+") + InStr(ee, "-") > 0 Then
            ' follow 'my dear aunt sally' method if multiple operations found in same ( ) group
            ppe = 0
            If ppe = 0 Then If InStr(ee, "^") > 0 Then ppe = InStr(ee, "^"): tp = "^"
            If ppe = 0 Then If InStr(ee, "*") > 0 Then ppe = InStr(ee, "*"): tp = "*"
            If ppe = 0 Then If InStr(ee, "/") > 0 Then ppe = InStr(ee, "/"): tp = "/"
            If ppe = 0 Then If InStr(ee, "+") > 0 Then ppe = InStr(ee, "+"): tp = "+"
            If ppe = 0 Then If InStr(ee, "-") > 0 Then ppe = InStr(ee, "-"): tp = "-"
            If ppe > 0 Then

               s = GetValues(ee, ppe, vlist)
               l1 = Val(Delim(s, "|"))
               l2 = Val(Delim(s, "|"))
               v1 = Val(Delim(s, "|"))
               v2 = Val(s)
               If tp = "^" Then R = v1 ^ v2
               If tp = "*" Then R = v1 * v2
               If tp = "/" Then R = v1 / v2
               If tp = "+" Then R = v1 + v2
               If tp = "-" Then R = v1 - v2
               R = Trim$(Str$(R))
               ' replace original expression with final value
               ee = Left$(ee, l1 - 1) & R & Right$(ee, Len(ee) - l2)
            End If
         Else
            ' replace entire expression with final value - check if a math function exists
            ' immediately outside () expression or () is just used to force the calculation order
            R = Val(ee)
            pp3 = pp2 - 3
            If pp3 >= 1 Then
               ppf = UCase$(Mid$(E, pp3, 3))
               If InStr("/ABS/ATN/COS/EXP/FIX/INT/LOG/SGN/SIN/SQR/TAN/", "/" & ppf & "/") > 0 Then
                  pp2 = pp2 - 3
                  If ppf = "ABS" Then R = Abs(R)
                  If ppf = "ATN" Then R = Atn(R)
                  If ppf = "COS" Then R = Cos(R)
                  If ppf = "EXP" Then R = Exp(R)
                  If ppf = "FIX" Then R = Fix(R)
                  If ppf = "INT" Then R = Int(R)
                  If ppf = "LOG" Then R = Log(R)
                  If ppf = "SGN" Then R = Sgn(R)
                  If ppf = "SIN" Then R = Sin(R)
                  If ppf = "SQR" Then R = Sqr(R)
                  If ppf = "TAN" Then R = Tan(R)
               End If
            End If
            R = Trim$(Str$(R))
            If pp2 > 0 And pp > 0 Then
               E = Left$(E, pp2 - 1) & R & Right$(E, Len(E) - pp)
            Else
               E = R
            End If
            Exit Do
         End If
      Loop
   Loop

   
endeval:
   On Error GoTo 0
   
   Exit Sub
   
matherr:
   MsgBox "An error occurred '" & Error$ & "' while trying to evaluate this formula.", 16, "Stein Seal Advisor Message"
  
   
   Resume endeval
   
End Sub
Function Delim(s, ByVal d)

   ' return left portion of string 's' prior to first
   ' occurance of delimiting character 'd'
   '
   ' strip string of leftmost portion, including
   ' delimiting character to prepare for next function call

   p = InStr(s, d)
   If p > 0 Then
      l = Left$(s, InStr(s, d) - 1)
      s = Right$(s, Len(s) - InStr(s, d))
   Else
      l = "" ' error - delimiter char not found, return empty string
   End If

   Delim = l

End Function
Function GetValues(ByVal ee, ByVal ppe, ByVal vlist) As String

   ' get variable or value to left of operand
   pp1 = ppe - 1
   vflag1 = False
   Do While pp1 > 0
      A = Asc(UCase$(Mid$(ee, pp1)))
      If Not ((A >= 65 And A <= 97) Or (A >= 48 And A <= 57) Or A = 46) Then pp1 = pp1 + 1: Exit Do
      If A >= 65 And A <= 97 Then vflag1 = True
      pp1 = pp1 - 1
   Loop
   If pp1 = 0 Then pp1 = 1
   VLeft = Mid$(ee, pp1, (ppe - pp1))
   If vflag1 Then
      ' alpha variable found - locate corrosponding value
      ppp = InStr(vlist, "|" & UCase$(VLeft) & "=")
      If ppp > 0 Then
         eee = Right$(vlist, Len(vlist) - ppp)
         xxx = Delim(eee, "=")
         VLeft = Delim(eee, "|")
      End If
   End If
   
   ' get variable or value to right of operand
   pp2 = ppe + 1
   vflag2 = False
   Do While pp2 <= Len(ee)
      A = Asc(UCase$(Mid$(ee, pp2)))
      If Not ((A >= 65 And A <= 97) Or (A >= 48 And A <= 57) Or A = 46) Then pp2 = pp2 - 1: Exit Do
      If A >= 65 And A <= 97 Then vflag2 = True
      pp2 = pp2 + 1
   Loop
   If pp2 > Len(ee) Then pp2 = Len(ee)
   VRight = Mid$(ee, ppe + 1, pp2 - ppe)
   If vflag2 Then
      ' alpha variable found - locate corrosponding value
      ppp = InStr(vlist, "|" & UCase$(VRight) & "=")
      If ppp > 0 Then
         eee = Right$(vlist, Len(vlist) - ppp)
         xxx = Delim(eee, "=")
         VRight = Delim(eee, "|")
      End If
   End If

   GetValues = pp1 & "|" & pp2 & "|" & VLeft & "|" & VRight
   
End Function
Private Function ReplaceCouplings_Auto(MyString As String, sLen As Single, intType As Integer) As Boolean


On Error GoTo ERRHANDLER
Dim x As Integer
Dim xx As Integer, Y As Integer, yy As Integer, strStart As String, strEnd As String
Dim strCoupler As String, X1 As Integer, strCPath As String, strCouplingFile As String
Dim Z As Integer, zz As Integer, q As Integer, strComment As String, qq As Integer


strCPath = App.Path & "\Couplings\TurboBills Couplers V3.0\"
Select Case intType
Case 1   'Steam
Rem ************** Put engine lengths here............................................................
strCouplingFile = strCPath & "Pass-Roadrailer-Coupler.txt"
Case 2   'Diesel
If sLen >= 19 Then
strCouplingFile = strCPath & "Large-Road-Engine-Coupler.txt"
ElseIf sLen < 19 And sLen >= 15 Then
strCouplingFile = strCPath & "Small-Road-Engine-Coupler.txt"
Else
strCouplingFile = strCPath & "Switch-Engine-Coupler.txt"
End If
Case 3   'Electric
If sLen >= 19 Then
strCouplingFile = strCPath & "Large-Road-Engine-Coupler.txt"
ElseIf sLen < 19 And sLen >= 15 Then
strCouplingFile = strCPath & "Small-Road-Engine-Coupler.txt"
Else
strCouplingFile = strCPath & "Switch-Engine-Coupler.txt"
End If
Case 4   'Freight
If sLen > 21.64 Then
strCouplingFile = strCPath & "75+_Ft_Rollingstock-Coupler.txt"
Else
strCouplingFile = strCPath & "RR_Short-Railcar-Coupler.txt"
End If
Case 5   'Carriage
strCouplingFile = strCPath & "Pass-Roadrailer-Coupler.txt"
Case 6   'Tender
strCouplingFile = strCPath & "RR_Short-Railcar-Coupler.txt"
Case 7   'Steam Tank Loco
strCouplingFile = strCPath & "Switch-Engine-Coupler.txt"


End Select

MyString = Replace(MyString, "Coupling" & vbCrLf & " (", "Coupling (")
DoEvents
MyString = Replace(MyString, "Lights    (", "Lights (")
DoEvents
MyString = Replace(MyString, "Lights   (", "Lights (")
DoEvents
MyString = Replace(MyString, "Lights  (", "Lights (")
DoEvents
strCoupler = ReadUniFile(strCouplingFile)
strCoupler = Trim(strCoupler)
If Left(strCoupler, 7) = "Comment" Or Left(strCoupler, 8) = vbTab & "Comment" Then
qq = InStr(strCoupler, ")")
strComment = Left(strCoupler, qq)
End If

x = InStr(MyString, "Coupling (")
If x > 0 Then
q = InStr(x, MyString, "Lights (")
xx = InStr(x + 1, MyString, "Coupling (")
If q > 0 Then
    If xx > q Then
    xx = 0
    End If
End If
x = InStrRev(MyString, vbLf, x)
x = x + 1
End If

Z = InStr(MyString, "Buffers (")
    If Z > 0 Then
    zz = InStr(Z + 1, MyString, "Buffers (")
    End If


Rem ********* if xx>0 then two couplers else only 1 ************
Rem *********if zz>0 then there are also two sets of Buffers *********
'Y = InStr(x, myString, "Buffers (")
If zz > 0 Then
Y = zz
Else
Y = Z
End If
If Y > 0 Then
'Buffers exist

yy = InStr(Y, MyString, "Angle")
yy = InStr(yy + 1, MyString, ")")
yy = InStr(yy + 1, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, yy + 1)
X1 = InStr(strEnd, ")")

qq = InStrRev(strStart, strComment)
If qq > 0 Then
strStart = Left(strStart, qq - 1)
End If


If X1 > 0 And X1 < 9 Then
strEnd = Mid$(strEnd, X1 + 1)
End If

MyString = strStart & strCoupler & strEnd
Else
    If xx > 0 And xx < (x + 1000) Then
    Y = InStr(xx, MyString, "Velocity (")
    q = InStr(xx, MyString, "CouplingHasRigidConnection")
    Else
    Y = InStr(x, MyString, "Velocity (")
    q = InStr(x, MyString, "CouplingHasRigidConnection")
    End If
    
    If Y > 0 Then
            If Mid$(MyString, Y - 3, 3) = "Max" Then
            Y = 0
            End If
    End If
    If q > Y Then Y = q
    If xx = 0 Or xx > (x + 1000) Then
    xx = x
    End If
   
    If Y > 0 Then
        Y = InStr(Y + 1, MyString, ")")
        Y = InStr(Y + 1, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & strCoupler & strEnd
    Else
        Y = InStr(xx, MyString, "r0")
        Y = InStr(Y + 1, MyString, ")")
        Y = InStr(Y + 1, MyString, ")")
        Y = InStr(Y + 1, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & strCoupler & strEnd
    End If
    'End If
End If



Exit Function
ERRHANDLER:
  Call MsgBox("An error " & Err.Description & " occurred in the function ReplaceCouplings_Auto ", vbExclamation, App.Title)

End Function
Private Function ReplaceCouplings(MyString As String, strCouplingFile As String) As Boolean


On Error GoTo ERRHANDLER
Dim x As Integer
Dim xx As Integer, Y As Integer, yy As Integer, strStart As String, strEnd As String
Dim strCoupler As String, X1 As Integer, Z As Integer, zz As Integer, q As Integer

MyString = Replace(MyString, "Coupling" & vbCrLf & " (", "Coupling (")
DoEvents
MyString = Replace(MyString, "Lights    (", "Lights (")
DoEvents
MyString = Replace(MyString, "Lights   (", "Lights (")
DoEvents
MyString = Replace(MyString, "Lights  (", "Lights (")
DoEvents
strCoupler = ReadUniFile(strCouplingFile)
x = InStr(MyString, "Coupling (")
If x = 0 Then
Call MsgBox("Current rolling-stock does not have a Coupling entry in the file being checked.", vbCritical, App.Title)

Exit Function
End If
If x > 0 Then
q = InStr(x, MyString, "Lights (")
xx = InStr(x + 1, MyString, "Coupling (")
If q > 0 Then
    If xx > q Then
    xx = 0
    End If
End If
x = InStrRev(MyString, vbLf, x)
x = x + 1
End If

Z = InStr(MyString, "Buffers (")
    If Z > 0 Then
    zz = InStr(Z + 1, MyString, "Buffers (")
    End If


Rem ********* if xx>0 then two couplers else only 1 ************
Rem *********if zz>0 then there are also two sets of Buffers *********
'Y = InStr(x, myString, "Buffers (")
If zz > 0 Then
Y = zz
Else
Y = Z
End If
If Y > 0 Then
'Buffers exist

    yy = InStr(Y, MyString, "Angle")
    yy = InStr(yy + 1, MyString, ")")
    yy = InStr(yy + 1, MyString, ")")
    strStart = Left$(MyString, x - 1)
    strEnd = Mid$(MyString, yy + 1)
    X1 = InStr(strEnd, ")")
        If X1 > 0 And X1 < 9 Then
        strEnd = Mid$(strEnd, X1 + 1)
        End If
    MyString = strStart & strCoupler & strEnd
Else
    If xx > 0 And xx < (x + 1000) Then
    Y = InStr(xx, MyString, "Velocity (")
    q = InStr(xx, MyString, "CouplingHasRigidConnection")
    Else
    Y = InStr(x, MyString, "Velocity (")
    q = InStr(x, MyString, "CouplingHasRigidConnection")
    End If
    
    If Y > 0 Then
            If Mid$(MyString, Y - 3, 3) = "Max" Then
            Y = 0
            End If
    End If
    If q > Y Then Y = q
    If xx = 0 Or xx > (x + 1000) Then
    xx = x
    End If
    If Y > 0 Then
        Y = InStr(Y + 1, MyString, ")")
        Y = InStr(Y + 1, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & strCoupler & strEnd
    Else
        Y = InStr(xx, MyString, "r0")
        Y = InStr(Y + 1, MyString, ")")
        Y = InStr(Y + 1, MyString, ")")
        Y = InStr(Y + 1, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & strCoupler & strEnd
    End If
    'End If
End If



Exit Function
ERRHANDLER:
  Call MsgBox("An error " & Err.Description & " occurred in the function ReplaceCouplings ", vbExclamation, App.Title)

End Function

Private Sub GetFileType(strPath As String, FileType As Integer)
Dim xx As Long, X1 As Long, strTemp As String

FileType = 0
MyString = ReadUniFile(strPath)
MyString = Replace(MyString, vbTab, " ")
MyString = Replace(MyString, ")", " )")
MyString = Replace(MyString, "(", "( ")
MyString = Replace(MyString, "          ", " ")
MyString = Replace(MyString, "         ", " ")
MyString = Replace(MyString, "        ", " ")
MyString = Replace(MyString, "       ", " ")
MyString = Replace(MyString, "      ", " ")
MyString = Replace(MyString, "     ", " ")
MyString = Replace(MyString, "    ", " ")
MyString = Replace(MyString, "   ", " ")
MyString = Replace(MyString, "  ", " ")
MyString = Replace(MyString, " ", " ")

xx = InStr(MyString, "Type ( Steam )")
If xx > 0 Then FileType = 1
xx = InStr(MyString, "Type ( Diesel )")
If xx > 0 Then FileType = 2
xx = InStr(MyString, "Type ( Electric )")
If xx > 0 Then FileType = 3
xx = InStr(MyString, "Type ( Freight )")
If xx > 0 Then FileType = 4
xx = InStr(MyString, "Type ( Carriage )")
If xx > 0 Then FileType = 5
xx = InStr(MyString, "Type ( Tender )")
If xx > 0 Then FileType = 6
If FileType = 1 Then

xx = InStr(MyString, "IsTenderRequired")
If xx > 0 Then
xx = InStr(xx, MyString, "(")
X1 = InStr(xx, MyString, ")")
strTemp = Mid$(MyString, xx + 1, X1 - (xx + 1))
If Val(strTemp) <> 1 Then
FileType = 7     'Steam Tank Engine
End If
End If
End If



End Sub

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


Private Sub Check1_Click()
If Check1.value = 0 Then
Label6.Caption = "Insert Before Line"
ElseIf Check1.value = 1 Then
Label6.Caption = "Change or Delete Values in:-"
End If
End Sub

Private Sub cmdApplyAll_Click()
Dim MyString As String, x As Integer, xx As Integer, strStart As String
Dim strEnd As String, strSD As String, i As Integer
Dim strShape As String, strWagName As String


For i = 0 To List1.ListCount - 1

If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)

List1.TopIndex = i
List1.Selected(i) = False

If Right$(List1.List(i), 4) <> ".eng" And Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not an .eng or .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
MyString = ReadUniFile(fullpath$)

x = InStr(MyString, "WagonShape")
If x = 0 Then
strReport = strReport & "WagonShape entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strShape = Mid$(MyString, x + 1, xx - (x + 1))
strShape = Trim$(strShape)
strShape = Replace(strShape, ChrW$(34), "")
If Right$(strShape, 2) <> ".s" Then
strReport = strReport & "File " & List1.List(i) & _
    " has an invalid WagonShape entry so could not be processed" & vbCrLf
GoTo CarryON
End If

x = InStrRev(List1.List(i), "\")
strShapePath = MSTSPath & "\Trains\Trainset\" & Left$(List1.List(i), x)
strWagName = MSTSPath & "\Trains\Trainset\" & Mid$(List1.List(i), x + 1)
strPicView = strShapePath & strShape
strSDFile = strPicView & "d"

Rem *****************************************************


FileCopy strSDFile, strSDFile & ".bak"
MyString = ReadUniFile(strSDFile)
x = InStr(MyString, "ESD_Bounding")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
strSD = Text1(5)
MyString = strStart & strSD & strEnd
Call WriteUniFile(strSDFile, MyString)
DoEvents
Rem **************** Update .wag/.eng file

FileCopy fullpath$, fullpath$ & ".bak"
MyString = ReadUniFile(fullpath$)
x = InStr(MyString, "Size")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(0) & strEnd

x = InStr(MyString, "InertiaTensor")
xx = InStr(x, MyString, ")")
xx = InStr(xx + 1, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(2) & strEnd

'x = InStr(myString, "DerailRailForce")
'xx = InStr(x, myString, ")")
'strStart = Left$(myString, x - 1)
'strEnd = Mid$(myString, xx + 1)
'myString = strStart & Text1(3) & strEnd
'
'x = InStr(myString, "DerailBufferForce")
'xx = InStr(x, myString, ")")
'strStart = Left$(myString, x - 1)
'strEnd = Mid$(myString, xx + 1)
'myString = strStart & Text1(4) & strEnd
'
'x = InStr(myString, "Mass")
'xx = InStr(x, myString, ")")
'strStart = Left$(myString, x - 1)
'strEnd = Mid$(myString, xx + 1)
'myString = strStart & Text1(1) & vbCrLf & strEnd
Call WriteUniFile(fullpath$, MyString)
DoEvents
End If
CarryON:
Next i
End Sub

Private Sub Command1_Click(Index As Integer)
Dim result As Variant, DirCount As Integer, FirstPath As String, intType As Integer

MousePointer = 11
Dim i As Integer

If Index <> 3 Then
Command28.Visible = False
Command29.Visible = False
Command30.Visible = False
Else
Command28.Visible = True
Command29.Visible = True
Command30.Visible = True
End If
Command23.Visible = False
Command24.Visible = False
Command25.Visible = False
For i = List1.ListCount - 1 To 0 Step -1
List1.RemoveItem i
Next i
lblCount(0).Caption = List1.ListCount
lblCount(1).Caption = List1.SelCount
Select Case Index
Case 0


flagType = 1
frmUtils.Text1(0).Text = "*.eng"
frmUtils.File1(0).Pattern = "*.eng"
    frmUtils.Dir1(0).Path = MSTSPath & "\Trains\Trainset"
    DoEvents
    FirstPath = frmUtils.Dir1(0).Path
    DirCount = frmUtils.Dir1(0).ListCount
    intType = 1
    result = DirDiver(FirstPath, DirCount, "", intType)
    intType = 7
    result = DirDiver(FirstPath, DirCount, "", intType)
Case 1
Command23.Visible = True
Command24.Visible = True
flagType = 2
frmUtils.Text1(0).Text = "*.eng"
    frmUtils.Dir1(0).Path = MSTSPath & "\Trains\Trainset"
    DoEvents
    FirstPath = frmUtils.Dir1(0).Path
    DirCount = frmUtils.Dir1(0).ListCount
    intType = 2
    result = DirDiver(FirstPath, DirCount, "", intType)
Case 2
flagType = 3
frmUtils.Text1(0).Text = "*.eng"
    frmUtils.Dir1(0).Path = MSTSPath & "\Trains\Trainset"
    DoEvents
    FirstPath = frmUtils.Dir1(0).Path
    DirCount = frmUtils.Dir1(0).ListCount
    intType = 3
    result = DirDiver(FirstPath, DirCount, "", intType)
Case 3
Command25.Visible = True
flagType = 4
frmUtils.Text1(0).Text = "*.wag"
    frmUtils.Dir1(0).Path = MSTSPath & "\Trains\Trainset"
    DoEvents
    FirstPath = frmUtils.Dir1(0).Path
    DirCount = frmUtils.Dir1(0).ListCount
    intType = 4
    result = DirDiver(FirstPath, DirCount, "", intType)
Case 4
Command25.Visible = True
flagType = 5
frmUtils.Text1(0).Text = "*.wag"
    frmUtils.Dir1(0).Path = MSTSPath & "\Trains\Trainset"
    DoEvents
    FirstPath = frmUtils.Dir1(0).Path
    DirCount = frmUtils.Dir1(0).ListCount
    intType = 5
    result = DirDiver(FirstPath, DirCount, "", intType)
Case 5
flagType = 6
frmUtils.Text1(0).Text = "*.wag"
    frmUtils.Dir1(0).Path = MSTSPath & "\Trains\Trainset"
    DoEvents
    FirstPath = frmUtils.Dir1(0).Path
    DirCount = frmUtils.Dir1(0).ListCount
    intType = 6
    result = DirDiver(FirstPath, DirCount, "", intType)
    Case 6
flagType = 9
frmUtils.Text1(0).Text = "*.eng;*.wag"
    frmUtils.Dir1(0).Path = MSTSPath & "\Trains\Trainset"
    DoEvents
    FirstPath = frmUtils.Dir1(0).Path
    DirCount = frmUtils.Dir1(0).ListCount
    intType = 9
    result = DirDiver(FirstPath, DirCount, "", intType)
    Case 7
flagType = 8
frmUtils.Text1(0).Text = "*.bak*"
    frmUtils.Dir1(0).Path = MSTSPath & "\Trains\Trainset"
    DoEvents
    FirstPath = frmUtils.Dir1(0).Path
    DirCount = frmUtils.Dir1(0).ListCount
    intType = 8
   
    result = DirDiver(FirstPath, DirCount, "", intType)
End Select
DoEvents
frmEngEdit.Show
MousePointer = 0
If strReport <> vbNullString Then
frmReport.Rich1.Text = strReport
 
     
 strReport = vbNullString

End If
End Sub

Private Sub Command10_Click()
Dim i As Integer, x As Integer, xx As Integer, j As Integer
Dim strBatText As String, strTemp As String
Dim strShape As String, strWagName As String
Dim NewFile As Integer, intSteam As Integer
Dim X1 As Single, X2 As Single, y1 As Single, y2 As Single, z1 As Single, z2 As _
    Single, x3 As Single, sngMass As Single
Dim strDerail As String, strDerailBuf As String, strMass As String
Dim booMass As Boolean, strDBF As String, strType As String
Dim strCouplingType As String, booBin As Boolean, dDBF As Double, strCOG As String
Dim strDRF As String, strESDBB As String, dDRF As Double, sTemp1 As Single, sTemp2 As Single
Dim booFormula As Boolean

On Error GoTo Errtrap

If flagType > 7 Then
Call MsgBox("You have to select a specific Rolling-Stock type for this option.", vbExclamation, App.Title)
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
List1.Selected(i) = False
End If
Next i
lblCount(1).Caption = List1.SelCount
Exit Sub
End If
If Check4.value = 1 Then booFormula = True

Frame2.Visible = False
If Not DirExists(App.Path & "\TempFiles") Then
   MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents
If booFormula = True Then
If numFormula > 0 Then
For i = 1 To numFormula
If strParam(i) = "DerailBufferForce" Then
strDBF = strFormula(i, 1)
ElseIf strParam(i) = "DerailRailForce" Then
strDRF = strFormula(i, 1)
ElseIf strParam(i) = "ESD_Bounding_Box" Then
strESDBB = strParam(i)
End If
Next i
End If
End If
If strDBF = "" Then
strDBF = InputBox("DerailBufferForce will be set to 1000 kN for all items, enter a new value if you wish to change it", "DerailBufferForce", "1000")
End If
If strDBF = vbNullString Then Exit Sub
MousePointer = 11
Select Case MsgBox("Are the selected Models to be used with MSTSbin ?", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
booBin = True
    Case vbNo
booBin = False
End Select
For i = 0 To List1.ListCount - 1
If Check4.value = 1 Then booFormula = True

If List1.Selected(i) = True Then
j = InStr(List1.List(i), "Fred")
If j > 0 Then GoTo TryAgain
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
FileCopy fullpath$, fullpath$ & ".bak"
Else
For j = 1 To 50
    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
    Exit For
    End If
Next j
End If
    End If
List1.TopIndex = i
List1.Selected(i) = False
x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
If Right$(List1.List(i), 4) <> ".eng" And Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not an .eng or .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

Call GetFileType(fullpath$, intSteam)

MyString = ReadUniFile(fullpath$)
x = InStr(MyString, "Mass")
    If x = 0 Then
    strReport = strReport & "Mass entry not found in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
x = InStr(MyString, "WagonShape")
If x = 0 Then
strReport = strReport & "WagonShape entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strShape = Mid$(MyString, x + 1, xx - (x + 1))
strShape = Trim$(strShape)
strShape = Replace(strShape, ChrW$(34), "")
If Right$(strShape, 2) <> ".s" Then
strReport = strReport & "File " & List1.List(i) & _
    " has an invalid WagonShape entry so could not be processed" & vbCrLf
GoTo CarryON
End If

x = InStr(MyString, "Type")
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strType = Mid$(MyString, x + 1, xx - (x + 1))
strType = Trim$(strType)
If Left$(strType, 1) = ChrW$(34) Then
strType = Mid$(strType, 2)
End If
If Right$(strType, 1) = ChrW$(34) Then
strType = Left$(strType, Len(strType) - 1)
End If
If strType <> "Engine" And strType <> "Tender" And strType <> "Freight" And strType <> "Carriage" Then
strReport = strReport & fullpath$ & " contains an invalid Type statement. (Should be Engine, Tender, Freight or Carriage)." & vbCrLf
GoTo CarryON
End If

MyString = Replace(MyString, "- 0.1", "-0.1")

If Check3.value = 1 Then GoTo NoBBox
x = InStrRev(List1.List(i), "\")
strShapePath = MSTSPath & "\Trains\Trainset\" & Left$(List1.List(i), x)
strWagName = MSTSPath & "\Trains\Trainset\" & Mid$(List1.List(i), x + 1)
strPicView = strShapePath & strShape
strSDFile = strPicView & "d"
If Not FileExists(strSDFile) Then
strReport = strReport & strShape & "d not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If



strPicView = strPicView & ";2"

strBatText = ChrW$(34) & App.Path & "\sviewRR4.exe" & ChrW$(34) & " " & ChrW$(34) & _
    strPicView & ChrW$(34)
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)

DoEvents
NewFile = FreeFile
Open App.Path & "\tempfiles\tempsize.txt" For Input As #NewFile
Input #NewFile, A$
X1 = CSng(A$)
Input #NewFile, A$
y1 = CSng(A$)
Input #NewFile, A$
z1 = CSng(A$)
Input #NewFile, A$
X2 = CSng(A$)
Input #NewFile, A$
y2 = CSng(A$)
Input #NewFile, A$
z2 = CSng(A$)
Close #NewFile

x3 = X2 - X1
x3 = x3 / 2
If y2 < 2.5 Then
y2 = 2.5
End If

strSize = "Size ( " & Round(X2 - X1, 2) & "m " & Round(y2, 2) & "m " & Round(z2 - z1 - 0.4, 2) & "m )"
strInertia = "InertiaTensor ( Box ( " & Round(X2 - X1, 2) & "m " & Round(y2, 2) & "m " & Round(z2 - z1 - 1, 2) & "m ) )"
strSD = "ESD_Bounding_Box ( " & Round(-x3, 2) & " 0.9 " & Round(z1 + 0.5, 2) & " " & Round(x3, 2) & " " & Round(y2, 2) & " " & Round(z2 - 0.5, 2) & " )"
NewFile = FreeFile

NoBBox:
Rem ************ Replace Bar Couplers ****************
x = InStr(MyString, "( Automatic )")
    If x = 0 Then
    x = InStr(MyString, "( Chain )")
        If x > 0 Then
        strCouplingType = "Chain"
        End If
    Else
    strCouplingType = "Automatic"
    End If
Y = InStr(MyString, "( Bar )")
If Y > 0 Then
    If strCouplingType = vbNullString Then
    strCouplingType = "Automatic"
    End If
    MyString = Replace(MyString, "( Bar )", "(" & strCouplingType & ")")
End If
Rem ****************Check for front coupler
Rem ****************Check length/type

If booBin = False Then
If strType = "Freight" Or strType = "Carriage" Or strType = "Tender" Then
    If z2 - z1 < 15 Then
    ' "Short wagon"
        If strCouplingType = "Automatic" Then
        Call ReplaceCouplings(MyString, App.Path & "\Couplings\Coupling_Freight.txt")
        Else
        Call ReplaceCouplings(MyString, App.Path & "\Couplings\Coupling_Freight_Chain.txt")
        End If
    Else
    'Long Wagons
        If strCouplingType = "Automatic" Then
        Call ReplaceCouplings(MyString, App.Path & "\Couplings\Coupling_Long_Freight.txt")
        Else
        Call ReplaceCouplings(MyString, App.Path & "\Couplings\Coupling_Long_Freight_Chain.txt")
        End If
    End If
Else
    'Locos
    If strCouplingType = "Automatic" Then
    Call ReplaceCouplings(MyString, App.Path & "\Couplings\Coupling.txt")
    Else
    Call ReplaceCouplings(MyString, App.Path & "\Couplings\Coupling_Chain.txt")
    End If
    End If
ElseIf booBin = True Then
If strType = "Freight" Or strType = "Carriage" Or strType = "Tender" Then
    If z2 - z1 < 15 Then
    ' "Short wagon"
        If strCouplingType = "Automatic" Then
        Call ReplaceCouplings(MyString, App.Path & "\Couplings\Bin_Coupling_Freight.txt")
        Else
        Call ReplaceCouplings(MyString, App.Path & "\Couplings\Bin_Coupling_Freight_Chain.txt")
        End If
    Else
    'Long Wagons
        If strCouplingType = "Automatic" Then
        Call ReplaceCouplings(MyString, App.Path & "\Couplings\Bin_Coupling_Long_Freight.txt")
        Else
        Call ReplaceCouplings(MyString, App.Path & "\Couplings\Bin_Coupling_Long_Freight_Chain.txt")
        End If
    End If
Else
    'Locos
    If strCouplingType = "Automatic" Then
    Call ReplaceCouplings(MyString, App.Path & "\Couplings\Bin_Coupling.txt")
    Else
    Call ReplaceCouplings(MyString, App.Path & "\Couplings\Bin_Coupling_Chain.txt")
    End If
End If
End If
Rem *****************************************************

Text1(6) = List1.List(i)
x = InStr(MyString, "Mass")
    If x = 0 Then
    strReport = strReport & "Mass entry not found in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
If x > 1 Then
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp = Mid$(MyString, x + 1, xx - (x + 1))
strTemp = Trim$(strTemp)
x = InStr(strTemp, ChrW$(34))
    If x > 0 Then
    booMass = True
    strTemp = Replace(strTemp, ChrW$(34), "")
    End If
x = InStr(strTemp, "t")
    If x > 0 And Len(strTemp) > x + 1 Then
    strTemp = Left$(strTemp, x - 1)
    booMass = True
    End If
    If booMass = True Then
    strMass = "Mass ( " & strTemp & "t )"
    End If
    If Right$(strTemp, 1) = "t" Then
    strTemp = Left$(strTemp, Len(strTemp) - 1)
    End If
    strMass = "Mass ( " & strTemp & "t )"
        Text1(1) = strMass
sngMass = Val(strTemp)
If sngMass < 21 Then booFormula = False
dDBF = sngMass
dDRF = sngMass
sngMass = Round(sngMass * 2.7)
    If sngMass < 1 Then
    strReport = strReport & "Mass entry invalid in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
strDerail = "DerailRailForce ( " & Str(sngMass) & "kN )"
Text1(3) = strDerail
End If
Rem *************** Calculate DBF **************************************
If booFormula = True Then
x = InStr(strDBF, "$mass")
If x <> 0 Then
strDBF = Replace(strDBF, "$mass", Trim(Str(Val(dDBF))))
Call EvalExpression(strDBF)
dDBF = Val(strDBF)
strDBF = Format(dDBF, "###0.00")
End If
End If
strDerailBuf = "DerailBufferForce ( " & strDBF & "kN )"
Text1(4) = strDerailBuf
Rem ***************** Calculate DRF *************

x = InStr(MyString, "CentreOfGravity")
If x > 0 And booFormula = True Then
x = InStr(x, MyString, "(")
xx = InStr(x + 1, MyString, ")")
strCOG = Mid(MyString, x + 1, xx - (x + 1))
strCOG = Trim(strCOG)
x = InStr(strCOG, " ")
xx = InStr(x + 1, strCOG, " ")
strCOG = Mid(strCOG, x + 1, xx - (x + 1))
strCOG = Trim(strCOG)

If Val(strCOG) = 0 Then
strCOG = Str(Round(y2, 2) / 2)
End If
Else
strCOG = Str(Round(y2, 2) / 2)
End If
If strDRF <> "" Then
strDRF = Replace(strDRF, "$mass", Trim(Str(Val(dDRF))))
strDRF = Replace(strDRF, "$CentreOfGravity", strCOG)
Call EvalExpression(strDRF)
dDRF = Val(strDRF)
strDRF = Format(dDRF, "###0.00")

strDerail = "DerailRailForce ( " & strDRF & "kN )"
Text1(3) = strDerail
Else
strDerail = "DerailRailForce ( " & Str(sngMass) & "t )"
Text1(3) = strDerail
End If

Rem **************************************************
x = InStr(MyString, "DerailRailForce")
If x = 0 Then
strReport = strReport & "DerailRailForce entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strDerail & strEnd
x = InStr(MyString, "DerailBufferForce")
If x = 0 Then
strReport = strReport & "DerailBufferForce entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strDerailBuf & strEnd
If booMass = True Then
booMass = False
End If
x = InStr(MyString, "Mass")
xx = InStr(x, MyString, vbLf)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strMass & vbCrLf & strEnd

If Check3.value = 1 Then GoTo NoBBox2
x = InStr(MyString, "Size")

If x = 0 Then
strReport = strReport & "Size entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strSize & strEnd
Text1(0) = strSize
x = InStr(MyString, "InertiaTensor")
If x = 0 Then
strReport = strReport & "InertiaTensor entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
xx = InStr(xx + 1, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strInertia & strEnd
Text1(2) = strInertia

'End If

NoBBox2:
Rem **********************************
Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), MyString)
DoEvents

If Check3.value = 1 Then
GoTo TryAgain
End If
FileCopy strSDFile, strSDFile & ".bak"
Rem ************ Calculate ESD_Bounding_Box
If booFormula = True Then

strTemp = Replace(strFormula(3, 1), "$Shape.MinX", Str(Val(x3)))
x = InStr(strTemp, "-")
xx = InStr(strTemp, "+")
If x > 0 Or xx > 0 Then
Call EvalExpression(strTemp)
End If

DoEvents
strFormula(3, 1) = Format(Val(strTemp), "#0.00##")

strTemp = strFormula(3, 2)
x = InStr(strTemp, "-")
xx = InStr(strTemp, "+")
If x > 0 Or xx > 0 Then
Call EvalExpression(strTemp)
End If


DoEvents
strFormula(3, 2) = Format(Val(strTemp), "#0.00##")

strTemp = Trim(Replace(strFormula(3, 3), "$Shape.MinZ", Str(Val(z1))))
x = InStr(strTemp, " ")
If x > 0 Then
sTemp1 = Val(Left(strTemp, x - 1))
sTemp2 = Val(Mid(strTemp, x + 1))
sTemp1 = sTemp1 + sTemp2
strTemp = Str(sTemp1)

End If

DoEvents
strFormula(3, 3) = Format(Val(strTemp), "#0.00##")

strTemp = Replace(strFormula(3, 4), "$Shape.MaxX", Str(Val(x3)))

x = InStr(strTemp, "-")
xx = InStr(strTemp, "+")
If x > 0 Or xx > 0 Then
Call EvalExpression(strTemp)
End If

DoEvents

strFormula(3, 4) = Format(Val(strTemp), "#0.00##")

If Left(strFormula(3, 5), 3) = "max" Then
x = InStr(strFormula(3, 5), ",")
xx = InStr(x, strFormula(3, 5), "m")
If xx = 0 Then
xx = InStr(x, strFormula(3, 5), ")")
End If
strTemp = Mid(strFormula(3, 5), x + 1, xx - x + 1)
strTemp = Trim(strTemp)
If y2 < Val(strTemp) Then
strTemp = Str(y2)
End If
End If


DoEvents
strFormula(3, 5) = Format(Val(strTemp), "#0.00##")

strTemp = Replace(strFormula(3, 6), "$Shape.MaxZ", Str(Val(z2)))
x = InStr(strTemp, "-")
xx = InStr(strTemp, "+")
If x > 0 Or xx > 0 Then
Call EvalExpression(strTemp)
End If

DoEvents
strFormula(3, 6) = Format(Val(strTemp), "#0.00##")



strSD = "ESD_Bounding_Box ( -" & strFormula(3, 1) & " " & strFormula(3, 2) & " " & strFormula(3, 3) & " " & strFormula(3, 4) & " " & strFormula(3, 5) & " " & strFormula(3, 6) & " )"

End If
MyString = ReadUniFile(strSDFile)
x = InStr(MyString, "ESD_Bounding")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)

MyString = strStart & strSD & strEnd
Text1(5) = strSD
Call WriteUniFile(strSDFile, MyString)
DoEvents


TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If
Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'Command10_CLick' processing" _
                       & vbCrLf & fullpath$ _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
      
    Resume Next
        Case vbCancel
       GoTo CarryON
    Exit Sub
    End Select

End Sub

Private Sub Command11_Click()
Dim i As Integer, x As Integer, xx As Integer
Dim strShape As String, strWagName As String
Dim X1 As Single, X2 As Single, y1 As Single
Dim strIT As String, strBBox As String, strLength As String, strLenZ As String


On Error GoTo Errtrap
MousePointer = 11
Frame2.Visible = True
If Not DirExists(App.Path & "\TempFiles") Then
  MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents

For i = 0 To List1.ListCount - 1

If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)

List1.TopIndex = i
List1.Selected(i) = False

If Right$(List1.List(i), 4) <> ".eng" And Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not an .eng or .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
MyString = ReadUniFile(fullpath$)

x = InStr(MyString, "WagonShape")
If x = 0 Then
strReport = strReport & "WagonShape entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strShape = Mid$(MyString, x + 1, xx - (x + 1))
strShape = Trim$(strShape)
strShape = Replace(strShape, ChrW$(34), "")
If Right$(strShape, 2) <> ".s" Then
strReport = strReport & "File " & List1.List(i) & _
    " has an invalid WagonShape entry so could not be processed" & vbCrLf
GoTo CarryON
End If
Text1(6) = List1.List(i)
x = InStrRev(List1.List(i), "\")
strShapePath = MSTSPath & "\Trains\Trainset\" & Left$(List1.List(i), x)
strWagName = MSTSPath & "\Trains\Trainset\" & Mid$(List1.List(i), x + 1)


strPicView = strShapePath & strShape
strSDFile = strPicView & "d"

Rem *****************************************************

x = InStr(MyString, "Size")
If x = 0 Then
strReport = strReport & "No Size entry in " & fullpath$ & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
strSize = Mid$(MyString, x, (xx - x) + 1)
Text1(0) = strSize
x = InStr(MyString, "Mass")
If x = 0 Then
strReport = strReport & "No Mass entry in " & fullpath$ & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
strMass = Mid$(MyString, x, (xx - x) + 1)
Text1(1) = strMass
x = InStr(MyString, "InertiaTensor")
If x = 0 Then
strReport = strReport & "No Size entry in " & fullpath$ & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
xx = InStr(xx + 1, MyString, ")")
strIT = Mid$(MyString, x, (xx - x) + 1)
Text1(2) = strIT
x = InStr(MyString, "DerailRailForce")
If x = 0 Then
strReport = strReport & "No DerailRailForce entry in " & fullpath$ & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
strDerail = Mid$(MyString, x, (xx - x) + 1)
Text1(3) = strDerail
x = InStr(MyString, "DerailBufferForce")
If x = 0 Then
strReport = strReport & "No DerailBufferForce entry in " & fullpath$ & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
strDerailBuf = Mid$(MyString, x, (xx - x) + 1)
Text1(4) = strDerailBuf
MyString = ReadUniFile(strSDFile)
x = InStr(MyString, "ESD_Bounding")
If x = 0 Then
strReport = strReport & "No Bounding_Box entry in " & strSDFile & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
strBBox = Mid$(MyString, x, (xx - x) + 1)
Text1(5) = strBBox

End If


CarryON:
Next i
strBound = Text1(5)

x = InStr(strBound, "(")
x = InStr(x + 1, strBound, " ")
X1 = InStr(x + 1, strBound, " ")
Text2(0) = Mid$(strBound, x + 1, X1 - x)
X2 = InStr(X1 + 1, strBound, " ")
Text2(1) = Mid$(strBound, X1 + 1, X2 - X1)
xx = InStr(X2 + 1, strBound, " ")
strMinusZ = Mid$(strBound, X2 + 1, xx - X2)
Text2(2) = strMinusZ
Y = InStr(xx + 1, strBound, " ")
Text2(3) = Mid$(strBound, xx + 1, Y - xx)
y1 = InStr(Y + 1, strBound, " ")
Text2(4) = Mid$(strBound, Y + 1, y1 - Y)
yy = InStr(y1 + 1, strBound, " ")
strPlusZ = Mid$(strBound, y1 + 1, yy - y1)
Text2(5) = strPlusZ


strLength = Text1(0)
x = InStr(strLength, "m")
x = InStr(x + 1, strLength, "m")
X2 = InStr(x + 1, strLength, ")")
strLenZ = Trim$(Mid$(strLength, x + 1, X2 - (x + 1)))
If Right$(strLenZ, 1) = "m" Then
strLenZ = Left$(strLenZ, Len(strLenZ) - 1)
End If
Text3 = strLenZ



DoEvents
MousePointer = 0
If strReport <> vbNullString Then

   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
End If
' Command14.Enabled = False
 
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If

End Sub


Private Sub Command12_Click(Index As Integer)
Dim strTemp As String, lngTemp As Single, strIT As String, x As Integer
Dim xx As Integer, lngIT As Single

strIT = Text1(2)
x = InStrRev(strIT, "m")
xx = InStrRev(strIT, "m", x - 1)
strIT = Left$(strIT, xx + 1)

Select Case Index
Case 0
strTemp = Text2(2)
lngTemp = Val(strTemp)
lngTemp = Round(lngTemp - 0.05, 2)
Text2(2) = Trim$(Str(lngTemp))
Case 1
strTemp = Text2(2)
lngTemp = Val(strTemp)
lngTemp = Round(lngTemp + 0.05, 2)
Text2(2) = Trim$(Str(lngTemp))
Case 2
strTemp = Text2(5)
lngTemp = Val(strTemp)
lngTemp = Round(lngTemp - 0.05, 2)
Text2(5) = Trim$(Str(lngTemp))
Case 3
strTemp = Text2(5)
lngTemp = Val(strTemp)
lngTemp = Round(lngTemp + 0.05, 2)
Text2(5) = Trim$(Str(lngTemp))
Case 4
booEdited = True
strTemp = Text3
lngTemp = Val(strTemp)
lngTemp = Round(lngTemp + 0.05, 2)
Text3 = Trim$(Str(lngTemp))
Case 5
booEdited = True
strTemp = Text3
lngTemp = Val(strTemp)
lngTemp = Round(lngTemp - 0.05, 2)
Text3 = Trim$(Str(lngTemp))
End Select
lngIT = -(Val(Text2(2))) + Val(Text2(5))
strIT = strIT & Trim$(Str(lngIT)) & "m ))"
Text1(2) = strIT

End Sub

Private Sub Command13_Click()
Dim MyString As String, x As Integer, xx As Integer, strStart As String
Dim strEnd As String, strSD As String

FileCopy strSDFile, strSDFile & ".bak"
MyString = ReadUniFile(strSDFile)
x = InStr(MyString, "ESD_Bounding")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
strSD = Text1(5)
MyString = strStart & strSD & strEnd
Call WriteUniFile(strSDFile, MyString)
DoEvents
Rem **************** Update .wag/.eng file
If booEdited = True Then
booEdited = False
FileCopy fullpath$, fullpath$ & ".bak"
MyString = ReadUniFile(fullpath$)
x = InStr(MyString, "Size")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(0) & strEnd

x = InStr(MyString, "InertiaTensor")
xx = InStr(x, MyString, ")")
xx = InStr(xx + 1, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(2) & strEnd

x = InStr(MyString, "DerailRailForce")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(3) & strEnd

x = InStr(MyString, "DerailBufferForce")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(4) & strEnd

x = InStr(MyString, "Mass")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(1) & vbCrLf & strEnd
Call WriteUniFile(fullpath$, MyString)
DoEvents
End If
Command14.Enabled = True

End Sub

Private Sub Command14_Click()
Dim strSFile As String, strBatText As String

Call MsgBox("When the Shape File Viewer screen appears, the measuring tool may not appear correctly aligned, (a single" _
            & vbCrLf & "layer of lines), if this is so, select the Tools\Orthagonal View from the Menu to correct it (you may need to click twice)." _
            , vbInformation, App.Title)


strSFile = Left$(strSDFile, Len(strSDFile) - 1)

  If strNewSize <> vbNullString Then
        
    strBatText = ChrW$(34) & App.Path & "\sviewBB.exe" & ChrW$(34) & " " & ChrW$(34) & strSFile & ";" & strNewSize & ChrW$(34)

Else
 strBatText = ChrW$(34) & App.Path & "\sviewBB.exe" & ChrW$(34) & " " & ChrW$(34) & strSFile & ChrW$(34)
 End If
  
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)
    
   
End Sub




Private Sub Command15_Click()
Dim z1 As Single, z2 As Single, z3 As Single

z1 = Val(Text2(2))
z2 = Val(Text2(5))
z3 = -(z1) + z2
z1 = z3 / 2
z2 = z3 / 2
z1 = -z1
z1 = Round(z1, 2)
z2 = Round(z2, 2)
Text2(2) = Str(z1)
Text2(5) = Str(z2)

End Sub

Private Sub Command16_Click()
Dim i As Integer, x As Integer, xx As Integer, j As Integer
Dim strBatText As String, strTemp As String
Dim strShape As String, strWagName As String
Dim NewFile As Integer, intSteam As Integer
Dim X1 As Single, X2 As Single, y1 As Single, y2 As Single, z1 As Single, z2 As _
    Single, x3 As Single, sngMass As Single
Dim strDerail As String, strDerailBuf As String, strMass As String
Dim booMass As Boolean, strDBF As String, strType As String
Dim strCouplingType As String, strCoupling As String, strLength As String
Dim booShort As Boolean, strShort As String, dDBF As Double

If flagType > 7 Then
Call MsgBox("You have not selected a specific Rolling-Stock type for this option.", vbExclamation, App.Title)
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
List1.Selected(i) = False
End If
Next i
lblCount(1).Caption = List1.SelCount
Exit Sub
End If
On Error GoTo Errtrap
MousePointer = 11
Frame2.Visible = False
If Not DirExists(App.Path & "\TempFiles") Then
   MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents
If Check3.value = 0 Then
strDBF = InputBox("DerailBufferForce will be set to 1000 kN for all items, enter a new value here if you wish to change it")
If strDBF = vbNullString Then
strDBF = "1000"
End If
End If
CDL1.Filter = "Coupling Files (*.txt)|*.txt"
CDL1.DialogTitle = "Select Coupling File to use"
CDL1.InitDir = App.Path & "\Couplings"
CDL1.FilterIndex = 1
CDL1.Action = 1
strCoupling = CDL1.Filename
DoEvents
CDL1.InitDir = ""
CDL1.Filename = ""
If flagType = 4 Then
Select Case MsgBox("If length of wagon is less than 11.89m (39 ft) then use the" _
                   & vbCrLf & "'Short Wagon' coupler instead of your selected coupler?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
booShort = True
strShort = App.Path & "\Couplings\RR_Short-Railcar-Coupler.txt"
    Case vbNo
booShort = False
strShort = vbNullString
End Select
End If
If strCoupling = vbNullString Then
MousePointer = 0
Exit Sub
End If
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
FileCopy fullpath$, fullpath$ & ".bak"
Else
For j = 1 To 50
    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
    Exit For
    End If
Next j
End If
    End If
List1.TopIndex = i
List1.Selected(i) = False
x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "-NC-")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "SBW-")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "SCN-")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
If Right$(List1.List(i), 4) <> ".eng" And Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not an .eng or .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

Call GetFileType(fullpath$, intSteam)
MyString = ReadUniFile(fullpath$)
x = InStr(MyString, "Mass")
    If x = 0 Then
    strReport = strReport & "Mass entry not found in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
x = InStr(MyString, "WagonShape")
If x = 0 Then
strReport = strReport & "WagonShape entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strShape = Mid$(MyString, x + 1, xx - (x + 1))
strShape = Trim$(strShape)
strShape = Replace(strShape, ChrW$(34), "")
If Right$(strShape, 2) <> ".s" Then
strReport = strReport & "File " & List1.List(i) & _
    " has an invalid WagonShape entry so could not be processed" & vbCrLf
GoTo CarryON
End If

x = InStr(MyString, "Type")
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strType = Mid$(MyString, x + 1, xx - (x + 1))
strType = Trim$(strType)
If Left$(strType, 1) = ChrW$(34) Then
strType = Mid$(strType, 2)
End If
If Right$(strType, 1) = ChrW$(34) Then
strType = Left$(strType, Len(strType) - 1)
End If
If strType <> "Engine" And strType <> "Tender" And strType <> "Freight" And strType <> "Carriage" Then
strReport = strReport & fullpath$ & " contains an invalid Type statement. (Should be Engine, Tender, Freight or Carriage)." & vbCrLf
GoTo CarryON
End If

MyString = Replace(MyString, "- 0.1", "-0.1")

If Check3.value = 1 And booShort = False Then GoTo NoBBox
x = InStrRev(List1.List(i), "\")
strShapePath = MSTSPath & "\Trains\Trainset\" & Left$(List1.List(i), x)
strWagName = MSTSPath & "\Trains\Trainset\" & Mid$(List1.List(i), x + 1)
strPicView = strShapePath & strShape
strSDFile = strPicView & "d"
strPicView = strPicView & ";2"

strBatText = ChrW$(34) & App.Path & "\sviewRR4.exe" & ChrW$(34) & " " & ChrW$(34) & _
    strPicView & ChrW$(34)
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)

DoEvents
NewFile = FreeFile
Open App.Path & "\tempfiles\tempsize.txt" For Input As #NewFile
Input #NewFile, A$
X1 = CSng(A$)
Input #NewFile, A$
y1 = CSng(A$)
Input #NewFile, A$
z1 = CSng(A$)
Input #NewFile, A$
X2 = CSng(A$)
Input #NewFile, A$
y2 = CSng(A$)
Input #NewFile, A$
z2 = CSng(A$)
Close #NewFile

x3 = X2 - X1
x3 = x3 / 2
If y2 < 2.5 Then
y2 = 2.5
End If
strSize = "Size ( " & Round(X2 - X1, 2) & "m " & Round(y2, 2) & "m " & Round(z2 - z1 - 0.4, 2) & "m )"
strInertia = "InertiaTensor ( Box ( " & Round(X2 - X1, 2) & "m " & Round(y2, 2) & "m " & Round(z2 - z1 - 1, 2) & "m ) )"
strSD = "ESD_Bounding_Box ( " & Round(-x3, 2) & " 0.9 " & Round(z1 + 0.5, 2) & " " & Round(x3, 2) & " " & Round(y2, 2) & " " & Round(z2 - 0.5, 2) & " )"
strLength = Round(z2 - z1 - 0.4, 2)
NoBBox:
Rem ************ Replace Bar Couplers ****************
x = InStr(MyString, "( Automatic )")
    If x = 0 Then
    x = InStr(MyString, "( Chain )")
        If x > 0 Then
        strCouplingType = "Chain"
        End If
    Else
    strCouplingType = "Automatic"
    End If
Y = InStr(MyString, "( Bar )")
If Y > 0 Then
    If strCouplingType = vbNullString Then
    strCouplingType = "Automatic"
    End If
    MyString = Replace(MyString, "( Bar )", "(" & strCouplingType & ")")
End If
Rem ****************Check for front coupler
Rem ****************Check length/type


If booShort = True And Val(strLength) < 11.89 Then
Call ReplaceCouplings(MyString, strShort)
Else
Call ReplaceCouplings(MyString, strCoupling)
End If

    
Rem *****************************************************

Text1(6) = List1.List(i)
x = InStr(MyString, "Mass")
    If x = 0 Then
    strReport = strReport & "Mass entry not found in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
If x > 1 Then
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp = Mid$(MyString, x + 1, xx - (x + 1))
strTemp = Trim$(strTemp)
x = InStr(strTemp, ChrW$(34))
    If x > 0 Then
    booMass = True
    strTemp = Replace(strTemp, ChrW$(34), "")
    End If
x = InStr(strTemp, "t")
    If x > 0 And Len(strTemp) > x + 1 Then
    strTemp = Left$(strTemp, x - 1)
    booMass = True
    End If
    If booMass = True Then
    strMass = "Mass ( " & strTemp & "t )"
    End If
    If Right$(strTemp, 1) = "t" Then
    strTemp = Left$(strTemp, Len(strTemp) - 1)
    End If
    strMass = "Mass ( " & strTemp & "t )"
        Text1(1) = strMass
sngMass = Val(strTemp)
dDBF = sngMass
sngMass = Round(sngMass * 2.7)
    If sngMass < 1 Then
    strReport = strReport & "Mass entry invalid in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
    If Check3.value = 1 Then GoTo NoBBX3
strDerail = "DerailRailForce ( " & Str(sngMass) & "t )"
Text1(3) = strDerail
End If
x = InStr(strDBF, "$mass")
If x <> 0 Then
strDBF = Replace(strDBF, "$mass", Trim(Str(Val(dDBF))))
Call EvalExpression(strDBF)
dDBF = Val(strDBF)
strDBF = Format(dDBF, "###0.00")
End If

strDerailBuf = "DerailBufferForce ( " & strDBF & "kN )"
Text1(4) = strDerailBuf
x = InStr(MyString, "DerailRailForce")
If x = 0 Then
strReport = strReport & "DerailRailForce entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strDerail & strEnd
x = InStr(MyString, "DerailBufferForce")
If x = 0 Then
strReport = strReport & "DerailBufferForce entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strDerailBuf & strEnd

NoBBX3:
If booMass = True Then
booMass = False
End If
x = InStr(MyString, "Mass")
xx = InStr(x, MyString, vbLf)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strMass & vbCrLf & strEnd

If Check3.value = 1 Then GoTo NoBBox2
x = InStr(MyString, "Size")

If x = 0 Then
strReport = strReport & "Size entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strSize & strEnd
Text1(0) = strSize
x = InStr(MyString, "InertiaTensor")
If x = 0 Then
strReport = strReport & "InertiaTensor entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
xx = InStr(xx + 1, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strInertia & strEnd
Text1(2) = strInertia

'End If

NoBBox2:
Rem **********************************
Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), MyString)
DoEvents

If Check3.value = 1 Then
GoTo TryAgain
End If
FileCopy strSDFile, strSDFile & ".bak"
MyString = ReadUniFile(strSDFile)
x = InStr(MyString, "ESD_Bounding")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strSD & strEnd
Text1(5) = strSD
Call WriteUniFile(strSDFile, MyString)
DoEvents
'End If

TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If


End Sub

Private Sub Command17_Click()
Dim i As Integer, x As Integer, xx As Integer, j As Integer
Dim strTemp As String
Dim intSteam As Integer
Dim sngMass As Single
Dim strMass As String
Dim booMass As Boolean, strType As String
Dim strBrake As String, booAuto As Boolean
Dim strLoco As String

booProBrakes = False
booLSD = False

If flagType > 7 Then
Call MsgBox("You have not selected a specific Rolling-Stock type for this option.", vbExclamation, App.Title)
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
List1.Selected(i) = False
End If
Next i
lblCount(1).Caption = List1.SelCount
Exit Sub
End If
On Error GoTo Errtrap
MousePointer = 11
Frame2.Visible = False
If Not DirExists(App.Path & "\TempFiles") Then
   MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents
If flagType = 2 Then
Select Case MsgBox("You are modifying the Brake settings on Diesel Locomotives, do you wish the program to automatically" _
                   & vbCrLf & "select the Brake Values based on the mass of the selected locomotives?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
    
    booAuto = True
    dlgBrakes.Show 1
GoTo Brakes1
    Case vbNo

End Select
End If
If flagType = 4 Then
Select Case MsgBox("You are modifying the Brake settings on Freight Wagons, do you wish the program to automatically" _
                   & vbCrLf & "select the Brake Values based on the mass of the selected wagons?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
    booAuto = True
    dlgBrakes.Show 1
GoTo Brakes1
    Case vbNo

End Select
End If
If flagType = 5 Then
Select Case MsgBox("You are modifying the Brake settings on Passenger Cars, do you wish the program to automatically" _
                   & vbCrLf & "select the Brake Values for you?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
    booAuto = True
    dlgBrakes.Show 1
GoTo Brakes1
    Case vbNo

End Select
End If
CDL1.InitDir = App.Path & "\BrakeFiles"
CDL1.Filter = "Brake Files (*.txt)|*.txt"
CDL1.DialogTitle = "Select Brake File"

CDL1.FilterIndex = 1
CDL1.Action = 1
CDL1.InitDir = App.Path & "\BrakeFiles"
DoEvents
strBrake = CDL1.Filename
CDL1.Filename = ""
CDL1.InitDir = ""

If strBrake = vbNullString Then
MousePointer = 0
Exit Sub
End If
Brakes1:
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    x = InStrRev(fullpath$, "\")
    strLoco = Mid$(fullpath$, x + 1)
    If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
FileCopy fullpath$, fullpath$ & ".bak"
Else
For j = 1 To 50
    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
    Exit For
    End If
Next j
End If
    End If
List1.TopIndex = i
List1.Selected(i) = False
x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
If Right$(List1.List(i), 4) <> ".eng" And Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not an .eng or .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

Call GetFileType(fullpath$, intSteam)
MyString = ReadUniFile(fullpath$)
x = InStr(MyString, "Mass")
    If x = 0 Then
    strReport = strReport & "Mass entry not found in " & fullpath$ & " File could not be processed automatically" & vbCrLf
    GoTo CarryON
    End If
x = InStr(MyString, "Type")
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strType = Mid$(MyString, x + 1, xx - (x + 1))
strType = Trim$(strType)
If Left$(strType, 1) = ChrW$(34) Then
strType = Mid$(strType, 2)
End If
If Right$(strType, 1) = ChrW$(34) Then
strType = Left$(strType, Len(strType) - 1)
End If
If strType <> "Engine" And strType <> "Tender" And strType <> "Freight" And strType <> "Carriage" Then
strReport = strReport & fullpath$ & " contains an invalid Type statement. (Should be Engine, Tender, Freight or Carriage)." & vbCrLf
GoTo CarryON
End If

MyString = Replace(MyString, "- 0.1", "-0.1")
Rem ***************** Get Mass *****************
If booAuto = False Then GoTo Brakes2
x = InStr(MyString, "Mass")
    If x = 0 Then
    strReport = strReport & "Mass entry not found in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
If x > 1 Then
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp = Mid$(MyString, x + 1, xx - (x + 1))
strTemp = Trim$(strTemp)
x = InStr(strTemp, ChrW$(34))
    If x > 0 Then
    booMass = True
    strTemp = Replace(strTemp, ChrW$(34), "")
    End If
x = InStr(strTemp, "t")
    If x > 0 And Len(strTemp) > x + 1 Then
    strTemp = Left$(strTemp, x - 1)
    booMass = True
    End If
    If booMass = True Then
    strMass = "Mass ( " & strTemp & "t )"
    End If
    If Right$(strTemp, 1) = "t" Then
    strTemp = Left$(strTemp, Len(strTemp) - 1)
    End If
    sngMass = Val(strTemp)
    If booIron = False Then
    If booAuto = True And flagType = 4 Then
    Select Case sngMass
     Case Is <= 18.75
     strBrake = App.Path & "\BrakeFiles\18.75t railcar brake values.txt"
    Case Is <= 20
     strBrake = App.Path & "\BrakeFiles\20t railcar brake values.txt"
     Case Is <= 22.5
     strBrake = App.Path & "\BrakeFiles\22.5t railcar brake values.txt"
     Case Is <= 23.75
     strBrake = App.Path & "\BrakeFiles\23.75t railcar brake values.txt"
     Case Is <= 25
     strBrake = App.Path & "\BrakeFiles\25.0t railcar brake values.txt"
     Case Is <= 27.5
     strBrake = App.Path & "\BrakeFiles\27.5t railcar brake values.txt"
     Case Is <= 30
     strBrake = App.Path & "\BrakeFiles\30t railcar brake values.txt"
     Case Is <= 31.25
     strBrake = App.Path & "\BrakeFiles\31.25t railcar brake values.txt"
     Case Is <= 35
     strBrake = App.Path & "\BrakeFiles\35t railcar brake values.txt"
     Case Is <= 39
     strBrake = App.Path & "\BrakeFiles\39t railcar brake values.txt"
     Case Is <= 40
     strBrake = App.Path & "\BrakeFiles\40t railcar brake values.txt"
     Case Is <= 45
     strBrake = App.Path & "\BrakeFiles\45t railcar brake values.txt"
     Case Is <= 50
     strBrake = App.Path & "\BrakeFiles\50t railcar brake values.txt"
     Case Is <= 60
     strBrake = App.Path & "\BrakeFiles\60t railcar brake values.txt"
     Case Is <= 70
     strBrake = App.Path & "\BrakeFiles\70t railcar brake values.txt"
     Case Is <= 75
     strBrake = App.Path & "\BrakeFiles\75t railcar brake values.txt"
     Case Is <= 80
     strBrake = App.Path & "\BrakeFiles\80.0t railcar brake values.txt"
     Case Is <= 85
     strBrake = App.Path & "\BrakeFiles\85t railcar brake values.txt"
     Case Is <= 90
     strBrake = App.Path & "\BrakeFiles\90t railcar brake values.txt"
     Case Is <= 95
     strBrake = App.Path & "\BrakeFiles\95t railcar brake values.txt"
     Case Is <= 100
     strBrake = App.Path & "\BrakeFiles\100t railcar brake values.txt"
     Case Is <= 105
     strBrake = App.Path & "\BrakeFiles\105.0t railcar brake values.txt"
     Case Is <= 110
     strBrake = App.Path & "\BrakeFiles\110t railcar brake values.txt"
     Case Is <= 115
     strBrake = App.Path & "\BrakeFiles\115t railcar brake values.txt"
     Case Is <= 120
     strBrake = App.Path & "\BrakeFiles\120t railcar brake values.txt"
     Case Is <= 125
     strBrake = App.Path & "\BrakeFiles\125t railcar brake values.txt"
     Case Is <= 130
     strBrake = App.Path & "\BrakeFiles\130t railcar brake values.txt"
     Case Is <= 135
     strBrake = App.Path & "\BrakeFiles\135t railcar brake values.txt"
     Case Is <= 140
     strBrake = App.Path & "\BrakeFiles\140t railcar brake values.txt"
     Case Is <= 145
     strBrake = App.Path & "\BrakeFiles\145t railcar brake values.txt"
     Case Is <= 150
     strBrake = App.Path & "\BrakeFiles\150t railcar brake values.txt"
     Case Is > 150
     strBrake = App.Path & "\BrakeFiles\155t railcar brake values.txt"
    End Select
    ElseIf booAuto = True And flagType = 2 Then
        If Left$(strLoco, 1) = "#" Or Left$(strLoco, 2) = "AI" Or Right$(strLoco, 6) = "AI.eng" Then
        strBrake = App.Path & "\BrakeFiles\AI engines brake values.txt"
        Else
        Select Case sngMass
        Case Is < 140
         strBrake = App.Path & "\BrakeFiles\Small Driving Engine Brake values.txt"
         Case Is < 200
         strBrake = App.Path & "\BrakeFiles\Large Driving Engine Brake values.txt"
         Case Is >= 200
         strBrake = App.Path & "\BrakeFiles\200+ Ton Driving Engine Brake values.txt"
         End Select
         End If
     ElseIf booAuto = True And flagType = 5 Then
     strBrake = App.Path & "\BrakeFiles\Pax Brake values.txt"
    End If
    Rem ******************************* Iron brakes ************************************
ElseIf booIron = True Then
     If booAuto = True And flagType = 4 Then
    Select Case sngMass
    Case Is <= 5.5
     strBrake = App.Path & "\BrakeFiles\CastIron\5.5t railcar brake values.txt"
     Case Is <= 6.6
     strBrake = App.Path & "\BrakeFiles\CastIron\6.6t railcar brake values.txt"
     Case Is <= 7.75
     strBrake = App.Path & "\BrakeFiles\CastIron\7.75t railcar brake values.txt"
     Case Is <= 8.8
     strBrake = App.Path & "\BrakeFiles\CastIron\8.8t railcar brake values.txt"
     Case Is <= 9.9
     strBrake = App.Path & "\BrakeFiles\CastIron\9.9t railcar brake values.txt"
     Case Is <= 11
     strBrake = App.Path & "\BrakeFiles\CastIron\11t railcar brake values.txt"
     Case Is <= 12.1
     strBrake = App.Path & "\BrakeFiles\CastIron\12.1t railcar brake values.txt"
     Case Is <= 13.2
     strBrake = App.Path & "\BrakeFiles\CastIron\13.2t railcar brake values.txt"
     Case Is <= 14.3
     strBrake = App.Path & "\BrakeFiles\CastIron\14.3t railcar brake values.txt"
     Case Is <= 15.5
     strBrake = App.Path & "\BrakeFiles\CastIron\15.5t railcar brake values.txt"
     Case Is <= 16.6
     strBrake = App.Path & "\BrakeFiles\CastIron\16.6t railcar brake values.txt"
     Case Is <= 17.7
     strBrake = App.Path & "\BrakeFiles\CastIron\17.7t railcar brake values.txt"
    
    
    
    Rem ********************* OK
     Case Is <= 18.75
     strBrake = App.Path & "\BrakeFiles\CastIron\18.75t railcar brake values.txt"
    Case Is <= 20
     strBrake = App.Path & "\BrakeFiles\CastIron\20t railcar brake values.txt"
     Case Is <= 22.5
     strBrake = App.Path & "\BrakeFiles\CastIron\22.5t railcar brake values.txt"
     Case Is <= 23.75
     strBrake = App.Path & "\BrakeFiles\CastIron\23.75t railcar brake values.txt"
     Case Is <= 25
     strBrake = App.Path & "\BrakeFiles\CastIron\25.0t railcar brake values.txt"
     Case Is <= 27.5
     strBrake = App.Path & "\BrakeFiles\CastIron\27.5t railcar brake values.txt"
     Case Is <= 30
     strBrake = App.Path & "\BrakeFiles\CastIron\30t railcar brake values.txt"
     Case Is <= 31.25
     strBrake = App.Path & "\BrakeFiles\CastIron\31.25t railcar brake values.txt"
     Case Is <= 35
     strBrake = App.Path & "\BrakeFiles\CastIron\35t railcar brake values.txt"
     Case Is <= 39
     strBrake = App.Path & "\BrakeFiles\CastIron\39t railcar brake values.txt"
     Case Is <= 40
     strBrake = App.Path & "\BrakeFiles\CastIron\40t railcar brake values.txt"
     Case Is <= 45
     strBrake = App.Path & "\BrakeFiles\CastIron\45t railcar brake values.txt"
     Case Is <= 50
     strBrake = App.Path & "\BrakeFiles\CastIron\50t railcar brake values.txt"
     Case Is <= 55
     strBrake = App.Path & "\BrakeFiles\CastIron\55t railcar brake values.txt"
     Rem ************** OK from here down *************************************
     Case Is <= 60
     strBrake = App.Path & "\BrakeFiles\CastIron\60t railcar brake values.txt"
     Case Is <= 70
     strBrake = App.Path & "\BrakeFiles\CastIron\70t railcar brake values.txt"
     Case Is <= 75
     strBrake = App.Path & "\BrakeFiles\CastIron\75t railcar brake values.txt"
     Case Is <= 80
     strBrake = App.Path & "\BrakeFiles\CastIron\80.0t railcar brake values.txt"
     Case Is <= 85
     strBrake = App.Path & "\BrakeFiles\CastIron\85t railcar brake values.txt"
     Case Is <= 90
     strBrake = App.Path & "\BrakeFiles\CastIron\90t railcar brake values.txt"
     Case Is <= 95
     strBrake = App.Path & "\BrakeFiles\CastIron\95t railcar brake values.txt"
     Case Is <= 100
     strBrake = App.Path & "\BrakeFiles\CastIron\100t railcar brake values.txt"
     Case Is <= 105
     strBrake = App.Path & "\BrakeFiles\CastIron\105.0t railcar brake values.txt"
     Case Is <= 110
     strBrake = App.Path & "\BrakeFiles\CastIron\110t railcar brake values.txt"
     Case Is <= 115
     strBrake = App.Path & "\BrakeFiles\CastIron\115t railcar brake values.txt"
     Case Is <= 120
     strBrake = App.Path & "\BrakeFiles\CastIron\120t railcar brake values.txt"
     Case Is > 120
     strBrake = App.Path & "\BrakeFiles\CastIron\125t railcar brake values.txt"
    End Select
    ElseIf booAuto = True And flagType = 2 Then
        If Left$(strLoco, 1) = "#" Or Left$(strLoco, 2) = "AI" Or Right$(strLoco, 6) = "AI.eng" Then
        strBrake = App.Path & "\BrakeFiles\CastIron\AI engines brake values.txt"
        Else
        Select Case sngMass
        Case Is < 140
         strBrake = App.Path & "\BrakeFiles\CastIron\Small Driving Engine Brake values.txt"
         Case Is < 200
         strBrake = App.Path & "\BrakeFiles\CastIron\Large Driving Engine Brake values.txt"
         Case Is >= 200
         strBrake = App.Path & "\BrakeFiles\CastIron\200+ Ton Driving Engine Brake values.txt"
         End Select
         End If
     ElseIf booAuto = True And flagType = 5 Then
     strBrake = App.Path & "\BrakeFiles\CastIron\Pax Brake values.txt"
    End If
    End If
End If

Brakes2:
'************************************************


Call ReplaceBrakes(MyString, strBrake, fullpath$)
    


NoBBox2:
Rem **********************************
Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), MyString)
DoEvents



TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If


End Sub


Private Sub Command18_Click()
Dim i As Integer, x As Long, MyString As String, fullpath$
Dim strStart As String, strEnd As String, Y As Long, yy As Long
Dim intLights As Integer, strTemp As String, Z As Integer, ii As Integer
Dim zz As Integer, intStates As Integer

On Error GoTo Errtrap
MousePointer = 11
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
List1.TopIndex = i
List1.Selected(i) = False
MyString = ReadUniFile(fullpath$)
MyString = Replace(MyString, vbTab, " ")
DoEvents
MyString = Replace(MyString, "          ", " ")
MyString = Replace(MyString, "         ", " ")
MyString = Replace(MyString, "        ", " ")
MyString = Replace(MyString, "       ", " ")
MyString = Replace(MyString, "      ", " ")
MyString = Replace(MyString, "     ", " ")
MyString = Replace(MyString, "    ", " ")
MyString = Replace(MyString, "   ", " ")
MyString = Replace(MyString, "  ", " ")
MyString = Replace(MyString, " ", " ")
For ii = 1 To 10
MyString = Replace(MyString, "  (", " (")
Next ii

x = InStr(MyString, "Lights (")
If x = 0 Then GoTo CarryON
Y = InStr(x, MyString, vbCr)
strTemp = Mid$(MyString, x + 8, Y - (x + 8))
strTemp = Trim$(strTemp)
intLights = Val(strTemp)
yy = x + 1
Z = 0

Do
TryAgain:
Y = InStr(yy, MyString, "Light (")
If Y > 0 Then
If Mid$(MyString, Y - 4, 4) = "Head" Then yy = Y + 1: GoTo TryAgain
Z = Z + 1
yy = Y + 1
End If
Loop While Y > 0

If intLights <> Z Then
strReport = strReport & "The Lights entry in " & List1.List(i) & " does not equal the number of light entries, file has not been altered" & vbCrLf
GoTo CarryON
End If

zz = InStr(yy, MyString, "States (")
If zz = 0 Then GoTo CarryOn2
Y = InStr(zz, MyString, vbCr)
strTemp = Mid$(MyString, zz + 8, Y - (zz + 8))
strTemp = Trim$(strTemp)
intStates = Val(strTemp)
Z = 0
Do

Y = InStr(yy, MyString, "State (")
If Y > 0 Then
Z = Z + 1
yy = Y + 1
End If
Loop While Y > 0
Rem ***********************
If intStates <> Z Then
strReport = strReport & "The States entry in " & List1.List(i) & " does not equal the number of State entries, file has not been altered" & vbCrLf
GoTo CarryOn2
End If



CarryOn2:
Z = InStr(yy, MyString, "Radius (")
If Z = 0 Then
strReport = strReport & "Unable to find end of Lights in " & List1.List(i) & " file has not been altered" & vbCrLf
GoTo CarryON
End If
zz = InStr(Z, MyString, "Elevation (")
If zz > Z Then Z = zz
For ii = 1 To 5
Y = InStr(Z, MyString, ")")
Z = Y + 1
Next
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, Y + 1)
MyString = strStart & strEnd
Call WriteUniFile(fullpath$, MyString)
DoEvents
End If
CarryON:
Next i
MousePointer = 0
If strReport <> vbNullString Then
frmReport.Rich1.Text = strReport
 
     frmReport.Show 1
 strReport = vbNullString

End If
Exit Sub
Errtrap:
Call MsgBox("An error - " & Err.Description & " occurred in" & List1.List(i) _
            & vbCrLf & "File not modified." _
            & vbCrLf & "" _
            , vbExclamation, App.Title)
            GoTo CarryON

End Sub

Private Sub Command19_Click()
Dim i As Integer, x As Integer, xx As Integer, j As Integer
Dim strBatText As String, strTemp As String
Dim strShape As String, strWagName As String
Dim NewFile As Integer, intSteam As Integer
Dim X1 As Single, X2 As Single, y1 As Single, y2 As Single, z1 As Single, z2 As _
    Single, x3 As Single, sngMass As Single
Dim strDerail As String, strDerailBuf As String, strMass As String
Dim booMass As Boolean, strDBF As String, strType As String
Dim strLength As String
Dim sLength As Single, dDBF As Double


On Error GoTo Errtrap
MousePointer = 11
Frame2.Visible = False
    If Not DirExists(App.Path & "\TempFiles") Then
       MkDir App.Path & "\TempFiles"
    End If
Kill App.Path & "\TempFiles\*.*"
DoEvents
If Check3.value = 0 Then
strDBF = InputBox("DerailBufferForce will be set to 1000 kN for all items, enter a new value here if you wish to change it")
    If strDBF = vbNullString Then
    strDBF = "1000"
    End If
End If
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
FileCopy fullpath$, fullpath$ & ".bak"
Else
For j = 1 To 50
    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
    Exit For
    End If
Next j
End If
    End If
List1.TopIndex = i
List1.Selected(i) = False
x = InStr(List1.List(i), "Common.")
    If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
    If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
    If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "-NC-")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "SBW-")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "SCN-")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
    If x > 0 Then GoTo CarryON
    If Right$(List1.List(i), 4) <> ".eng" And Right$(List1.List(i), 4) <> ".wag" Then
        Call MsgBox(List1.List(i) & vbCrLf & _
            "Is not an .eng or .wag file and has been ignored.", vbExclamation, App.Title)
        GoTo CarryON
    End If

Call GetFileType(fullpath$, intSteam)
MyString = ReadUniFile(fullpath$)
x = InStr(MyString, "Mass")
    If x = 0 Then
    strReport = strReport & "Mass entry not found in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
x = InStr(MyString, "WagonShape")
    If x = 0 Then
    strReport = strReport & "WagonShape entry not found in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strShape = Mid$(MyString, x + 1, xx - (x + 1))
strShape = Trim$(strShape)
strShape = Replace(strShape, ChrW$(34), "")
    If Right$(strShape, 2) <> ".s" Then
    strReport = strReport & "File " & List1.List(i) & _
        " has an invalid WagonShape entry so could not be processed" & vbCrLf
    GoTo CarryON
    End If

x = InStr(MyString, "Type")
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strType = Mid$(MyString, x + 1, xx - (x + 1))
strType = Trim$(strType)
    If Left$(strType, 1) = ChrW$(34) Then
    strType = Mid$(strType, 2)
    End If
    If Right$(strType, 1) = ChrW$(34) Then
    strType = Left$(strType, Len(strType) - 1)
    End If
    If strType <> "Engine" And strType <> "Tender" And strType <> "Freight" And strType <> "Carriage" Then
    strReport = strReport & fullpath$ & " contains an invalid Type statement. (Should be Engine, Tender, Freight or Carriage)." & vbCrLf
    GoTo CarryON
    End If

MyString = Replace(MyString, "- 0.1", "-0.1")


x = InStrRev(List1.List(i), "\")
strShapePath = MSTSPath & "\Trains\Trainset\" & Left$(List1.List(i), x)
strWagName = MSTSPath & "\Trains\Trainset\" & Mid$(List1.List(i), x + 1)
strPicView = strShapePath & strShape
strSDFile = strPicView & "d"
strPicView = strPicView & ";2"

strBatText = ChrW$(34) & App.Path & "\sviewRR4.exe" & ChrW$(34) & " " & ChrW$(34) & _
    strPicView & ChrW$(34)
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)

DoEvents
NewFile = FreeFile
Open App.Path & "\tempfiles\tempsize.txt" For Input As #NewFile
Input #NewFile, A$
X1 = CSng(A$)
Input #NewFile, A$
y1 = CSng(A$)
Input #NewFile, A$
z1 = CSng(A$)
Input #NewFile, A$
X2 = CSng(A$)
Input #NewFile, A$
y2 = CSng(A$)
Input #NewFile, A$
z2 = CSng(A$)
Close #NewFile

x3 = X2 - X1
x3 = x3 / 2
    If y2 < 2.5 Then
    y2 = 2.5
    End If
strSize = "Size ( " & Round(X2 - X1, 2) & "m " & Round(y2, 2) & "m " & Round(z2 - z1 - 0.4, 2) & "m )"
strInertia = "InertiaTensor ( Box ( " & Round(X2 - X1, 2) & "m " & Round(y2, 2) & "m " & Round(z2 - z1 - 1, 2) & "m ) )"
strSD = "ESD_Bounding_Box ( " & Round(-x3, 2) & " 0.9 " & Round(z1 + 0.5, 2) & " " & Round(x3, 2) & " " & Round(y2, 2) & " " & Round(z2 - 0.5, 2) & " )"
strLength = Round(z2 - z1 - 0.4, 2)
NoBBox:

Text1(6) = List1.List(i)
x = InStr(MyString, "Mass")
    If x = 0 Then
    strReport = strReport & "Mass entry not found in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
If x > 1 Then
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strTemp = Mid$(MyString, x + 1, xx - (x + 1))
strTemp = Trim$(strTemp)
x = InStr(strTemp, ChrW$(34))
    If x > 0 Then
    booMass = True
    strTemp = Replace(strTemp, ChrW$(34), "")
    End If
x = InStr(strTemp, "t")
    If x > 0 And Len(strTemp) > x + 1 Then
    strTemp = Left$(strTemp, x - 1)
    booMass = True
    End If
    If booMass = True Then
    strMass = "Mass ( " & strTemp & "t )"
    End If
    If Right$(strTemp, 1) = "t" Then
    strTemp = Left$(strTemp, Len(strTemp) - 1)
    End If
    strMass = "Mass ( " & strTemp & "t )"
        Text1(1) = strMass
        
sngMass = Val(strTemp)
dDBF = sngMass
sngMass = Round(sngMass * 2.7)
    If sngMass < 1 Then
    strReport = strReport & "Mass entry invalid in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
    
 If Check3.value = 1 Then GoTo NoBBX3
strDerail = "DerailRailForce ( " & Str(sngMass) & "t )"
Text1(3) = strDerail
End If
x = InStr(strDBF, "$mass")
If x <> 0 Then
strDBF = Replace(strDBF, "$mass", Trim(Str(Val(dDBF))))
Call EvalExpression(strDBF)
dDBF = Val(strDBF)
strDBF = Format(dDBF, "###0.00")
End If
strDerailBuf = "DerailBufferForce ( " & strDBF & "kN )"
Text1(4) = strDerailBuf
x = InStr(MyString, "DerailRailForce")
If x = 0 Then
strReport = strReport & "DerailRailForce entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strDerail & strEnd
x = InStr(MyString, "DerailBufferForce")
If x = 0 Then
strReport = strReport & "DerailBufferForce entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strDerailBuf & strEnd

NoBBX3:


If booMass = True Then
booMass = False
End If
x = InStr(MyString, "Mass")
xx = InStr(x, MyString, vbLf)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strMass & vbCrLf & strEnd

If Check3.value = 1 Then GoTo NoBBox2
x = InStr(MyString, "Size")

If x = 0 Then
strReport = strReport & "Size entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strSize & strEnd
Text1(0) = strSize
x = InStr(MyString, "InertiaTensor")
If x = 0 Then
strReport = strReport & "InertiaTensor entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
xx = InStr(x, MyString, ")")
xx = InStr(xx + 1, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strInertia & strEnd
Text1(2) = strInertia

'End If

NoBBox2:
Rem **********************************
sLength = Val(strLength)

Call ReplaceCouplings_Auto(MyString, sLength, intSteam)
DoEvents
Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), MyString)
DoEvents

If Check3.value = 1 Then
GoTo TryAgain
End If
FileCopy strSDFile, strSDFile & ".bak"
MyString = ReadUniFile(strSDFile)
x = InStr(MyString, "ESD_Bounding")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strSD & strEnd
Text1(5) = strSD
Call WriteUniFile(strSDFile, MyString)
DoEvents
'End If

TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:

Resume Next
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If


End Sub

Private Sub Command2_Click()
Unload Me

End Sub


Private Function DirDiver(NewPath As String, DirCount As Integer, Backup As String, intType As Integer) As Integer
'  Recursively search directories from NewPath down...
'  NewPath is searched on this recursion.
'  BackUp is origin of this recursion.
'  DirCount is number of subdirectories in this directory.

Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, Entry As String
Dim retval As Integer, intFile As Integer, x As Integer, xx As Integer

    SearchFlag = True           ' Set flag so the user can interrupt.
    DirDiver = False            ' Set to True if there is an error.
    retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
  
    On Local Error GoTo DirDriverHandler
    DirsToPeek = frmUtils.Dir1(cursouind).ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = frmUtils.Dir1(cursouind).Path                      ' Save old path for next recursion.
        frmUtils.Dir1(cursouind).Path = NewPath
        If frmUtils.Dir1(cursouind).ListCount > 0 Then
            ' Get to the node bottom.
            frmUtils.Dir1(cursouind).Path = frmUtils.Dir1(cursouind).List(DirsToPeek - 1)
            AbandonSearch = DirDiver((frmUtils.Dir1(cursouind).Path), DirCount%, OldPath, intType)
        End If
        ' Go up one level in directories.

        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
        
    Loop
    
    ' Call function to enumerate files.
    If frmUtils.File1(cursouind).ListCount Then
        If Len(frmUtils.Dir1(cursouind).Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = frmUtils.Dir1(cursouind).Path                  ' If at root level, leave as is...
        Else
            ThePath = frmUtils.Dir1(cursouind).Path + "\"            ' Otherwise put "\" before the filename.
        End If

        If Right$(ThePath, 14) = "SpareConsists\" Then GoTo SkipThis
        x = InStrRev(ThePath, "\", Len(ThePath) - 1)
        If Mid$(ThePath, x + 1, 7) = "Cabview" Then
        GoTo SkipThis
        End If
       x = InStr(ThePath, "Trainset")
       x = InStr(x, ThePath, "\")
       x = InStr(x + 1, ThePath, "\")
       xx = InStr(x + 1, ThePath, "\")
       If xx > 0 Then GoTo SkipThis
       
         x = InStrRev(ThePath, "\", Len(ThePath) - 1)
        strFolder = Mid$(ThePath, x + 1)
        For ind = 0 To frmUtils.File1(cursouind).ListCount - 1        ' Add conforming files in this directory to the list box.
            If intType <> 8 Then
            Call GetFileType(ThePath & frmUtils.File1(0).List(ind), intFile)
            
             If intFile = 0 And frmUtils.File1(0).List(ind) <> "default.wag" Then
            strReport = strReport & strFolder & frmUtils.File1(0).List(ind) & " did not contain a valid Type entry" & vbCrLf
            
            GoTo GetAnother
            End If
            End If
            If intType = intFile Or intType = 9 Or intType = 8 Then
            Entry = strFolder & frmUtils.File1(cursouind).List(ind)
            List1.AddItem Entry
            lblCount(0).Caption = Str(Val(lblCount(0).Caption) + 1)
            End If
            DoEvents
            If booAbort = True Then
            SearchFlag = False
            Exit Function
            End If
GetAnother:
        Next ind
SkipThis:
    End If
    If Backup <> vbNullString Then        ' If there is a superior directory, move it.
        frmUtils.Dir1(cursouind).Path = Backup
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiver = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    Else                        ' Otherwise display error message and quit.
        MsgBox Error
        End
    End If
End Function





Private Sub Command20_Click()
Dim i As Integer
On Error GoTo Errtrap
Select Case MsgBox("Please confirm you really wish to DELETE all selected files." _
                   & vbCrLf & "This operation is irreversible." _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
    
For i = List1.ListCount - 1 To 0 Step -1
If List1.Selected(i) = True Then
    If FileExists(MSTSPath & "\Trains\Trainset\" & List1.List(i)) Then
    Kill MSTSPath & "\Trains\Trainset\" & List1.List(i)
    List1.Selected(i) = False
    DoEvents
    End If
End If
CarryON:
Next i
    Case vbNo
Exit Sub
End Select
Exit Sub
Errtrap:
If Err = 75 Then GoTo CarryON
End Sub


Private Sub Command21_Click()
Dim i As Integer, strPath As String, strTemp As String, Y As Integer

strTemp = InputBox("Remove files containing this string from list", "Enter string")
For i = List1.ListCount - 1 To 0 Step -1

strPath = List1.List(i)
Y = InStr(strPath, strTemp)

If Y > 0 Then
List1.RemoveItem (i)
End If
Y = 0
Next i
lblCount(0).Caption = List1.ListCount
End Sub

Private Sub Command22_Click()
Dim i As Integer, x As Integer, xx As Integer, j As Integer
Dim strTemp As String
Dim intSteam As Integer
Dim sngMass As Single
Dim strMass As String
Dim booMass As Boolean, strType As String
Dim strBrake As String, booAuto As Boolean
Dim strLoco As String, Mystring2 As String

booProBrakes = True
booLSD = False
If flagType > 7 Then
Call MsgBox("You have not selected a specific Rolling-Stock type for this option.", vbExclamation, App.Title)
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
List1.Selected(i) = False
End If
Next i
lblCount(1).Caption = List1.SelCount
Exit Sub
End If
On Error GoTo Errtrap
MousePointer = 11
Frame2.Visible = False
If Not DirExists(App.Path & "\TempFiles") Then
   MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents
If flagType = 2 Then
Select Case MsgBox("You are modifying the Brake settings on Diesel Locomotives, do you wish the program to automatically" _
                   & vbCrLf & "select the Brake Values based on the mass of the selected locomotives?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
    
    booAuto = True
    dlgBrakes.Show 1
GoTo Brakes1
    Case vbNo

End Select
End If
If flagType = 4 Then
Select Case MsgBox("You are modifying the Brake settings on Freight Wagons, do you wish the program to automatically" _
                   & vbCrLf & "select the Brake Values based on the mass of the selected wagons?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
    booAuto = True
    dlgBrakes.Show 1
GoTo Brakes1
    Case vbNo

End Select
End If
If flagType = 5 Then
Select Case MsgBox("You are modifying the Brake settings on Passenger Cars, do you wish the program to automatically" _
                   & vbCrLf & "select the Brake Values for you?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
    booAuto = True
    dlgBrakes.Show 1
GoTo Brakes1
    Case vbNo

End Select
End If
CDL1.InitDir = App.Path & "\BrakeFiles_Pro"

CDL1.Filter = "Brake Files (*.txt)|*.txt"
CDL1.DialogTitle = "Select Brake File"

DoEvents
CDL1.FilterIndex = 1
CDL1.Action = 1
CDL1.InitDir = App.Path & "\BrakeFiles_Pro"
DoEvents
strBrake = CDL1.Filename

CDL1.Filename = ""
CDL1.InitDir = ""
If strBrake = vbNullString Then
MousePointer = 0
Exit Sub
End If
Brakes1:
If booIron = False Then
Select Case MsgBox("You are about to change the brake settings for COMPOSITION brakes in" _
                   & vbCrLf & "the selected rolling-stock. Do you really wish to do this?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

    Case vbNo
Exit Sub
End Select

ElseIf booIron = True Then
Select Case MsgBox("You are about to change the brake settings for CAST-IRON brakes in" _
                   & vbCrLf & "the selected rolling-stock. Do you really wish to do this?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

    Case vbNo
Exit Sub
End Select
End If
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    x = InStrRev(fullpath$, "\")
    strLoco = Mid$(fullpath$, x + 1)
    If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
FileCopy fullpath$, fullpath$ & ".bak"
Else
For j = 1 To 50
    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
    Exit For
    End If
Next j
End If
    End If
List1.TopIndex = i
List1.Selected(i) = False
x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
If Right$(List1.List(i), 4) <> ".eng" And Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not an .eng or .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

Call GetFileType(fullpath$, intSteam)
Mystring2 = ReadUniFile(fullpath$)
x = InStr(Mystring2, "Mass")
    If x = 0 Then
    strReport = strReport & "Mass entry not found in " & fullpath$ & " File could not be processed automatically" & vbCrLf
    GoTo CarryON
    End If
x = InStr(Mystring2, "Type")
x = InStr(x + 1, Mystring2, "(")
xx = InStr(x, Mystring2, ")")
strType = Mid$(Mystring2, x + 1, xx - (x + 1))
strType = Trim$(strType)
If Left$(strType, 1) = ChrW$(34) Then
strType = Mid$(strType, 2)
End If
If Right$(strType, 1) = ChrW$(34) Then
strType = Left$(strType, Len(strType) - 1)
End If
If strType <> "Engine" And strType <> "Tender" And strType <> "Freight" And strType <> "Carriage" Then
strReport = strReport & fullpath$ & " contains an invalid Type statement. (Should be Engine, Tender, Freight or Carriage)." & vbCrLf
GoTo CarryON
End If

Mystring2 = Replace(Mystring2, "- 0.1", "-0.1")
Rem ***************** Get Mass *****************
If booAuto = False Then GoTo Brakes2
x = InStr(Mystring2, "Mass")
    If x = 0 Then
    strReport = strReport & "Mass entry not found in " & fullpath$ & " File not processed" & vbCrLf
    GoTo CarryON
    End If
If x > 1 Then
x = InStr(x + 1, Mystring2, "(")
xx = InStr(x, Mystring2, ")")
strTemp = Mid$(Mystring2, x + 1, xx - (x + 1))
strTemp = Trim$(strTemp)
x = InStr(strTemp, ChrW$(34))
    If x > 0 Then
    booMass = True
    strTemp = Replace(strTemp, ChrW$(34), "")
    End If
x = InStr(strTemp, "t")
    If x > 0 And Len(strTemp) > x + 1 Then
    strTemp = Left$(strTemp, x - 1)
    booMass = True
    End If
    If booMass = True Then
    strMass = "Mass ( " & strTemp & "t )"
    End If
    If Right$(strTemp, 1) = "t" Then
    strTemp = Left$(strTemp, Len(strTemp) - 1)
    End If
    sngMass = Val(strTemp)
    If booIron = False And booLSD = False Then
    If booAuto = True And flagType = 4 Then
    Select Case sngMass
     Case Is <= 18.75
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\18.75t railcar brake values.txt"
    Case Is <= 20
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\20t railcar brake values.txt"
     Case Is <= 22.5
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\22.5t railcar brake values.txt"
     Case Is <= 23.75
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\23.75t railcar brake values.txt"
     Case Is <= 25
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\25.0t railcar brake values.txt"
     Case Is <= 27.5
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\27.5t railcar brake values.txt"
     Case Is <= 30
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\30t railcar brake values.txt"
     Case Is <= 31.25
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\31.25t railcar brake values.txt"
     Case Is <= 35
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\35t railcar brake values.txt"
     Case Is <= 39
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\39t railcar brake values.txt"
     Case Is <= 40
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\40t railcar brake values.txt"
     Case Is <= 45
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\45t railcar brake values.txt"
     Case Is <= 50
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\50t railcar brake values.txt"
     Case Is <= 60
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\60t railcar brake values.txt"
     Case Is <= 70
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\70t railcar brake values.txt"
     Case Is <= 75
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\75t railcar brake values.txt"
     Case Is <= 80
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\80.0t railcar brake values.txt"
     Case Is <= 85
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\85t railcar brake values.txt"
     Case Is <= 90
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\90t railcar brake values.txt"
     Case Is <= 95
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\95t railcar brake values.txt"
     Case Is <= 100
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\100t railcar brake values.txt"
     Case Is <= 105
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\105.0t railcar brake values.txt"
     Case Is <= 110
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\110t railcar brake values.txt"
     Case Is <= 115
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\115t railcar brake values.txt"
     Case Is <= 120
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\120t railcar brake values.txt"
     Case Is <= 125
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\125t railcar brake values.txt"
     Case Is <= 130
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\130t railcar brake values.txt"
     Case Is <= 135
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\135t railcar brake values.txt"
     Case Is <= 140
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\140t railcar brake values.txt"
     Case Is <= 145
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\145t railcar brake values.txt"
     Case Is <= 150
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\150t railcar brake values.txt"
     Case Is > 150
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\155t railcar brake values.txt"
    End Select
    ElseIf booAuto = True And flagType = 2 Then
        If Left$(strLoco, 1) = "#" Or Left$(strLoco, 2) = "AI" Or Right$(strLoco, 6) = "AI.eng" Then
        strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\AI engines brake values.txt"
        Else
        Select Case sngMass
        Case Is < 140
         strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\Small Driving Engine Brake values.txt"
         Case Is < 200
         strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\Large Driving Engine Brake values.txt"
         Case Is >= 200
         strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\200+ Ton Driving Engine Brake values.txt"
         End Select
         End If
     ElseIf booAuto = True And flagType = 5 Then
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_Composition\Pax Brake values.txt"
    End If
    Rem ******************************* Load sensing device ****************************
    ElseIf booIron = False And booLSD = True Then
    If booAuto = True And flagType = 4 Then
    Select Case sngMass
     Case Is <= 18.75
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\18.75t railcar brake values.txt"
    Case Is <= 20
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\20t railcar brake values.txt"
     Case Is <= 22.5
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\22.5t railcar brake values.txt"
     Case Is <= 23.75
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\23.75t railcar brake values.txt"
     Case Is <= 25
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\25.0t railcar brake values.txt"
     Case Is <= 27.5
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\27.5t railcar brake values.txt"
     Case Is <= 30
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\30t railcar brake values.txt"
     Case Is <= 31.25
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\31.25t railcar brake values.txt"
     Case Is <= 35
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\35t railcar brake values.txt"
     Case Is <= 39
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\39t railcar brake values.txt"
     Case Is <= 40
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\40t railcar brake values.txt"
     Case Is <= 45
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\45t railcar brake values.txt"
     Case Is <= 50
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\50t railcar brake values.txt"
     Case Is <= 60
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\60t railcar brake values.txt"
     Case Is <= 70
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\70t railcar brake values.txt"
     Case Is <= 75
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\75t railcar brake values.txt"
     Case Is <= 80
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\80.0t railcar brake values.txt"
     Case Is <= 85
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\85t railcar brake values.txt"
     Case Is <= 90
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\90t railcar brake values.txt"
     Case Is <= 95
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\95t railcar brake values.txt"
     Case Is <= 100
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\100t railcar brake values.txt"
     Case Is <= 105
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\105.0t railcar brake values.txt"
     Case Is <= 110
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\110t railcar brake values.txt"
     Case Is <= 115
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\115t railcar brake values.txt"
     Case Is <= 120
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\120t railcar brake values.txt"
     Case Is <= 125
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\125t railcar brake values.txt"
     Case Is <= 130
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\130t railcar brake values.txt"
     Case Is <= 135
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\135t railcar brake values.txt"
     Case Is <= 140
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\140t railcar brake values.txt"
     Case Is <= 145
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\145t railcar brake values.txt"
     Case Is <= 150
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\150t railcar brake values.txt"
     Case Is > 150
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\155t railcar brake values.txt"
    End Select
    ElseIf booAuto = True And flagType = 2 Then
        If Left$(strLoco, 1) = "#" Or Left$(strLoco, 2) = "AI" Or Right$(strLoco, 6) = "AI.eng" Then
        strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\AI engines brake values.txt"
        Else
        Select Case sngMass
        Case Is < 140
         strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\Small Driving Engine Brake values.txt"
         Case Is < 200
         strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\Large Driving Engine Brake values.txt"
         Case Is >= 200
         strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\200+ Ton Driving Engine Brake values.txt"
         End Select
         End If
     ElseIf booAuto = True And flagType = 5 Then
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_LSD\Pax Brake values.txt"
    End If
    Rem ******************************* Iron brakes ************************************
ElseIf booIron = True Then
     If booAuto = True And flagType = 4 Then
    Select Case sngMass
    Case Is <= 5.5
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\5.5t railcar brake values.txt"
     Case Is <= 6.6
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\6.6t railcar brake values.txt"
     Case Is <= 7.75
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\7.75t railcar brake values.txt"
     Case Is <= 8.8
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\8.8t railcar brake values.txt"
     Case Is <= 9.9
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\9.9t railcar brake values.txt"
     Case Is <= 11
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\11t railcar brake values.txt"
     Case Is <= 12.1
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\12.1t railcar brake values.txt"
     Case Is <= 13.2
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\13.2t railcar brake values.txt"
     Case Is <= 14.3
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\14.3t railcar brake values.txt"
     Case Is <= 15.5
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\15.5t railcar brake values.txt"
     Case Is <= 16.6
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\16.6t railcar brake values.txt"
     Case Is <= 17.7
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\17.7t railcar brake values.txt"
    
    
    
    Rem ********************* OK
     Case Is <= 18.75
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\18.75t railcar brake values.txt"
    Case Is <= 20
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\20t railcar brake values.txt"
     Case Is <= 22.5
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\22.5t railcar brake values.txt"
     Case Is <= 23.75
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\23.75t railcar brake values.txt"
     Case Is <= 25
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\25.0t railcar brake values.txt"
     Case Is <= 27.5
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\27.5t railcar brake values.txt"
     Case Is <= 30
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\30t railcar brake values.txt"
     Case Is <= 31.25
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\31.25t railcar brake values.txt"
     Case Is <= 35
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\35t railcar brake values.txt"
     Case Is <= 39
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\39t railcar brake values.txt"
     Case Is <= 40
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\40t railcar brake values.txt"
     Case Is <= 45
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\45t railcar brake values.txt"
     Case Is <= 50
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\50t railcar brake values.txt"
     Case Is <= 55
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\55t railcar brake values.txt"
     Rem ************** OK from here down *************************************
     Case Is <= 60
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\60t railcar brake values.txt"
     Case Is <= 70
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\70t railcar brake values.txt"
     Case Is <= 75
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\75t railcar brake values.txt"
     Case Is <= 80
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\80.0t railcar brake values.txt"
     Case Is <= 85
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\85t railcar brake values.txt"
     Case Is <= 90
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\90t railcar brake values.txt"
     Case Is <= 95
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\95t railcar brake values.txt"
     Case Is <= 100
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\100t railcar brake values.txt"
     Case Is <= 105
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\105.0t railcar brake values.txt"
     Case Is <= 110
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\110t railcar brake values.txt"
     Case Is <= 115
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\115t railcar brake values.txt"
     Case Is <= 120
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\120t railcar brake values.txt"
     Case Is > 120
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\125t railcar brake values.txt"
    End Select
    ElseIf booAuto = True And flagType = 2 Then
        If Left$(strLoco, 1) = "#" Or Left$(strLoco, 2) = "AI" Or Right$(strLoco, 6) = "AI.eng" Then
        strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\AI engines brake values.txt"
        Else
        Select Case sngMass
        Case Is < 140
         strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\Small Driving Engine Brake values.txt"
         Case Is < 200
         strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\Large Driving Engine Brake values.txt"
         Case Is >= 200
         strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\200+ Ton Driving Engine Brake values.txt"
         End Select
         End If
     ElseIf booAuto = True And flagType = 5 Then
     strBrake = App.Path & "\BrakeFiles_Pro\Pro_CastIron\Pax Brake values.txt"
    End If
    End If
End If

Brakes2:
'************************************************


Call ReplaceBrakes(Mystring2, strBrake, fullpath$)
    


NoBBox2:
Rem **********************************
Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), Mystring2)
DoEvents



TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If


End Sub

Private Sub Command23_Click()
Dim i As Integer, x As Integer, j As Integer

Dim intSteam As Integer
Dim strBrake As String
Dim strLoco As String

booProBrakes = False
booLSD = False

If flagType <> 2 Then
Call MsgBox("This option is only applicable to certain Diesel Locomotives.", vbExclamation, App.Title)
Exit Sub
End If

On Error GoTo Errtrap
MousePointer = 11
Frame2.Visible = False
If Not DirExists(App.Path & "\TempFiles") Then
   MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents

CDL1.Filter = "Brake Files (*.txt)|*.txt"
CDL1.DialogTitle = "Select Brake File"
CDL1.InitDir = App.Path & "\BrakeFiles_Pro\Pro_24RL"
CDL1.FilterIndex = 1
CDL1.Action = 1
strBrake = CDL1.Filename

CDL1.Filename = ""
CDL1.InitDir = ""
If strBrake = vbNullString Then
MousePointer = 0
Exit Sub
End If
Brakes1:
Select Case MsgBox("You are about to change the brakes to Pro_24RL values in the selected stock." _
                   & vbCrLf & "Do you really wish to do this?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

    Case vbNo
Exit Sub
End Select
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    x = InStrRev(fullpath$, "\")
    strLoco = Mid$(fullpath$, x + 1)
    If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
FileCopy fullpath$, fullpath$ & ".bak"
Else
For j = 1 To 50
    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
    Exit For
    End If
Next j
End If
    End If
List1.TopIndex = i
List1.Selected(i) = False
x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
If Right$(List1.List(i), 4) <> ".eng" And Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not an .eng or .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

Call GetFileType(fullpath$, intSteam)
MyString = ReadUniFile(fullpath$)


MyString = Replace(MyString, "- 0.1", "-0.1")

Brakes2:



Call ReplaceAirBrakes(MyString, strBrake, fullpath$)
    


NoBBox2:

Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), MyString)
DoEvents



TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If


End Sub

Private Sub Command24_Click()
Dim i As Integer, x As Integer, xx As Integer
Dim strTemp As String
Dim strBrake As String
Dim strLoco As String
Dim strStart As String, strEnd As String, Mystring2 As String
Dim Y As Long, yy As Long, strTemp2 As String, j As Integer

booProBrakes = False
booLSD = False

If flagType <> 2 Then
Call MsgBox("This option is only applicable to certain Diesel Locomotives.", vbExclamation, App.Title)
Exit Sub
End If

On Error GoTo Errtrap
MousePointer = 11
Frame2.Visible = False
If Not DirExists(App.Path & "\TempFiles") Then
   MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents

CDL1.Filter = "Brake Files (*.txt)|*.txt"
CDL1.DialogTitle = "Select Throttle_Brake Percentages File"
CDL1.InitDir = App.Path & "\BrakeFiles_Pro\Pro_Throttle_Brake Percentages"
CDL1.FilterIndex = 1
CDL1.Action = 1
strBrake = CDL1.Filename
If strBrake = vbNullString Then
MousePointer = 0
Exit Sub
End If
Brakes1:
Select Case MsgBox("You are about to change the Throttle-Brake percentage values in the selected Locomotive(s)." _
                   & vbCrLf & "Do you really wish to do this?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

    Case vbNo
Exit Sub
End Select
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    x = InStrRev(fullpath$, "\")
    strLoco = Mid$(fullpath$, x + 1)
    If Check2.value = 1 Then

If Not FileExists(fullpath$ & ".bak") Then
FileCopy fullpath$, fullpath$ & ".bak"
Else
For j = 1 To 50
    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
    Exit For
    End If
Next j
End If
    End If
List1.TopIndex = i
List1.Selected(i) = False

If Right$(List1.List(i), 4) <> ".eng" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not an .eng file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If


MyString = ReadUniFile(fullpath$)
MyString = Replace(MyString, "- 0.1", "-0.1")

Mystring2 = ReadUniFile(strBrake)
x = InStr(Mystring2, "EngineBrakesControllerMinSystemPressure")
xx = InStr(x, Mystring2, ")")
strTemp = Mid(Mystring2, x, xx - (x - 1))
Y = InStr(Mystring2, "EngineBrakesControllerMinSystemPressure")
yy = InStr(Y, Mystring2, ")")
strTemp2 = Mid(Mystring2, Y, yy - (Y - 1))
If strTemp <> strTemp2 Then
MyString = Replace(MyString, strTemp2, strTemp)
End If

x = InStr(Mystring2, "TrainBrakesControllerMinSystemPressure")
xx = InStr(x, Mystring2, ")")
strTemp = Mid(Mystring2, x, xx - (x - 1))
Y = InStr(Mystring2, "TrainBrakesControllerMinSystemPressure")
yy = InStr(Y, Mystring2, ")")
strTemp2 = Mid(Mystring2, Y, yy - (Y - 1))
If strTemp <> strTemp2 Then
MyString = Replace(MyString, strTemp2, strTemp)
End If

x = InStr(Mystring2, "EngineBrakesControllerMaxSystemPressure")
xx = InStr(x, Mystring2, ")")
strTemp = Mid(Mystring2, x, xx - (x - 1))
Y = InStr(Mystring2, "EngineBrakesControllerMaxSystemPressure")
yy = InStr(Y, Mystring2, ")")
strTemp2 = Mid(Mystring2, Y, yy - (Y - 1))
If strTemp <> strTemp2 Then
MyString = Replace(MyString, strTemp2, strTemp)
End If

x = InStr(Mystring2, "TrainBrakesControllerMaxSystemPressure")
xx = InStr(x, Mystring2, ")")
strTemp = Mid(Mystring2, x, xx - (x - 1))
Y = InStr(Mystring2, "TrainBrakesControllerMaxSystemPressure")
yy = InStr(Y, Mystring2, ")")
strTemp2 = Mid(Mystring2, Y, yy - (Y - 1))
If strTemp <> strTemp2 Then
MyString = Replace(MyString, strTemp2, strTemp)
End If
DoEvents
x = InStr(Mystring2, "Comment")
strTemp = Mid(Mystring2, x)
Y = InStr(MyString, "EngineControllers")
If Mid(MyString, Y - 6, 5) = "Brake" Then

Y = InStr(Y + 1, MyString, "EngineControllers")
End If
yy = InStr(MyString, "Brake_Train")

If yy < Y Then
Call MsgBox("Unable to fix " & strBrake _
            & vbCrLf & "Format unsuitable for automatic update, please fix manually." _
            , vbExclamation, App.Title)

GoTo CarryON
End If
strStart = Left(MyString, Y - 1)
strEnd = Mid(MyString, yy)
MyString = strStart & strTemp & strEnd
strStart = ""
strEnd = ""
strTemp = ""
Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), MyString)
DoEvents



TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If

End Sub


Private Sub Command25_Click()
Dim strStart As String, strEnd As String, MyString As String, i As Integer
Dim fullpath$, x As Long

Text4.Text = ""
dlgFriction.Show 1
DoEvents
If Text4.Text = "" Then Exit Sub
Select Case MsgBox("You are about to change the Friction values in the selected Wagons." _
                   & vbCrLf & "Do you really wish to do this?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

    Case vbNo
Exit Sub
End Select
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    x = InStrRev(fullpath$, "\")
    strLoco = Mid$(fullpath$, x + 1)
    If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
FileCopy fullpath$, fullpath$ & ".bak"
Else
For j = 1 To 50
    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
    Exit For
    End If
Next j
End If
    End If
List1.TopIndex = i
List1.Selected(i) = False
x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
If Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not a .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

MyString = ReadUniFile(fullpath$)
MyString = Replace(MyString, "- 0.1", "-0.1")

x = InStr(MyString, "Friction (")
If x = 0 Then
Call MsgBox("No 'Friction (' entry was found in" _
            & vbCrLf & fullpath$ _
            , vbExclamation, App.Title)


Exit Sub
End If

strStart = Left(MyString, x - 1)
x = InStr(x, MyString, vbCr)
x = InStr(x + 3, MyString, vbCr)
strEnd = Mid(MyString, x)
MyString = strStart & Text4.Text & strEnd



Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), MyString)
DoEvents



TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If

End Sub

Private Sub Command26_Click()
Dim MyString As String, x As Integer, xx As Integer, strStart As String
Dim strEnd As String, strSD As String


If Not FileExists(App.Path & "\Formulae.txt") Then
Call MsgBox("No Formulae.txt file found in your Route_Riter folder?" _
            & vbCrLf & "" _
            , vbExclamation, App.Title)

Exit Sub
End If

FileCopy strSDFile, strSDFile & ".bak"
MyString = ReadUniFile(strSDFile)
x = InStr(MyString, "ESD_Bounding")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
strSD = Text1(5)
MyString = strStart & strSD & strEnd
Call WriteUniFile(strSDFile, MyString)
DoEvents
Rem **************** Update .wag/.eng file
If booEdited = True Then
booEdited = False
FileCopy fullpath$, fullpath$ & ".bak"
MyString = ReadUniFile(fullpath$)
x = InStr(MyString, "Size")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(0) & strEnd

x = InStr(MyString, "InertiaTensor")
xx = InStr(x, MyString, ")")
xx = InStr(xx + 1, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(2) & strEnd

x = InStr(MyString, "DerailRailForce")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(3) & strEnd

x = InStr(MyString, "DerailBufferForce")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(4) & strEnd

x = InStr(MyString, "Mass")
xx = InStr(x, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & Text1(1) & vbCrLf & strEnd
Call WriteUniFile(fullpath$, MyString)
DoEvents
End If
Command14.Enabled = True

End Sub

Private Sub Command27_Click()
Dim i As Integer, x As Integer, xx As Integer, j As Integer
Dim strBatText As String
Dim strShape As String, strWagName As String
Dim NewFile As Integer, intSteam As Integer
Dim X1 As Single, X2 As Single, y1 As Single, y2 As Single, z1 As Single, z2 As _
    Single, x3 As Single
Dim strType As String
Dim strLength As String
Dim TrueY As Single, TrueZ As Single, strCOG As String, strTrueY As String, strTrueZ As String

If flagType > 7 Then
Call MsgBox("You have not selected a specific Rolling-Stock type for this option.", vbExclamation, App.Title)
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
List1.Selected(i) = False
End If
Next i
lblCount(1).Caption = List1.SelCount
Exit Sub
End If
On Error GoTo Errtrap
MousePointer = 11
Frame2.Visible = False
If Not DirExists(App.Path & "\TempFiles") Then
   MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents


For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
FileCopy fullpath$, fullpath$ & ".bak"
Else
For j = 1 To 50
    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
    Exit For
    End If
Next j
End If
    End If
List1.TopIndex = i
List1.Selected(i) = False
x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "-NC-")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "SBW-")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "SCN-")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
If Right$(List1.List(i), 4) <> ".eng" And Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not an .eng or .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

Call GetFileType(fullpath$, intSteam)
MyString = ReadUniFile(fullpath$)

x = InStr(MyString, "WagonShape")
If x = 0 Then
strReport = strReport & "WagonShape entry not found in " & fullpath$ & " File not processed" & vbCrLf
GoTo CarryON
End If
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strShape = Mid$(MyString, x + 1, xx - (x + 1))
strShape = Trim$(strShape)
strShape = Replace(strShape, ChrW$(34), "")
If Right$(strShape, 2) <> ".s" Then
strReport = strReport & "File " & List1.List(i) & _
    " has an invalid WagonShape entry so could not be processed" & vbCrLf
GoTo CarryON
End If

x = InStr(MyString, "Type")
x = InStr(x + 1, MyString, "(")
xx = InStr(x, MyString, ")")
strType = Mid$(MyString, x + 1, xx - (x + 1))
strType = Trim$(strType)
If Left$(strType, 1) = ChrW$(34) Then
strType = Mid$(strType, 2)
End If
If Right$(strType, 1) = ChrW$(34) Then
strType = Left$(strType, Len(strType) - 1)
End If
If strType <> "Engine" And strType <> "Tender" And strType <> "Freight" And strType <> "Carriage" Then
strReport = strReport & fullpath$ & " contains an invalid Type statement. (Should be Engine, Tender, Freight or Carriage)." & vbCrLf
GoTo CarryON
End If

MyString = Replace(MyString, "- 0.1", "-0.1")

'If Check3.Value = 1 And booShort = False Then GoTo NoBBox
x = InStrRev(List1.List(i), "\")
strShapePath = MSTSPath & "\Trains\Trainset\" & Left$(List1.List(i), x)
strWagName = MSTSPath & "\Trains\Trainset\" & Mid$(List1.List(i), x + 1)
strPicView = strShapePath & strShape
strSDFile = strPicView & "d"
strPicView = strPicView & ";2"

strBatText = ChrW$(34) & App.Path & "\sviewRR4.exe" & ChrW$(34) & " " & ChrW$(34) & _
    strPicView & ChrW$(34)
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)

DoEvents
NewFile = FreeFile

Open App.Path & "\tempfiles\tempsize.txt" For Input As #NewFile
Input #NewFile, A$
X1 = CSng(A$)
Input #NewFile, A$
y1 = CSng(A$)
Input #NewFile, A$
A$ = Format(A$, "0.000")
z1 = CSng(A$)
Input #NewFile, A$
X2 = CSng(A$)
Input #NewFile, A$
y2 = CSng(A$)
Input #NewFile, A$
A$ = Format(A$, "0.000")
z2 = CSng(A$)
Close #NewFile

x3 = X2 - X1
x3 = x3 / 2
If y2 < 2.5 Then
y2 = 2.5
End If
strSize = "Size ( " & Round(X2 - X1, 2) & "m " & Round(y2, 2) & "m " & Round(z2 - z1 - 0.4, 2) & "m )"
strInertia = "InertiaTensor ( Box ( " & Round(X2 - X1, 2) & "m " & Round(y2, 2) & "m " & Round(z2 - z1 - 1, 2) & "m ) )"
strSD = "ESD_Bounding_Box ( " & Round(-x3, 2) & " 0.9 " & Round(z1 + 0.5, 2) & " " & Round(x3, 2) & " " & Round(y2, 2) & " " & Round(z2 - 0.5, 2) & " )"
strLength = Round(z2 - z1 - 0.4, 2)
TrueY = Round((y2 / 2), 3)
TrueZ = Round(((z1 + z2) / 2), 3)
strTrueY = Trim(Str(TrueY))
strTrueZ = Trim(Str(TrueZ))
strTrueY = Format(strTrueY, "0.000")
strTrueZ = Format(strTrueZ, "0.000")

strCOG = "CentreOfGravity ( 0m " & strTrueY & "m " & strTrueZ & "m )"


x = InStr(MyString, "CentreOfGravity")
If x > 0 Then
xx = InStr(x + 1, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & strCOG & strEnd
Else
x = InStr(MyString, "Size")
xx = InStr(x, MyString, ")")
strStart = Left(MyString, xx + 1)
strEnd = Mid(MyString, xx + 1)
MyString = strStart & vbCrLf & strCOG & strEnd

End If
Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), MyString)
DoEvents


TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If


End Sub

Private Sub Command28_Click()
Dim MyString As String, i As Integer
Dim fullpath$, x As Long


Select Case MsgBox("You are about to Enable Shunting on the selected Wagons." _
                   & vbCrLf & "Do you really wish to do this? It has no effect on their performance" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

    Case vbNo
Exit Sub
End Select
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    x = InStrRev(fullpath$, "\")
    strLoco = Mid$(fullpath$, x + 1)
'    If Check2.Value = 1 Then
'    If Not FileExists(fullpath$ & ".bak") Then
'FileCopy fullpath$, fullpath$ & ".bak"
'Else
'For j = 1 To 50
'    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
'    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
'    Exit For
'    End If
'Next j
'End If
'    End If
List1.TopIndex = i
List1.Selected(i) = False
x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
If Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not a .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

MyString = ReadUniFile(fullpath$)

x = InStr(MyString, "ShuntMaxBrakeForce (")
If x = 0 Then
strReport = strReport & "Wagon has not been converted for Shunting" _
            & vbCrLf & fullpath$ & vbCrLf
            GoTo CarryON



End If

MyString = Replace(MyString, "MaxBrakeForce", "RoadMaxBrakeForce", 1, 1)
DoEvents
MyString = Replace(MyString, "ShuntMaxBrakeForce", "MaxBrakeForce", 1, 1)
DoEvents



Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), MyString)
DoEvents



TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If

End Sub

Private Sub Command29_Click()
Dim MyString As String, i As Integer, strStart As String, strEnd As String
Dim fullpath$, x As Long


Select Case MsgBox("You are about to Disable Shunting on the selected Wagons." _
                   & vbCrLf & "Do you really wish to do this? " _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

    Case vbNo
Exit Sub
End Select
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    x = InStrRev(fullpath$, "\")
    strLoco = Mid$(fullpath$, x + 1)
'    If Check2.Value = 1 Then
'    If Not FileExists(fullpath$ & ".bak") Then
'FileCopy fullpath$, fullpath$ & ".bak"
'Else
'For j = 1 To 50
'    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
'    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
'    Exit For
'    End If
'Next j
'End If
'    End If
List1.TopIndex = i
List1.Selected(i) = False
x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
If Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not a .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

MyString = ReadUniFile(fullpath$)

x = InStr(MyString, "RoadMaxBrakeForce (")
If x = 0 Then
strReport = strReport & "Wagon has not been converted for Shunting" _
            & vbCrLf & fullpath$ & vbCrLf
            GoTo CarryON



End If

MyString = Replace(MyString, "RoadMaxBrakeForce", "MaxBrakeForce", 1, 1)
DoEvents
x = InStr(MyString, "MaxBrakeForce")
x = InStr(x + 5, MyString, "MaxBrakeForce")
strStart = Left(MyString, x - 1)
strEnd = Mid(MyString, x)
MyString = strStart & "Shunt" & strEnd

DoEvents



Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), MyString)
DoEvents



TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If

End Sub


Private Sub Command3_Click()
Dim strKey As String, i As Integer, fullpath$
Dim x As Long, xx As Long, strStart As String, strEnd As String, Y As Long
Dim strInsert As String


strKey = cbEng.Text
For i = 1 To intIgnore
If strKey = EngIgnore(i) Then
Call MsgBox("The selected KEY is a multi-line key, and can not be edited with this program." _
            & vbCrLf & "Suggest you edit manually." _
            , vbExclamation, App.Title)

Exit Sub
End If
Next i
strInsert = cbInsert.Text
If Option1(0).value = True And Check1.value = 0 Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)

If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
MyString = ReadUniFile(fullpath$)
x = InStr(MyString, strKey)
If x = 0 Then GoTo TryAnother
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strEnd
Call WriteUniFile(fullpath$, MyString)
End If
TryAnother:
Next i
End If
If Option1(1).value = True And Check1.value = 0 Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
MyString = ReadUniFile(fullpath$)
Y = InStr(MyString, strKey)
If Y = 0 Then GoTo TryAnother2
x = InStr(Y + 5, MyString, strKey)
If x = 0 Then GoTo TryAnother2
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strEnd
Call WriteUniFile(fullpath$, MyString)
End If
TryAnother2:
Next i
End If
If Option1(2).value = True And Check1.value = 0 Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
MyString = ReadUniFile(fullpath$)
x = 1
FindMore:
x = InStr(x, MyString, strKey)
If x = 0 Then GoTo TryAnother3
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strEnd
Call WriteUniFile(fullpath$, MyString)
x = x + 5
GoTo FindMore
End If
TryAnother3:
Next i
End If
Rem******************Only in selected area***************
If Option1(0).value = True And Check1.value = 1 Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
MyString = ReadUniFile(fullpath$)
Y = InStr(MyString, strInsert & " (")
x = InStr(Y, MyString, strKey)
If x = 0 Then GoTo TryAnother4
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strEnd
Call WriteUniFile(fullpath$, MyString)
End If
TryAnother4:
Next i
End If
If Option1(1).value = True And Check1.value = 1 Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
MyString = ReadUniFile(fullpath$)
Y = InStr(MyString, strInsert & " (")
Y = InStr(Y + 5, MyString, strInsert & " (")
If Y = 0 Then GoTo TryAnother5
x = InStr(Y, MyString, strKey)
If x = 0 Then GoTo TryAnother5
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strEnd
Call WriteUniFile(fullpath$, MyString)
End If
TryAnother5:
Next i
End If
If Option1(2).value = True And Check1.value = 1 Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
MyString = ReadUniFile(fullpath$)
Y = 1
FindMore2:
Y = InStr(Y, MyString, strInsert & " (")
x = InStr(Y, MyString, strKey)
If x = 0 Then GoTo TryAnother6
xx = InStr(x, MyString, vbCr)
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx)
MyString = strStart & strEnd
Call WriteUniFile(fullpath$, MyString)
Y = Y + 5
GoTo FindMore2
End If
TryAnother6:
Next i
End If
txtValue.Text = "All Changes Made"
End Sub

Private Sub Command30_Click()
Dim strStart As String, strEnd As String, MyString As String, i As Integer
Dim fullpath$, x As Long


Select Case MsgBox("You are about to add a ShuntMaxBrakeForce entry to the selected Wagons." _
                   & vbCrLf & "Do you really wish to do this? It has no effect on their performance" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

    Case vbNo
Exit Sub
End Select
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = False Then GoTo CarryON
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
    x = InStrRev(fullpath$, "\")
    strLoco = Mid$(fullpath$, x + 1)

    End If
List1.TopIndex = i
List1.Selected(i) = False
x = InStr(List1.List(i), "Common.")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Default.wag")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Invisocar")
If x > 0 Then GoTo CarryON
x = InStr(List1.List(i), "Sample")
If x > 0 Then GoTo CarryON
If Right$(List1.List(i), 4) <> ".wag" Then
    Call MsgBox(List1.List(i) & vbCrLf & _
        "Is not a .wag file and has been ignored.", vbExclamation, App.Title)
    GoTo CarryON
End If

MyString = ReadUniFile(fullpath$)
x = InStr(MyString, "ShuntMaxBrakeForce")
If x > 0 Then GoTo CarryON

MyString = Replace(MyString, "MaxBrakeForce(", "MaxBrakeForce (")

x = InStr(MyString, "MaxBrakeForce (")
If x = 0 Then

strReport = strReport & "No 'MaxBrakeForce (' entry was found in" _
            & vbCrLf & fullpath$ & vbCrLf
            GoTo CarryON
End If
x = InStr(x, MyString, ")")

strStart = Left(MyString, x)

strEnd = Mid(MyString, x + 1)
MyString = strStart & vbCrLf & vbTab & "ShuntMaxBrakeForce ( 0.0kN )" & strEnd

    If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
FileCopy fullpath$, fullpath$ & ".bak"
Else
For j = 1 To 50
    If Not FileExists(fullpath$ & ".bak" & Trim(Str(j))) Then
    FileCopy fullpath$, fullpath$ & ".bak" & Trim(Str(j))
    Exit For
    End If
Next j
End If

Call WriteUniFile(MSTSPath & "\Trains\Trainset\" & List1.List(i), MyString)
DoEvents



TryAgain:

 End If
DoEvents
CarryON:
Set fLoad = Nothing
If booAbort = True Then
booAbort = False
i = List1.ListCount - 1
End If
DoEvents

Next i
MousePointer = 0
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
   
   
Exit Sub
Errtrap:
If Err = 53 Then
Rem ***** tempfiles is empty ***********
Resume Next
End If

End Sub

Private Sub Command4_Click()
Dim strKey As String, i As Integer, fullpath$
Dim x As Long, xx As Long, strStart As String, strEnd As String
Dim strValue As String, xy As Long, Y As Long, strInsert As String

strKey = cbEng.Text
For i = 1 To intIgnore
If strKey = EngIgnore(i) Then
Call MsgBox("The selected KEY is a multi-line key, and can not be edited with this program." _
            & vbCrLf & "Suggest you edit manually." _
            , vbExclamation, App.Title)

Exit Sub
End If
Next i
strValue = txtValue.Text
strInsert = cbInsert.Text
If strValue = vbNullString Then Exit Sub
strValue = Trim$(strValue)
If Left$(strValue, 1) <> "(" Then
strValue = "( " & strValue
End If
If Right$(strValue, 1) <> ")" Then
strValue = strValue & " )"
End If
If Option1(0).value = True And Check1.value = 0 Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
    MyString = ReadUniFile(fullpath$)
    x = InStr(MyString, strKey)
    If x = 0 Then GoTo TryAnother
    xx = InStr(x, MyString, "(")
    xy = InStr(xx, MyString, vbCr)
    strStart = Left$(MyString, xx - 1)
    strEnd = Mid$(MyString, xy)
    MyString = strStart & strValue & strEnd
    Call WriteUniFile(fullpath$, MyString)
End If
TryAnother:
Next i
End If
If Option1(1).value = True And Check1.value = 0 Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
MyString = ReadUniFile(fullpath$)
x = InStr(MyString, strKey)
x = InStr(x + 5, MyString, strKey)
If x = 0 Then GoTo TryAnother2
xx = InStr(x, MyString, "(")
xy = InStr(xx, MyString, vbCr)
strStart = Left$(MyString, xx - 1)
strEnd = Mid$(MyString, xy)
MyString = strStart & strValue & strEnd
Call WriteUniFile(fullpath$, MyString)
End If
TryAnother2:
Next i
End If
If Option1(2).value = True And Check1.value = 0 Then
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
MyString = ReadUniFile(fullpath$)
x = 1
FindMore:
x = InStr(x, MyString, strKey)
If x = 0 Then GoTo TryAnother3
xx = InStr(x, MyString, "(")
xy = InStr(xx, MyString, vbCr)
strStart = Left$(MyString, xx - 1)
strEnd = Mid$(MyString, xy)
MyString = strStart & strValue & strEnd
Call WriteUniFile(fullpath$, MyString)
x = x + 5
GoTo FindMore
End If
TryAnother3:
Next i
End If
Rem ********************** Change Specific Item ************
If Option1(0).value = True And Check1.value = 1 Then
If strInsert = vbNullString Then
Call MsgBox("You do not appear to have selected an item to change, but the" _
            & vbCrLf & "Check Box has been selected?" _
            , vbExclamation, App.Title)

Exit Sub
End If
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
    MyString = ReadUniFile(fullpath$)
    Y = InStr(MyString, strInsert & " (")
    x = InStr(Y, MyString, strKey)
    If x = 0 Then GoTo TryAnother4
    xx = InStr(x, MyString, "(")
    xy = InStr(xx, MyString, vbCr)
    strStart = Left$(MyString, xx - 1)
    strEnd = Mid$(MyString, xy)
    MyString = strStart & strValue & strEnd
    Call WriteUniFile(fullpath$, MyString)
End If
TryAnother4:
Next i
End If
If Option1(1).value = True And Check1.value = 1 Then
If strInsert = vbNullString Then
Call MsgBox("You do not appear to have selected an item to change, but the" _
            & vbCrLf & "Check Box has been selected?" _
            , vbExclamation, App.Title)

Exit Sub
End If
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
MyString = ReadUniFile(fullpath$)
Y = InStr(MyString, strInsert & " (")
Y = InStr(Y + 5, strInsert & " (")
If Y = 0 Then GoTo TryAnother5
x = InStr(Y, MyString, strKey)
If x = 0 Then GoTo TryAnother5
xx = InStr(x, MyString, "(")
xy = InStr(xx, MyString, vbCr)
strStart = Left$(MyString, xx - 1)
strEnd = Mid$(MyString, xy)
MyString = strStart & strValue & strEnd
Call WriteUniFile(fullpath$, MyString)
End If
TryAnother5:
Next i
End If
If Option1(2).value = True And Check1.value = 1 Then
If strInsert = vbNullString Then
Call MsgBox("You do not appear to have selected an item to change, but the" _
            & vbCrLf & "Check Box has been selected?" _
            , vbExclamation, App.Title)

Exit Sub
End If
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
MyString = ReadUniFile(fullpath$)
Y = 1
FindMore2:
Y = InStr(Y, MyString, strInsert & " (")
If Y = 0 Then GoTo TryAnother6
x = InStr(Y, MyString, strKey)

xx = InStr(x, MyString, "(")
xy = InStr(xx, MyString, vbCr)
strStart = Left$(MyString, xx - 1)
strEnd = Mid$(MyString, xy)
MyString = strStart & strValue & strEnd
Call WriteUniFile(fullpath$, MyString)
Y = Y + 5
GoTo FindMore2
End If
TryAnother6:
Next i
End If
txtValue.Text = "All Changes Made"
End Sub


Private Sub Command5_Click()
Dim strKey As String, i As Integer, fullpath$
Dim x As Long, strStart As String, strEnd As String
Dim strInsert As String, strValue As String

strKey = cbEng.Text
For i = 1 To intIgnore
If strKey = EngIgnore(i) Then
Call MsgBox("The selected KEY is a multi-line key, and can not be edited with this program." _
            & vbCrLf & "Suggest you edit manually." _
            , vbExclamation, App.Title)

Exit Sub
End If
Next i
strInsert = cbInsert.Text
If strInsert = vbNullString Then Exit Sub

strValue = txtValue.Text
If strValue = vbNullString Then
Select Case MsgBox("You have not provided a value for this KeyWord?" _
                   & vbCrLf & "If this Key does not have parameters, then press OK, else Cancel." _
                   , vbOKCancel Or vbExclamation Or vbDefaultButton1, "No Value Given?")

    Case vbOK
Rem Carry On
    Case vbCancel
Exit Sub
End Select
End If
strValue = Trim$(strValue)
If Left$(strValue, 1) <> "(" Then
strValue = "( " & strValue
End If
If Right$(strValue, 1) <> ")" Then
strValue = strValue & " )"
End If
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
fullpath$ = MSTSPath & "\Trains\Trainset\" & List1.List(i)
If Check2.value = 1 Then
    If Not FileExists(fullpath$ & ".bak") Then
    FileCopy fullpath$, fullpath$ & ".bak"
    ElseIf Not FileExists(fullpath$ & ".bak1") Then
    FileCopy fullpath$, fullpath$ & ".bak1"
    ElseIf Not FileExists(fullpath$ & ".bak2") Then
    FileCopy fullpath$, fullpath$ & ".bak2"
    ElseIf Not FileExists(fullpath$ & ".bak3") Then
    FileCopy fullpath$, fullpath$ & ".bak3"
    ElseIf Not FileExists(fullpath$ & ".bak4") Then
    FileCopy fullpath$, fullpath$ & ".bak4"
    End If
    End If
MyString = ReadUniFile(fullpath$)
x = InStr(MyString, strInsert)
If x = 0 Then GoTo TryAnother

strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, x)
MyString = strStart & strKey & " " & strValue & vbCrLf & strEnd
Call WriteUniFile(fullpath$, MyString)
End If
TryAnother:
Next i
txtValue.Text = "All Changes Made"
End Sub

Private Sub Command6_Click()
Dim i As Integer

For i = 0 To List1.ListCount - 1
If List1.Selected(i) = False Then
List1.Selected(i) = True
End If
Next i
lblCount(1).Caption = List1.SelCount
End Sub

Private Sub Command7_Click()
Dim i As Integer

For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
List1.Selected(i) = False
End If
Next i
lblCount(1).Caption = List1.SelCount
End Sub


Private Sub Command8_Click()
Dim i As Integer

For i = List1.ListCount - 1 To 0 Step -1
List1.RemoveItem i
Next i
lblCount(1).Caption = List1.SelCount
End Sub

Private Sub Command9_Click()
Dim i As Integer, strEng As String, strPath As String

For i = List1.ListCount - 1 To 0 Step -1

strPath = List1.List(i)
x = InStrRev(strPath, "\")
strEng = Mid$(strPath, x + 1)
If Left$(strEng, 1) = "#" Or Left$(strEng, 1) = "$" Or Left$(strEng, 2) = "AI" Or Left$(strEng, 4) = "Dead" Then
List1.RemoveItem (i)
End If
'End If
Next i
lblCount(0).Caption = List1.ListCount
End Sub

Private Sub Form_Load()
Dim strTemp As String, i As Integer, A$, j As Integer

Close
Frame2.Visible = False



Label1:
If Not FileExists(App.Path & "\enginekeys.txt") Then
Call MsgBox("The file 'EngineKeys.txt' is missing from the Route_Riter folder," _
            & vbCrLf & "without it this option will not run." _
            , vbExclamation, App.Title)

Exit Sub
End If
If Not FileExists(App.Path & "\engineignore.txt") Then
Call MsgBox("The file 'EngineIgnore.txt' is missing from the Route_Riter folder," _
            & vbCrLf & "without it this option will not run." _
            , vbExclamation, App.Title)

Exit Sub
End If
Me.Caption = "Engine/Wagon Editor     You are Editing - " & MSTSPath & "\Trains\Trainset"
Open App.Path & "\enginekeys.txt" For Input As #1
Do While Not EOF(1)
Input #1, strTemp

cbEng.AddItem strTemp
cbInsert.AddItem strTemp
Loop

Close #1
Open App.Path & "\engineignore.txt" For Input As #2
Input #2, intIgnore
ReDim EngIgnore(1 To intIgnore)
For i = 1 To intIgnore
Input #2, EngIgnore(i)
Next i
Close #2
If FileExists(App.Path & "\Formulae.txt") Then

i = 0: j = 0
Open App.Path & "\Formulae.txt" For Input As #3
Do While Not EOF(3)

Line Input #3, A$

If Left(A$, 7) = "Comment" Then GoTo CarryON
If A$ = "" Then GoTo CarryON
If Left(A$, 1) = "#" Then
i = i + 1
j = 0
strParam(i) = Mid(A$, 2)
Else
j = j + 1
strFormula(i, j) = A$

End If


CarryON:
Loop
numFormula = i
numItems = j

End If
Call CheckDefaultWag

End Sub





Private Sub Form_Resize()
On Error GoTo Errtrap

Frame2.Left = 200
Frame2.width = Me.width - 700
Label7(3).Left = Frame2.Left + (Frame2.width / 2)
Label7(4).Left = Frame2.Left + (Frame2.width / 2)
Label7(5).Left = Frame2.Left + (Frame2.width / 2)
Text1(3).Left = Label7(3).Left + Label7(3).width + 50
Text1(4).Left = Label7(4).Left + Label7(4).width + 50
Text1(5).Left = Label7(5).Left + Label7(5).width + 50
Text1(0).width = Label7(3).Left - (835 + Text1(0).Left)
Text1(1).width = Label7(3).Left - (835 + Text1(0).Left)
Text1(2).width = Label7(3).Left - (835 + Text1(0).Left)
Text1(3).width = Frame2.width - (Text1(3).Left + 100)
Text1(4).width = Frame2.width - (Text1(4).Left + 100)
Text1(5).width = Frame2.width - (Text1(5).Left + 100)
Exit Sub
Errtrap:
If Err = 380 Then
Exit Sub
End If
End Sub


Private Sub List1_Click()
lblCount(1).Caption = List1.SelCount
End Sub













Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
booEdited = True
End Sub


Private Sub Text2_Change(Index As Integer)
Dim strESD As String

For i = 0 To 5
strESD = strESD & Trim$(Text2(i)) & " "
Next i
strESD = Trim$(strESD)
strESD = "ESD_Bounding_Box ( " & strESD & " )"
Text1(5) = strESD
DoEvents

End Sub


Private Sub Text3_Change()
Dim strTemp As String, x As Integer

strTemp = Text1(0)
strNewSize = Trim$(Text3)
x = InStr(strTemp, "m")
x = InStr(x + 1, strTemp, "m")
strTemp = Left$(strTemp, x + 1) & Trim$(Text3) & "m )"
Text1(0) = strTemp
End Sub


