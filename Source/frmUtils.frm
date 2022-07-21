VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmUtils 
   Caption         =   "Route_Riter v7"
   ClientHeight    =   9420
   ClientLeft      =   1860
   ClientTop       =   -360
   ClientWidth     =   12075
   HelpContextID   =   1050
   Icon            =   "frmUtils.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9420
   ScaleWidth      =   12075
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   10680
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   12255
      Begin VB.CommandButton Command124 
         Caption         =   "E/W"
         Height          =   255
         Left            =   3840
         TabIndex        =   199
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command78 
         Caption         =   "sms"
         Height          =   255
         Left            =   3360
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command59 
         Caption         =   "Con"
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command43 
         Caption         =   "*"
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
         Left            =   5520
         TabIndex        =   22
         ToolTipText     =   "Filter to show all files"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command10 
         Caption         =   "MSTS"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   19
         ToolTipText     =   "Show main MSTS folder"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Caption         =   "MSTS"
         Height          =   255
         Index           =   1
         Left            =   9240
         TabIndex        =   18
         ToolTipText     =   "Show main MSTS folder"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Same View"
         Height          =   255
         Left            =   5760
         TabIndex        =   17
         ToolTipText     =   "Makes both File View windows the same."
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command42 
         Caption         =   "All"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         ToolTipText     =   "Selects All files in file window"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command43 
         Caption         =   "S"
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   15
         ToolTipText     =   "Filter to show only .S files"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command43 
         Caption         =   "T"
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   14
         ToolTipText     =   "Filter to show only .T files"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command43 
         Caption         =   "W"
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   13
         ToolTipText     =   "Filter to show only .W files"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command53 
         Caption         =   "Ace"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Source Directory"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Target Directory"
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Leave Program"
      Top             =   8760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   9
      Text            =   "*.*"
      ToolTipText     =   "File Type Filters"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   8
      Text            =   "*.*"
      ToolTipText     =   "File  Type Filters"
      Top             =   600
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Index           =   1
      Left            =   5520
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      ToolTipText     =   "Select File(s)"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Index           =   1
      Left            =   7440
      TabIndex        =   4
      ToolTipText     =   "Select Directory"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Index           =   1
      Left            =   7440
      TabIndex        =   3
      ToolTipText     =   "Select Disk Drive"
      Top             =   600
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Index           =   0
      Left            =   2160
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      ToolTipText     =   "Select File(s)"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Index           =   0
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Select Directory"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Select Disk Drive"
      Top             =   600
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   25
      ToolTipText     =   "Select TAB for operations you wish to use"
      Top             =   4440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "Route Utils"
      TabPicture(0)   =   "frmUtils.frx":406A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SB1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Activities/Stock"
      TabPicture(1)   =   "frmUtils.frx":4086
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8(5)"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label8(3)"
      Tab(1).Control(3)=   "Label8(2)"
      Tab(1).Control(4)=   "Label8(4)"
      Tab(1).Control(5)=   "Label7(4)"
      Tab(1).Control(6)=   "Label7(3)"
      Tab(1).Control(7)=   "Label7(2)"
      Tab(1).Control(8)=   "Label8(1)"
      Tab(1).Control(9)=   "Label8(0)"
      Tab(1).Control(10)=   "Label7(1)"
      Tab(1).Control(11)=   "Label7(0)"
      Tab(1).Control(12)=   "Label7(5)"
      Tab(1).Control(13)=   "Label8(6)"
      Tab(1).Control(14)=   "SB2"
      Tab(1).Control(15)=   "Frame4"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "MSTS File Utils"
      TabPicture(2)   =   "frmUtils.frx":40A2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "General Utils"
      TabPicture(3)   =   "frmUtils.frx":40BE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Mini-Routes"
      TabPicture(4)   =   "frmUtils.frx":40DA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame14"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Misc. Options"
      TabPicture(5)   =   "frmUtils.frx":40F6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame15"
      Tab(5).Control(1)=   "Frame6"
      Tab(5).Control(2)=   "Frame13"
      Tab(5).Control(3)=   "Frame11"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Graphics"
      TabPicture(6)   =   "frmUtils.frx":4112
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label11"
      Tab(6).Control(1)=   "Frame8"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "TsUtils"
      TabPicture(7)   =   "frmUtils.frx":412E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label12"
      Tab(7).Control(1)=   "Frame10"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Hardlink Files"
      TabPicture(8)   =   "frmUtils.frx":414A
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame16"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "RailDriver/Registry"
      TabPicture(9)   =   "frmUtils.frx":4166
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame17"
      Tab(9).Control(1)=   "Frame18"
      Tab(9).ControlCount=   2
      Begin VB.Frame Frame18 
         Caption         =   "Registry Options"
         Height          =   1095
         Left            =   -74520
         TabIndex        =   218
         Top             =   720
         Width           =   10575
         Begin VB.CommandButton Command111 
            Caption         =   "Change MSTS Registry Path"
            Height          =   495
            Left            =   2520
            TabIndex        =   223
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command110 
            Caption         =   "Revert to Original Registry Settings"
            Height          =   495
            Left            =   6360
            TabIndex        =   222
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command115 
            Caption         =   "Read Original MSTS Registry Path"
            Height          =   495
            Left            =   600
            TabIndex        =   221
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command117 
            Caption         =   "Read Current MSTS Registry Path"
            Height          =   495
            Left            =   4440
            TabIndex        =   220
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command114 
            Caption         =   "Show type of .eng"
            Height          =   495
            Left            =   8280
            TabIndex        =   219
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Raildriver Options"
         Height          =   1815
         Left            =   -74520
         TabIndex        =   185
         Top             =   1920
         Width           =   10575
         Begin VB.CommandButton Command113 
            Caption         =   "Modify ComboThrottle .eng for RailDriver"
            Height          =   495
            Left            =   8280
            TabIndex        =   191
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command112 
            Caption         =   "Modify Diesel .eng for Raildriver"
            Height          =   495
            Left            =   4440
            TabIndex        =   190
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command119 
            Caption         =   "Modify Electric.eng for RailDriver"
            Height          =   495
            Left            =   600
            TabIndex        =   189
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command118 
            Caption         =   "Modify Steam .eng for RailDriver"
            Height          =   495
            Left            =   2520
            TabIndex        =   188
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command116 
            Caption         =   "Modify Diesel with Gears (e.g. Kiha)"
            Height          =   495
            Left            =   6360
            TabIndex        =   187
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   840
            TabIndex        =   186
            Top             =   1080
            Width           =   9015
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "This option requires Windows XP with NTFS file system."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2295
         Left            =   -74280
         TabIndex        =   177
         Top             =   1380
         Width           =   9735
         Begin VB.CommandButton Command132 
            Caption         =   "Link Global\Shapes to Common Files"
            Height          =   495
            Left            =   7560
            TabIndex        =   208
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton Command131 
            Caption         =   "Copy Global\Shapes to Common Files"
            Height          =   495
            Left            =   5760
            TabIndex        =   207
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton Command65 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Confirm Route"
            Height          =   495
            Index           =   1
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   183
            ToolTipText     =   "Click to confirm Route"
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton Command107 
            Caption         =   "List Linked Files in Selected Route"
            Height          =   495
            Left            =   4920
            TabIndex        =   182
            ToolTipText     =   "Lists all linked files in Selected route."
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton Command106 
            Caption         =   "List Hard-LInks in Common Files"
            Height          =   495
            Left            =   3120
            TabIndex        =   181
            ToolTipText     =   "Lists all hard-linked files in the Common folder. Allows you to delete unlinked files."
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton Command104 
            Caption         =   "Copy Route Files to Common Folder"
            Height          =   495
            Left            =   2160
            TabIndex        =   179
            ToolTipText     =   "Copy shapes/textures/sounds from Selected Route to Common folder."
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton Command105 
            Caption         =   "Link All Route Files to Common Files"
            Height          =   495
            Left            =   3960
            TabIndex        =   178
            ToolTipText     =   "Link all possible files in Selected route to Common files."
            Top             =   480
            Width           =   1695
         End
         Begin MSComctlLib.StatusBar SB3 
            Height          =   375
            Left            =   120
            TabIndex        =   180
            Top             =   1800
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   2
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   1764
                  MinWidth        =   1764
                  Text            =   "Processing:"
                  TextSave        =   "Processing:"
               EndProperty
               BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   9596
                  MinWidth        =   9596
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Convert Tile Names"
         Height          =   3255
         Left            =   -67800
         TabIndex        =   161
         Top             =   600
         Width           =   4095
         Begin VB.CommandButton Command99 
            Caption         =   "Clear"
            Height          =   375
            Left            =   2160
            TabIndex        =   172
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton Command98 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   840
            TabIndex        =   167
            ToolTipText     =   "Will convert either Long/Lat or WorldTile coords to Tile File name."
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   4
            Left            =   2160
            TabIndex        =   166
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   165
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   2
            Left            =   1560
            TabIndex        =   164
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   1
            Left            =   2160
            TabIndex        =   163
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   162
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "World Tile Co-ordinates"
            Height          =   255
            Left            =   600
            TabIndex        =   171
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "Tile file name"
            Height          =   375
            Left            =   600
            TabIndex        =   170
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Latitude"
            Height          =   255
            Left            =   2040
            TabIndex        =   169
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Longitude"
            Height          =   255
            Left            =   360
            TabIndex        =   168
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame14 
         Height          =   3255
         Left            =   -74280
         TabIndex        =   156
         Top             =   720
         Width           =   9375
         Begin VB.CommandButton Command129 
            Caption         =   "Mini-Route Compact Sounds"
            Height          =   495
            Left            =   7440
            TabIndex        =   204
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command101 
            Caption         =   "Check Route"
            Height          =   495
            Left            =   4680
            TabIndex        =   174
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command100 
            Caption         =   "Confirm Route"
            Height          =   495
            Left            =   2880
            TabIndex        =   173
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command87 
            Caption         =   "Mini-Route Setup"
            Height          =   495
            Left            =   240
            TabIndex        =   160
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command88 
            Caption         =   "Mini-Route Compact Tracks"
            Height          =   495
            Left            =   5640
            TabIndex        =   159
            ToolTipText     =   "Removes track sections from Global\Shapes not used by the mini-route"
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command89 
            Caption         =   "Mini-Route get Stock"
            Height          =   495
            Left            =   3840
            TabIndex        =   158
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command90 
            Caption         =   "Mini-Route Copy Route"
            Height          =   495
            Left            =   2040
            TabIndex        =   157
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Common Files"
         Height          =   3255
         Left            =   -70080
         TabIndex        =   152
         Top             =   600
         Width           =   2175
         Begin VB.CommandButton Command33 
            Caption         =   "Copy Default Files to Common"
            Height          =   495
            Left            =   360
            TabIndex        =   155
            ToolTipText     =   "Copies default files to Common folder"
            Top             =   1020
            Width           =   1455
         End
         Begin VB.CommandButton Command32 
            Caption         =   "Convert Install File"
            Height          =   495
            Left            =   360
            TabIndex        =   154
            ToolTipText     =   "Converts an existing InstallMe.bat to point to Common files."
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Set Up Common Files"
            Height          =   495
            Left            =   360
            TabIndex        =   153
            ToolTipText     =   "Sets up a Common folder for default files"
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Environment"
         Height          =   3255
         Left            =   -72360
         TabIndex        =   147
         Top             =   600
         Width           =   2175
         Begin VB.CommandButton Command19 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Set Up New Env Files"
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
            Left            =   240
            MaskColor       =   &H00C0FFFF&
            Style           =   1  'Graphical
            TabIndex        =   151
            ToolTipText     =   "Writes a new set of .env files for the selected route"
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Correct Sun Rise/Set"
            Height          =   495
            Left            =   240
            TabIndex        =   150
            ToolTipText     =   "Corrects sunrise/set for route."
            Top             =   1005
            Width           =   1575
         End
         Begin VB.CommandButton Command25 
            BackColor       =   &H00FFFF00&
            Caption         =   "Improve Snow"
            Height          =   495
            Left            =   240
            TabIndex        =   149
            ToolTipText     =   "Changes Routes Snow textures to give deeper snow."
            Top             =   2280
            Width           =   1575
         End
         Begin VB.CommandButton Command28 
            Caption         =   "Replace BlendATexDiff"
            Height          =   495
            Left            =   240
            TabIndex        =   148
            ToolTipText     =   "Replace BlendATexDiff in .env to stop Flashing Water."
            Top             =   1635
            Width           =   1575
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Misc"
         Height          =   3255
         Left            =   -74640
         TabIndex        =   144
         Top             =   600
         Width           =   2175
         Begin VB.CommandButton Command95 
            Caption         =   "Fix Bad .S File Format"
            Height          =   495
            Left            =   240
            TabIndex        =   217
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CommandButton Command92 
            Caption         =   "Raise Bounding Box Minimum"
            Height          =   495
            Left            =   240
            TabIndex        =   146
            ToolTipText     =   "Raises the base of all rolling-stock Bounding Boxes to 0.9 to give greater clearance at crossings etc."
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton Command91 
            Caption         =   "Set up Editing Folder"
            Height          =   495
            Left            =   240
            TabIndex        =   145
            ToolTipText     =   "Backs up the Trainset folder .eng/.wag/.sd files only"
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "General Utilities"
         Height          =   2655
         Left            =   -74520
         TabIndex        =   111
         Top             =   720
         Width           =   9855
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dir to Txt"
            Height          =   495
            Index           =   17
            Left            =   3360
            TabIndex        =   128
            ToolTipText     =   "Print out selected Directory."
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Delete"
            Height          =   495
            Index           =   2
            Left            =   3360
            TabIndex        =   127
            ToolTipText     =   "Delete selected files"
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Rename"
            Height          =   495
            Index           =   7
            Left            =   1800
            TabIndex        =   126
            ToolTipText     =   "Rename selected file"
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&MakeDir"
            Height          =   495
            Index           =   8
            Left            =   3360
            TabIndex        =   125
            ToolTipText     =   "Make new directory"
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "R&emDir"
            Height          =   495
            Index           =   9
            Left            =   4920
            TabIndex        =   124
            ToolTipText     =   "Remove selected directory"
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&All"
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   123
            ToolTipText     =   "Select all files "
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&None"
            Height          =   495
            Index           =   1
            Left            =   1800
            TabIndex        =   122
            ToolTipText     =   "De-Select all selected files"
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Copy"
            Height          =   495
            Index           =   3
            Left            =   4920
            TabIndex        =   121
            ToolTipText     =   "Copy selected files to Target Dir."
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Parent"
            Height          =   495
            Index           =   11
            Left            =   6480
            TabIndex        =   120
            ToolTipText     =   "Go to parent directory"
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "R&oot"
            Height          =   495
            Index           =   12
            Left            =   8040
            TabIndex        =   119
            ToolTipText     =   "Go to root directory"
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Mo&ve"
            Height          =   495
            Index           =   4
            Left            =   6480
            TabIndex        =   118
            ToolTipText     =   "Move selected files to Target Dir."
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Edit"
            Height          =   495
            Index           =   5
            Left            =   8040
            TabIndex        =   117
            ToolTipText     =   "Edit File"
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "E&xecute"
            Height          =   495
            Index           =   13
            Left            =   240
            TabIndex        =   116
            ToolTipText     =   "Execute selected file"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Hide/Sys"
            Height          =   495
            Index           =   14
            Left            =   1800
            TabIndex        =   115
            ToolTipText     =   "Toggles whether Hidden or System files appear"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Pr&operties"
            Height          =   495
            Index           =   6
            Left            =   240
            TabIndex        =   114
            ToolTipText     =   "Show/adjust properties of selected file"
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton Command34 
            Caption         =   "Move Folder"
            Height          =   495
            Left            =   4920
            TabIndex        =   113
            ToolTipText     =   "Move selected folder"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton Command37 
            Caption         =   "My&Zipp"
            Height          =   495
            Left            =   6480
            TabIndex        =   112
            ToolTipText     =   "Zip tool similar to WinZip"
            Top             =   1560
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Activity Checking"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   98
         Top             =   600
         Width           =   9375
         Begin VB.CommandButton Command144 
            Caption         =   "Consist Editor"
            Height          =   495
            Left            =   7680
            TabIndex        =   231
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton Command70 
            Caption         =   "Check Acts for Selected Route"
            Height          =   495
            Left            =   1680
            TabIndex        =   216
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command55 
            Caption         =   "Make unpowered Loco"
            Height          =   495
            Left            =   7680
            TabIndex        =   214
            ToolTipText     =   "Turns Loco into a Wagon which may be pulled by another train etc."
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Make AI or MU .eng"
            Height          =   495
            Left            =   6120
            TabIndex        =   213
            ToolTipText     =   "Makes an AI or MU version of a Locomotive"
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command123 
            Caption         =   "List Rolling Stock for Selected Route"
            Height          =   495
            Left            =   4560
            TabIndex        =   198
            ToolTipText     =   "Lists the rolling stock for the Route in the Left hand window, Click 'Confirm Route' before using this option."
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command83 
            Caption         =   ".ENG/.WAG file Editor"
            Height          =   495
            Left            =   6120
            TabIndex        =   195
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton Command108 
            Caption         =   "List Stock for Selected Consists"
            Height          =   495
            Left            =   4560
            TabIndex        =   184
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton Command103 
            Caption         =   "Fix .CVF Files"
            Height          =   495
            Left            =   6840
            TabIndex        =   176
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Check All Act"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   110
            ToolTipText     =   "Checks all the Activities"
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Check Consists"
            Height          =   495
            Left            =   1680
            TabIndex        =   109
            ToolTipText     =   "Checks the Consists"
            Top             =   960
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Check  Rolling Stock"
            Height          =   495
            Left            =   3000
            TabIndex        =   108
            ToolTipText     =   "Checks rolling stock and fixes some errors."
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton Command35 
            Caption         =   "Check Selected Activity"
            Height          =   495
            Left            =   120
            TabIndex        =   107
            ToolTipText     =   "Checks the Selected Activity only"
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command75 
            Caption         =   "Quick consist check"
            Height          =   495
            Left            =   3000
            TabIndex        =   106
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command85 
            Caption         =   "Fix .SMS Files"
            Height          =   495
            Left            =   6840
            TabIndex        =   105
            ToolTipText     =   "Fixes some alias and filename errors in sms files"
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Frame Frame12 
            Caption         =   "The Following Six Options MUST be used sequentially 1 thru 6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1695
            Left            =   240
            TabIndex        =   99
            Top             =   1560
            Width           =   6375
            Begin VB.CommandButton Command142 
               Caption         =   "[2] Fix Wag Names"
               Height          =   495
               Left            =   2160
               TabIndex        =   230
               Top             =   360
               Width           =   1815
            End
            Begin VB.CommandButton Command84 
               Caption         =   "[ 3 ]   Fix .SD Files"
               Height          =   495
               Left            =   4080
               TabIndex        =   104
               ToolTipText     =   "Corrects some Case errors in .SD files"
               Top             =   360
               Width           =   1815
            End
            Begin VB.CommandButton Command82 
               Caption         =   "[ 1 ]   Fix .ENG names"
               Height          =   495
               Left            =   240
               TabIndex        =   103
               ToolTipText     =   "Corrects case errors in .eng files"
               Top             =   360
               Width           =   1815
            End
            Begin VB.CommandButton Command81 
               Caption         =   "[6 ]   Fix .SRV Names"
               Height          =   495
               Left            =   4080
               TabIndex        =   102
               ToolTipText     =   "Fixes case errors in .srv files"
               Top             =   960
               Width           =   1815
            End
            Begin VB.CommandButton Command76 
               Caption         =   "[ 4 ]   Fix .CON names"
               Height          =   495
               Left            =   240
               TabIndex        =   101
               ToolTipText     =   "Fixes case errors in .con files"
               Top             =   960
               Width           =   1815
            End
            Begin VB.CommandButton Command79 
               Caption         =   "[5 ]   Fix .ACT Names"
               Height          =   495
               Left            =   2160
               TabIndex        =   100
               ToolTipText     =   "Fixes case errors in .act files"
               Top             =   960
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "The following options are for selected files only"
         ForeColor       =   &H000000FF&
         Height          =   3255
         Left            =   -74880
         TabIndex        =   78
         Top             =   480
         Width           =   11055
         Begin VB.CommandButton Command140 
            Caption         =   "Remove ViewDbSphere from Selected .W files"
            Height          =   495
            Left            =   8760
            TabIndex        =   228
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CommandButton Command125 
            Caption         =   "Correct Stuck Points"
            Height          =   495
            Left            =   6600
            TabIndex        =   225
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CommandButton Command109 
            Caption         =   "Find Stuck Points in Route"
            Height          =   495
            Left            =   4440
            TabIndex        =   224
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CommandButton Command130 
            Caption         =   "Del Shape or Transfer from Selected .W files"
            Height          =   495
            Left            =   8760
            TabIndex        =   206
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CommandButton Command126 
            Caption         =   "Change .T ErrorBias on tiles with track/roads"
            Height          =   495
            Left            =   120
            TabIndex        =   200
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H0080FF80&
            Caption         =   "Compress all selected     S,T && W Files"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   97
            ToolTipText     =   "Compress the selected  S,T or W files"
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Unicode to ASCII"
            Height          =   495
            Index           =   0
            Left            =   6600
            TabIndex        =   96
            ToolTipText     =   "Convert Unicode to Ascii"
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton Command6 
            Caption         =   "ASCII to Unicode"
            Height          =   495
            Index           =   1
            Left            =   8760
            TabIndex        =   95
            ToolTipText     =   "Convert ASCII to Unicode"
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Count Objects per tile in Selected Route"
            Height          =   495
            Left            =   120
            TabIndex        =   94
            ToolTipText     =   "Count all route objects"
            Top             =   1440
            Width           =   2055
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Add LoadAllWaves"
            Height          =   495
            Left            =   2280
            TabIndex        =   93
            ToolTipText     =   "Adds LoadAllWaves to selected Sound folder"
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Edit Unicode"
            Height          =   495
            Left            =   4440
            TabIndex        =   92
            ToolTipText     =   "Edit Unicode files"
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Delete LoadAllWaves"
            Height          =   495
            Left            =   4440
            TabIndex        =   91
            ToolTipText     =   "Removes LoadAllWaves from folder"
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Uncompress all selected  S,T or W Files"
            Height          =   495
            Left            =   2280
            TabIndex        =   90
            ToolTipText     =   "Uncompresses selected .T files"
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Del Shape or Transfer from all .W files"
            Height          =   495
            Left            =   6600
            TabIndex        =   89
            ToolTipText     =   "Deletes a named shape or transfer from selected route's .W files"
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CommandButton Command30 
            Caption         =   "Copy Shape to another route"
            Height          =   495
            Left            =   6600
            TabIndex        =   88
            ToolTipText     =   "Copy selected shape and all associated files."
            Top             =   1440
            Width           =   2055
         End
         Begin VB.CommandButton Command38 
            Caption         =   "List all instances of a shape in .W"
            Height          =   495
            Left            =   120
            TabIndex        =   87
            ToolTipText     =   "Lists all instances of a shape in a route."
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CommandButton Command39 
            Caption         =   "Make tsection.dat"
            Height          =   495
            Left            =   6600
            TabIndex        =   86
            ToolTipText     =   "Makes a tsection.dat which only includes shapes you have"
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton Command40 
            Caption         =   "Make all files in folder  read/write"
            Height          =   495
            Left            =   120
            TabIndex        =   85
            ToolTipText     =   "Makes all  files in selected folder and sub-folders  read/write"
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton Command44 
            Caption         =   "Replace Shapes in .W files"
            Height          =   495
            Left            =   4440
            TabIndex        =   84
            ToolTipText     =   "Replace Shape A with Shape B throughout a Route"
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CommandButton Command45 
            Caption         =   "Set Up a New Route"
            Height          =   495
            Left            =   4440
            TabIndex        =   83
            ToolTipText     =   "Copies all the files from the Default routes to your new routes and sets up a .ref file"
            Top             =   1440
            Width           =   2055
         End
         Begin VB.CommandButton Command46 
            Caption         =   "Duplicate route with new name"
            Height          =   495
            Left            =   2280
            TabIndex        =   82
            Top             =   1440
            Width           =   2055
         End
         Begin VB.CommandButton Command60 
            Caption         =   "List .ACE files for Selected .S file"
            Height          =   495
            Left            =   8760
            TabIndex        =   81
            Top             =   1440
            Width           =   2055
         End
         Begin VB.CommandButton Command74 
            Caption         =   "Replace Forest Textures in .W Files"
            Height          =   495
            Left            =   2280
            TabIndex        =   80
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Change .T ErrorBias"
            Height          =   495
            Left            =   2280
            TabIndex        =   79
            ToolTipText     =   "Make all Error Bias entries the same."
            Top             =   2640
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "MSTS Utilities"
         Height          =   3735
         Left            =   120
         TabIndex        =   61
         ToolTipText     =   "Makes a .ref file for a route where no .ref exists"
         Top             =   480
         Width           =   11175
         Begin VB.CommandButton Command141 
            Caption         =   "Change altitude of objects in .W files"
            Height          =   495
            Left            =   9240
            TabIndex        =   229
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CommandButton Command136 
            Caption         =   "Remove Auto-Gantry items from .w files"
            Height          =   495
            Left            =   9240
            TabIndex        =   212
            ToolTipText     =   "Remove auto-gantry items from .w files - for Route Authors only."
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command134 
            Caption         =   "Convert Google .kml or .gpx file to  .mkr"
            Height          =   495
            Left            =   7440
            TabIndex        =   210
            ToolTipText     =   "Converts Google .kml or .gpx files into MSTS Marker files."
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command127 
            Caption         =   "Run Conbuilder"
            Height          =   495
            Left            =   5640
            TabIndex        =   205
            ToolTipText     =   "Runs Conbuilder on the Instance of MSTS currently running in Route_Riter"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command128 
            Caption         =   "Compact Global\Shapes"
            Height          =   495
            Left            =   3840
            TabIndex        =   203
            ToolTipText     =   "Checks the Track items used in all routes and stores those not in use."
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command86 
            Caption         =   "Count objects in World tiles"
            Height          =   495
            Left            =   2040
            TabIndex        =   202
            ToolTipText     =   "Lists and Counts all object in the selected World tile and the 8 adjacent tiles"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command77 
            Caption         =   "List Track Sections in Route"
            Height          =   495
            Left            =   240
            TabIndex        =   201
            ToolTipText     =   "Displays a list of all Track/Road sections used in selected route."
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command121 
            Caption         =   "List Filtered Files"
            Height          =   495
            Left            =   9240
            TabIndex        =   196
            ToolTipText     =   "Lists all files in accordance with the pattern in the filter box"
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton Command102 
            Caption         =   "Activate/Deactivate Routes"
            Height          =   495
            Left            =   9240
            TabIndex        =   175
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H00FFFF00&
            Caption         =   "Compress .ACE as DXT1"
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Compress any uncompressed .ace files in route."
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00FFFF00&
            Caption         =   "Compress .S"
            Height          =   495
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Compresses any uncompressed .S files in selected route"
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFF00&
            Caption         =   "Compress .W"
            Height          =   495
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Compresses any uncompressed .W files in selected route."
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Compact Route"
            Height          =   495
            Left            =   5640
            TabIndex        =   74
            ToolTipText     =   "Compact the Route by removing Unused files."
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Confirm Route"
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
            Left            =   240
            MaskColor       =   &H00C0FFFF&
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Click to confirm Route"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Frame Frame5 
            Caption         =   "For Route Developers Only:-  "
            Height          =   975
            Left            =   120
            TabIndex        =   67
            Top             =   2640
            Width           =   10935
            Begin VB.CommandButton Command122 
               BackColor       =   &H0080FFFF&
               Caption         =   "Compress with UHARC"
               Height          =   495
               Left            =   7320
               Style           =   1  'Graphical
               TabIndex        =   197
               ToolTipText     =   "Packs the route as an .exe file using UHARC"
               Top             =   360
               Width           =   1695
            End
            Begin VB.CommandButton Command12 
               BackColor       =   &H0080FFFF&
               Caption         =   "Write .BAT"
               Height          =   495
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   72
               ToolTipText     =   "Write an installation batch file for your route."
               Top             =   360
               Width           =   1695
            End
            Begin VB.CommandButton Command13 
               BackColor       =   &H0080FFFF&
               Caption         =   "Delete Raw"
               Height          =   495
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   71
               ToolTipText     =   "Deletes e.raw/n.raw/.bk files"
               Top             =   360
               Width           =   1695
            End
            Begin VB.CommandButton Command2 
               BackColor       =   &H0080FFFF&
               Caption         =   "ZIP Route"
               Height          =   495
               Left            =   5520
               Style           =   1  'Graphical
               TabIndex        =   70
               ToolTipText     =   "Make ZIP file(s) of the route."
               Top             =   360
               Width           =   1695
            End
            Begin VB.CommandButton Command36 
               BackColor       =   &H0080FFFF&
               Caption         =   "Make Update"
               Height          =   495
               Left            =   3720
               Style           =   1  'Graphical
               TabIndex        =   69
               ToolTipText     =   "Make a Route Updater to next version."
               Top             =   360
               Width           =   1695
            End
            Begin VB.CommandButton Command62 
               BackColor       =   &H0080FFFF&
               Caption         =   "Backup Route"
               Height          =   495
               Left            =   9120
               Style           =   1  'Graphical
               TabIndex        =   68
               ToolTipText     =   "Backup World\Tiles\TD\Root folders of Route"
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Make Read/Write"
            Height          =   495
            Left            =   2040
            TabIndex        =   66
            ToolTipText     =   "Make all files on selected route Read/Write"
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command58 
            Caption         =   ".Ref File Editor"
            Height          =   495
            Left            =   5640
            TabIndex        =   65
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton Command63 
            Caption         =   "Make .REF File"
            Height          =   495
            Left            =   7440
            TabIndex        =   64
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton Command80 
            Caption         =   "Check Route"
            Height          =   495
            Left            =   3840
            TabIndex        =   63
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command97 
            Caption         =   "Make tsection.dat for use with TrainStore"
            Height          =   495
            Left            =   7440
            TabIndex        =   62
            ToolTipText     =   "Builds a route-specific tsection.dat file for selected route to use with Train-Store."
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame8 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   40
         ToolTipText     =   "View the selected .s files in MSTSView"
         Top             =   480
         Width           =   11055
         Begin VB.CommandButton Command133 
            Caption         =   "List Shapes which use this .ACE"
            Height          =   495
            Left            =   9240
            TabIndex        =   209
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton Command48 
            Caption         =   "Show Selected Pictures"
            Height          =   495
            Left            =   4920
            TabIndex        =   60
            ToolTipText     =   "Display selected graphics files."
            Top             =   2160
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CommandButton Command49 
            Caption         =   "ACE to TGA"
            Height          =   495
            Index           =   0
            Left            =   480
            TabIndex        =   59
            Top             =   960
            Width           =   1935
         End
         Begin VB.CommandButton Command49 
            Caption         =   "ACE to BMP"
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   58
            Top             =   1560
            Width           =   1935
         End
         Begin VB.CommandButton Command47 
            Caption         =   "List ACE types"
            Height          =   495
            Left            =   4920
            TabIndex        =   57
            ToolTipText     =   "Lists all selected .ace files and shows whether or not compressed and file type."
            Top             =   960
            Width           =   1935
         End
         Begin VB.CommandButton Command24 
            Caption         =   "Compress all .ACE files in selected folder."
            Height          =   495
            Left            =   2640
            TabIndex        =   56
            ToolTipText     =   "Compress the selected .ACE files"
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H0080FF80&
            Caption         =   "Make uncompressed ACE"
            Height          =   495
            Index           =   3
            Left            =   4920
            TabIndex        =   55
            ToolTipText     =   "Make a standard .ACE file from .bmp or .tga file."
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H0080FF80&
            Caption         =   "Make ACE Zlib"
            Height          =   495
            Index           =   4
            Left            =   2640
            TabIndex        =   54
            ToolTipText     =   "Make a compressed .ACE file from .bmp or .tga file."
            Top             =   960
            Width           =   1935
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H0080FF80&
            Caption         =   "Convert Selected .ace to DXT1 (ignore if DXT)"
            Height          =   495
            Index           =   5
            Left            =   2640
            TabIndex        =   53
            ToolTipText     =   "Make a compressed .ACE file from .bmp or .tga file."
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Frame Frame9 
            Caption         =   "Slide-Show Settings"
            Height          =   975
            Left            =   480
            TabIndex        =   49
            Top             =   2160
            Visible         =   0   'False
            Width           =   4095
            Begin VB.CommandButton Command50 
               Caption         =   "Slide-Show of Selected Folder"
               Height          =   615
               Left            =   1800
               TabIndex        =   51
               ToolTipText     =   "Show all graphics in selected folder in slide-show."
               Top             =   240
               Width           =   2055
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   120
               TabIndex        =   50
               Text            =   "4"
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label10 
               Caption         =   "Delay"
               Height          =   495
               Left            =   600
               TabIndex        =   52
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.CommandButton Command51 
            Caption         =   "View .S File"
            Height          =   495
            Left            =   4920
            TabIndex        =   48
            ToolTipText     =   "Display .s file in 3D mode including its texture"
            Top             =   1560
            Width           =   1935
         End
         Begin VB.CommandButton Command52 
            Caption         =   "Convert Selected files to DXT1"
            Height          =   495
            Left            =   480
            TabIndex        =   47
            ToolTipText     =   "Convert selected .ace files to DXT1 compression"
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton Command56 
            Caption         =   "Run T-View"
            Height          =   495
            Index           =   0
            Left            =   7080
            TabIndex        =   46
            ToolTipText     =   "Display all graphics in folder using T-View"
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CommandButton Command56 
            Caption         =   "Run TGATool2"
            Height          =   495
            Index           =   1
            Left            =   7080
            TabIndex        =   45
            ToolTipText     =   "Display all graphics in folder using T-View"
            Top             =   1560
            Width           =   1935
         End
         Begin VB.CommandButton Command61 
            Caption         =   "Retrieve Saved Pictures"
            Height          =   495
            Left            =   7080
            TabIndex        =   44
            ToolTipText     =   "Display pictures of shapes which have been saved."
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton Command54 
            Caption         =   "List all Filtered Files"
            Height          =   495
            Left            =   7080
            TabIndex        =   43
            ToolTipText     =   "Lists all files according to the filter box, includes sub-folders."
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton comAbortFil 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Abort"
            Height          =   495
            Left            =   7080
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   360
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Include CabView Folder"
            Height          =   375
            Left            =   7080
            TabIndex        =   41
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.Frame Frame10 
         Height          =   2895
         Left            =   -74400
         TabIndex        =   28
         Top             =   1200
         Width           =   9375
         Begin VB.CommandButton Command57 
            Caption         =   "Raise or Lower Track (mveobj)"
            Height          =   495
            Left            =   3840
            TabIndex        =   233
            Top             =   2040
            Width           =   1695
         End
         Begin VB.CommandButton Command145 
            BackColor       =   &H0080FFFF&
            Caption         =   "Open TsUtil Stand Alone Version"
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
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   232
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CommandButton Command138 
            Caption         =   "Log Signals in Route (srchsig)"
            Height          =   495
            Left            =   7440
            TabIndex        =   227
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton Command137 
            Caption         =   "Remove .t where no .w exists"
            Height          =   495
            Left            =   7440
            TabIndex        =   226
            ToolTipText     =   "If no .w file exists, the corresponding .t file is moved to RRBackups folder"
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command69 
            BackColor       =   &H0080FFFF&
            Caption         =   "Manually enter TsUtil commands"
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
            Left            =   7080
            Style           =   1  'Graphical
            TabIndex        =   215
            ToolTipText     =   "Allows you to use all of the TsUtils commands manually"
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CommandButton Command135 
            Caption         =   "Use CVRT to fix Soundsources"
            Height          =   495
            Left            =   7440
            TabIndex        =   211
            ToolTipText     =   "Where possible correctly places SoundSources in a route"
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton Command120 
            Caption         =   "Show TsUtil Version"
            Height          =   495
            Left            =   5640
            TabIndex        =   194
            ToolTipText     =   "Displays the version numbers of all TsUtil files"
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton Command64 
            Caption         =   "Check Integrity of Route (ICHK)"
            Height          =   495
            Left            =   2040
            TabIndex        =   39
            ToolTipText     =   "Checks .tdb .rdb .tit .rit etc files"
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton Command65 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Confirm Route"
            Height          =   495
            Index           =   0
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Click to confirm Route"
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton Command66 
            Caption         =   "Move a Route to new Lat/Long (MOVE)"
            Height          =   495
            Left            =   240
            TabIndex        =   37
            ToolTipText     =   "Moves a route to a new position"
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton Command67 
            Caption         =   "Modify RDB/TDB (RENDB)"
            Height          =   495
            Left            =   3840
            TabIndex        =   36
            ToolTipText     =   "Repair tdb/rdb which has had  faulty Track Nodes manually deleted."
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command68 
            Caption         =   "Change refs in TDB (CHGDB)"
            Height          =   495
            Left            =   240
            TabIndex        =   35
            ToolTipText     =   "Attempts to repair a faulty .tdb file"
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command71 
            Caption         =   "Change Route Altitude (ADJH)"
            Height          =   495
            Left            =   2040
            TabIndex        =   34
            ToolTipText     =   "Changes the altitude of a route"
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton Command72 
            Caption         =   "Check new tsection.dat (CHKUP)"
            Height          =   495
            Left            =   2040
            TabIndex        =   33
            ToolTipText     =   "Checks that a new Global tsection.dat is compatible with the original one used by a route."
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command73 
            Caption         =   "Modify Route for new tsection.dat (cvrt)"
            Height          =   495
            Left            =   5640
            TabIndex        =   32
            ToolTipText     =   "Similar to the Horace.exe program"
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command93 
            Caption         =   "Check Tile Definitions (Filter)"
            Height          =   495
            Left            =   3840
            TabIndex        =   31
            ToolTipText     =   "Removes unused tiles from the TD file."
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton Command94 
            Caption         =   "Merge Routes (Merge)"
            Height          =   495
            Left            =   5640
            TabIndex        =   30
            ToolTipText     =   "Merges two routes."
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton Command96 
            Caption         =   "Reorg .tdb/.rdb (CLRDB)"
            Height          =   495
            Left            =   3840
            TabIndex        =   29
            ToolTipText     =   "Reorganizes .tdb/.rdb files"
            Top             =   240
            Width           =   1695
         End
      End
      Begin MSComctlLib.StatusBar SB2 
         Height          =   375
         Left            =   -74640
         TabIndex        =   26
         Top             =   4080
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1764
               MinWidth        =   1764
               Text            =   "Processing:"
               TextSave        =   "Processing:"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   9596
               MinWidth        =   9596
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.StatusBar SB1 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   4320
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   7
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1588
               MinWidth        =   1588
               Text            =   "Processing:"
               TextSave        =   "Processing:"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
               MinWidth        =   5292
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1235
               MinWidth        =   1235
               Text            =   "Start:"
               TextSave        =   "Start:"
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1235
               MinWidth        =   1235
               Text            =   "End:"
               TextSave        =   "End:"
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "Traffic"
         Height          =   255
         Index           =   6
         Left            =   -64440
         TabIndex        =   193
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   -65400
         TabIndex        =   192
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   -65400
         TabIndex        =   143
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   -65400
         TabIndex        =   142
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Locos"
         Height          =   255
         Index           =   0
         Left            =   -64440
         TabIndex        =   141
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Wagons"
         Height          =   255
         Index           =   1
         Left            =   -64440
         TabIndex        =   140
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   -65400
         TabIndex        =   139
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   -65400
         TabIndex        =   138
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   -65400
         TabIndex        =   137
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Consists"
         Height          =   255
         Index           =   4
         Left            =   -64440
         TabIndex        =   136
         Top             =   2460
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Activities"
         Height          =   255
         Index           =   2
         Left            =   -64440
         TabIndex        =   135
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Services"
         Height          =   255
         Index           =   3
         Left            =   -64440
         TabIndex        =   134
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -65400
         TabIndex        =   133
         Top             =   2820
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Paths"
         Height          =   255
         Index           =   5
         Left            =   -64440
         TabIndex        =   132
         Top             =   2820
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -74160
         TabIndex        =   131
         Top             =   3960
         Visible         =   0   'False
         Width           =   8775
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -72600
         TabIndex        =   130
         ToolTipText     =   "File being processed"
         Top             =   5280
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -72720
         TabIndex        =   129
         Top             =   4560
         Width           =   5895
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   8400
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   7
      ToolTipText     =   "Path of Selected File"
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Path of Selected File"
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      X1              =   4920
      X2              =   4920
      Y1              =   120
      Y2              =   4320
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "Files"
      Begin VB.Menu mnuPath 
         Caption         =   "MSTS Path"
      End
      Begin VB.Menu mnuCommon 
         Caption         =   "Common File Path"
      End
      Begin VB.Menu mnuTrainstore 
         Caption         =   "Run Trainstore"
      End
      Begin VB.Menu mnuTER 
         Caption         =   "Run Text Editor"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuLan 
      Caption         =   "Languages"
      Begin VB.Menu mnu1 
         Caption         =   "Lang1"
         Index           =   0
      End
      Begin VB.Menu mnu1 
         Caption         =   "Lang2"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnu1 
         Caption         =   "Lang3"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnu1 
         Caption         =   "Lang4"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnu1 
         Caption         =   "Lang5"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnu1 
         Caption         =   "Lang6"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnu1 
         Caption         =   "Lang7"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnu1 
         Caption         =   "Lang8"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnu1 
         Caption         =   "Lang9"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnu1 
         Caption         =   "Lang10"
         Index           =   9
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Options"
      Begin VB.Menu mnuConPath 
         Caption         =   "Conbuilder Path"
      End
      Begin VB.Menu mnuTE 
         Caption         =   "Select Text Editor"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuCont 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuHome 
         Caption         =   "Home Page"
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFAQ 
         Caption         =   "FAQ"
      End
      Begin VB.Menu mnuUpdates 
         Caption         =   "Check for Updates"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

    
Const CHUNK = 500
Const REF_CHUNK = 100
Const Car_CHUNK = 50
Dim Rname() As String
Dim wTile(1 To 9) As String
Dim booEndFile As Boolean, booFixSound As Boolean
Dim gtotShp As Integer, gtotAce As Integer, gtotGlobShp As Integer
Dim gtotTerr As Integer, gtotTrans As Integer, gtotHaz As Integer, gtotFor As Integer
Dim HazShp() As Variant, strGetAce() As String
Dim HazShp2() As Variant
Dim strShp() As Variant
Dim strShapes() As Variant, GlobalSparePath As String
Dim strGlobShp() As Variant
Dim strGlobShp2() As Variant
Dim numGlobShp As Long
Dim numShp As Long
Dim strWave(0 To 500) As String
Dim numWave As Integer
Dim ForTex() As Variant, strPoints() As String
Dim ForTex2() As Variant
Dim Ace1() As Variant
Dim Ace2() As Variant
Dim TerrTex() As Variant
Dim TerrTex2() As Variant
Dim Transfer() As Variant
Dim Transfer2() As Variant
Dim MiscAce(0 To 500) As String
Dim numMisc As Integer, strLogPath As String
Dim numAce As Long
Dim numFor As Long
Dim numHaz As Long
Dim numTerr As Long
Dim numTrans As Long
Dim strBatFile As String
Dim totAce As Integer
Dim totFor As Integer
Dim totHaz As Integer
Dim totTerr As Integer
Dim totTrans As Integer, strConPath As String, strTrainset As String
Dim totShp As Integer
Dim totGlobshp As Integer
Const Shp_Chunk = 5000
Const For_Chunk = 500
Rem *********
Dim Cars() As Variant, intCars As Long, Cars2() As Variant
Dim KillEm() As String, KE As Integer, KillEnvBat(1 To 100) As String
Dim EnvAceFile(1 To 100), envAceNumber As Integer, strLanguage As String
Dim booElectric As Boolean, strEnv(1 To 13) As String, conShort As String
Dim bdata() As Byte, MainRoute As Integer
Dim ConEng() As Integer, actOnly As Boolean
Dim conWag() As Integer, ConIntName() As String, strGraphic As String
Dim ActEng() As Integer, ActWag() As Integer
Dim ConEngNumber As Integer, ConWagNumber As Integer
Dim hazNumber As Integer, Hazard(0 To 10) As String, HazName(0 To 10) As String
Dim hazAceNumber As Integer, HazAce(0 To 10) As String
Dim flagNoRef As Boolean, booHaz As Boolean
Dim booCross As Boolean, booSig As Boolean, booWat As Boolean
Dim booCoal As Boolean, booDies As Boolean
Dim DefCross As String, DefSig As String, DefWat As String
Dim DefCoal As String, DefDies As String
Dim flagTerr As Boolean
Dim SearchFlag As Boolean, flagFull As Boolean
Dim AceNumber As Long
Dim AceFile(0 To 5000) As String, OldRouteName As String, ESD(0 To 5000) As String
Dim MasterFile() As String
Dim MasterIndex() As String, NumRoutes As Integer
Dim ActChecked As Boolean
Dim strSeason As String, TerrPath As String
Dim intTempFile As Integer
Dim RouteListed As Boolean, ShapePath As String
Dim TexturePath As String, TexSnowPath As String, TexNightPath As String
Dim TexAutPath As String, TexAutSnowPath As String, TexSprPath As String
Dim TexSprSnowPath As String, TexWinPath As String, TexWinSnowPath As String
Dim WorldPath As String
Dim SoundPath As String
Dim TilePath As String
Dim OriginalRef As String
Dim Soundfile(0 To 1000) As String, SoundNumber As Long
Dim GlobalSoundPath As String
Dim WavFile(0 To 1000) As String, WavNumber As Long
Dim MainRoutePath As String, Eur1Path As String
Dim Eur2Path As String, Jap1Path As String, Jap2Path As String
Dim USA1Path As String, USA2Path As String
Dim TransferFile(0 To 100) As String, TransferNumber As Integer
Dim booPlayer As Boolean, EnvPath As String
Dim OldEnv(1 To 12) As String

Dim strAnimUsed() As String
Dim AllRoutes() As String

Dim strWorldTiles() As String, numWorldTiles As Integer









Private Sub ChangeConPath(strConName As String, strEngname As String, strEngpath As String, strFoundPath As String, MyString As String)
Dim ConsistPath As String, i As Integer

ConsistPath = MSTSPath & "\Trains\Consists\"


i = InStr(strFoundPath, "\")
strConPath = Left(strFoundPath, i - 1)


strEngname = Left(strEngname, Len(strEngname) - 4)

strConPath = ChrW$(34) & strEngname & ChrW$(34) & " " & ChrW$(34) & strConPath & ChrW$(34)
MyString = ReadUniFile(ConsistPath & strConName)
DoEvents
MyString = Replace(MyString, ChrW$(34) & strEngname & ChrW$(34), strEngname)
DoEvents
MyString = Replace(MyString, ChrW$(34) & strEngpath & ChrW$(34), strEngname)
DoEvents
strEngpath = strEngname & " " & strEngpath

MyString = Replace(MyString, strEngpath, strConPath)
DoEvents
Call WriteUniFile(ConsistPath & strConName, MyString)
DoEvents



End Sub


Private Sub CheckMissFolders(strRoutePath As String)
If Not DirExists(strRoutePath & "\Services") Then
MkDir strRoutePath & "\Services"
End If
If Not DirExists(strRoutePath & "\Activities") Then
MkDir strRoutePath & "\Activities"
End If
If Not DirExists(strRoutePath & "\Traffic") Then
MkDir strRoutePath & "\Traffic"
End If
If Not DirExists(strRoutePath & "\Paths") Then
MkDir strRoutePath & "\Paths"
End If
If Not DirExists(strRoutePath & "\LO_TILES") Then
MkDir strRoutePath & "\LO_TILES"
End If
End Sub

Private Sub CheckTrkForEnv(strRoute As String)
Dim NewFile As Integer, strNew As String, x As Integer, Y As Integer, Z As Integer
Dim i As Integer, EnvTag(1 To 12), xx As Integer

On Error GoTo Errtrap

EnvTag(1) = "springclear "
EnvTag(2) = "springrain "
EnvTag(3) = "springsnow "
EnvTag(4) = "summerclear "
EnvTag(5) = "summerrain "
EnvTag(6) = "summersnow "
EnvTag(7) = "autumnclear "
EnvTag(8) = "autumnrain "
EnvTag(9) = "autumnsnow "
EnvTag(10) = "winterclear "
EnvTag(11) = "winterrain "
EnvTag(12) = "wintersnow "

NewFile = FreeFile
Open strRoute For Input As #NewFile
Do While Not EOF(NewFile)
   
   Line Input #NewFile, strNew
  
   strNew = Trim$(strNew)
   For i = 1 To 12
  
   xx = InStr(strNew, EnvTag(i))
  
   If xx = 0 Then
   x = 0
   End If
   If Mid$(strNew, xx + Len(EnvTag(i)), 4) <> ".env" Then
   x = xx
   Else
   x = 0
   End If
   If x > 0 Then
    Y = InStr(x, strNew, "(")
    Z = InStr(Y, strNew, ")")
    OldEnv(i) = Trim$(Mid$(strNew, Y + 1, Z - Y - 1))
    If Left$(OldEnv(i), 1) = ChrW$(34) Then
    OldEnv(i) = Mid$(OldEnv(i), 2)
    End If
    If Right$(OldEnv(i), 1) = ChrW$(34) Then
    OldEnv(i) = Left$(OldEnv(i), Len(OldEnv(i)) - 1)
    End If
    
    If Not FileExists(RoutePath & "\Envfiles\" & OldEnv(i)) Then
    
     strReport = strReport & "Envfile - " & OldEnv(i) & " is missing - please check" & vbCrLf
    End If
    Exit For
    End If
 Next i
 Loop
 Close #NewFile
 
Exit Sub
Errtrap:
Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'CheckTrkForEnv' in " & strRoute & " please advise" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
                             Case vbRetry
    Resume Next
        Case vbCancel
       'Resume Next
    Exit Sub
    End Select
        
End Sub

Private Sub ConvertTrk3(strTrkFile As String, strOld As String, strNew As String, strNewIntName As String)
Dim MyString As String, i As Integer, ii As Integer, strStart As String, strEnd As String

MyString = ReadUniFile(strTrkFile)

i = InStr(MyString, "RouteID")
i = InStr(i, MyString, "(")
ii = InStr(i, MyString, ")")
strStart = Left(MyString, i)
strEnd = Mid(MyString, ii)
MyString = strStart & " " & strNew & " " & strEnd
DoEvents
i = InStr(i, MyString, "Name")
i = InStr(i, MyString, "(")
ii = InStr(i, MyString, ")")
strStart = Left(MyString, i)
strEnd = Mid(MyString, ii)
MyString = strStart & " " & Chr(34) & strNewIntName & Chr(34) & " " & strEnd
DoEvents
Call WriteUniFile(strTrkFile, MyString)
End Sub

Private Sub DoDeComp3(strFile As Variant, strFPath As String, strSparePath As String)
Dim strBatText As String, strSuffix As String

strSuffix = "-" & Right$(strFile, 1)


   ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java TSUtil fmgr " & strSuffix & " -e " & ChrW$(34) & "-n" & strFile & ChrW$(34) & " " & ChrW$(34) & strFPath & ChrW$(34) & " " & ChrW$(34) & strSparePath & ChrW$(34)


  Call ShellAndWait(strBatText, True, vbHide)

 DoEvents

End Sub

Private Sub DoDeComp2(strFile As String, strFPath As String, strSparePath As String)
Dim strBatText As String, strSuffix As String, booRetry As Boolean, strUni As String, NewFile As Integer

strSuffix = "-" & Right$(strFile, 1)
TryAgain:

   ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
 

strBatText = "java TSUtil fmgr " & strSuffix & " -e " & ChrW$(34) & "-n" & strFile & ChrW$(34) & " " & ChrW$(34) & strFPath & ChrW$(34) & " " & ChrW$(34) & strSparePath & ChrW$(34)
NewFile = FreeFile
   Open App.Path & "\TempFiles2\doBat.bat" For Output As #NewFile
   Print #NewFile, strBatText
   Close #NewFile
     ChDrive Left(App.Path, 1)
   ChDir App.Path & "\TempFiles2"
  DoEvents
    Call ShellAndWait("doBat.bat", True, vbHide)
        DoEvents
 
 DoEvents
 If strSuffix <> "-s" Then Exit Sub
  Rem ************ Check for uncompressed
  NewFile = FreeFile
    Open strSparePath & "\" & strFile For Binary As #NewFile
    strUni = String(2, " ")
    Get #NewFile, , strUni
     Close #NewFile

 If Asc(Mid$(strUni, 1, 1)) <> 255 Then
 If Asc(Mid$(strUni, 2, 1)) <> 254 Then
 If booRetry = False Then
 Rem ************* Uncompress
 
 Call DoDeCompFFEdit(strFPath & "\" & strFile)
 booRetry = True
 GoTo TryAgain
 End If
 End If
 End If
 If booRetry = False Then Exit Sub
 NewFile = FreeFile
   Open strFPath & "\" & strFile For Binary As #NewFile
    strUni = String(2, " ")
    Get #NewFile, , strUni
 Close #NewFile
 Rem ************* Uncompress
 If Asc(Mid$(strUni, 1, 1)) <> 255 Then
 If Asc(Mid$(strUni, 2, 1)) <> 254 Then
 Call MsgBox(strFile & " Did not uncompress", vbExclamation, App.Title)
 End If
 End If
End Sub
Private Sub DoComp(strFile As String, strFPath As String, strSparePath As String)
Dim strBatText As String, strSuffix As String

strSuffix = "-" & Right$(strFile, 1)


   ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
 If strSuffix = "-t" Then
 strBatText = "java TSUtil fmgr " & strSuffix & " -r -n" & ChrW$(34) & strFile & ChrW$(34) & " " & ChrW$(34) & strFPath & ChrW$(34) & " " & ChrW$(34) & strSparePath & ChrW$(34)
 ElseIf strSuffix <> "-t" Then

 
   strBatText = "java TSUtil fmgr " & strSuffix & " -c " & ChrW$(34) & "-n" & strFile & ChrW$(34) & " " & ChrW$(34) & strFPath & ChrW$(34) & " " & ChrW$(34) & strSparePath & ChrW$(34)

End If

  Call ShellAndWait(strBatText, True, vbHide)

 DoEvents
 End Sub
 
Private Sub DoDeCompFFEdit(strFile As String)
Dim strBatText As String, strOrigFile As String, NewFile As Integer, x As Integer
Dim result As String, strTempFile As String

On Error GoTo Errtrap

If Not FileExists(App.Path & "\TempFiles2\ffeditc_unicode.exe") Then
    If DirExists(MSTSPath & "\utils\ffedit") Then
    FileCopy MSTSPath & "\utils\ffedit\appids.tok", App.Path & "\TempFiles2\appids.tok"
    FileCopy MSTSPath & "\utils\ffedit\coreids.tok", App.Path & "\TempFiles2\coreids.tok"
    FileCopy MSTSPath & "\utils\ffedit\ffedit.cfg", App.Path & "\TempFiles2\ffedit.cfg"
    FileCopy MSTSPath & "\utils\ffedit\ffeditc_unicode.exe", App.Path & "\TempFiles2\ffeditc_unicode.exe"
    FileCopy MSTSPath & "\utils\ffedit\forms.hdr", App.Path & "\TempFiles2\forms.hdr"
    FileCopy MSTSPath & "\utils\ffedit\loadstr.hdr", App.Path & "\TempFiles2\loadstr.hdr"
    FileCopy MSTSPath & "\utils\ffedit\sidn.txt", App.Path & "\TempFiles2\sidn.txt"
    FileCopy MSTSPath & "\utils\ffedit\worldfile.bnf", App.Path & "\TempFiles2\worldfile.bnf"
    FileCopy MSTSPath & "\utils\ffedit\newshape.bnf", App.Path & "\TempFiles2\newshape.bnf"
    Else
    
    result = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "Path")
    
        If DirExists(result & "\utils\ffedit") Then
        FileCopy result & "\utils\ffedit\appids.tok", App.Path & "\TempFiles2\appids.tok"
        FileCopy result & "\utils\ffedit\coreids.tok", App.Path & "\TempFiles2\coreids.tok"
        FileCopy result & "\utils\ffedit\ffedit.cfg", App.Path & "\TempFiles2\ffedit.cfg"
        FileCopy result & "\utils\ffedit\ffeditc_unicode.exe", App.Path & "\TempFiles2\ffeditc_unicode.exe"
        FileCopy result & "\utils\ffedit\forms.hdr", App.Path & "\TempFiles2\forms.hdr"
        FileCopy result & "\utils\ffedit\loadstr.hdr", App.Path & "\TempFiles2\loadstr.hdr"
        FileCopy result & "\utils\ffedit\sidn.txt", App.Path & "\TempFiles2\sidn.txt"
        FileCopy result & "\utils\ffedit\worldfile.bnf", App.Path & "\TempFiles2\worldfile.bnf"
        FileCopy result & "\utils\ffedit\newshape.bnf", App.Path & "\TempFiles2\newshape.bnf"
        Else
           
        Call MsgBox("Could not find the Utils\FFEDIT folder in MSTS, this folder is required to process this file.", vbExclamation, App.Title)
        Exit Sub
        End If
    
    End If
    End If

MousePointer = 11
x = InStrRev(strFile, "\")
strOrigFile = Mid(strFile, x + 1)

        FileCopy strFile, App.Path & "\TempFiles2\" & strOrigFile
   Rem ************************
   Name App.Path & "\TempFiles2\" & strOrigFile As App.Path & "\TempFiles2\tempfile.s"
   strTempFile = App.Path & "\TempFiles2\tempfile.s"
   Rem *************************
   strBatText = "ffeditc_unicode.exe " & ChrW$(34) & "tempfile.s" & ChrW$(34) & " /c " & ChrW$(34) & "/o:" & "tempfile.s" & ChrW$(34) & vbCrLf
 
   
   NewFile = FreeFile
   Open App.Path & "\TempFiles2\doFfeditc.bat" For Output As #NewFile
   Print #NewFile, strBatText
   Close #NewFile
       
   ChDrive Left(App.Path, 1)
   ChDir App.Path & "\TempFiles2"
  DoEvents
    Call ShellAndWait("doffeditc.bat", True, vbHide)
        DoEvents


    DoEvents
    Kill strFile
    DoEvents
    Name App.Path & "\TempFiles2\tempfile.s" As App.Path & "\TempFiles2\" & strOrigFile
    DoEvents

     FileCopy App.Path & "\TempFiles2\" & strOrigFile, strFile
     DoEvents

    MousePointer = 0
    
Exit Sub
Errtrap:

Resume Next
End Sub

Private Sub DoDeCompFolder(strSuffix As String, strFPath As String, strSparePath As String)
Dim strBatText As String, fullpath$, strUni As String, strFile As String

On Error GoTo Errtrap

strSuffix = "-" & strSuffix


   ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java TSUtil fmgr " & strSuffix & " -e -o " & ChrW$(34) & strFPath & ChrW$(34) & " " & ChrW$(34) & strSparePath & ChrW$(34)


  Call ShellAndWait(strBatText, True, vbHide)

 DoEvents
 Rem ****************** Check for .s files which did not uncompress
 If strSuffix <> "-s" Then Exit Sub
 Close
     cursouind = 1
SparePath = App.Path & "\TempFiles"
frmUtils.Drive1(1).Drive = Left$(SparePath, 2)
frmUtils.Dir1(1).Path = SparePath
frmUtils.Text1(1).Text = "*.S"
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i
For i = 0 To frmUtils.File1(cursouind).ListCount - 1

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = frmUtils.File1(cursouind).Path
   strFile = frmUtils.File1(cursouind).List(i)

    Open fullpath$ & "\" & strFile For Binary As #5
    strUni = String(2, " ")
    Get #5, , strUni
     Close #5

 If Asc(Mid$(strUni, 1, 1)) <> 255 Then
 If Asc(Mid$(strUni, 2, 1)) <> 254 Then
 Rem ************* Uncompress

 Call DoDeCompFFEdit(fullpath$ & "\" & strFile)
 DoEvents
 strSuffix = "-s"
 ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java TSUtil fmgr " & strSuffix & " -e " & ChrW$(34) & "-n" & strFile & ChrW$(34) & " " & ChrW$(34) & strFPath & ChrW$(34) & " " & ChrW$(34) & strSparePath & ChrW$(34)


  Call ShellAndWait(strBatText, True, vbHide)
DoEvents

    Open fullpath$ & "\" & strFile For Binary As #5
    strUni = String(2, " ")
    Get #5, , strUni
     Close #5

 If Asc(Mid$(strUni, 1, 1)) <> 255 Then
 If Asc(Mid$(strUni, 2, 1)) <> 254 Then
 Call MsgBox(strFile & " was not decompressed, so could not be processed. This may cause its corresponding" _
             & vbCrLf & ".ace file to be missing from your route." _
             , vbExclamation, App.Title)
 
 End If
End If
 End If
 End If
 End If
 Next i
Exit Sub
Errtrap:
Call MsgBox("An error #" & Err & " occurred in subroutine 'DoDecompFolder' while checking" _
            & vbCrLf & strFile _
            , vbExclamation, frmConsists)
Resume Next
 End Sub
 
Private Sub DoCompFolder(strSuffix As String, strFPath As String, strSparePath As String)
Dim strBatText As String

strSuffix = "-" & strSuffix


   ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
 If strSuffix <> "-t" Then
   strBatText = "java TSUtil fmgr " & strSuffix & " -c -o " & ChrW$(34) & strFPath & ChrW$(34) & " " & ChrW$(34) & strSparePath & ChrW$(34)
ElseIf strSuffix = "-t" Then
strBatText = "java TSUtil fmgr " & strSuffix & " -r -o " & ChrW$(34) & strFPath & ChrW$(34) & " " & ChrW$(34) & strSparePath & ChrW$(34)
End If

  Call ShellAndWait(strBatText, True, vbHide)

 DoEvents
 End Sub
 
Private Sub FindStuck()
Dim i As Integer, filepath1$, fullpath$, x As Long, Z As Long, Y As Long, xx As Long
Dim strNew As String
Dim MyString As String, yy As Long, zz As Integer, q As Integer
Dim strUiD As String, ix As Integer, NewFile As Integer, A$, intPoints As Integer
Rem

On Error GoTo Errtrap
MousePointer = 11
cursouind = 1
filepath1$ = App.Path & "\TempFiles"
Rem ******************
If Not FileExists(App.Path & "\SetupFiles\Points.txt") Then

     Call MsgBox(App.Path & "\SetupFiles\Points.txt" & vbCrLf & " is missing and must be installed before this option can be used", vbCritical, App.Path)
Exit Sub
End If
 NewFile = FreeFile

Open App.Path & "\SetupFiles\Points.txt" For Input As #NewFile
Input #NewFile, intPoints
ReDim strPoints(0 To intPoints - 1)
For i = 0 To intPoints - 1
Input #NewFile, A$
strPoints(i) = A$
Next i
Close #NewFile
Rem *******************************
'strPoints(0) = "Pnt"
'strPoints(1) = "swt_"
'strPoints(2) = "45dYard"
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.W"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    Label9.Caption = "Processing: " & File1(cursouind).List(i)
   
    MyString = ReadUniFile(fullpath$)
    For ix = 0 To intPoints - 1
    Y = 1
    Do
    x = InStr(Y, MyString, strPoints(ix))
    If x > 0 Then
    If Mid(MyString, x - 3, 3) <> "End" And Mid(MyString, x - 5, 5) <> "Dummy" And Mid(MyString, x + 7, 3) <> "Crv" Then
    Z = InStrRev(MyString, "TrackObj", x)
    If Z = 0 Or (x - Z > 250) Then GoTo CarryON
    xx = InStr(x, MyString, "StaticDetailLevel")
        If xx = 0 Then
        xx = InStr(x, MyString, "TrackObj (")
            If xx = 0 Then
            xx = InStr(x, MyString, "Static (")
                If xx = 0 Then
            xx = InStr(x, MyString, "Forest (")
            End If
            End If
        End If
        If xx = 0 Then GoTo CarryON
    strNew = Mid(MyString, Z, xx - Z)
    yy = InStr(strNew, "JNodePosn")
    If yy = 0 Then
    zz = InStr(strNew, "UiD")
    q = InStr(zz, strNew, ")")
    strUiD = Mid(strNew, zz, (q - zz) + 1)
    strReport = strReport & "No JNodePosn entry in " & File1(cursouind).List(i) & " " & strUiD & vbCrLf
    End If
    End If
    End If
CarryON:
    Y = x + 1
    Loop While x > 0
  Next ix

    End If
    Next i
    MousePointer = 0
    If strReport = "" Then
    strReport = "No Stuck Points Found"
    End If
    
    frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
Exit Sub
Errtrap:
 Call MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'FindStuck' " _
            & vbCrLf & "while processing " & File1(cursouind).List(i) _
            , vbExclamation, App.Title)
            
            Resume Next
   
    
End Sub

Private Sub FixStuck(strSparePath As String, strWorldPath As String)
Dim i As Integer, filepath1$, fullpath$, x As Long, Z As Long, Y As Long, xx As Long
Dim strNew As String, strPos As String
Dim MyString As String, yy As Long, zz As Integer, q As Integer, qq As Long
Dim strUiD As String, strA As String, strB As String, strTemp As String
Dim strStart As String, strEnd As String, booStuck As Boolean
Dim ix As Integer, NewFile As Integer, A$, intPoints As Integer

cursouind = 1
MousePointer = 11
filepath1$ = App.Path & "\TempFiles"
'strPoints(0) = "Pnt"
'strPoints(1) = "swt_"
'strPoints(2) = "45dYard"
If Not FileExists(App.Path & "\SetupFiles\Points.txt") Then

     Call MsgBox(App.Path & "\SetupFiles\Points.txt" & vbCrLf & " is missing and must be installed before this option can be used", vbCritical, App.Path)
Exit Sub
End If
 NewFile = FreeFile

Open App.Path & "\SetupFiles\Points.txt" For Input As #NewFile
Input #NewFile, intPoints
ReDim strPoints(0 To intPoints - 1)
For i = 0 To intPoints - 1
Input #NewFile, A$
strPoints(i) = A$
Next i
Close #NewFile
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.W"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    MyString = ReadUniFile(fullpath$)
    Label9.Caption = "Processing: " & File1(cursouind).List(i)
    DoEvents
   For ix = 0 To intPoints - 1
    Y = 1
    Do

 x = InStr(Y, MyString, strPoints(ix))
    If x > 0 Then
    If Mid(MyString, x - 3, 3) <> "End" And Mid(MyString, x - 5, 5) <> "Dummy" And Mid(MyString, x + 7, 3) <> "Crv" Then
    Z = InStrRev(MyString, "TrackObj", x)
    If Z = 0 Or (x - Z > 250) Then GoTo CarryON
    xx = InStr(x, MyString, "StaticDetailLevel")
        If xx = 0 Then
        xx = InStr(x, MyString, "TrackObj (")
            If xx = 0 Then
            xx = InStr(x, MyString, "Static (")
                If xx = 0 Then
            xx = InStr(x, MyString, "Forest (")
            End If
            End If
        End If
        If xx = 0 Then GoTo CarryON
    qq = InStrRev(MyString, "CollideFlags", x)
    strNew = Mid(MyString, Z, xx - Z)
    yy = InStr(strNew, "JNodePosn")
    If yy = 0 Then
    zz = InStr(strNew, "UiD")
    q = InStr(zz, strNew, ")")
    strUiD = Mid(strNew, zz, (q - zz) + 1)
    zz = InStr(strNew, "Position")
    zz = InStr(zz, strNew, "(")
    q = InStr(zz, strNew, ")")
    strPos = Mid(strNew, zz + 1, (q - zz) - 1)
    strPos = Trim(strPos)
    strA = File1(cursouind).List(i)
    strA = Mid(strA, 2)
    strA = Left(strA, Len(strA) - 2)
    strB = Mid(strA, 8)
    strA = Left(strA, 7)
    If Left(strA, 1) = "+" Then
    strTemp = ""
    Else
    strTemp = "-"
    End If
  
    strA = Mid(strA, 2)
    Do
    If Left(strA, 1) = "0" Then
    strA = Mid(strA, 2)
    End If
    Loop Until Left(strA, 1) <> "0"
    strA = strTemp & strA
    If Left(strB, 1) = "+" Then
    strTemp = ""
    Else
    strTemp = "-"
    End If
    strB = Mid(strB, 2)
    Do
    If Left(strB, 1) = "0" Then
    strB = Mid(strB, 2)
    End If
    Loop Until Left(strB, 1) <> "0"
    strB = strTemp & strB
    strTemp = "JNodePosn ( " & strA & " " & strB & " " & strPos & " )" & vbCrLf & Chr(9) & vbTab
    strStart = Left(MyString, qq - 1)
    strEnd = Mid(MyString, qq)
    MyString = strStart & strTemp & strEnd
    Call WriteUniFile(fullpath$, MyString)
   
    strReport = strReport & "No JNodePosn entry in " & File1(cursouind).List(i) & " " & strUiD & " fixed" & vbCrLf
   booStuck = True
    End If
    End If
    End If
CarryON:
    Y = x + 1
    Loop While x > 0
  Next ix

    End If
    If booStuck = True Then
    Call DoComp(File1(cursouind).List(i), strSparePath, strWorldPath)
    DoEvents
    booStuck = False
    End If
    Next i
    MousePointer = 0
   If strReport = "" Then
    strReport = "No Stuck Points Found"
    End If
    
    frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
End Sub



Private Sub FixVDbId()
Dim i As Integer, filepath1$, fullpath$, x As Long, Z As Long, Y As Long
Dim MyString As String
Dim strStart As String, strEnd As String, strTemp As String, ii As Integer, X1(1 To 16) As String, y1(1 To 16) As Double
Dim dLow As Double, dHigh As Double
Rem
X1(1) = "Static ("
X1(2) = "Gantry ("
X1(3) = "Trackobj ("
X1(4) = "Forest ("
X1(5) = "Signal ("
X1(6) = "Speedpost ("
X1(7) = "tr_watermark"
X1(8) = "Levelcr ("
X1(9) = "Pickup ("
X1(10) = "CollideObject ("
X1(11) = "Hazard ("
X1(12) = "Dyntrack ("
X1(13) = "Siding ("
X1(14) = "Platform ("
X1(15) = "CarSpawner ("
X1(16) = "Transfer ("

cursouind = 1
filepath1$ = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.W"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

If File1(cursouind).Selected(i) Then
fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
 MyString = ReadUniFile(fullpath$)
 Y = 1
 Do
 x = InStr(Y, MyString, "VDbId (")
 If x = 0 Then Exit Do
 Z = InStr(x, MyString, ")")
 strTemp = Mid(MyString, x + 7, Z - (x + 7))
 If strTemp <> " 4294967294 " Then
 strStart = Left(MyString, x + 6)
 strEnd = Mid(MyString, Z)
 MyString = strStart & " 4294967294 " & strEnd
 End If
 Y = Z
 Loop
 x = InStr(MyString, "VDbIdCount")
 If x = 0 Then GoTo CarryON
 For ii = 1 To 16
 y1(ii) = InStr(MyString, X1(ii))
 Next ii
 Call FindMinMax(y1(), dLow, dHigh)
 
 strStart = Left(MyString, x - 1)
 strEnd = Mid(MyString, dLow)
 MyString = strStart & strEnd
CarryON:
End If
Call WriteUniFile(fullpath$, MyString)
Rem ************ Do not change yet ***********************
ii = InStrRev(fullpath$, "\")
strTemp = Mid(fullpath$, ii)
Kill WorldPath & strTemp
FileCopy fullpath$, WorldPath & strTemp
DoEvents
Next i


  
     
     DoEvents
     
End Sub

Private Sub GetAllRoutes()
Dim j As Integer, i As Integer, x As Integer

Drive1(cursouind).Drive = Left$(MSTSPath, 2)
Dir1(cursouind).Path = MSTSPath & "\Routes"

j = 0
For i = 0 To Dir1(cursouind).ListCount - 1
x = InStrRev(Dir1(cursouind).List(i), "\")

AllRoutes(j) = Dir1(cursouind).List(i)
'a$ = AllRoutes(j)
'Print #27, a$
j = j + 1
      
              
              If j > UBound(AllRoutes) Then
           ReDim Preserve AllRoutes(0 To j + REF_CHUNK)
           End If
Next i

NumRoutes = j
ReDim Preserve AllRoutes(0 To j - 1)
Rem ******** Get All Routes ****************************

Drive1(cursouind).Drive = Left$(MSTSPath, 2)
Dir1(cursouind).Path = MSTSPath & "\Routes"

j = 0
For i = 0 To Dir1(cursouind).ListCount - 1
x = InStrRev(Dir1(cursouind).List(i), "\")
If Right$(Dir1(cursouind).List(i), 6) = "Common" Then GoTo GetAnother
AllRoutes2(j) = Dir1(cursouind).List(i)
j = j + 1


              If j > UBound(AllRoutes2) Then
           ReDim Preserve AllRoutes2(0 To j + REF_CHUNK)
           End If
GetAnother:
Next i

NumRoutes = j
ReDim Preserve AllRoutes2(0 To j - 1)

End Sub


Private Sub LiftW(strHeight As String)
Dim i As Integer, filepath1$, fullpath$, ii As Integer, j As Integer
Dim MyString As String, x As Long, Y As Long, xx As Long
Dim dHeight As Double, strPosn As String, strStart As String, strEnd As String
Dim strPosition As String, dposn As Double

cursouind = 1
dHeight = Val(strHeight)
filepath1$ = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.W"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
  MyString = ReadUniFile(fullpath$)
  x = 1
  Do
  x = InStr(x, MyString, "Position (")
  If x = 0 Then Exit Do
  xx = InStr(x, MyString, "Radius (")
  If xx > 0 Then GoTo CarryON
  Rem *********** OK got a good one *******

  Y = InStr(x, MyString, ")")
  strPosn = Trim(Mid(MyString, x + 10, Y - (x + 10)))
  j = InStr(strPosn, " ")
  ii = InStr(j + 1, strPosn, " ")
  strPosition = Mid(strPosn, j + 1, ii - j)
  dposn = Val(strPosition) + dHeight
  strStart = Left(MyString, x + 10 + j)
  strEnd = Mid(MyString, x + 10 + ii)
  strPosition = Trim(Str(dposn))
  MyString = strStart & strPosition & strEnd

  
  
  '**************************
CarryON:
x = x + 1
  Loop
DoEvents
Call WriteUniFile(fullpath$, MyString)


    Kill WorldPath & "\" & File1(cursouind).List(i)
    
DoEvents
FileCopy fullpath$, WorldPath & "\" & File1(cursouind).List(i)
DoEvents
End If
    Next i
End Sub

Public Sub ListTrainsCopy()
Dim strActPath As String
Dim strSrvPath As String, Trainspath As String, ConPath As String
Dim ConsistPath As String, i As Integer

MousePointer = 11

Trainspath = MSTSPath & "\Trains\"
ConPath = MSTSPath & "\Trains\Consists\"
 Drive1(1).Drive = Left$(Trainspath, 2)
Dir1(1).Path = Trainspath
Text1(1).Text = "*.*"
strReport = strReport & "ROLLING STOCK FOR " & RoutePath & vbCrLf & vbCrLf
strReport = UCase(strReport)

 strActPath = RoutePath & "\Activities\"
 strSrvPath = RoutePath & "\Services\"
 strConPath = Dir1(1).Path & "\Consists\"
 strTrainset = Dir1(1).Path & "\Trainset\"

 
 ConsistPath = MSTSPath & "\Trains\Consists"
Dir1(1).Path = strConPath
Text1(1).Text = "*.con"
            
Dir1(0).Path = strActPath
Text1(0).Text = "*.act"
For i = 0 To File1(0).ListCount - 1
    File1(0).Selected(i) = True
Next i

For i = 0 To File1(0).ListCount - 1
Call ListLooseActConsists(strActPath & "\" & File1(0).List(i))
Next i


DoEvents
Dir1(0).Path = strSrvPath
Text1(0).Text = "*.srv"
For i = 0 To File1(0).ListCount - 1
    File1(0).Selected(i) = True
Next i

For i = 0 To File1(0).ListCount - 1
Call ListCheckService(strSrvPath & File1(0).List(i))
Next i


DoEvents

DoEvents
frmReport.Rich1.Text = strReport
 
     frmReport.Show 1
 strReport = vbNullString
MousePointer = 0
End Sub


Public Sub MiniTrainsCopy()
Dim strActPath As String, strTrain As String
Dim strSrvPath As String, i As Integer
Dim ConsistPath As String

On Error GoTo Errtrap

MousePointer = 11


strTrain = Dir1(1).Path
If Right$(strTrain, 6) <> "Trains" Then
Call MsgBox("Select your Mini-Route in the left hand window and" _
            & vbCrLf & "your Main MSTS Trains folder in the right hand window." _
            , vbExclamation, App.Title)
            Exit Sub
            End If
  
 strActPath = RoutePath & "\Activities\"
 strSrvPath = RoutePath & "\Services\"
 strConPath = Dir1(1).Path & "\Consists\"
 strTrainset = Dir1(1).Path & "\Trainset\"

 If Not DirExists(MSTSPath & "\Trains") Then
 MkDir MSTSPath & "\Trains"
 End If
 If Not DirExists(MSTSPath & "\Trains\Trainset") Then
 MkDir MSTSPath & "\Trains\Trainset"
 End If
 If Not DirExists(MSTSPath & "\Trains\Consists") Then
 MkDir MSTSPath & "\Trains\Consists"
 End If
 
 ConsistPath = MSTSPath & "\Trains\Consists"
Dir1(1).Path = strConPath
Text1(1).Text = "*.con"
            
Dir1(0).Path = strActPath
Text1(0).Text = "*.act"
For i = 0 To File1(0).ListCount - 1
    File1(0).Selected(i) = True
Next i

For i = 0 To File1(0).ListCount - 1
Call MiniLooseActConsists(strActPath & "\" & File1(0).List(i))
Next i


DoEvents
Dir1(0).Path = strSrvPath
Text1(0).Text = "*.srv"
For i = 0 To File1(0).ListCount - 1
    File1(0).Selected(i) = True
Next i

For i = 0 To File1(0).ListCount - 1
Call MiniCheckService(strSrvPath & File1(0).List(i))
Next i


DoEvents
Dir1(0).Path = ConsistPath
Text1(0).Text = "*.con"
DoEvents
For i = 0 To File1(0).ListCount - 1
    File1(0).Selected(i) = True
Next i

For i = 0 To File1(0).ListCount - 1
Call MiniLooseActConsists(ConsistPath & "\" & File1(0).List(i))
Next i

DoEvents
If Not DirExists(MSTSPath & "\Trains\Trainset\Default") Then
MkDir MSTSPath & "\Trains\Trainset\Default"
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "Default\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\Default" & ChrW$(34) & " /S /I /C" & vbCrLf
End If
If DirExists(strTrainset & "common.sound") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\Common.sound") Then
    MkDir MSTSPath & "\Trains\Trainset\Common.sound"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "common.sound\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\common.sound" & ChrW$(34) & " /S /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "common.cab") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\Common.cab") Then
    MkDir MSTSPath & "\Trains\Trainset\Common.cab"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "common.cab\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\common.cab" & ChrW$(34) & " /S /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "common.snd") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\Common.snd") Then
    MkDir MSTSPath & "\Trains\Trainset\Common.snd"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "common.snd\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\common.snd" & ChrW$(34) & " /S /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "common.crew") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\Common.crew") Then
    MkDir MSTSPath & "\Trains\Trainset\Common.crew"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "common.crew\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\common.crew" & ChrW$(34) & " /S /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "Common.Loads") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\Common.Loads") Then
    MkDir MSTSPath & "\Trains\Trainset\Common.Loads"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "Common.Loads\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\Common.Loads" & ChrW$(34) & " /S /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "ACELA") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\ACELA") Then
    MkDir MSTSPath & "\Trains\Trainset\ACELA"
    MkDir MSTSPath & "\Trains\Trainset\ACELA\Cabview"
    MkDir MSTSPath & "\Trains\Trainset\ACELA\Sound"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "ACELA\Sound\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\ACELA\Sound" & ChrW$(34) & " /I /C" & vbCrLf
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "ACELA\Cabview\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\ACELA\Cabview" & ChrW$(34) & " /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "DASH9") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\DASH9") Then
    MkDir MSTSPath & "\Trains\Trainset\DASH9"
    MkDir MSTSPath & "\Trains\Trainset\DASH9\Cabview"
    MkDir MSTSPath & "\Trains\Trainset\DASH9\Sound"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "DASH9\Cabview\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\DASH9\Cabview" & ChrW$(34) & " /I /C" & vbCrLf
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "DASH9\Sound\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\DASH9\Sound" & ChrW$(34) & " /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "GP38") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\GP38") Then
    MkDir MSTSPath & "\Trains\Trainset\GP38"
    MkDir MSTSPath & "\Trains\Trainset\GP38\Cabview"
    MkDir MSTSPath & "\Trains\Trainset\GP38\Sound"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "GP38\Cabview\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\GP38\Cabview" & ChrW$(34) & " /I /C" & vbCrLf
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "GP38\Sound\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\GP38\Sound" & ChrW$(34) & " /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "SCOTSMAN") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\SCOTSMAN") Then
    MkDir MSTSPath & "\Trains\Trainset\SCOTSMAN"
    MkDir MSTSPath & "\Trains\Trainset\SCOTSMAN\Cabview"
    MkDir MSTSPath & "\Trains\Trainset\SCOTSMAN\Sound"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "SCOTSMAN\Cabview\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\SCOTSMAN\Cabview" & ChrW$(34) & " /I /C" & vbCrLf
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "SCOTSMAN\Sound\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\SCOTSMAN\Sound" & ChrW$(34) & " /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "KIHA31") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\KIHA31") Then
    MkDir MSTSPath & "\Trains\Trainset\KIHA31"
    MkDir MSTSPath & "\Trains\Trainset\KIHA31\Cabview"
    MkDir MSTSPath & "\Trains\Trainset\KIHA31\Sound"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "KIHA31\Cabview\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\KIHA31\Cabview" & ChrW$(34) & " /I /C" & vbCrLf
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "KIHA31\Sound\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\KIHA31\Sound" & ChrW$(34) & " /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "HHP") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\HHP") Then
    MkDir MSTSPath & "\Trains\Trainset\HHP"
    MkDir MSTSPath & "\Trains\Trainset\HHP\Cabview"
    MkDir MSTSPath & "\Trains\Trainset\HHP\Sound"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "HHP\Cabview\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\HHP\Cabview" & ChrW$(34) & " /I /C" & vbCrLf
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "HHP\Sound\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\HHP\Sound" & ChrW$(34) & " /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "Series2000") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\Series2000") Then
    MkDir MSTSPath & "\Trains\Trainset\Series2000"
    MkDir MSTSPath & "\Trains\Trainset\Series2000\Cabview"
    MkDir MSTSPath & "\Trains\Trainset\Series2000\Sound"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "Series2000\Cabview\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\Series2000\Cabview" & ChrW$(34) & " /I /C" & vbCrLf
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "Series2000\Sound\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\Series2000\Sound" & ChrW$(34) & " /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "Series7000") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\Series7000") Then
    MkDir MSTSPath & "\Trains\Trainset\Series7000"
    MkDir MSTSPath & "\Trains\Trainset\Series7000\Cabview"
    MkDir MSTSPath & "\Trains\Trainset\Series7000\Sound"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "Series7000\Cabview\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\Series7000\Cabview" & ChrW$(34) & " /I /C" & vbCrLf
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "Series7000\Sound\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\Series7000\Sound" & ChrW$(34) & " /I /C" & vbCrLf
    End If
End If
If DirExists(strTrainset & "380") Then
    If Not DirExists(MSTSPath & "\Trains\Trainset\380") Then
    MkDir MSTSPath & "\Trains\Trainset\380"
    MkDir MSTSPath & "\Trains\Trainset\380\Cabview"
    MkDir MSTSPath & "\Trains\Trainset\380\Sound"
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "380\Cabview\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\380\Cabview" & ChrW$(34) & " /I /C" & vbCrLf
    strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & "380\Sound\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\380\Sound" & ChrW$(34) & " /I /C" & vbCrLf
    End If
End If

If strBatFile <> vbNullString Then

Open App.Path & "\TempFiles\debug.bat" For Output As #1
Print #1, strBatFile
Close #1
ChDrive Left$(App.Path, 1)
 ChDir App.Path & "\TempFiles"

DoEvents
Call ShellAndWait("debug.bat", True, vbNormalFocus)

DoEvents
End If
MousePointer = 0
Exit Sub
Errtrap:
Call MsgBox("Error " & Err.Description & " occurred in MiniTrainsCopy ", vbExclamation, App.Title)

End Sub



Private Sub CompactTransfer(strNew As Variant)
Dim j As Integer, strESD As String, itExists As Boolean, JNumber As Long



On Error GoTo Errtrap
  For j = 1 To numAce
   If strNew = Ace1(j) Then
   itExists = True
   Exit For
   End If
  Next j

   If itExists = False Then
   numAce = numAce + 1
   
                If numAce > UBound(Ace1) Then
                    ReDim Preserve Ace1(1 To numAce + Shp_Chunk)
                    End If
                DoEvents
   
   Ace1(numAce) = strNew
   JNumber = numAce
   
   ESD(numAce) = "1"
   
   ElseIf itExists = True Then
   JNumber = j
                     
            strESD = "1"
            End If
   ' End If
    If ESD(JNumber) = "252" And strESD = "257" Then
    ESD(JNumber) = "259"
    ElseIf ESD(JNumber) = "257" And strESD = "252" Then
    ESD(JNumber) = "259"
    ElseIf ESD(JNumber) = "252" And strESD <> "257" Then
    ESD(JNumber) = "252"
    ElseIf Val(strESD) > Val(ESD(JNumber)) Then
   
   ESD(JNumber) = strESD
   
   End If

 
  
  itExists = False
 Exit Sub
 
Errtrap:
 Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'Compact Transfer' please advise" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
    Exit Sub
    End Select
   
   
End Sub
Private Sub CompactAllAce(strNew As String, strSDFile As String, booCar As Boolean)
Dim j As Integer, strESD As String, JNumber As Long, itExists As Boolean
Dim GlobalPath As String

On Error GoTo Errtrap
GlobalPath = MSTSPath & "\Global\Shapes\"
  For j = 1 To numAce
   If strNew = Ace1(j) Then
        itExists = True
        Exit For
   End If
  Next j

   If itExists = False Then
   numAce = numAce + 1
   
                If numAce > UBound(Ace1) Then
                    ReDim Preserve Ace1(1 To numAce + Shp_Chunk)
                End If
                DoEvents
   
   Ace1(numAce) = strNew
   JNumber = numAce
   If booCar = True Then
   ESD(numAce) = "0"
   Exit Sub
   End If
   

         If FileExists(RoutePath & "\shapes\" & strSDFile) And Not FileExists(GlobalPath & strSDFile) Then
         Call GetSeasons(RoutePath & "\shapes\" & strSDFile, strESD)
        
         ESD(numAce) = strESD
         ElseIf Not FileExists(RoutePath & "\shapes\" & strSDFile) And Not FileExists(GlobalPath & strSDFile) Then
         strESD = "0"
         Else
           strESD = "2"
         End If
   ElseIf itExists = True Then
   JNumber = j
            If FileExists(RoutePath & "\shapes\" & strSDFile) And Not FileExists(GlobalPath & strSDFile) Then
            Call GetSeasons(RoutePath & "\shapes\" & strSDFile, strESD)
            ElseIf Not FileExists(RoutePath & "\shapes\" & strSDFile) And Not FileExists(GlobalPath & strSDFile) Then
         strESD = "0"
         Else
          
            strESD = "2"
            End If
    End If
    If ESD(JNumber) = "252" And strESD = "257" Then
    ESD(JNumber) = "259"
    ElseIf ESD(JNumber) = "257" And strESD = "252" Then
    ESD(JNumber) = "259"
    ElseIf ESD(JNumber) = "252" And strESD = "1" Or strESD = "2" Then
    ESD(JNumber) = "259"
    ElseIf ESD(JNumber) = "1" And strESD = "252" Then
    ESD(JNumber) = "259"
    ElseIf ESD(JNumber) = "2" And strESD = "252" Then
    ESD(JNumber) = "259"
    ElseIf ESD(JNumber) = "252" And strESD <> "257" Then
    ESD(JNumber) = "252"
    ElseIf Val(strESD) > Val(ESD(JNumber)) Then
   
   ESD(JNumber) = strESD
   
   End If
   
   
 
   strNew = vbNullString
  itExists = False
   
  Exit Sub
Errtrap:
Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'CompactAllAce' please advise" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
        'Resume Next
    Exit Sub
    End Select
   
End Sub


Private Sub CompactRef()
Dim Filpath1$
Dim strNew As String
Dim intShapes As Integer, booFound As Boolean, MyString As String
Dim x As Long, Y As Integer, yy As Long, strRef As String, i As Integer, ii As Integer
Dim strRef1 As String, strRef2 As String, strRef3 As String, GlobalPath As String
Dim strTrans1 As String, strTrans2 As String, strTrans3 As String, strTrans4 As String
Dim tempHaz As String, Newfile3 As Integer, NewFile As Integer, strTemp2 As String, NewFile2 As Integer
Dim TertexPath As String, Min As Double, q As Integer, AceTemp(1 To 100) As String, strSpare As String
Dim MinIndex As Long, tokStart As Double, xx As Long, MinNext As Double, strS As String, intAce As Integer
Dim Z As Long, zz As Long, strFName As String, strTerr As String, strGlobal As String, strTemp As String

ReDim MasterIndex(0 To REF_CHUNK)
ReDim MasterFile(0 To REF_CHUNK)

On Error GoTo Errtrap


MousePointer = 11
cursouind = 1
TertexPath = RoutePath & "\Terrtex"


GlobalPath = MSTSPath & "\Global\Shapes\"
strGlobal = MSTSPath & "\Global\Shapes"
Filpath1$ = App.Path & "\setupfiles\"

If FileExists(Filpath1$ & "tempref.ref") Then
Kill Filpath1$ & "tempref.ref"
End If
If FileExists(Filpath1$ & "tempunk.ref") Then

Kill Filpath1$ & "tempunk.ref"
End If
SB1.Panels(2).Text = "Reference file"


FileCopy Filpath1$ & "reffilestart.txt", Filpath1$ & "tempref.ref"

strRef = ReadUniFile(Filpath1$ & "master.ref")
strRef = Replace(strRef, vbTab, " ")
strRef = Replace(strRef, "        (", " (")
strRef = Replace(strRef, "       (", " (")
strRef = Replace(strRef, "      (", " (")
strRef = Replace(strRef, "     (", " (")
strRef = Replace(strRef, "    (", " (")
strRef = Replace(strRef, "   (", " (")
strRef = Replace(strRef, "  (", " (")
strRef = Replace(strRef, "         (", " (")
strRef = Replace(strRef, "          (", " (")
strRef = Replace(strRef, "Static(", "Static (")

x = 1

GetAnother:
Call GetFirstToken(strRef, Min, MinIndex, x)


                    

tokStart = Min
x = Min + 1

Call GetNextToken(strRef, MinNext, MinIndex, x)


If MinNext = 0 Then GoTo Label1
strTemp2 = Mid$(strRef, tokStart, MinNext - tokStart)


              Rem***************
              xx = InStr(strTemp2, "FileName")
             
                     If xx > 0 Then
                     Y = InStr(xx, strTemp2, "(")
                     yy = InStr(Y, strTemp2, ")")
                     strNew = Mid$(strTemp2, Y + 1, yy - (Y + 1))
                     strNew = Trim$(strNew)
                     
                   
                              
                              If Left$(strNew, 1) = ChrW$(34) Then
                                strNew = Mid$(strNew, 2)
                                Y = InStr(strNew, ChrW$(34))
                                 If Y > 0 Then
                                 strNew = Left$(strNew, Y - 1)
                                 End If
                                End If
                    End If
              Rem****************
         SB1.Panels(2).Text = strNew
              MasterIndex(i) = strNew
              MasterFile(i) = strTemp2
              strTemp = vbNullString
              strTemp2 = vbNullString
            
              i = i + 1
              
              If i > UBound(MasterIndex) Then
           ReDim Preserve MasterIndex(0 To i + REF_CHUNK)
           ReDim Preserve MasterFile(0 To i + REF_CHUNK)
           End If
            
TryAgain:
      
       x = MinNext - 1
       If booEndFile = False Then
       GoTo GetAnother
       End If
  

Label1:
 
 
 ReDim Preserve MasterIndex(0 To i)
 ReDim Preserve MasterFile(0 To i)
 intShapes = i - 1
 'Call FillGrid
 booEndFile = False

strRef1 = "Static (" & vbCrLf & "        FileName ( " & ChrW$(34)
strRef2 = ChrW$(34) & ")" & vbCrLf & "        Shadow ( RECT )" & vbCrLf & "        Class ( " & ChrW$(34) & "Misc" & ChrW$(34) & " )"
strRef3 = "        Align ( None )" & vbCrLf & "        Description ( " & ChrW$(34)

GlobalPath = MSTSPath & "\global\shapes\"
For i = 0 To numShp - 1
 

            For ii = 0 To intShapes
                If MasterIndex(ii) = strShp(i) Then
                strTemp = MasterFile(ii)
                NewFile = FreeFile
                Open Filpath1$ & "\TempRef.ref" For Append As #NewFile
                   Print #NewFile, strTemp
                   Close NewFile
                   booFound = True
                GoTo Label7
                End If
            Next ii
    
   
            Rem ******************* No .ref

            If Not FileExists(GlobalPath & strShp(i)) And strShp(i) <> vbNullString Then
            strTemp = strRef1 & strShp(i) & strRef2 & vbCrLf & strRef3 & Mid$(strShp(i), 1, Len(strShp(i)) - 2) & ChrW$(34) & " )" & vbCrLf & ")"
            NewFile = FreeFile
            Open Filpath1$ & "\TempRef.ref" For Append As #NewFile
               Print #NewFile, strTemp
               Close NewFile
               booFound = True
            GoTo Label7
            
            End If
           ' End If
Label7:

    If booFound = False And strShp(i) <> vbNullString Then
    Newfile3 = FreeFile
    Open Filpath1$ & "\TempUnk.ref" For Append As #Newfile3
    Print #Newfile3, strShp(i) & vbCrLf
    Close Newfile3
    End If
    
    booFound = False

 '  End If
   
Next i



   Rem ************ Find transfers **********
   strTrans1 = "Transfer (" & vbCrLf & "     Class (" & ChrW$(34) & "Transfers" & ChrW$(34) & ")"
strTrans2 = "     Filename ("
strTrans3 = ")" & vbCrLf & "     Align (None)" & vbCrLf & "     Description ("
strTrans4 = ")" & vbCrLf & ")" & vbCrLf
   For i = 0 To numTrans - 1
  
   For ii = 1 To intShapes
   If MasterIndex(ii) = Transfer2(i) Then
   strTemp = MasterFile(ii)
   NewFile = FreeFile
   Open Filpath1$ & "\TempRef.ref" For Append As #NewFile
               Print #NewFile, strTemp
               Close NewFile
               
            GoTo Label9
            End If
            Next ii
    
            Rem ******************* No .ref
           
          
            strTemp = strTrans1 & vbCrLf & strTrans2 & Transfer2(i) & strTrans3 & Mid$(Transfer2(i), 1, Len(Transfer2(i)) - 4) & strTrans4
            
            NewFile = FreeFile
            Open Filpath1$ & "\TempRef.ref" For Append As #NewFile
               Print #NewFile, strTemp
               Close NewFile
               booFound = True
           ' Exit For
           ' End If
            
Label9:
   Next i
   Rem ************** Deal with any Hazards


 If numHaz > 0 Then

 If Not DirExists(RoutePath & "\Global\Shapes") Then
 
 MkDir (RoutePath & "\Global")
 MkDir (RoutePath & "\Global\Shapes")
 End If
 If Not DirExists(RoutePath & "\Global\textures") Then
 MkDir (RoutePath & "\Global\Textures")
 End If
 booHaz = True
If numHaz > 0 Then
For i = 0 To numHaz - 1

Call CompactCheckForS(RoutePath & "\" & HazShp2(i), booHaz)
DoEvents

If HazShp2(i) <> "deer.haz" And HazShp2(i) <> "spotter.haz" Then
 FileCopy GlobalPath & Hazard(i), RoutePath & "\Global\Shapes\" & Hazard(i)
FileCopy GlobalPath & Hazard(i) & "d", RoutePath & "\Global\Shapes\" & Hazard(i) & "d"
End If
   Rem ****************** Hazards
   strSpare = App.Path & "\tempfiles"
   NewFile2 = FreeFile
   strTerr = GlobalPath & Hazard(i)
 
   Call DoDeComp2(Hazard(i), strGlobal, strSpare)
   DoEvents
   
 MyString = ReadUniFile(strSpare & "\" & Hazard(i))


      yy = 1
 Do
 
 yy = InStr(yy, MyString, "image (")
 If yy > 0 Then
 Z = InStr(yy, MyString, "(")
 zz = InStr(Z, MyString, ")")
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Mid$(strFName, 2)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
 intAce = intAce + 1
  AceTemp(intAce) = strFName
     yy = Z
     End If
    Loop While yy > 0
            For q = 1 To intAce
                    
              strNew = AceTemp(q)
   If HazShp2(i) <> "deer.haz" And HazShp2(i) <> "spotter.haz" Then
  FileCopy MSTSPath & "\Global\Textures\" & strNew, RoutePath & "\Global\textures\" & strNew
  End If
             Next q
       ' End If
 DoEvents
 If FileExists(strSpare & "\" & strS) Then
Kill strSpare & "\" & strS
End If

NextOne:
Next i
End If


 

 For i = 0 To numHaz - 1
For ii = 0 To intShapes
tempHaz = HazShp2(i)

                If MasterIndex(ii) = tempHaz Then
              
                strTemp = MasterFile(ii)

                NewFile = FreeFile
                Open Filpath1$ & "\TempRef.ref" For Append As #NewFile
                   Print #NewFile, strTemp
                   Close NewFile
                   booFound = True
                GoTo label18
               
                End If
            Next ii
label18:
            Next i
            
 End If
 Exit Sub
Errtrap:
Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'Compact.ref file' please advise" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
   
        Case vbRetry
       
    Resume Next
        Case vbCancel
   
    Exit Sub
    End Select

 
End Sub

Private Sub GetFirstToken(strRef As String, Min As Double, MinIndex As Long, x As Long)
Dim refToken(0 To 11) As Double, Max As Double, maxIndex As Long



refToken(0) = InStr(x, strRef, "Static (")
refToken(1) = InStr(x, strRef, "Forest (")
refToken(2) = InStr(x, strRef, "Hazard (")
refToken(3) = InStr(x, strRef, "LevelCr (")
refToken(4) = InStr(x, strRef, "Transfer (")
refToken(5) = InStr(x, strRef, "Pickup (")
refToken(6) = InStr(x, strRef, "Dyntrack (")
refToken(7) = InStr(x, strRef, "Platform (")
refToken(8) = InStr(x, strRef, "Siding (")
refToken(9) = InStr(x, strRef, "Carspawner (")
refToken(10) = InStr(x, strRef, "Skip(")
refToken(11) = InStr(x, strRef, "Skip (")

Call FindMinMax2(refToken(), Min, Max, MinIndex, maxIndex)

If Min = 0 And Max > 0 Then
Min = Max
MinIndex = maxIndex
End If
End Sub

Private Sub GetNextToken(strRef As String, Min As Double, MinIndex As Long, x As Long)
Dim refToken(0 To 11) As Double, Max As Double, maxIndex As Long, i As Integer

Dim booToken As Boolean


refToken(0) = InStr(x, strRef, "Static (")
refToken(1) = InStr(x, strRef, "Forest (")
refToken(2) = InStr(x, strRef, "Hazard (")
refToken(3) = InStr(x, strRef, "LevelCr (")
refToken(4) = InStr(x, strRef, "Transfer (")
refToken(5) = InStr(x, strRef, "Pickup (")
refToken(6) = InStr(x, strRef, "Dyntrack (")
refToken(7) = InStr(x, strRef, "Platform (")
refToken(8) = InStr(x, strRef, "Siding (")
refToken(9) = InStr(x, strRef, "Carspawner (")
refToken(10) = InStr(x, strRef, "Skip(")
refToken(11) = InStr(x, strRef, "Skip (")
For i = 0 To 11
If refToken(i) > 0 Then
booToken = True
End If
Next i

Call FindMinMax2(refToken(), Min, Max, MinIndex, maxIndex)
If Min = 0 And Max > 0 Then
Min = Max
MinIndex = maxIndex
ElseIf Min = 0 And Max = 0 Then  'end of file
Min = Len(strRef)
booEndFile = True
End If
End Sub
Private Sub CompactRoute()


Dim i As Integer, j As Long, jj As Long, TertexPath As String
Dim booCar As Boolean, booFound As Boolean, strTemp As String, strOrig As String

Dim GlobalPath As String, Lo_Tilepath As String, strSpare As String



On Error GoTo Errtrap

ReDim Preserve strShp(0 To Shp_Chunk)
ReDim Preserve strGlobShp(0 To Shp_Chunk)
ReDim Preserve Ace1(0 To Shp_Chunk)
ReDim ForTex(0 To For_Chunk)
ReDim HazShp(0 To For_Chunk)
ReDim TerrTex(0 To For_Chunk)
ReDim Transfer(0 To For_Chunk)

If numShp < 1 Then
numShp = 0
End If
numHaz = 0
numFor = 0
numAce = 0
numTerr = 0
numTrans = 0
WorldPath = RoutePath & "\world"
TertexPath = RoutePath & "\Terrtex"
TilePath = RoutePath & "\Tiles"
Lo_Tilepath = RoutePath & "\Lo_Tiles"
GlobalPath = MSTSPath & "\Global"
frmUtils.Dir1(0).Path = WorldPath
frmUtils.Text1(0).Text = "*.w"
cursouind = 0
strSpare = App.Path & "\TempFiles"

booOKAll = False
Rem*******************
 SB1.Panels(2) = "Uncompressing .w files"
 DoEvents
Call DoDeCompFolder("w", WorldPath, strSpare)
DoEvents
frmUtils.Dir1(0).Path = strSpare
frmUtils.Text1(0).Text = "*.w"
cursouind = 0
DoEvents

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i
Rem copy to Spares **************************************

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
frmUtils.Label12.Caption = "Reading: " & frmUtils.File1(cursouind).List(i)
'
  If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = strSpare & "\" & frmUtils.File1(cursouind).List(i)

Call ReadWorld(fullpath$)
'

   End If
'
   Next i
Call KillSpare("*.w")
DoEvents

   Rem ********** Read the Tiles *******************
  SB1.Panels(2) = "Uncompressing .t files"
 DoEvents
   frmUtils.Dir1(0).Path = TilePath
Call DoDeCompFolder("t", TilePath, strSpare)
DoEvents
frmUtils.Dir1(0).Path = strSpare
frmUtils.Text1(0).Text = "*.t"
 cursouind = 0

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i

For i = 0 To frmUtils.File1(cursouind).ListCount - 1

frmUtils.Label12.Caption = "Reading: " & frmUtils.File1(cursouind).List(i)
frmUtils.Label12.Refresh

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = strSpare & "\" & frmUtils.File1(cursouind).List(i)
  
Call ReadTerrain(fullpath$)

   End If

   Next i
Call KillSpare("*.t")
DoEvents
 Rem *********** Read any Lo_Tiles **********
 If DirExists(Lo_Tilepath) Then
     frmUtils.Dir1(0).Path = Lo_Tilepath
Call DoDeCompFolder("t", Lo_Tilepath, strSpare)
DoEvents
frmUtils.Dir1(0).Path = strSpare
frmUtils.Text1(0).Text = "*.t"
 cursouind = 0

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i

For i = 0 To frmUtils.File1(cursouind).ListCount - 1

frmUtils.Label12.Caption = "Reading: " & frmUtils.File1(cursouind).List(i)
frmUtils.Label12.Refresh

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = strSpare & "\" & frmUtils.File1(cursouind).List(i)
  
Call ReadTerrain(fullpath$)

   End If

   Next i
Call KillSpare("*.t")
 End If

   Rem *********************************************
   
   QSort3 Cars(), 0, intCars - 1
   DoEvents
   RemD2 Cars(), Cars2()
   DoEvents
   For j = 0 To intCars - 1
   Cars(j) = vbNullString
   Next j
   intCars = UBound(Cars2)
   
   
   
    
    ReDim Preserve strShp(0 To numShp - 1)
  
    QSort3 strShp(), 0, numShp - 1
    DoEvents
    
    RemD2 strShp(), strShapes()
    DoEvents
  
    For j = 0 To UBound(strShp)
    strShp(j) = vbNullString
    Next j
   
    ReDim Preserve strGlobShp(0 To numGlobShp - 1)
 
    QSort3 strGlobShp(), 0, numGlobShp - 1
    DoEvents
    
    RemD2 strGlobShp(), strGlobShp2()
    DoEvents
  
    For j = 0 To UBound(strGlobShp)
    strGlobShp(j) = vbNullString
    Next j
    
    ReDim Preserve ForTex(0 To numFor - 1)
    QSort3 ForTex(), 0, numFor - 1
    DoEvents
    RemD2 ForTex(), ForTex2()
    DoEvents
    For j = 0 To numFor - 1
    ForTex(j) = vbNullString
    Next j
   


    If numHaz > 0 Then
    ReDim Preserve HazShp(0 To numHaz - 1)
    QSort3 HazShp(), 0, numHaz - 1
    DoEvents
    RemD2 HazShp(), HazShp2()
    DoEvents
  
    numHaz = UBound(HazShp2)
    End If
    

    If numTrans > 0 Then
    ReDim Preserve Transfer(0 To numTrans - 1)
    QSort3 Transfer(), 0, numTrans - 1
    DoEvents
   
    RemD2 Transfer(), Transfer2()
    DoEvents
    numTrans = UBound(Transfer2)
    End If

    
    ReDim Preserve TerrTex(0 To numTerr - 1)
    QSort3 TerrTex(), 0, numTerr - 1
    DoEvents
    RemD2 TerrTex(), TerrTex2()
    DoEvents
    For j = 0 To numTerr - 1
    TerrTex(j) = vbNullString
    Next j
    DoEvents
    
    QSort3 Ace1(), 0, numAce - 1
    DoEvents
    
    RemD2 Ace1(), Ace2()
    DoEvents
    
    For j = 0 To UBound(Ace2)
    Ace1(j) = Ace2(j)
    Ace2(j) = vbNullString
    ESD(j) = 0
    Next j
    DoEvents
    numAce = UBound(Ace2) + 1
    numShp = UBound(strShapes) + 1
    numGlobShp = UBound(strGlobShp2) + 1
    numFor = UBound(ForTex2) + 1
    numTerr = UBound(TerrTex2) + 1
   
    If numTrans > 0 Then
   
    For j = 0 To numTrans
   
    Call CompactTransfer(Transfer2(j))
    Next j
    End If
    
    If numMisc > 0 Then
    For j = 0 To numMisc - 1
    booCar = True
    Call CompactAllAce(MiscAce(j), "", booCar)
    Next j
    End If
   If intCars > 0 Then

    For j = 0 To intCars
   SB1.Panels(2) = Cars2(j)
    If FileExists(RoutePath & "\shapes\" & Cars2(j)) Then
      FileCopy RoutePath & "\shapes\" & Cars2(j), strSpare & "\" & Cars2(j)
      End If
    Next j
    strOrig = RoutePath & "\shapes\"
  
    DoEvents
    End If
 If numShp > 0 Then
    For j = 0 To numShp - 1
   SB1.Panels(2) = strShapes(j)
    If FileExists(RoutePath & "\shapes\" & strShapes(j)) Then
      FileCopy RoutePath & "\shapes\" & strShapes(j), strSpare & "\" & strShapes(j)
      End If
    Next j
    SB1.Panels(2) = "Uncompressing .s files"
 DoEvents
 End If
    If numGlobShp > 0 Then
    For j = 0 To numGlobShp - 1
   SB1.Panels(2) = strGlobShp2(j)
    If FileExists(GlobalPath & "\shapes\" & strGlobShp2(j)) Then
      FileCopy GlobalPath & "\shapes\" & strGlobShp2(j), strSpare & "\" & strGlobShp2(j)
      End If
    Next j
    SB1.Panels(2) = "Uncompressing .s files"
 DoEvents
 End If
 strOrig = RoutePath & "\shapes\"
    Call DoDeCompFolder("s", strSpare, strSpare)
    DoEvents

For j = 0 To intCars
If FileExists(strSpare & "\" & Cars2(j)) And Cars2(j) <> vbNullString Then
booCar = True
    Call CompactAceSeasons(strSpare & "\" & Cars2(j), booCar)
End If
Next j
For j = 0 To intCars
    For jj = 0 To numShp - 1
    If Cars2(j) = strShapes(jj) Then
    Cars2(j) = vbNullString
    Exit For
    End If
    Next jj
Next j
'End If
DoEvents
QSort3 Cars2(), 0, intCars
   DoEvents
   RemD2 Cars2(), Cars()
   DoEvents
   For j = 0 To intCars
   Cars2(j) = vbNullString
   Next j
   intCars = UBound(Cars)

Rem *****************************************************
 If numShp > 0 Then

 strOrig = RoutePath & "\shapes\"
   
    DoEvents

    For j = 0 To numShp - 1
    booCar = False
   If FileExists(strSpare & "\" & strShapes(j)) And strShapes(j) <> vbNullString Then
    Call CompactAceSeasons(strSpare & "\" & strShapes(j), booCar)
    End If
    Next j

    ReDim Preserve strShapes(0 To (numShp) + intCars)

    For j = 0 To intCars
    strShapes(numShp) = Cars(j)
  
    
    numShp = numShp + 1
    Next j
    DoEvents
    Call QSort3(strShapes(), 0, numShp)
    
    DoEvents
   
    RemD2 strShapes(), strShp()
    DoEvents
  
    numShp = UBound(strShp) + 1
    End If
    Rem ***********************************
    If numGlobShp > 0 Then


    For j = 0 To numGlobShp - 1
    booCar = False
   If FileExists(strSpare & "\" & strGlobShp2(j)) And strGlobShp2(j) <> vbNullString Then
    Call CompactAceSeasons(strSpare & "\" & strGlobShp2(j), booCar)
    End If
    Next j
    
   End If
   
    Call QSort3(strGlobShp2(), 0, numGlobShp - 1)
    
    DoEvents
   
    RemD2 strGlobShp2(), strGlobShp()
    DoEvents
    numGlobShp = UBound(strGlobShp) + 1
    

   For j = 0 To numAce
   strTemp = Trim$(ESD(j))
   If Len(strTemp) = 1 Then
   strTemp = "00" & strTemp
   ElseIf Len(strTemp) = 2 Then
   strTemp = "0" & strTemp
   End If
   strTemp = "|" & strTemp
   
   Ace1(j) = Ace1(j) & strTemp
   Next j
   
    ReDim Preserve Ace1(0 To numAce + numFor + numTrans + 10)
    numAce = numAce + 1
    For j = 0 To numFor - 1
    Ace1(numAce) = ForTex2(j) & "|252"
    numAce = numAce + 1
    Next j
    
    For j = 0 To numTrans
    Ace1(numAce) = Transfer2(j) & "|001"
    numAce = numAce + 1
    Next j
    For j = 0 To numAce - 1
    If Ace1(j) = "acleantrack1.ace" & "|001" Then
    booFound = True
    Exit For
    End If
    Next j
    If booFound = False Then
    Ace1(numAce) = "acleantrack1.ace" & "|001"
    ESD(numAce) = "1"
    numAce = numAce + 1
    End If
    booFound = False
    For j = 0 To numAce - 1
    If Ace1(j) = "acleantrack2.ace" & "|001" Then
    booFound = True
    Exit For
    End If
    Next j
    If booFound = False Then
    Ace1(numAce) = "acleantrack2.ace" & "|001"
    ESD(numAce) = "1"
    numAce = numAce + 1
    End If
    booFound = False
  
    Call QSort3(Ace1(), 0, numAce - 1)
    DoEvents
    
    RemD2 Ace1(), Ace2()
    
    numAce = UBound(Ace2) + 1

    

Exit Sub
Errtrap:
If Err = 9 Then
Resume Next
End If

    Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'Compact route' please advise" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
      ' Resume Next
    Exit Sub
    End Select
        


End Sub






Private Sub GetConsists2()
Dim ConsistPath As String, i As Integer

ConsistPath = Trainspath & "Consists"
cursouind = 0
Drive1(cursouind).Drive = Left$(ConsistPath, 2)
Dir1(cursouind).Path = ConsistPath
Text1(cursouind).Text = "*.con"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
lngCon = File1(cursouind).ListCount
ReDim Consists(0 To lngCon - 1)
ReDim ConIntName(0 To lngCon - 1)
ReDim ConIntWagName(0 To lngCon - 1)
DoEvents
Label7(4).Caption = lngCon
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
   Consists(i) = File1(cursouind).List(i)

   End If
   
   Next i

End Sub


Private Sub ReadAce(strAce As String, intESD As Integer, sFile As String)
Dim booTop As Boolean


booTop = False
Select Case intESD
   Case 0
   If Not FileExists(TexturePath & "\" & strAce) Then
   booTop = True
   Call LookForACESD(strAce, sFile, booTop)
  
   End If
   Case 2
  If Not FileExists(TexturePath & "\" & strAce) Then
   booTop = True
   Call LookForACESD(strAce, sFile, booTop)
   End If
   If Not FileExists(TexSnowPath & "\" & strAce) Then
   
   Call LookForACESD("snow\" & strAce, sFile, booTop)
   End If
   
   Case 1
   
   If Not FileExists(TexturePath & "\" & strAce) Then
   booTop = True
   Call LookForACESD(strAce, sFile, booTop)
   
   End If
   If Not FileExists(TexSnowPath & "\" & strAce) Then
   Call LookForACESD("snow\" & strAce, sFile, booTop)
  
   End If
   Case 252
    If Not FileExists(TexturePath & "\" & strAce) Then
    booTop = True
   Call LookForACESD(strAce, sFile, booTop)
    End If

   If Not FileExists(TexAutPath & "\" & strAce) Then
   Call LookForACESD("autumn\" & strAce, sFile, booTop)
  
   End If
   If Not FileExists(TexAutSnowPath & "\" & strAce) Then
    Call LookForACESD("autumnsnow\" & strAce, sFile, booTop)
   
   End If
   If Not FileExists(TexSprPath & "\" & strAce) Then
   Call LookForACESD("spring\" & strAce, sFile, booTop)
  
   End If
   If Not FileExists(TexSprSnowPath & "\" & strAce) Then
   Call LookForACESD("springsnow\" & strAce, sFile, booTop)
  
   End If
   If Not FileExists(TexWinPath & "\" & strAce) Then
   Call LookForACESD("winter\" & strAce, sFile, booTop)
  
   End If
   If Not FileExists(TexWinSnowPath & "\" & strAce) Then
   Call LookForACESD("wintersnow\" & strAce, sFile, booTop)
  
   End If
   Rem *********************** ESD 256 ********************************
   Case 256
   
   If Not FileExists(TexturePath & "\" & strAce) Then
   booTop = True
   Call LookForACESD(strAce, sFile, booTop)
   
   End If
   
   
   
   If Not FileExists(TexNightPath & "\" & strAce) Then
   Call LookForACESD("night\" & strAce, sFile, booTop)
   
   End If
   strReport = strReport & " " & sFile & "d has an ESD of 256 - This has a bug in it and will cause an error if the route is run in the snow" & vbCrLf
   strReport = strReport & "It is recommended you include a snow texture and change the ESD to 257 to avoid this problem" & vbCrLf
   Rem ************************************************************************
   Case 257
   If Not FileExists(TexturePath & "\" & strAce) Then
   booTop = True
   Call LookForACESD(strAce, sFile, booTop)
  
   End If
   If Not FileExists(TexSnowPath & "\" & strAce) Then
   Call LookForACESD("snow\" & strAce, sFile, booTop)
   
   End If
   If Not FileExists(TexNightPath & "\" & strAce) Then
   Call LookForACESD("night\" & strAce, sFile, booTop)
   
   End If
   Case Else
  
   End Select
End Sub

Private Sub ReadSD(fullpath$, intESD As Integer, intPath As Integer)
Dim CurrentSD As String, flagway As Integer
Dim x As Integer, xx As Integer, xy As Integer, strESD As String
Dim j As Integer, strTemp As String, booBadESD As Boolean
Dim booMiss As Boolean, strCorrShape As String, strSearch As String
Dim ShapeOK As Boolean, booGotEsd As Boolean, MyString As String, B$

On Error GoTo Errtrap


MousePointer = 11


strSearch = "esd_a"
  
  CurrentSD = fullpath$ & "d"
  x = InStrRev(fullpath$, "\")
  strCorrShape = Mid$(fullpath$, x + 1)
   
  Rem *************
  
  
  If CurrentSD = ShapePath & "\jp1multicarpark.sd" Then
 
  flagway = 0
  Call FixCarPark(ShapePath & "\jp1multicarpark.sd", flagway)
  flagway = 1
  Call FixCarPark(ShapePath & "\jp1multicarpark.sd", flagway)
  
  End If
  
 
  If Not FileExists(CurrentSD) Then
  
  j = InStrRev(CurrentSD, "\")
  B$ = Mid$(CurrentSD, j + 1)
  For j = 0 To UBound(Cars)
  If B$ = Cars(j) & "d" Then GoTo TryAgain3
  Next j

    Call LookForSD(CurrentSD, booMiss)
  
  If booMiss = False Then
 Rem
  Else
  booMiss = False
  End If
  Else
  If FileExists(CurrentSD) Then
  MyString = ReadUniFile(CurrentSD)
  MyString = LCase(MyString)

x = 1
Do While x > 0
  x = InStr(x, MyString, strSearch, 0)
      If x = 0 Then GoTo TryAgain
      booGotEsd = True
      If Mid$(MyString, x, 23) <> "esd_alternative_texture" Then
      booBadESD = True
      End If
      xx = InStr(x, MyString, "(", 0)
      xy = InStr(xx, MyString, ")", 0)
      strESD = Mid$(MyString, xx + 1, xy - xx - 1)
      strESD = Trim$(strESD)
      intESD = Val(strESD)
      x = x + 1
      Exit Do
      
TryAgain:
      Loop
    '  Close #NewFile
      If booGotEsd = True Then
     
      booGotEsd = False
      Else
       strReport = strReport & strCorrShape & "d did not have a ESD_Alternative_Texture entry" & vbCrLf
      GoTo EndPart
      End If
      If booBadESD = True Then
      booBadESD = False
      Call ConvertESD(CurrentSD, 0, strESD)
    DoEvents
    Call ConvertESD(CurrentSD, 1, strESD)
      End If
      If intPath = 1 Then
        If strESD <> "0" And strESD <> "1" And strESD <> "2" And strESD <> "252" And strESD <> "256" And strESD <> "257" Then
                strReport = strReport & strCorrShape & " " & Lang(556) & strESD & vbCrLf
        End If
        If strESD = "2" Then
        strReport = strReport & strCorrShape & " " & Lang(556) & strESD & " This value is only valid for Shapes in the Global\Shapes folder - It has been changed to '1'" & vbCrLf
        strESD = "1"
    Call ConvertESD(CurrentSD, 0, strESD)
    DoEvents
    Call ConvertESD(CurrentSD, 1, strESD)
    strReport = strReport & CurrentSD & Lang(559) & strESD & vbCrLf
        End If
        End If
       ' If strESD = "0" Then strESD = "1"
        If strESD = "256" Then
     
        If intResponse2 = 0 Or intResponse2 = 1 Or intResponse2 = 3 Then
        strTemp = CurrentSD & Lang(557)
   strTemp = strTemp & Lang(558)
   frmDialog2.OKButton(0).Caption = "Make 0"
   frmDialog2.OKButton(2).Caption = "All 0"
   frmDialog2.OKButton(1).Caption = "Make 257"
   frmDialog2.OKButton(3).Caption = "All 257"
   frmDialog2.CancelButton.Caption = "Leave"
   frmDialog2.Label1.Caption = strTemp
   frmDialog2.Caption = "ESD 256 Error"
   frmDialog2.Show 1
   
     DoEvents
     
   End If
   Select Case intResponse2
   Case 1, 2
    strESD = "0"
    Call ConvertESD(CurrentSD, 0, strESD)
    DoEvents
    Call ConvertESD(CurrentSD, 1, strESD)
    strReport = strReport & CurrentSD & Lang(559) & strESD & vbCrLf
    Case 3, 4
    strESD = "257"
    Call ConvertESD(CurrentSD, 0, strESD)
    DoEvents
    Call ConvertESD(CurrentSD, 1, strESD)
    strReport = strReport & CurrentSD & Lang(559) & strESD & vbCrLf
    End Select
     End If
        
     
EndPart:

      ShapeOK = False

     End If
      
  
  End If
 
TryAgain3:

Exit Sub

Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'ReadSD' while reading " _
            & vbCrLf & fullpath$ _
            , vbExclamation, App.Title)

End Sub


Private Sub ReadTerrain(strTerr As String)
 Dim x As Integer, strS As String, MyString As String, strSpare As String, yy As Long
 Dim Z As Long, zz As Long, strFName As String
    
    On Error GoTo Errtrap
    
   x = InStrRev(strTerr, "\")
   strS = Mid$(strTerr, x + 1)

   SB1.Panels(2).Text = strS
   
      strSpare = App.Path & "\TempFiles"

 MyString = ReadUniFile(strTerr)
MyString = LCase(MyString)

      yy = 1
 Do

 yy = InStr(yy, MyString, "terrain_texslot (", 0)
 If yy > 0 Then
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ".ace", 0)
 strFName = Mid$(MyString, Z + 1, (zz + 4) - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Mid$(strFName, 2)
 End If
  If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Mid$(strFName, 2)
 End If

 
                TerrTex(numTerr) = strFName
                numTerr = numTerr + 1
                If numTerr > UBound(TerrTex) Then
                    ReDim Preserve TerrTex(0 To numTerr + For_Chunk)
                    End If
            
                    yy = zz
                    End If

    
    Loop While yy > 0
   
 DoEvents

Exit Sub

Errtrap:
Call MsgBox("An error occurred while attempting to read " & strTerr, vbExclamation, App.Title)

Resume Next
End Sub


Private Sub CompactTrack()
Dim i As Integer, j As Long, GlobalPath As String, FoundIt As Boolean
Dim strTemp As String, ii As Integer, strSpare As String

On Error GoTo Errtrap

MousePointer = 11
ReDim strGlobShp(0 To Shp_Chunk)
numGlobShp = 1

GlobalPath = MSTSPath & "\Global\Shapes\"
GlobalSparePath = MSTSPath & "\Global\SpareTrack\"
If Not DirExists(GlobalSparePath) Then
MkDir GlobalSparePath
End If
GlobalSparePath = MSTSPath & "\Global\SpareTrack\"
strSpare = App.Path & "\tempfiles"

For ii = 0 To NumRoutes - 1
WorldPath = AllRoutes(ii) & "\world"

If Not DirExists(WorldPath) Then GoTo GetAnother
SB1.Panels(2).Text = "Uncompressing .w files"
Call DoDeCompFolder("w", WorldPath, strSpare)
DoEvents
frmUtils.Dir1(0).Path = strSpare
frmUtils.Text1(0).Text = "*.w"
cursouind = 0
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
SB1.Panels(2).Text = frmUtils.File1(cursouind).List(i)
DoEvents

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = strSpare & "\" & frmUtils.File1(cursouind).List(i)


Call ReadWorld2(fullpath$)

   End If

   Next i
   Call KillSpare("*.w")
GetAnother:
Next ii


    
    ReDim Preserve strGlobShp(0 To numGlobShp - 1)
   SB1.Panels(2).Text = "Sorting Shapes"
    QSort3 strGlobShp(), 0, numGlobShp - 1
    DoEvents
    SB1.Panels(2).Text = "Removing Dupes from List"
    RemD2 strGlobShp(), strGlobShp2()
    DoEvents
    For j = 0 To numGlobShp - 1
    strGlobShp(j) = vbNullString
    Next j
   
   numGlobShp = UBound(strGlobShp2)
   
 
   Dir1(0).Path = GlobalPath
Text1(0).Text = "*.s"

cursouind = 0

DoEvents

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
strTemp = File1(cursouind).List(i)

If Left$(strTemp, 2) = "im" Or Left$(strTemp, 4) = "mark" Or Left$(strTemp, 4) = "mile" Then
FoundIt = True
GoTo CarryON
End If
If Left$(strTemp, 3) = "us2" Or Left$(strTemp, 4) = "yard" Or strTemp = "cow.s" Or strTemp = "us1deer.s" Then
FoundIt = True
GoTo CarryON
End If
If strTemp = "platform.s" Or strTemp = "workman.s" Or strTemp = "crosshair.s" Or strTemp = "Csend.s" Or strTemp = "csstart.s" Then
FoundIt = True
GoTo CarryON
End If
If Left$(strTemp, 6) = "female" Or Left$(strTemp, 6) = "forest" Or Left$(strTemp, 3) = "Tml" Then
FoundIt = True
GoTo CarryON
End If



For j = 0 To numGlobShp
If File1(cursouind).List(i) = strGlobShp2(j) Then
FoundIt = True
Exit For
End If
Next j
If FoundIt = False Then

SB1.Panels(2).Text = "Moving " & strTemp
    If FileExists(GlobalPath & strTemp) Then
    
    FileCopy GlobalPath & strTemp, GlobalSparePath & strTemp
    DoEvents
    Kill GlobalPath & strTemp
    DoEvents
    End If
    If FileExists(GlobalPath & strTemp & "d") Then
    FileCopy GlobalPath & strTemp & "d", GlobalSparePath & strTemp & "d"
    DoEvents
    Kill GlobalPath & strTemp & "d"
    DoEvents
    End If
    strTemp = Left$(strTemp, Len(strTemp) - 1)
    strTemp = strTemp & "thm"
    If FileExists(GlobalPath & strTemp) Then
    Kill GlobalPath & strTemp
    DoEvents
    End If
    
End If
CarryON:
FoundIt = False
Next i
MousePointer = 0

 strReport = vbNullString
 Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'CompactTrack' while" _
            & vbCrLf & "reading " & fullpath$ _
            , vbExclamation, App.Title)
            

End Sub

Private Sub MiniTrack()
Dim i As Integer, j As Long, GlobalPath As String, FoundIt As Boolean
Dim strTemp As String, strSpare As String

On Error GoTo Errtrap

MousePointer = 11
ReDim Preserve strGlobShp(0 To Shp_Chunk)
numGlobShp = 1
GlobalPath = MSTSPath & "\Global\Shapes\"
strSpare = App.Path & "\Tempfiles"
WorldPath = RoutePath & "\world"
SB1.Panels(2).Text = "Uncompressing .w files"
Call DoDeCompFolder("w", WorldPath, strSpare)

cursouind = 0

booOKAll = False

frmUtils.Dir1(0).Path = strSpare
frmUtils.Text1(0).Text = "*.w"
cursouind = 0
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
SB1.Panels(2).Text = frmUtils.File1(cursouind).List(i)
DoEvents

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = strSpare & "\" & frmUtils.File1(cursouind).List(i)


Call ReadWorld2(fullpath$)

   End If

   Next i

 
    
    ReDim Preserve strGlobShp(0 To numGlobShp - 1)
   
    QSort3 strGlobShp(), 0, numGlobShp - 1
    DoEvents
    
    RemD2 strGlobShp(), strGlobShp2()
    DoEvents
    For j = 0 To numGlobShp - 1
    strGlobShp(j) = vbNullString
    Next j
   
   numGlobShp = UBound(strGlobShp2)
   
   
   Dir1(0).Path = GlobalPath
Text1(0).Text = "*.s"

cursouind = 0

DoEvents

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
strTemp = File1(cursouind).List(i)

If Left$(strTemp, 2) = "im" Or Left$(strTemp, 4) = "mark" Or Left$(strTemp, 4) = "mile" Then
FoundIt = True
GoTo CarryON
End If
If Left$(strTemp, 3) = "us2" Or Left$(strTemp, 4) = "yard" Or strTemp = "cow.s" Or strTemp = "us1deer.s" Then
FoundIt = True
GoTo CarryON
End If
If strTemp = "platform.s" Or strTemp = "workman.s" Or strTemp = "crosshair.s" Or strTemp = "Csend.s" Or strTemp = "csstart.s" Then
FoundIt = True
GoTo CarryON
End If
If Left$(strTemp, 6) = "female" Or Left$(strTemp, 6) = "forest" Or Left$(strTemp, 3) = "Tml" Then
FoundIt = True
GoTo CarryON
End If



For j = 0 To numGlobShp
If File1(cursouind).List(i) = strGlobShp2(j) Then
FoundIt = True
Exit For
End If
Next j
If FoundIt = False Then
    If FileExists(GlobalPath & strTemp) Then
    Kill GlobalPath & strTemp
    DoEvents
    End If
    If FileExists(GlobalPath & strTemp & "d") Then
    Kill GlobalPath & strTemp & "d"
    DoEvents
    End If
    strTemp = Left$(strTemp, Len(strTemp) - 1)
    strTemp = strTemp & "thm"
    If FileExists(GlobalPath & strTemp) Then
    Kill GlobalPath & strTemp
    DoEvents
    End If
    
End If
CarryON:
FoundIt = False
Next i
MousePointer = 0

 strReport = vbNullString
 Exit Sub
Errtrap:

End Sub


Private Sub ReadTrack()
Dim i As Integer, j As Long, strSpare As String


On Error GoTo Errtrap

MousePointer = 11
ReDim Preserve strGlobShp(0 To Shp_Chunk)
numGlobShp = 1


WorldPath = RoutePath & "\world"
strSpare = App.Path & "\Tempfiles"


booOKAll = False
SB1.Panels(2).Text = "Uncompressing .w files"
Call DoDeCompFolder("w", WorldPath, strSpare)


DoEvents

frmUtils.Dir1(0).Path = strSpare
frmUtils.Text1(0).Text = "*.w"
cursouind = 0
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
SB1.Panels(2).Text = frmUtils.File1(cursouind).List(i)
DoEvents

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = strSpare & "\" & frmUtils.File1(cursouind).List(i)


Call ReadWorld2(fullpath$)

   End If

   Next i

 
    
    ReDim Preserve strGlobShp(0 To numGlobShp - 1)
   
    QSort3 strGlobShp(), 0, numGlobShp - 1
    DoEvents
    RemD2 strGlobShp(), strGlobShp2()
    DoEvents
    For j = 0 To numGlobShp - 1
    strGlobShp(j) = vbNullString
    Next j
   numGlobShp = UBound(strGlobShp2)
   
For j = 0 To numGlobShp
strReport = strReport & strGlobShp2(j) & vbCrLf
Next
MousePointer = 0
frmReport.Rich1.Text = strReport
 
     frmReport.Show 1
 strReport = vbNullString
 Exit Sub
Errtrap:

End Sub


Private Sub WorldCount()
Dim i As Integer, j As Long
Dim booCar As Boolean, booFound As Boolean, strTemp As String, x As Integer

Dim GlobalPath As String, Lo_Tilepath As String, wt_ew As Integer
Dim wt_ns As Integer, strTName As String, strEW As String, strNS As String
Dim strOrigFile As String, jj As Integer, strSpare As String, TertexPath As String


On Error GoTo Errtrap


WorldPath = RoutePath & "\world"
TertexPath = RoutePath & "\Terrtex"
TilePath = RoutePath & "\Tiles"
Lo_Tilepath = RoutePath & "\Lo_Tiles"
GlobalPath = MSTSPath & "\Global"
strSpare = App.Path & "\TempFiles"


cursouind = 0
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
   Exit For
   End If
  Next i
    
   strOrigFile = File1(cursouind).List(i)
   If Right$(strOrigFile, 2) <> ".w" Then
   Call MsgBox("You do not appear to have selected a .W file ?", vbCritical, App.Title)
   
   Exit Sub
   End If
   Call Worldtiles(strOrigFile)
   
   For jj = 1 To 9
   
   ReDim strShp(0 To Shp_Chunk)
ReDim strGlobShp(0 To Shp_Chunk)
ReDim Ace1(0 To Shp_Chunk)
ReDim ForTex(0 To For_Chunk)
ReDim HazShp(0 To For_Chunk)
ReDim TerrTex(0 To For_Chunk)
ReDim Transfer(0 To For_Chunk)

numShp = 0
numGlobShp = 0
numHaz = 0
numFor = 0
numAce = 0
numTerr = 0
numTrans = 0

totShp = 0
totGlobshp = 0
totHaz = 0
totFor = 0
totAce = 0
totTerr = 0
totTrans = 0
   strOrigFile = wTile(jj)
   fullpath$ = File1(cursouind).Path & "\" & wTile(jj)
   If Not FileExists(fullpath$) Then
   strReport = strReport & vbCrLf & "Tile #" & Str(jj) & "  " & wTile(jj) & " - THIS TILE IS NOT IN THE ROUTE" & vbCrLf
strReport = strReport & "---------------------------" & vbCrLf
   GoTo GetAnother
   End If
   x = InStr(3, strOrigFile, "-")
   If x = 0 Then
   x = InStr(3, strOrigFile, "+")
   End If
   strEW = Mid$(strOrigFile, 2, x - 2)
   strNS = Mid$(strOrigFile, x)
   strNS = Left$(strNS, Len(strNS) - 2)
   wt_ew = Val(strEW)
   wt_ns = Val(strNS)
   Call TileName2(wt_ew, wt_ns, strTName)
   strTName = strTName & ".t"

booOKAll = False
Call DoDeComp2(wTile(jj), File1(cursouind).Path, strSpare)
DoEvents
 
Call ReadWorld(strSpare & "\" & wTile(jj))
DoEvents
If Not FileExists(TilePath & "\" & strTName) Then
Call MsgBox(strTName & " is missing, your World .w and .t folders do not match." _
            & vbCrLf & "Suggest you run the Route Integrity check." _
            , vbExclamation, App.Title)
GoTo GetAnother
End If
Call DoDeComp2(strTName, TilePath, strSpare)
DoEvents
Call ReadTerrain(strSpare & "\" & strTName)
DoEvents

   totShp = totShp + numShp
   totGlobshp = totGlobshp + numGlobShp
   totFor = totFor + numFor
   totHaz = totHaz + numHaz
   totTrans = totTrans + numTrans
    If numShp > 0 Then
    ReDim Preserve strShp(0 To numShp - 1)

    QSort3 strShp(), 0, numShp - 1
    DoEvents
    
    RemD2 strShp(), strShapes()
    DoEvents
    numShp = UBound(strShapes) + 1
   
    For j = 0 To UBound(strShp)
    strShp(j) = vbNullString
    Next j
    
    End If
   
   If numGlobShp > 0 Then
    ReDim Preserve strGlobShp(0 To numGlobShp - 1)
  
    QSort3 strGlobShp(), 0, numGlobShp - 1
    DoEvents
    
    RemD2 strGlobShp(), strGlobShp2()
    DoEvents
     numGlobShp = UBound(strGlobShp2) + 1
    For j = 0 To UBound(strGlobShp)
    strGlobShp(j) = vbNullString
    Next j
     
    End If
    
    If numFor > 0 Then
    ReDim Preserve ForTex(0 To numFor - 1)
    QSort3 ForTex(), 0, numFor - 1
    DoEvents
    RemD2 ForTex(), ForTex2()
    DoEvents
    numFor = UBound(ForTex2) + 1
    For j = 0 To numFor - 1
    ForTex(j) = vbNullString
    Next j
    
   End If
    
    

    If numHaz > 0 Then
    ReDim Preserve HazShp(0 To numHaz - 1)
    QSort3 HazShp(), 0, numHaz - 1
    DoEvents
    RemD2 HazShp(), HazShp2()
    DoEvents
  
    numHaz = UBound(HazShp2)
    End If
    
    
    If numTrans > 0 Then
    ReDim Preserve Transfer(0 To numTrans - 1)
    QSort3 Transfer(), 0, numTrans - 1
    DoEvents
   
    RemD2 Transfer(), Transfer2()
    DoEvents
    numTrans = UBound(Transfer2)
    End If
 
    totTerr = numTerr
    If numTerr > 0 Then
    ReDim Preserve TerrTex(0 To numTerr - 1)
    QSort3 TerrTex(), 0, numTerr - 1
    DoEvents
    RemD2 TerrTex(), TerrTex2()
    DoEvents
    numTerr = UBound(TerrTex2) + 1
    For j = 0 To numTerr - 1
    TerrTex(j) = vbNullString
    Next j
    
    DoEvents
   End If
   If numAce > 0 Then
   
    QSort3 Ace1(), 0, numAce - 1
    DoEvents
    RemD2 Ace1(), Ace2()
    DoEvents
    numAce = UBound(Ace2) + 1
    
    For j = 0 To numAce - 1
    Ace1(j) = Ace2(j)
    Ace2(j) = vbNullString
    ESD(j) = 0
    Next j
    DoEvents
    End If
    Rem ***********************************************
    
    
  
    
    
    Rem ***********************************************
    If numTrans > 0 Then
    For j = 0 To numTrans - 1

    Call CompactTransfer(Transfer2(j))
    Next j
    End If

    If numMisc > 0 Then
    For j = 0 To numMisc - 1
    booCar = True
    Call CompactAllAce(MiscAce(j), "", booCar)
    Next j
    End If
    If intCars > 0 Then
For j = 0 To intCars - 1
If FileExists(RoutePath & "\shapes\" & Cars(j)) And Cars(j) <> vbNullString Then
Call DoDeComp3(Cars(j), RoutePath & "\shapes\", strSpare)
'Call DoDeComp2(Cars(j), RoutePath & "\shapes\", strSpare)
booCar = True
    Call CompactAceSeasons(strSpare & "\" & Cars(j), booCar)
End If
Next j
End If
If numShp > 0 Then
    For j = 0 To numShp - 1
    booCar = False
   If FileExists(RoutePath & "\shapes\" & strShapes(j)) And strShapes(j) <> vbNullString Then
   'Call DoDeComp2(strShapes(j), RoutePath & "\shapes\", strSpare)
   Call DoDeComp3(strShapes(j), RoutePath & "\shapes\", strSpare)
   ' Call CompactAceSeasons(strSpare & "\" & strShapes(j), booCar)
    End If
    Next j
End If
If numShp > 0 Or intCars > 0 Then
    ReDim Preserve strShapes(0 To (numShp - 1) + intCars)
    End If
If intCars > 0 Then
    For j = 0 To intCars - 1
    strShapes(numShp) = Cars(j)
    numShp = numShp + 1
    Next j
    DoEvents
    End If
    Call QSort3(strShapes(), 0, numShp - 1)

    DoEvents

    RemD2 strShapes(), strShp()
    DoEvents
    numShp = UBound(strShp) + 1

    For j = 0 To numGlobShp - 1
    booCar = False

   If FileExists(GlobalPath & "\shapes\" & strGlobShp2(j)) And strGlobShp2(j) <> vbNullString Then
   Call DoDeComp3(strGlobShp2(j), GlobalPath & "\shapes\", strSpare)
  ' Call DoDeComp2(strGlobShp2(j), GlobalPath & "\shapes\", strSpare)
    Call CompactAceSeasons(strSpare & "\" & strGlobShp2(j), booCar)
    End If
    Next j
    Call QSort3(strGlobShp2(), 0, numGlobShp - 1)

    DoEvents

    RemD2 strGlobShp2(), strGlobShp()
    DoEvents
    numGlobShp = UBound(strGlobShp) + 1

totAce = numAce
   For j = 0 To numAce - 1
   strTemp = Trim$(ESD(j))
   If Len(strTemp) = 1 Then
   strTemp = "00" & strTemp
   ElseIf Len(strTemp) = 2 Then
   strTemp = "0" & strTemp
   End If
   strTemp = "|" & strTemp

   Ace1(j) = Ace1(j) & strTemp
   Next j

    ReDim Preserve Ace1(0 To numAce + numFor + numTrans + 10)
    numAce = numAce + 1
    For j = 0 To numFor - 1
    Ace1(numAce) = ForTex2(j) & "|252"
    numAce = numAce + 1
    Next j
    For j = 0 To numTrans - 1
    Ace1(numAce) = Transfer2(j) & "|001"
    numAce = numAce + 1
    Next j
    For j = 0 To numAce - 1
    If Ace1(j) = "acleantrack1.ace" & "|001" Then
    booFound = True
    Exit For
    End If
    Next j
    If booFound = False Then
    Ace1(numAce) = "acleantrack1.ace" & "|001"
    ESD(numAce) = "1"
    numAce = numAce + 1
    End If
    booFound = False
    For j = 0 To numAce - 1
    If Ace1(j) = "acleantrack2.ace" & "|001" Then
    booFound = True
    Exit For
    End If
    Next j
    If booFound = False Then
    Ace1(numAce) = "acleantrack2.ace" & "|001"
    ESD(numAce) = "1"
    numAce = numAce + 1
    End If
    booFound = False
   
  totAce = numAce
    Call QSort3(Ace1(), 0, numAce - 1)
    DoEvents

    RemD2 Ace1(), Ace2()

    numAce = UBound(Ace2) + 1
   
Call WorldACESMS(wTile(jj), jj)
GetAnother:
Next jj
strReport = strReport & vbCrLf & "Grand Total  Shapes = " & Str(gtotShp) & vbCrLf & vbCrLf
strReport = strReport & vbCrLf & "Grand Total  Track Shapes = " & Str(gtotGlobShp) & vbCrLf & vbCrLf
strReport = strReport & vbCrLf & "Grand Total  Textures = " & Str(gtotAce) & vbCrLf & vbCrLf
strReport = strReport & vbCrLf & "Grand Total  Forest Textures = " & Str(gtotFor) & vbCrLf & vbCrLf
strReport = strReport & vbCrLf & "Grand Total  Terrain Textures = " & Str(gtotTerr) & vbCrLf & vbCrLf
strReport = strReport & vbCrLf & "Grand Total  Hazards = " & Str(gtotHaz) & vbCrLf & vbCrLf
strReport = strReport & vbCrLf & "Grand Total  Transfers = " & Str(gtotTrans) & vbCrLf & vbCrLf

Exit Sub
Errtrap:

Call MsgBox("An error " & Err & " occurred in subroutine 'WorldCount' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub

Private Sub Worldtiles(strOrigFile As String)
Dim x As Integer, strEW As Variant, strNS As Variant
Dim EW(1 To 9) As Variant, ns(1 To 9) As String, i As Integer


x = InStr(3, strOrigFile, "-")
   If x = 0 Then
   x = InStr(3, strOrigFile, "+")
   End If
 strEW = Mid$(strOrigFile, 2, x - 2)
   strNS = Mid$(strOrigFile, x)
   strNS = Left$(strNS, Len(strNS) - 2)
   wTile(5) = strOrigFile
   EW(1) = strEW - 1
   EW(2) = strEW
   EW(3) = strEW + 1
   EW(4) = strEW - 1
   EW(5) = strEW
   EW(6) = strEW + 1
   EW(7) = strEW - 1
   EW(8) = strEW
   EW(9) = strEW + 1
   ns(1) = strNS + 1
   ns(2) = strNS + 1
   ns(3) = strNS + 1
   ns(4) = strNS
   ns(5) = strNS
   ns(6) = strNS
   ns(7) = strNS - 1
   ns(8) = strNS - 1
   ns(9) = strNS - 1
   For i = 1 To 9
   EW(i) = Format(EW(i), "000000")
   If Left$(EW(i), 1) <> "-" Then
   EW(i) = "+" & EW(i)
   End If
   ns(i) = Format(ns(i), "000000")
   If Left$(ns(i), 1) <> "-" Then
   ns(i) = "+" & ns(i)
   wTile(i) = "w" & EW(i) & ns(i) & ".w"
   End If
   Next i
 
End Sub


Private Sub Command100_Click()
Command7.value = True
End Sub

Private Sub Command101_Click()
Command80.value = True
End Sub


Private Sub Command102_Click()
Dim x As Integer, i As Integer

ReDim Rname(0 To NumRoutes - 1)

For i = 0 To NumRoutes - 1
frmUtils.Dir1(0).Path = AllRoutes2(i)
x = InStrRev(AllRoutes2(i), "\")
Rname(i) = Mid$(AllRoutes2(i), x + 1)
Text1(0) = "*.trk"
If File1(0).ListCount > 0 Then
frmActive.List1.AddItem Rname(i)
frmActive.List1.Selected(i) = True
GoTo NextOne
End If
Text1(0) = "*.off"
If File1(0).ListCount > 0 Then
frmActive.List1.AddItem Rname(i)
frmActive.List1.Selected(i) = False
GoTo NextOne
End If

frmActive.List1.AddItem Rname(i) & "  --> Train-Store"

NextOne:
Next i
frmActive.Show 1
DoEvents
Text1(0) = "*.*"
End Sub

Private Sub Command103_Click()
Dim strFirst As String, TrainsetPath As String

MousePointer = 11
TrainsetPath = MSTSPath & "\Trains\Trainset\"

strReport = vbNullString
strFirst = Dir1(0).Path
'Call CountStock
DoEvents
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = TrainsetPath
Text1(0) = "*.cvf"
frmUtils.Refresh
DoEvents
Check1.value = 1


booFixCVF = True

frmSearch.Show
DoEvents
Dir1(0).Path = strFirst
booFixCVF = False
MousePointer = 0
Select Case MsgBox("You must re-start Route_Riter after using the Fix .CVF files option.", vbOKCancel Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbOK
Command1(15).value = True
    Case vbCancel
Exit Sub
End Select

End Sub

Private Sub Command104_Click()

Dim i As Integer, strComShape As String
Dim strComFile As String, strPath As String


If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If

On Error GoTo Errtrap

If strComPath = vbNullString Then
Call MsgBox("You do not appear to have selected your Common Path" _
            & vbCrLf & "from the Files menu." _
            , vbExclamation, App.Title)

Exit Sub
End If
Command7.value = True
DoEvents
MousePointer = 11

 cursouind = 0
'**************** Shapes ***************
Drive1(0).Drive = Left$(ShapePath, 2)
Dir1(0).Path = ShapePath
Text1(0).Text = "*.s*"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Shapes\"
'ShapePath = Dir1(0).path
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Sounds ***************
strPath = RoutePath & "\Sound"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.*"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Sound\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Env Textures ***************
strPath = RoutePath & "\Envfiles\Textures"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Envfiles\Textures\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Terrtex ***************
strPath = RoutePath & "\Terrtex"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Terrtex\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Terrtex Snow***************
strPath = RoutePath & "\Terrtex\Snow"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Terrtex\Snow\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Textures ***************
strPath = RoutePath & "\Textures"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Textures Snow***************
strPath = RoutePath & "\Textures\Snow"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\Snow\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Textures Night***************
strPath = RoutePath & "\Textures\Night"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\Night\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Textures Autumn***************
strPath = RoutePath & "\Textures\Autumn"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\Autumn\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Textures AutumnSnow***************
strPath = RoutePath & "\Textures\AutumnSnow"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\AutumnSnow\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Textures Spring***************
strPath = RoutePath & "\Textures\Spring"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\Spring\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Textures SpringSnow***************
strPath = RoutePath & "\Textures\SpringSnow"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\SpringSnow\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Textures Winter***************
strPath = RoutePath & "\Textures\Winter"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\Winter\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
'**************** Textures WinterSnow***************
strPath = RoutePath & "\Textures\WinterSnow"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\WinterSnow\"
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
MousePointer = 0
SB3.Panels(2).Text = "Finished"
  Drive1(0).Drive = Left$(RoutePath, 2)
Dir1(0).Path = RoutePath
Text1(0).Text = "*.*"
Exit Sub
Errtrap:

 
 Call MsgBox(Err.Description & " occurred while copying the" _
             & vbCrLf & "file - " & File1(cursouind).List(i) _
             , vbExclamation, App.Title)
 
  
Resume Next
End Sub

Private Sub Command105_Click()
Dim strBatText As String, i As Integer, strComShape As String, Newfile3 As Integer
Dim strComFile As String, strPath As String, strDrive As String


If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If

On Error GoTo Errtrap

If strComPath = vbNullString Then
Call MsgBox("You do not appear to have selected your Common Path" _
            & vbCrLf & "from the Files menu." _
            , vbExclamation, App.Title)

Exit Sub
End If
'Command7.value = True
DoEvents
MousePointer = 11

 cursouind = 0
 '**************** Shapes ***************
 strPath = RoutePath & "\Shapes"
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.*"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Shapes\"
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
  
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
  
  strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
  
  Kill fullpath$
  DoEvents
  End If
  End If
DoEvents

   End If


   Next i

 '**************** Textures **************
 strPath = RoutePath & "\Textures"
 Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
 
  '************* Snow ***************
   strPath = RoutePath & "\Textures\Snow"
   Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\Snow"
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
  
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
     '************* Night ***************
      strPath = RoutePath & "\Textures\Night"
   Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\Night"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
        '************* Autumn ***************
         strPath = RoutePath & "\Textures\Autumn"
   Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\Autumn"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
          '************* Autumn Snow***************
           strPath = RoutePath & "\Textures\AutumnSnow"
   Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\AutumnSnow"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
            '************* Spring***************
             strPath = RoutePath & "\Textures\Spring"
   Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\Spring"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
              '************* SpringSnow***************
               strPath = RoutePath & "\Textures\SpringSnow"
   Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\SpringSnow"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
''************* Winter***************
 strPath = RoutePath & "\Textures\Winter"
   Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\Winter"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
              '************* WinterSnow***************
               strPath = RoutePath & "\Textures\WinterSnow"
   Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Textures\WinterSnow"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
 '************** Terrtex ***************
 strPath = RoutePath & "\TerrTex"
    Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Terrtex"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
   '************** Terrtex\Snow ***************
 strPath = RoutePath & "\TerrTex\Snow"
    Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Terrtex\Snow"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
     '************** Envfile\Textures ***************
 strPath = RoutePath & "\EnvFiles\Textures"
    Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Envfiles\Textures"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i
       '************** Sound  ***************
 strPath = RoutePath & "\Sound"
    Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.*"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\Sound"
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & "\" & File1(cursouind).List(i)
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
    strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
      Kill fullpath$
    DoEvents
    End If
  End If
  DoEvents
   End If
  Next i

  '******************** Set up batch file ********************************
   Newfile3 = FreeFile

Open App.Path & "\TempFiles\do_read.bat" For Output As #Newfile3

   Print #Newfile3, strBatText
   
   Close Newfile3
  
   strDrive = Left$(App.Path, 1)
      ChDrive strDrive
ChDir App.Path & "\TempFiles"

  DoEvents

Call ShellAndWait("do_read.bat", True, vbNormalFocus)
 MousePointer = 0
SB3.Panels(2).Text = "Finished"
  Drive1(0).Drive = Left$(RoutePath, 2)
Dir1(0).Path = RoutePath
Text1(0).Text = "*.*"
  Exit Sub
Errtrap:
  If Err = 70 Then
  MousePointer = 0
  Call MsgBox(Err.Description & " " & Err & " occurred while processing " & File1(cursouind).List(i) _
             & vbCrLf & "File is locked by another program - restart Route_Riter and try again." _
             , vbExclamation, App.Title)
  
 Exit Sub
 Else
  Call MsgBox(Err.Description & " occurred while processing the" _
             & vbCrLf & "file - " & File1(cursouind).List(i) _
             , vbExclamation, App.Title)
Resume Next
End If
End Sub




Private Sub Command106_Click()
Dim FirstPath As String, DirCount As Integer, result As String

MousePointer = 11
cursouind = 0
Dir1(cursouind).Path = strComPath
Text1(cursouind) = "*.*"
booLink = True
lLink = 0
DoEvents
FirstPath = frmUtils.Dir1(cursouind).Path
    DirCount = frmUtils.Dir1(cursouind).ListCount
    result = DirDiver(FirstPath, DirCount, "")
 DoEvents
 frmLinks.Label3.Caption = lLink
 
frmLinks.Show

MousePointer = 0
End Sub


Private Sub Command107_Click()


Dim FirstPath As String, DirCount As Integer, result As String

cursouind = 0

Text1(cursouind) = "*.*"
 booLink = False
 lLink = 0
DoEvents
FirstPath = frmUtils.Dir1(cursouind).Path
    DirCount = frmUtils.Dir1(cursouind).ListCount
    result = DirDiver(FirstPath, DirCount, "")
DoEvents
 frmLinks.Label3.Caption = lLink
frmLinks.Show
End Sub


Private Sub Command108_Click()
frmConsists.Show
End Sub

 



Private Sub Command109_Click()
Dim i As Integer, NewRouteName As String
Dim booExists As Boolean, OldRouteName As String

On Error GoTo Errtrap
MousePointer = 11
Label9.Visible = True
SparePath = App.Path & "\TempFiles"
strReport = vbNullString
RoutePath = File1(0).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If
Text1(0) = "*.trk"
RouteName = File1(0).List(i)
OldRouteName = RouteName
Call FindRouteName(RoutePath & "\" & RouteName, NewRouteName, booExists)
If booExists = False Then
MsgBox Lang(339) & vbCr & Lang(340), 16, Lang(341)
Exit Sub
Else
RouteName = NewRouteName
End If
WorldPath = RoutePath & "\World"

RouteListed = True
Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
Close

cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************


If RouteName = vbNullString Then
Call MsgBox(Lang(463), vbExclamation, App.Title)

Exit Sub
End If

numPoints = 1

Label9.Caption = "Uncompressing .w files"
DoEvents
Call UncompressAllW(WorldPath)
Call FindStuck
Exit Sub
Errtrap:
 Call MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'Command109' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
           
            Resume Next

End Sub
Private Sub Command110_Click()
Dim strNew As String, MySettings As Variant
MySettings = GetAllSettings("Route_Riter6", "OldMainPath")

If Not IsEmpty(MySettings) Then
strOldRegPath = MySettings(0, 1)
Label5 = strOldRegPath
Else
Call MsgBox("No old registry path has been stored.", vbExclamation Or vbDefaultButton1, App.Title)

Exit Sub
End If
Select Case MsgBox("Please confirm you wish to reset Registry entries to:-" & vbCrLf & strOldRegPath _
                   & vbCrLf & "" _
                   , vbOKCancel Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbOK
strNew = strOldRegPath & Chr$(0)

Call SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "Path", strNew, REG_SZ)
DoEvents

Call SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "EXE Path", strNew, REG_SZ)

    Case vbCancel

End Select
End Sub

Private Sub Command111_Click()


frmDialog4.Show
If booRaildriver = False Then
Exit Sub
End If



End Sub


Private Sub Command112_Click()
Dim strTemp As String, MyString As String
Dim x As Long, MyASCFile As String, xx As Long, strStart As String, strEnd As String
Dim Y As Long, strBrakeForce As String, booWrong As Boolean, i As Integer
strReport = vbNullString
Label5.Caption = vbNullString
For i = 0 To File1(cursouind).ListCount - 1
   
  
   If File1(cursouind).Selected(i) Then
       
        fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
        If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryON
        MyString = ReadUniFile(fullpath$)
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
        x = InStr(MyString, "Type ( Diesel")
        If x = 0 Then booWrong = True: GoTo CarryON
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDDiesel1.txt")
        Rem ********* Get Brake force *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strBrakeForce = Mid$(MyString, x, (xx + 1) - x)
        '*************Brake Equipt************************
        x = InStr(MyString, "BrakeEquipmentType")
        xx = InStr(MyString, "BrakeCylinderPressureForMaxBrakeBrakeForce")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '*************Vacuum****************************
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDDiesel2.txt")
        x = InStr(Y, MyString, "AirBrakesAirCompressorPowerRating")
        If x = 0 Then                 'This loco is not vacuum braked*****************
        x = InStr(Y, MyString, "EngineBrakesControllerMinPressureReduction")
        End If
        If x = 0 Then GoTo EngControl
        xx = InStr(MyString, "BrakesEngineControllers")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '***************EngineControllers ***************
EngControl:
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDDiesel3.txt")
        Y = InStr(MyString, "BrakesEngineControllers")
        x = InStr(Y + 20, MyString, "EngineControllers")
        xx = InStr(x, MyString, "BailOffButton")
        If xx <> 0 Then
        xx = InStr(xx, MyString, ")")
        xx = InStr(xx + 1, MyString, ")")
        xx = xx + 1
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "Sound")
        End If
        If x > 0 And xx > x Then
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx)
        MyString = strStart & MyASCFile & strEnd
        '************ Replace MaxBrakeForce *************
        x = 1
CheckAgain:
        x = InStr(x, MyString, "MaxBrakeForce")
        If x = 0 Then GoTo GotIt
        xx = InStr(x, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx + 1)
        MyString = strStart & strBrakeForce & strEnd
        x = x + 1
        GoTo CheckAgain
GotIt:
           
        '***********************************************
        strTemp = Left$(fullpath$, Len(fullpath$) - 1) & "x"
        If FileExists(strTemp) Then
        Kill strTemp
        DoEvents
        End If
        Name fullpath$ As strTemp
        DoEvents
        Call WriteUniFile(fullpath$, MyString)
   DoEvents
      End If
   End If
CarryON:
If booWrong = True Then
booWrong = False
strReport = strReport & fullpath$ & " was not an Diesel loco" & vbCrLf
End If
Next i
Label5.Caption = "Finished"
If strReport <> vbNullString Then
frmReport.Rich1.Text = strReport
 
     frmReport.Show 1

End If
End Sub

Private Sub Command113_Click()
Dim strTemp As String, MyString As String
Dim x As Long, MyASCFile As String, xx As Long, strStart As String, strEnd As String
Dim Y As Long, strBrakeForce As String, booWrong As Boolean, i As Integer
strReport = vbNullString
Label5.Caption = vbNullString
For i = 0 To File1(cursouind).ListCount - 1
   
  
   If File1(cursouind).Selected(i) Then
       
        fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
        If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryON
        MyString = ReadUniFile(fullpath$)
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
'        X = InStr(Mystring, "Type ( Diesel")
'        If X = 0 Then booWrong = True: GoTo CarryOn
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDCombo1.txt")
        Rem ********* Get Brake force *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strBrakeForce = Mid$(MyString, x, (xx + 1) - x)
        '*************Brake Equipt************************
        x = InStr(MyString, "BrakeEquipmentType")
        xx = InStr(MyString, "BrakeCylinderPressureForMaxBrakeBrakeForce")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '*************Vacuum****************************
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDCombo2.txt")
        x = InStr(Y, MyString, "AirBrakesAirCompressorPowerRating")
        If x = 0 Then                 'This loco is not vacuum braked*****************
        x = InStr(Y, MyString, "TrainBrakesControllerMinPressureReduction")
        End If
        If x = 0 Then GoTo EngControl
        xx = InStr(MyString, "BrakesEngineControllers")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '***************EngineControllers ***************
EngControl:
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDCombo3.txt")
        Y = InStr(MyString, "BrakesEngineControllers")
        x = InStr(Y + 20, MyString, "EngineControllers")
        xx = InStr(x, MyString, "Headlights")
        If xx > 0 Then
        Y = InStr(xx, MyString, "Wipers (")
        If Y > 0 Then
        xx = Y
        End If
        End If
        If xx <> 0 Then
        xx = InStr(xx, MyString, ")")
        xx = InStr(xx + 1, MyString, ")")
        xx = xx + 1
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "Sound")
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "Name")
        End If
        If x > 0 And xx > x Then
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx)
        MyString = strStart & MyASCFile & strEnd
        '************ Replace MaxBrakeForce *************
        x = 1
CheckAgain:
        x = InStr(x, MyString, "MaxBrakeForce")
        If x = 0 Then GoTo GotIt
        xx = InStr(x, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx + 1)
        MyString = strStart & strBrakeForce & strEnd
        x = x + 1
        GoTo CheckAgain
GotIt:
           
        '***********************************************
        strTemp = Left$(fullpath$, Len(fullpath$) - 1) & "x"
        If FileExists(strTemp) Then
        Kill strTemp
        DoEvents
        End If
        Name fullpath$ As strTemp
        DoEvents
        Call WriteUniFile(fullpath$, MyString)
   DoEvents
      End If
   End If
CarryON:
If booWrong = True Then
booWrong = False
strReport = strReport & fullpath$ & " was not an Combo loco" & vbCrLf
End If
Next i
Label5.Caption = "Finished"
If strReport <> vbNullString Then
frmReport.Rich1.Text = strReport
 
     frmReport.Show 1

End If
End Sub


Private Sub Command114_Click()
Dim i As Integer, x As Long, MyString As String, fullpath$, Y As Long
Dim strCab As String, strEngpath As String, strCSV As String, booAlias As Boolean
Dim booAliased As Boolean

For i = 0 To File1(cursouind).ListCount - 1
   
  
   If File1(cursouind).Selected(i) Then
  
       booAliased = False
        fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
        If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryON
        x = InStrRev(fullpath$, "\")
        strEngpath = Left$(fullpath$, x)
        MyString = ReadUniFile(fullpath$)
        x = InStr(MyString, ".cvf")
        If x > 0 Then
        Y = InStrRev(MyString, "(", x)
        strCab = Mid$(MyString, Y + 1, x - (Y + 1) + 4)
        strCab = Trim$(strCab)
        If Left$(strCab, 2) = ".." Then
        booAliased = True
        GoTo AliasedCab
        End If
        If FileExists(strEngpath & "CabView\" & strCab) Then
        strCSV = ReadUniFile(strEngpath & "CabView\" & strCab)
        x = InStr(strCSV, "combinedcontrol")
        If x > 0 Then
        Label5.Caption = "This loco has a Combined Control handle"
        GoTo CarryON
        End If
        Else
        booAliased = True
        End If
        End If
AliasedCab:
        
        x = InStr(MyString, "GearBox")
        If x > 0 Then
        Label5.Caption = "Geared Diesel Loco"
        GoTo CarryON
        End If
        x = InStr(MyString, "Type ( Electric")
        If x > 0 Then
        Label5.Caption = "Electric Loco"
        GoTo CarryON
        End If
        x = InStr(MyString, "Type ( Steam")
        If x > 0 Then
        Label5.Caption = "Steam Loco"
        GoTo CarryON
        End If
        x = InStr(MyString, "Type ( Diesel")
        If x > 0 Then
        Label5.Caption = "Diesel Loco"
        GoTo CarryON
        End If
        If x = 0 Then
        Label5.Caption = "Unable to determine Loco type"
        End If
    'End If
CarryON:
If booAlias = True Then
  Label5.Caption = Label5.Caption & " Loco has aliased cab so unable determine if Combo"
  End If
  End If
Next i
        
        
        
End Sub

Private Sub Command115_Click()
Dim result As String, MySettings As Variant

MySettings = GetAllSettings("Route_Riter6", "OldMainPath")

If Not IsEmpty(MySettings) Then
strOldRegPath = MySettings(0, 1)
Label5 = strOldRegPath
Else

result = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "Path")
strOldRegPath = result
If Right$(strOldRegPath, 1) = Chr$(0) Then
strOldRegPath = Left$(strOldRegPath, Len(strOldRegPath) - 1)
End If
Label5 = strOldRegPath
SaveSetting "Route_Riter6", "OldMainPath", "OldMainPath", strOldRegPath
End If
End Sub

Private Sub Command116_Click()
Dim strTemp As String, MyString As String
Dim x As Long, MyASCFile As String, xx As Long, strStart As String, strEnd As String
Dim Y As Long, strBrakeForce As String, booWrong As Boolean, i As Integer

strReport = vbNullString
Label5.Caption = vbNullString
For i = 0 To File1(cursouind).ListCount - 1
   
  
   If File1(cursouind).Selected(i) Then
        
        fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
        If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryON
        MyString = ReadUniFile(fullpath$)
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
        x = InStr(MyString, "GearBox")
        If x = 0 Then booWrong = True: GoTo CarryON
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDGeared1.txt")
        Rem ********* Get Brake force *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strBrakeForce = Mid$(MyString, x, (xx + 1) - x)
        '*************Brake Equipt************************
        x = InStr(MyString, "BrakeEquipmentType")
        xx = InStr(MyString, "BrakeCylinderPressureForMaxBrakeBrakeForce")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '*************Brakes****************************
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDGeared2.txt")
        x = InStr(Y, MyString, "TrainBrakesControllerMaxApplicationRate")
        If x = 0 Then
        x = InStr(Y, MyString, "TrainBrakesControllerMaxReleaseRate")
        End If
        If x = 0 Then GoTo EngControl
        xx = InStr(MyString, "BrakesEngineControllers")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '***************EngineControllers ***************
EngControl:
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDGeared3.txt")
        Y = InStr(MyString, "BrakesEngineControllers")
        x = InStr(Y + 20, MyString, "EngineControllers")
        xx = InStr(x, MyString, "Headlights")
        If xx <> 0 Then
        xx = InStr(xx, MyString, ")")
        xx = InStr(xx + 1, MyString, ")")
        xx = xx + 1
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "Name")
        End If
        If x > 0 And xx > x Then
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx)
        MyString = strStart & MyASCFile & strEnd
        '************ Replace MaxBrakeForce *************
        x = 1
CheckAgain:
        x = InStr(x, MyString, "MaxBrakeForce")
        If x = 0 Then GoTo GotIt
        xx = InStr(x, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx + 1)
        MyString = strStart & strBrakeForce & strEnd
        x = x + 1
        GoTo CheckAgain
GotIt:
           
        '***********************************************
        strTemp = Left$(fullpath$, Len(fullpath$) - 1) & "x"
        If FileExists(strTemp) Then
        Kill strTemp
        DoEvents
        End If
        Name fullpath$ As strTemp
        DoEvents
        Call WriteUniFile(fullpath$, MyString)
   DoEvents
      End If
   End If
CarryON:
If booWrong = True Then
booWrong = False
strReport = strReport & fullpath$ & " was not an Geared loco" & vbCrLf
End If
Next i
Label5.Caption = "Finished"
If strReport <> vbNullString Then
frmReport.Rich1.Text = strReport
 
     frmReport.Show 1

End If
End Sub

Private Sub Command117_Click()
Dim result As String

result = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "Path")

Label5.Caption = result


End Sub

Private Sub Command118_Click()
Dim strTemp As String, MyString As String
Dim x As Long, MyASCFile As String, xx As Long, strStart As String, strEnd As String
Dim Y As Long, strBrakeForce As String, booWrong As Boolean, i As Integer

strReport = vbNullString
Label5.Caption = vbNullString
For i = 0 To File1(cursouind).ListCount - 1
   
  
   If File1(cursouind).Selected(i) Then
        
        fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
        If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryON
        MyString = ReadUniFile(fullpath$)
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
        x = InStr(MyString, "Type ( Steam")
        If x = 0 Then booWrong = True: GoTo CarryON
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDSteam1.txt")
        Rem ********* Get Brake force *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strBrakeForce = Mid$(MyString, x, (xx + 1) - x)
        '*************Brake Equipt************************
        x = InStr(MyString, "BrakeEquipmentType")
        xx = InStr(MyString, "BrakeCylinderPressureForMaxBrakeBrakeForce")
        Y = InStr(xx, MyString, ")")
        
        
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '*************Vacuum****************************
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDSteam2.txt")
        x = InStr(Y, MyString, "VacuumBrakes")
        If x = 0 Then
        x = InStr(Y, MyString, "AirBrakesAirCompressorPowerRating")
        End If
        xx = InStr(MyString, "BrakesEngineControllers")
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '***************EngineControllers ***************
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDSteam3.txt")
        Y = InStr(MyString, "BrakesEngineControllers")
        x = InStr(Y + 20, MyString, "EngineControllers")
        xx = InStr(x, MyString, "Brake_Hand")
        If xx <> 0 Then
        xx = InStr(xx, MyString, ")")
        xx = InStr(xx + 1, MyString, ")")
        xx = xx + 1
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "FireDoor")
        End If
        If x > 0 And xx > x Then
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx)
        MyString = strStart & MyASCFile & strEnd
        '************ Replace MaxBrakeForce *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx + 1)
        MyString = strStart & strBrakeForce & strEnd
        '***********************************************
        strTemp = Left$(fullpath$, Len(fullpath$) - 1) & "x"
        If FileExists(strTemp) Then
        Kill strTemp
        DoEvents
        End If
        Name fullpath$ As strTemp
        DoEvents
        Call WriteUniFile(fullpath$, MyString)
   DoEvents
      End If
   End If
CarryON:
If booWrong = True Then
booWrong = False
strReport = strReport & fullpath$ & " was not an Steam loco - no changes were made" & vbCrLf
End If
Next i
Label5.Caption = "Finished"
If strReport <> vbNullString Then
frmReport.Rich1.Text = strReport
 
     frmReport.Show 1

End If
End Sub

Private Sub Command119_Click()
Dim strTemp As String, MyString As String
Dim x As Long, MyASCFile As String, xx As Long, strStart As String, strEnd As String
Dim Y As Long, strBrakeForce As String, booWrong As Boolean, strWrong As String, i As Integer

strReport = vbNullString
Label5.Caption = vbNullString
For i = 0 To File1(cursouind).ListCount - 1
   
  
   If File1(cursouind).Selected(i) Then
       
        fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
        If Right$(fullpath$, 4) <> ".eng" Then GoTo CarryON
        MyString = ReadUniFile(fullpath$)
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
        x = InStr(MyString, "Type ( Electric")
        If x = 0 Then
        booWrong = True
        strWrong = " does not appear to be an Electric loco "
        GoTo CarryON
        End If
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDElectric1.txt")
        Rem ********* Get Brake force *************
        x = InStr(MyString, "MaxBrakeForce")
        xx = InStr(x, MyString, ")")
        strBrakeForce = Mid$(MyString, x, (xx + 1) - x)
        '*************Brake Equipt************************
        x = InStr(MyString, "BrakeEquipmentType")
        xx = InStr(MyString, "BrakeDistributorNormalFullReleasePressure")
        If xx = 0 Then
        xx = InStr(MyString, "BrakeCylinderPressureForMaxBrakeBrakeForce")
        End If
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '*************Vacuum****************************
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDElectric2.txt")
        x = InStr(Y, MyString, "VacuumBrakes")
        If x = 0 Then                 'This loco is not vacuum braked*****************
        x = InStr(Y, MyString, "AirBrakesAirCompressorPowerRating")
        End If
        If x = 0 Then                 'This loco is not vacuum braked*****************
        x = InStr(Y, MyString, "EngineBrakesControllerDirectControlExponent")
        End If
        xx = InStr(MyString, "BrakesEngineControllers")
        If xx = 0 Then
        xx = InStr(MyString, "BrakesTrainBrakeType")
        End If
        Y = InStr(xx, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, Y + 1)
        MyString = strStart & MyASCFile & strEnd
        '***************EngineControllers ***************
        MyASCFile = ReadASCIIFile(App.Path & "\RailDriver\RDElectric3.txt")
        Y = InStr(MyString, "BrakesEngineControllers")
        x = InStr(Y + 20, MyString, "EngineControllers")
        xx = InStr(x, MyString, "Brake_Hand")
        If xx <> 0 Then
        xx = InStr(xx, MyString, ")")
        xx = InStr(xx + 1, MyString, ")")
        xx = xx + 1
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "DirControl")
        End If
        If xx = 0 Then
        xx = InStr(x, MyString, "EmergencyStopResetToggle")
        End If
        If xx = 0 Then
        booWrong = True
        strWrong = " could not be modified automatically"
        GoTo CarryON
        End If
        If x > 0 And xx > x Then
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx)
        MyString = strStart & MyASCFile & strEnd
        '************ Replace MaxBrakeForce *************
        x = 1
CheckAgain:
        x = InStr(x, MyString, "MaxBrakeForce")
        If x = 0 Then GoTo GotIt
        xx = InStr(x, MyString, ")")
        strStart = Left$(MyString, x - 1)
        strEnd = Mid$(MyString, xx + 1)
        MyString = strStart & strBrakeForce & strEnd
        x = x + 1
        GoTo CheckAgain
GotIt:
           
        '***********************************************
        strTemp = Left$(fullpath$, Len(fullpath$) - 1) & "x"
        If FileExists(strTemp) Then
        Kill strTemp
        DoEvents
        End If
        Name fullpath$ As strTemp
        DoEvents
        Call WriteUniFile(fullpath$, MyString)
   DoEvents
      End If
   End If
CarryON:
If booWrong = True Then
booWrong = False
strReport = strReport & fullpath$ & strWrong & vbCrLf
End If
Next i
Label5.Caption = "Finished"
If strReport <> vbNullString Then
frmReport.Rich1.Text = strReport
 
     frmReport.Show 1

End If
End Sub


Private Sub Command120_Click()
Dim strBatText As String

strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "TsUtil_version.log" & ChrW$(34) & "  version"


ChDrive Left$(App.Path, 1)
' 'ChDir App.Path & "\TSUtil"
  

 Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 If Not FileExists(App.Path & "\Reports\" & RouteName & "TsUtil_version.log") Then
Call MsgBox("As no .log file has been created, you can assume that TsUtils" _
            & vbCrLf & "did not complete the option you were attempting to run." _
            , vbCritical, App.Title)

MousePointer = 0
Exit Sub
End If
 
 
  MousePointer = 0
 frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "TsUtil_version.log"
 frmReport.Show 1
 DoEvents

End Sub

Private Sub Command121_Click()
Command54.value = True
End Sub

Private Sub Command122_Click()
Dim strTemp As String, x As Integer, ExeName As String



    If RouteListed = False Then
Select Case MsgBox(Lang(451) & vbCrLf & Dir1(cursouind).Path, vbOKCancel + vbExclamation + vbDefaultButton1, "Route Selection")

    Case vbOK

    Case vbCancel
Exit Sub
End Select

End If

''    'Open an archive  exist will create a new archive
'



    On Error Resume Next
    FromZip = 1

x = InStrRev(RoutePath, "\")
NewZipPath = Left$(RoutePath, x - 1)
ExeName = Mid$(RoutePath, x + 1)
ZipName = RouteName
ChDrive Left$(RoutePath, 1)
 ChDir NewZipPath
'strTemp = App.path & "\UHARC a -m3 -r -ed+ -pr -sfx -tRoutes " & chrw$(34) & RouteName & chrw$(34) & " " & chrw$(34) & RoutePath & "\*.*" & chrw$(34)
'strTemp = App.path & "\UHARC a -m3 -r -ed+ -p- -sfx -tRoutes " & chrw$(34) & RouteName & chrw$(34) & " " & chrw$(34) & RoutePath & "\*.*" & chrw$(34)
strTemp = App.Path & "\UHARC a -m3 -r -ed+ -pr -sfx " & RouteName & ".exe" & " " & ExeName & "\*.*"

Call ShellAndWait(strTemp, True, vbNormalFocus)
DoEvents
End Sub

Private Sub Command123_Click()
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If

Call ListTrainsCopy
Text1(0) = "*.*"

End Sub

Private Sub Command124_Click()
Text1(0).Text = "*.eng;*.wag"
End Sub


 


Private Sub Command125_Click()
Dim i As Integer, NewRouteName As String
Dim booExists As Boolean, OldRouteName As String, SparePath As String
MousePointer = 11
Label9.Visible = True
SparePath = App.Path & "\TempFiles"
strReport = vbNullString
RoutePath = File1(0).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If
Text1(0) = "*.trk"
RouteName = File1(0).List(i)
OldRouteName = RouteName
Call FindRouteName(RoutePath & "\" & RouteName, NewRouteName, booExists)
If booExists = False Then
MsgBox Lang(339) & vbCr & Lang(340), 16, Lang(341)
Exit Sub
Else
RouteName = NewRouteName
End If
WorldPath = RoutePath & "\World"

RouteListed = True
Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
Close

cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************


If RouteName = vbNullString Then
Call MsgBox(Lang(463), vbExclamation, App.Title)

Exit Sub
End If
Label9.Caption = "Uncompressing World Files"
DoEvents
Call UncompressAllW(WorldPath)
Call FixStuck(SparePath, WorldPath)

End Sub

Private Sub Command126_Click()
Dim Filpath1$, fullpath$, flagway As Integer, i As Integer
Dim strOrigPath As String, strErrorBias As String


On Error GoTo Errtrap
ReDim strWorldTiles(1 To CHUNK)
MousePointer = 11
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If
Text1(0) = "*.trk"
If File1(cursouind).ListCount = 0 Then
Call MsgBox("There does not appear to be a valid .trk entry for this route?", vbExclamation, App.Title)
Exit Sub
End If
Label9.Visible = True
 strOrigPath = File1(0).Path & "\Tiles"

Filpath1$ = App.Path & "\TempFiles"
Call KillSpare("*.t")
Call KillSpare("*.w")
strErrorBias = InputBox("Enter Bias Value ( 0 or 1 )")
If strErrorBias = vbNullString Then
MousePointer = 0
Exit Sub
End If

TokMode = 2

Rem *************** Check world ********
Call DoDeCompFolder("w", WorldPath, Filpath1$)

 cursouind = 0
 Drive1(0).Drive = Left$(Filpath1$, 2)
Dir1(0).Path = Filpath1$
Text1(0).Text = "*.w"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i



For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
  
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
Call ReadWorld3(fullpath$)
If numWorldTiles > 0 Then
Label9.Caption = strWorldTiles(numWorldTiles)
DoEvents
End If
End If
Next i

ReDim Preserve strWorldTiles(1 To numWorldTiles)

'Rem *********************************
For i = 1 To numWorldTiles
If Not FileExists(Filpath1$ & "\" & strWorldTiles(i)) Then

Call DoDeComp2(strWorldTiles(i), strOrigPath, Filpath1$)
End If

Next i

cursouind = 1
Drive1(cursouind).Drive = Left$(Filpath1$, 2)
Dir1(cursouind).Path = Filpath1$
Text1(cursouind) = "*.t"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
Label9.Visible = True

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
     Label9.Caption = "Processing:  " & File1(cursouind).List(i)
flagway = 0
Call ConvertT(fullpath$, strErrorBias, flagway)
DoEvents
flagway = 1
Call ConvertT(fullpath$, strErrorBias, flagway)
DoEvents

Kill strOrigPath & "\" & File1(cursouind).List(i)
DoEvents

Call DoComp(File1(cursouind).List(i), File1(cursouind).Path, strOrigPath)
DoEvents
Kill fullpath$
DoEvents
End If
Next i
MousePointer = 0
Label9.Caption = "Processing Finished"

Exit Sub
Errtrap:
Call MsgBox("An error " & Err.Description & " occurred while processing " & File1(cursouind).List(i), vbExclamation, App.Title)

End Sub

Private Sub Command127_Click()
Dim myrecord(1 To 8) As String, myTemp As String, i As Integer, strBatText As String, strCB As String

If strConbuilder = vbNullString Then
CDL1.DialogTitle = "Select Conbuilder.exe"
CDL1.Flags = cdlOFNExplorer
If strConbuilder <> vbNullString Then
CDL1.Filename = strConbuilder
End If
CDL1.ShowOpen
strConbuilder = CDL1.Filename
SaveSetting "Route_Riter6", "Conbuilder", "Conbuilder", strConbuilder
End If
'End If
i = InStrRev(strConbuilder, "conbuilder.exe")
strCB = Left$(strConbuilder, i - 1)
Open strCB & "settings.bin" For Binary As #1
myTemp = String(200, " ")
Get #1, , myTemp
myrecord(1) = myTemp
myTemp = String(200, " ")
Get #1, , myTemp
myrecord(2) = myTemp

myTemp = String(200, " ")
Get #1, , myTemp
myrecord(3) = myTemp
myTemp = String(200, " ")
Get #1, , myTemp
myrecord(4) = myTemp

myTemp = String(200, " ")
Get #1, , myTemp
myrecord(5) = myTemp
myTemp = String(200, " ")
Get #1, , myTemp
myrecord(6) = myTemp

myTemp = String(200, " ")
Get #1, , myTemp
myrecord(7) = myTemp
myTemp = String(264, " ")
Get #1, , myTemp
myrecord(8) = myTemp
Close #1

Open strCB & "settings.bin" For Binary As #1
myrecord(2) = MSTSPath & "\Trains\Consists"
myrecord(2) = myrecord(2) & String(200 - Len(myrecord(2)), " ")
myrecord(4) = MSTSPath
myrecord(4) = myrecord(4) & String(200 - Len(myrecord(4)), " ")
myrecord(6) = MSTSPath & "\Sound"
myrecord(6) = myrecord(6) & String(200 - Len(myrecord(6)), " ")
myrecord(7) = MSTSPath & "\Trains\Trainset"
myrecord(7) = myrecord(7) & String(200 - Len(myrecord(7)), " ")
For i = 1 To 7
Put #1, , myrecord(i)
Next
Close #1
DoEvents
 strBatText = ChrW$(34) & strCB & "conbuilder.exe" & ChrW$(34)
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)
    DoEvents
End Sub

Private Sub Command128_Click()
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If
MousePointer = 11
SB1.Panels(4).Text = time
SB1.Panels(6).Text = vbNullString

strReport = vbNullString
Call CompactTrack
DoEvents

SB1.Panels(6).Text = time
booWorldCount = True
DoEvents
MousePointer = 0

End Sub

Private Sub Command129_Click()
 Dim fullpath$, tempPath As String, smsFound As Boolean, wavFound As Boolean
 Dim strKillPath As String, i As Integer, ii As Integer, x As Integer
 
 
 
 strReport = vbNullString
 WorldPath = RoutePath & "\world"
 strKillPath = RoutePath & "\RemovedSounds"
 If Not DirExists(strKillPath) Then
 MkDir strKillPath
 End If
 If Not DirExists(strKillPath & "\Sound") Then
 MkDir strKillPath & "\Sound"
 End If
 If Not DirExists(strKillPath & "\GlobalSound") Then
 MkDir strKillPath & "\GlobalSound"
 End If
 Rem ************** Find Sound Files ***************
 Call CheckDefaultSounds2
 cursouind = 1
Drive1(cursouind).Drive = Left$(WorldPath, 2)
Dir1(cursouind).Path = WorldPath
Text1(cursouind).Text = "*.ws"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
   
   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   
   Call CheckForSounds3(fullpath$)
   File1(cursouind).Selected(i) = False


   End If
   
   Next i
If FileExists(RoutePath & "\ttype.dat") Then
Call CheckForSounds3(RoutePath & "\ttype.dat")
End If
For i = 0 To SoundNumber
If Soundfile(i) <> vbNullString Then
If FileExists(SoundPath & "\" & Soundfile(i)) Then
  tempPath = SoundPath & "\" & Soundfile(i)
  
  ElseIf FileExists(GlobalSoundPath & "\" & Soundfile(i)) Then
  tempPath = GlobalSoundPath & "\" & Soundfile(i)
  Else
  Call LookForSound(Soundfile(i), "")
  tempPath = SoundPath & "\" & Soundfile(i)
  End If
  strReport = strReport & tempPath & vbCrLf
  Call CheckForWav4(tempPath)
  Soundfile(i) = tempPath
  End If
Next i
For i = 0 To numWave
  If FileExists(SoundPath & "\" & strWave(i)) Then
     strReport = strReport & SoundPath & "\" & strWave(i) & vbCrLf
     strWave(i) = SoundPath & "\" & strWave(i)
     ElseIf FileExists(GlobalSoundPath & "\" & strWave(i)) Then
    strReport = strReport & GlobalSoundPath & "\" & strWave(i) & vbCrLf
    strWave(i) = GlobalSoundPath & "\" & strWave(i)
    End If
    Next i

Rem ************* Kill sounds
booKillMove = True
cursouind = 0
Drive1(cursouind).Drive = Left$(SoundPath, 2)
Dir1(cursouind).Path = SoundPath
Text1(cursouind).Text = "*.sms"

DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
smsFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To SoundNumber

   If SoundPath & "\" & File1(cursouind).List(i) = Soundfile(ii) Then
   smsFound = True


   Exit For
   End If
   Next ii
 If smsFound = False Then

 If booKillMove = False Then
 Kill SoundPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy SoundPath & "\" & File1(cursouind).List(i), strKillPath & "\Sound\" & File1(cursouind).List(i)
 DoEvents
 Kill SoundPath & "\" & File1(cursouind).List(i)
 End If
 End If
 End If
 smsFound = False
 Next i

 Text1(cursouind).Text = "*.wav"
 DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numWave

   If SoundPath & "\" & File1(cursouind).List(i) = strWave(ii) Then
   wavFound = True

   Exit For
   End If
   Next ii
 If wavFound = False Then

 If booKillMove = False Then
 Kill SoundPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy SoundPath & "\" & File1(cursouind).List(i), strKillPath & "\Sound\" & File1(cursouind).List(i)
 DoEvents
 Kill SoundPath & "\" & File1(cursouind).List(i)
 End If

 End If
 End If
 wavFound = False
 Next i
 Text1(cursouind).Text = "*.*"


Rem ************* Kill Global Sounds **************
Drive1(cursouind).Drive = Left$(GlobalSoundPath, 2)
Dir1(cursouind).Path = GlobalSoundPath
Text1(cursouind).Text = "*.sms"

DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
smsFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To SoundNumber

   If GlobalSoundPath & "\" & File1(cursouind).List(i) = Soundfile(ii) Then
   smsFound = True
   Exit For
   End If
   Next ii

 If smsFound = False Then
 x = InStr(File1(cursouind).List(i), "Track")
 If x > 0 Then GoTo GetMore3
 x = InStr(File1(cursouind).List(i), "Rail")
 If x > 0 Then GoTo GetMore3
 x = InStr(File1(cursouind).List(i), "Joint")
 If x > 0 Then GoTo GetMore3
 x = InStr(File1(cursouind).List(i), "Town_")
 If x > 0 Then GoTo GetMore3
 x = InStr(File1(cursouind).List(i), "Statio_")
 If x > 0 Then GoTo GetMore3
 GoTo GetMore
 
GetMore3:

 If booKillMove = False Then
 Kill GlobalSoundPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy GlobalSoundPath & "\" & File1(cursouind).List(i), strKillPath & "\GlobalSound\" & File1(cursouind).List(i)
 DoEvents
 Kill GlobalSoundPath & "\" & File1(cursouind).List(i)
 End If

 End If
 End If
GetMore:
 smsFound = False
 Next i

 Text1(cursouind).Text = "*.wav"
 DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numWave

   If GlobalSoundPath & "\" & File1(cursouind).List(i) = strWave(ii) Then
   wavFound = True

   Exit For
   End If
   Next ii
 If wavFound = False Then
x = InStr(File1(cursouind).List(i), "Track")
 If x > 0 Then GoTo GetMore2
 x = InStr(File1(cursouind).List(i), "Rail")
 If x > 0 Then GoTo GetMore2
 x = InStr(File1(cursouind).List(i), "Joint")
 If x > 0 Then GoTo GetMore2
 x = InStr(File1(cursouind).List(i), "ITR_")
 If x > 0 Then GoTo GetMore2
 x = InStr(File1(cursouind).List(i), "Gen_")
 If x > 0 Then GoTo GetMore2
 GoTo GetMore4
GetMore2:
 If booKillMove = False Then
 Kill GlobalSoundPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy GlobalSoundPath & "\" & File1(cursouind).List(i), strKillPath & "\GlobalSound\" & File1(cursouind).List(i)
 DoEvents
 Kill GlobalSoundPath & "\" & File1(cursouind).List(i)
 End If

 End If
 End If
GetMore4:
 wavFound = False
 Next i
 Text1(cursouind).Text = "*.*"


    
    
MousePointer = 0
If strReport <> vbNullString Then
strReport = "SOUND-FILES USED IN THIS MINI-ROUTE :-" & vbCrLf & vbCrLf & strReport
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString
End If

 
End Sub

Private Sub Command130_Click()
Dim SparePath As String, MSG As String

cursouind = 0
SparePath = App.Path & "\TempFiles"

WorldPath = File1(cursouind).Path
If Right$(WorldPath, 5) <> "World" Then
Call MsgBox("You do not appear to have selected any World tiles yet?" _
            & vbCrLf & "" _
            , vbExclamation, App.Title)

Exit Sub
End If


Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************

MousePointer = 11
Call UncompressSelectedW

GetAnother:
Close

frmDelShape.Show 1

     DoEvents
     
cursouind = 0

 If strDelShape <> vbNullString Then
 
MSG = Lang(110) & strDelShape & " from all selected .W files in " & RouteName
Response = MsgBox(MSG, 36, Lang(464))
Select Case Response
Case vbYes

Call StripW(strDelShape)

strDelShape = vbNullString
Case vbNo
strDelShape = vbNullString
End Select
End If
Select Case MsgBox("Do you wish to delete any more shapes from this Route?", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
GoTo GetAnother
    Case vbNo

End Select
MousePointer = 0
Text1(1).Text = "*.*"
End Sub

Private Sub Command131_Click()
Dim i As Integer, strComShape As String
Dim strComFile As String, GSPath As String, x As Integer


If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
MousePointer = 0
Exit Sub
End If

On Error GoTo Errtrap

If strComPath = vbNullString Then
Call MsgBox("You do not appear to have selected your Common Path" _
            & vbCrLf & "from the Files menu." _
            , vbExclamation, App.Title)
MousePointer = 0
Exit Sub
End If
Command7.value = True
DoEvents
x = InStrRev(ShapePath, "Routes")
GSPath = Left$(ShapePath, x - 1) & "Global\Shapes"

MousePointer = 11

 cursouind = 0
'**************** Shapes ***************
Drive1(0).Drive = Left$(GSPath, 2)

Dir1(0).Path = GSPath
Text1(0).Text = "*.s*"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
If Not DirExists(strComPath & "\GlobalShapes") Then
MkDir strComPath & "\GlobalShapes"
End If

strComShape = strComPath & "\GlobalShapes\"

For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) Then
    SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  If Not FileExists(strComFile) Then
    FileCopy fullpath$, strComFile
    DoEvents
  End If
End If
Next i
MousePointer = 0
SB3.Panels(2).Text = "Finished"
  Drive1(0).Drive = Left$(RoutePath, 2)
Dir1(0).Path = RoutePath
Text1(0).Text = "*.*"
Exit Sub
Errtrap:

 
 Call MsgBox(Err.Description & " occurred while copying the" _
             & vbCrLf & "file - " & File1(cursouind).List(i) _
             , vbExclamation, App.Title)
 
  
Resume Next
End Sub

Private Sub Command132_Click()
Dim strBatText As String, i As Integer, strComShape As String, Newfile3 As Integer
Dim strComFile As String, strPath As String, GSPath As String, x As Integer


If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If

On Error GoTo Errtrap

If strComPath = vbNullString Then
Call MsgBox("You do not appear to have selected your Common Path" _
            & vbCrLf & "from the Files menu." _
            , vbExclamation, App.Title)

Exit Sub
End If
'Command7.value = True
x = InStrRev(ShapePath, "Routes")
GSPath = Left$(ShapePath, x - 1) & "Global\Shapes"
DoEvents
MousePointer = 11

 cursouind = 0
 '**************** Shapes ***************
 strPath = GSPath
Drive1(0).Drive = Left$(strPath, 2)
Dir1(0).Path = strPath
Text1(0).Text = "*.*"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
strComShape = strComPath & "\GlobalShapes\"

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
  
  SB3.Panels(2).Text = File1(cursouind).List(i)
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    strComFile = strComShape & File1(cursouind).List(i)
  
  If FileExists(strComFile) Then
    If GetCRC(strComFile) = GetCRC(fullpath$) Then
 
  strBatText = strBatText & "fsutil hardlink create " & ChrW$(34) & fullpath$ & ChrW$(34) & " " & ChrW$(34) & strComFile & ChrW$(34) & vbCrLf
  
 Kill fullpath$
  DoEvents
  End If
  End If
DoEvents

   End If


   Next i
  
   Newfile3 = FreeFile

Open App.Path & "\TempFiles\do_read.bat" For Output As #Newfile3

   Print #Newfile3, strBatText
   
   Close Newfile3
  
   strDrive = Left$(App.Path, 1)
      ChDrive strDrive
ChDir App.Path & "\TempFiles"

  DoEvents

Call ShellAndWait("do_read.bat", True, vbNormalFocus)
 MousePointer = 0
SB3.Panels(2).Text = "Finished"
  
  Exit Sub
Errtrap:
  If Err = 70 Then
  MousePointer = 0
  Call MsgBox(Err.Description & " " & Err & " occurred while processing " & File1(cursouind).List(i) _
             & vbCrLf & "File is locked by another program - restart Route_Riter and try again." _
             , vbExclamation, App.Title)
  
 Exit Sub
 Else
  Call MsgBox(Err.Description & " occurred while processing the" _
             & vbCrLf & "file - " & File1(cursouind).List(i) _
             , vbExclamation, App.Title)
Resume Next
End If
End Sub

Private Sub Command133_Click()
Dim i As Integer, j As Integer, x As Integer, filepath1$, jj As Integer, strSpare As String

MousePointer = 11
strSpare = App.Path & "\Tempfiles"

jj = 0
For i = 0 To File1(0).ListCount - 1
  If File1(0).Selected(i) Then
 jj = jj + 1
 ReDim Preserve strGetAce(1 To jj)
 
    Filpath$ = File1(0).Path
   If Right$(Filpath$, 1) = "\" Then
   Filpath$ = Left$(Filpath$, Len(Filpath$) - 1)
   End If
        If Right$(filepath1$ & "\" & File1(0).List(i), 4) <> ".ace" Then
        Call MsgBox("The selected file is not an .ace file", vbExclamation, App.Title)
        Exit Sub
        End If
x = InStrRev(filepath1$ & "\" & File1(0).List(i), "\")
strGetAce(jj) = Mid$(filepath1$ & "\" & File1(0).List(i), x + 1)
End If
Next i

DoEvents
filepath1$ = ShapePath
Call DoDeCompFolder("s", ShapePath, strSpare)


Drive1(1).Drive = Left$(strSpare, 2)
Dir1(1).Path = strSpare
Text1(1).Text = "*.s"
For j = 0 To File1(1).ListCount - 1
    File1(1).Selected(j) = True
Next j

SB1.Visible = True
For j = 0 To File1(1).ListCount - 1
  If File1(1).Selected(j) Then
     DoEvents

    Call ReadShape4(strSpare & "\" & File1(1).List(j), jj)
   End If
   Next j
MousePointer = 0
DoEvents

If strReport <> vbNullString Then
 frmReport.Rich1.Text = strReport
 
     frmReport.Show 1
     
     DoEvents
End If
End Sub

Private Sub Command134_Click()
Dim strKMLFile As String, NewFile As Integer, strTemp As String, strMkr As String
Dim x As Long, Y As Long, strPlace As String, strLong As String, strLat As String
Dim yy As Long, strMarker As String, tit1 As String, zz As Integer
Dim strP As String, q As Long

strKMLFile = Label2(0).Caption
If strKMLFile = vbNullString Or (Right$(strKMLFile, 4) <> ".kml" And Right$(strKMLFile, 4) <> ".gpx") Then
Call MsgBox("You have not selected a .kml or .gpx file?", vbExclamation, App.Title)
Exit Sub
End If
NewFile = FreeFile
If Right$(strKMLFile, 4) = ".kml" Then
  Open strKMLFile For Input As #NewFile
  strTemp = Input(lOf(NewFile), #NewFile)

 Close #NewFile
 strMkr = "SIMISA@@@@@@@@@@JINX0I0t______" & vbCrLf & vbCrLf
 x = 1
 Do
 
x = InStr(x, strTemp, "<Placemark>")

If x = 0 Then GoTo EndIt
Y = InStr(x, strTemp, "name>")
yy = InStr(Y, strTemp, "<")
strPlace = Mid$(strTemp, Y + 5, yy - (Y + 5))
strPlace = Trim$(strPlace)
strPlace = Replace(strPlace, " ", "_")
Y = InStr(yy, strTemp, "<coordinates>")
yy = InStr(Y + 1, strTemp, ",")
strLong = Mid$(strTemp, Y + 13, yy - (Y + 13))
Y = yy + 1
yy = InStr(Y + 1, strTemp, ",")
q = InStr(Y + 1, strTemp, "<")
If yy = 0 Or yy > q Then
yy = q
End If
strLat = Mid$(strTemp, Y, yy - (Y))
strMarker = "Marker ( " & strLong & " " & strLat & " " & strPlace & " 2 )" & vbCrLf
strMkr = strMkr & strMarker
x = yy
EndIt:
Loop While x > 0
ElseIf Right$(strKMLFile, 4) = ".gpx" Then
  Open strKMLFile For Input As #NewFile
  strTemp = Input(lOf(NewFile), #NewFile)
 Close #NewFile
 strMkr = "SIMISA@@@@@@@@@@JINX0I0t______" & vbCrLf & vbCrLf
 x = InStr(strTemp, "<trkseg>")
zz = 1
strP = "Mark"
 Do
x = InStr(x, strTemp, "<trkpt")
If x = 0 Then GoTo EndIt2
strPlace = strP & Trim(Str(zz))
zz = zz + 1
xx = InStr(x, strTemp, "lat=")
Y = InStr(xx + 5, strTemp, ChrW$(34))
strLat = Mid(strTemp, xx + 5, Y - (xx + 5))
xx = InStr(xx, strTemp, "lon=")
Y = InStr(xx + 5, strTemp, ChrW$(34))
strLong = Mid(strTemp, xx + 5, Y - (xx + 5))
strMarker = "Marker ( " & strLong & " " & strLat & " " & strPlace & " 2 )" & vbCrLf
strMkr = strMkr & strMarker
x = Y
EndIt2:
Loop While x > 0
End If

MousePointer = 11

CDL1.InitDir = RoutePath
CDL1.Filter = "Marker Files (*.mkr)|*.mkr"
CDL1.DialogTitle = "Save Marker File"
CDL1.FilterIndex = 1
CDL1.Action = 2
tit1 = CDL1.Filename
If tit1 <> vbNullString Then
Call WriteUniFile(tit1, strMkr)
End If
DoEvents
 MousePointer = 0

End Sub

Private Sub Command135_Click()
Dim Filpath1$, strRP As String, x As Integer, strBatText As String
Dim NewFile As Integer, strTemp As String

Label12.Caption = vbNullString
Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
MousePointer = 11
'Call UncompW
Rem ******* Delete Target Folder *******
If FileExists(RoutePath & "\CleanUp.bat") Then
Kill RoutePath & "\CleanUp.bat"
End If
DoEvents
If DirExists(RoutePath & "\newRoute") Then
NewFile = FreeFile
Open RoutePath & "\CleanUp.bat" For Append As #NewFile
Print #NewFile, "If not exist c:\windows\command\deltree.exe goto tryxp"
Print #NewFile, "Deltree /y " & ChrW$(34) & "newRoute" & ChrW$(34)
Print #NewFile, "GoTo CarryOn"
Print #NewFile, ":tryxp"
Print #NewFile, "RD /s /q " & ChrW$(34) & "newRoute" & ChrW$(34)
Print #NewFile, ":CarryOn"
Print #NewFile, "echo *"
Print #NewFile, "echo *"
Print #NewFile, "echo All done - Please close any open DOS windows"
Print #NewFile, "echo *"
Close #NewFile
ChDrive (Left$(RoutePath, 1))
ChDir RoutePath
 Call ShellAndWait(ChrW$(34) & RoutePath & "\CleanUp.bat" & ChrW$(34), True, vbNormalFocus)
End If
Rem ************************************
DoEvents
'CDL1.DialogTitle = "Select Original tsection.dat"
'CDL1.Flags = cdlOFNExplorer
'CDL1.InitDir = MSTSPath & "\Global"
'CDL1.Filter = "Tsection Files (*.dat)|*.dat"
'CDL1.ShowOpen
'
'strOrigTS = CDL1.FileName

Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If
ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_cvrt.log" & ChrW$(34) & "  cvrt " & " -v99:99 " & ChrW$(34) & RoutePath & ChrW$(34)

  Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_cvrt.log"
 frmReport.Show 1
 DoEvents
 Select Case MsgBox("New Files have been prepared in the Newroute folder of your route. If the report indicated an error in preparing the route then Click NO, otherwise it should be safe to click YES to complete the move.", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

 
 strTemp = "xcopy " & ChrW$(34) & RoutePath & "\newroute\*.*" & ChrW$(34) & " " & ChrW$(34) & RoutePath & ChrW$(34) & " /S /Y"
Call ShellAndWait(strTemp, True, vbNormalFocus)
 DoEvents
  Call CheckMissFolders(RoutePath)
DoEvents
 Case vbNo
 End Select

  MousePointer = 0
 
 Label12.Caption = "Finished."

End Sub

Private Sub Command136_Click()


cursouind = 0
SparePath = App.Path & "\TempFiles"

WorldPath = File1(cursouind).Path
If Right$(WorldPath, 5) <> "World" Then
Call MsgBox("You do not appear to have selected any World tiles yet?" _
            & vbCrLf & "" _
            , vbExclamation, App.Title)

Exit Sub
End If


Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************

MousePointer = 11
Call UncompressSelectedW

GetAnother:
Close

frmGantry.Show 1

     DoEvents
     
cursouind = 0

 If strDelShape <> vbNullString Then
 
MSG = Lang(110) & strDelShape & " from all selected .W files in " & RouteName
Response = MsgBox(MSG, 36, Lang(464))
Select Case Response
Case vbYes

Call StripW2(strDelShape)

strDelShape = vbNullString
Case vbNo
strDelShape = vbNullString
End Select
End If
Select Case MsgBox("Do you wish to delete any more shapes from this Route?", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
GoTo GetAnother
    Case vbNo

End Select
MousePointer = 0
Text1(1).Text = "*.*"


Exit Sub
Errtrap:
End Sub

Private Sub Command137_Click()
Dim strT As String, strW As String, i As Integer, strA As String, strB As String
Dim strTemp As String, strKillPath As String, strBatText As String

If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
Select Case MsgBox("Do not use this option:-" _
                   & vbCrLf & "1. Unless you have run option Check Integrity first" _
                   & vbCrLf & "2. If 1. fails, run 'cvrt' to update tsection.dat" _
                   & vbCrLf & "3. If the route has a lot of sea/desert areas as these may be left blank" _
                   , vbOKCancel Or vbInformation Or vbDefaultButton1, App.Title)

    Case vbOK

    Case vbCancel
Exit Sub
End Select



MousePointer = 11
 cursouind = 0

 Drive1(0).Drive = Left$(TilePath, 2)
Dir1(0).Path = TilePath
Text1(0).Text = "*.t"
SetVariables
strKillPath = RoutePath & "\RRBackups"
strMoveTiles = strKillPath & "\Tiles"
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
End If

If Not DirExists(strKillPath) Then
MkDir strKillPath
End If
If Not DirExists(strKillPath & "\Tiles") Then
MkDir strKillPath & "\Tiles"
End If
DoEvents
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i

For i = 0 To frmUtils.File1(cursouind).ListCount - 1

   If frmUtils.File1(cursouind).Selected(i) Then
   FileCopy File1(cursouind).Path & "\" & File1(cursouind).List(i), strKillPath & "\Tiles\" & File1(cursouind).List(i)
   DoEvents
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   DoEvents
   End If
   strTemp = Left(File1(cursouind).List(i), Len(File1(cursouind).List(i)) - 2)
If FileExists(File1(cursouind).Path & "\" & strTemp & "_e.raw") Then
FileCopy File1(cursouind).Path & "\" & strTemp & "_e.raw", strKillPath & "\Tiles\" & strTemp & "_e.raw"
DoEvents
Kill File1(cursouind).Path & "\" & strTemp & "_e.raw"
DoEvents
End If
If FileExists(File1(cursouind).Path & "\" & strTemp & "_n.raw") Then
FileCopy File1(cursouind).Path & "\" & strTemp & "_n.raw", strKillPath & "\Tiles\" & strTemp & "_n.raw"
DoEvents
Kill File1(cursouind).Path & "\" & strTemp & "_n.raw"
DoEvents
End If
If FileExists(File1(cursouind).Path & "\" & strTemp & "_y.raw") Then
FileCopy File1(cursouind).Path & "\" & strTemp & "_y.raw", strKillPath & "\Tiles\" & strTemp & "_y.raw"
DoEvents
Kill File1(cursouind).Path & "\" & strTemp & "_y.raw"
DoEvents
End If
If FileExists(File1(cursouind).Path & "\" & strTemp & "_f.raw") Then
FileCopy File1(cursouind).Path & "\" & strTemp & "_f.raw", strKillPath & "\Tiles\" & strTemp & "_f.raw"
DoEvents
Kill File1(cursouind).Path & "\" & strTemp & "_f.raw"
DoEvents
End If
Next i
DoEvents



 Drive1(0).Drive = Left$(WorldPath, 2)
Dir1(0).Path = WorldPath
Text1(0).Text = "*.w"
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i

For i = 0 To frmUtils.File1(cursouind).ListCount - 1

   If frmUtils.File1(cursouind).Selected(i) Then
   
   strW = File1(cursouind).List(i)
   If Len(strW) > 17 Then GoTo NextOne
   strA = Mid(strW, 2, 7)
   strB = Mid(strW, 9, 7)
Call TileName2(Val(strA), Val(strB), strT)

If FileExists(strKillPath & "\tiles\" & strT & ".t") Then
FileCopy strKillPath & "\tiles\" & strT & ".t", RoutePath & "\Tiles\" & strT & ".t"
End If
If FileExists(strKillPath & "\tiles\" & strT & "_e.raw") Then
FileCopy strKillPath & "\tiles\" & strT & "_e.raw", RoutePath & "\Tiles\" & strT & "_e.raw"
End If
If FileExists(strKillPath & "\tiles\" & strT & "_n.raw") Then
FileCopy strKillPath & "\tiles\" & strT & "_n.raw", RoutePath & "\Tiles\" & strT & "_n.raw"
End If
If FileExists(strKillPath & "\tiles\" & strT & "_y.raw") Then
FileCopy strKillPath & "\tiles\" & strT & "_y.raw", RoutePath & "\Tiles\" & strT & "_y.raw"
End If
If FileExists(strKillPath & "\tiles\" & strT & "_f.raw") Then
FileCopy strKillPath & "\tiles\" & strT & "_f.raw", RoutePath & "\Tiles\" & strT & "_f.raw"
End If

NextOne:
    End If
Next i
DoEvents
 strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_filter.log" & ChrW$(34) & "  filter " & ChrW$(34) & RoutePath & ChrW$(34)


ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
  

 Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 If Not FileExists(App.Path & "\Reports\" & RouteName & "_filter.log") Then
Call MsgBox("As no .log file has been created, you can assume that TsUtils" _
            & vbCrLf & "did not complete the option you were attempting to run." _
            , vbCritical, App.Title)

MousePointer = 0
Exit Sub
End If
 
 
  MousePointer = 0
 frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_filter.log"
 frmReport.Show 1
 DoEvents
 Call MsgBox("A new 'TD' folder has been placed within the 'newRoute' folder in your route. Back up your TD folder and copy the new TD folder contents into it.", vbInformation, App.Title)
 
End Sub

Private Sub Command138_Click()
Dim Filpath1$, strRP As String, x As Integer, strBatText As String


Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
MousePointer = 11
If Not DirExists(App.Path & "\Reports") Then
MkDir App.Path & "\Reports"
End If

Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If

 strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_srchsig.log" & ChrW$(34) & "  srchsig " & ChrW$(34) & RoutePath & ChrW$(34)


ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
  

 Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 If Not FileExists(App.Path & "\Reports\" & RouteName & "_srchsig.log") Then
 Call MsgBox("As no .log file has been created, you can assume that TsUtils" _
            & vbCrLf & "did not complete the option you were attempting to run." _
            , vbCritical, App.Title)



MousePointer = 0
Exit Sub
Else
frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_srchsig.log"
frmReport.Rich1.Text = frmReport.Rich1.Text & vbCrLf & vbCrLf & "Locations of signals have been placed in " & RoutePath & "\Newroute" & vbCrLf
 frmReport.Show 1
 DoEvents

End If
 
 
  MousePointer = 0
 

End Sub

Private Sub Command140_Click()
Dim SparePath As String

cursouind = 0
SparePath = App.Path & "\TempFiles"

WorldPath = File1(cursouind).Path
If Right$(WorldPath, 5) <> "World" Then
Call MsgBox("You do not appear to have selected any World tiles yet?" _
            & vbCrLf & "" _
            , vbExclamation, App.Title)

Exit Sub
End If
Select Case MsgBox("This option will delete all ViewDbSphere entries from selected .W files" _
                   & vbCrLf & "It will also change all VDbId entries to VDbId ( 4294967294 ) - Do you wish to continue?" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

    Case vbNo
MousePointer = 0
Exit Sub
End Select

Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************

MousePointer = 11
Call UncompressSelectedW

Call FixVDbId



MousePointer = 0
Text1(1).Text = "*.*"
End Sub

Private Sub Command141_Click()
Dim SparePath As String, MSG As String, strResult As String

cursouind = 0
SparePath = App.Path & "\TempFiles"

WorldPath = File1(cursouind).Path
If Right$(WorldPath, 5) <> "World" Then
Call MsgBox("You do not appear to have selected any World tiles yet?" _
            & vbCrLf & "" _
            , vbExclamation, App.Title)

Exit Sub
End If


Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************

MousePointer = 11
Call UncompressSelectedW

GetAnother:
Close
strResult = InputBox("Enter altitude adjustment for Shapes in metres. (Negative to lower)", "Adjust Shape's Altitude")
'frmDelShape.Show 1

     DoEvents
     
cursouind = 0


 
MSG = "Adjust altitude of all Shapes in selected .w files by " & strResult
Response = MsgBox(MSG, 36, Lang(464))
Select Case Response
Case vbYes

Call LiftW(strResult)


Case vbNo
MousePointer = 0
Text1(1).Text = "*.*"
Exit Sub
End Select


MousePointer = 0
Text1(1).Text = "*.*"
End Sub

Private Sub Command142_Click()
Dim strFirst As String, TrainsetPath As String
MousePointer = 11
TrainsetPath = MSTSPath & "\Trains\Trainset\"
strReport = vbNullString

strFirst = Dir1(0).Path
'Call CountStock
DoEvents
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = TrainsetPath
Text1(0) = "*.wag"
frmUtils.Refresh
DoEvents


booFixEng = True
frmSearch.Show
DoEvents
Dir1(0).Path = strFirst
booFixEng = False
MousePointer = 0
End Sub


Private Sub Command144_Click()
Dim strMPath As String

On Error GoTo Errtrap
MousePointer = 11
strMPath = MSTSPath
Text1(0) = "*.*"
Drive1(0).Drive = Left$(strMPath, 2)
Dir1(0).Path = strMPath & "\Trains\Trainset"
frmConEdit.Show
MousePointer = 0
Exit Sub
Errtrap:
Call MsgBox(Err.Description & " occurred in Command144_click", vbExclamation, App.Title)

End Sub

Private Sub Command145_Click()
Dim strBatText As String

strBatText = "java -Xmx512m TSUtilDlg"
ChDrive Left$(App.Path, 1)
'ChDir App.Path & "\TSUtil"
Call ShellAndWait(strBatText, True, vbHide)
DoEvents
End Sub

Private Sub Command27_Click()
Dim Filpath1$, fullpath$, flagway As Integer
Dim strOrigPath As String, strErrorBias As String


On Error GoTo Errtrap
MousePointer = 11
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If
Text1(0) = "*.trk"
If File1(cursouind).ListCount = 0 Then
Call MsgBox("There does not appear to be a valid .trk entry for this route?", vbExclamation, App.Title)
MousePointer = 0
Exit Sub
End If
'strTrk = File1(cursouind).path & "\" & File1(cursouind).list(0)


 strOrigPath = File1(0).Path & "\Tiles"

Filpath1$ = App.Path & "\TempFiles"
Call KillSpare("*.t")
strErrorBias = InputBox("Enter Bias Value ( 0, 1 or 2 )")
If strErrorBias = vbNullString Then
MousePointer = 0
Exit Sub
End If


TokMode = 2

Call UncompressTFiles
cursouind = 1
Drive1(cursouind).Drive = Left$(Filpath1$, 2)
Dir1(cursouind).Path = Filpath1$
Text1(cursouind) = "*.t"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
Label9.Visible = True

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
     Label9.Caption = "Processing:  " & File1(cursouind).List(i)
flagway = 0
Call ConvertT(fullpath$, strErrorBias, flagway)
DoEvents
flagway = 1
Call ConvertT(fullpath$, strErrorBias, flagway)
DoEvents

Kill strOrigPath & "\" & File1(cursouind).List(i)
DoEvents

Call DoDeComp2(File1(cursouind).List(i), File1(cursouind).Path, strOrigPath)
DoEvents
Kill fullpath$
DoEvents
End If
Next i
MousePointer = 0
Label9.Caption = "Processing Finished"

Exit Sub
Errtrap:
Call MsgBox("An error occurred while processing " & File1(cursouind).List(i), vbExclamation, App.Title)

End Sub

Private Sub Command41_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Select Case Button
Case 1
Text1(1) = Text1(0)
Drive1(1).Drive = Drive1(0).Drive
Dir1(1).Path = Dir1(0).Path
Case 2
Dim strD1 As String, strD0 As String, strT1 As String, strT0 As String
Dim strP1 As String, strP0 As String

strD1 = Drive1(0).Drive
strD0 = Drive1(1).Drive
strT1 = Text1(0)
strT0 = Text1(1)
strP1 = Dir1(0).Path
strP0 = Dir1(1).Path
Dir1(0).Path = strP0
Dir1(1).Path = strP1
Drive1(0).Drive = strD0
Drive1(1).Drive = strD1
Text1(0) = strT0
Text1(1) = strT1


End Select
End Sub

Private Sub Command57_Click()
Dim Filpath1$, strRP As String, x As Integer, strBatText As String
Dim strTemp As String
Dim NewFile As Integer, strResult As String, DH As Single, WorldPath As String, DX As Single, DZ As Single


On Error GoTo Errtrap
Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If

If FileExists(RoutePath & "\CleanUp.bat") Then
Kill RoutePath & "\CleanUp.bat"
End If
If FileExists(RoutePath & "\CleanUpNew.bat") Then
Kill RoutePath & "\CleanUpNew.bat"
End If

MousePointer = 11

TryAgain:
strResult = InputBox("Enter height adjustment for Track in cm.", "Adjust Track Height")


DH = Val(strResult)
If DH = 0 Then Exit Sub
DH = DH / 100

Rem ******** Uncompress files ***********
WorldPath = RoutePath & "\World"

Call DoDeCompFolder("w", WorldPath, WorldPath)
DoEvents

DoEvents
Rem ******* Delete Target Folder *******

If DirExists(RoutePath & "\newRoute") Then
NewFile = FreeFile
Open RoutePath & "\CleanUp.bat" For Append As #NewFile
Print #NewFile, "If not exist c:\windows\command\deltree.exe goto tryxp"
Print #NewFile, "Deltree /y " & ChrW$(34) & "newRoute" & ChrW$(34)
Print #NewFile, "GoTo CarryOn"
Print #NewFile, ":tryxp"
Print #NewFile, "RD /s /q " & ChrW$(34) & "newRoute" & ChrW$(34)
Print #NewFile, ":CarryOn"
Print #NewFile, "echo *"
Print #NewFile, "echo *"
Print #NewFile, "echo All done - Please close any open DOS windows"
Print #NewFile, "echo *"
Close #NewFile
ChDrive (Left$(RoutePath, 1))
ChDir RoutePath
 Call ShellAndWait(ChrW$(34) & RoutePath & "\CleanUp.bat" & ChrW$(34), True, vbNormalFocus)
End If
Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
'RoutePath = File1(cursouind).path
'frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If
ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_Mveobj.log" & ChrW$(34) & " mveobj -t -r " & ChrW$(34) & RoutePath & ChrW$(34) & " " & DX & " " & DH & " " & DZ

  Call ShellAndWait(strBatText, True, vbNormalFocus)
 MousePointer = 0
 DoEvents
frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_Mveobj.log"
 frmReport.Show 1
 DoEvents
Select Case MsgBox("New Files have been prepared in the Newroute folder of your route. If the report indicated an error in preparing the route then Click NO, otherwise it should be safe to click YES to complete the move.", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

strTemp = "xcopy " & ChrW$(34) & RoutePath & "\newroute\*.*" & ChrW$(34) & " " & ChrW$(34) & RoutePath & ChrW$(34) & " /S /Y"
Call ShellAndWait(strTemp, True, vbNormalFocus)
 DoEvents
  Call CheckMissFolders(RoutePath)
DoEvents
 Rem ********************************

    Case vbNo

 End Select

 
 Label12.Caption = "Compressing .W tiles."
 Call DoCompFolder("w", RoutePath & "\World", RoutePath & "\World")
 
 
Label12.Caption = "Finished."

  MousePointer = 0
 
 Exit Sub
Errtrap:

Call MsgBox("An error " & Err & " occurred in subroutine 'Change Track Height' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
 Resume Next
 
 
End Sub

Private Sub Command61_Click()
flagPrint = 15
fEZPrint.Show
DoEvents
End Sub


Private Sub Command62_Click()


If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
Filpath1$ = App.Path & "\TempFiles"
strBackupPath = GetSetting(App.Title, "Backup", "Path", strBackupPath)
booBackup = True

frmBackup.Show

End Sub

Private Sub Command63_Click()
Dim flagway As Integer

strKillFiles = vbNullString

On Error GoTo Errtrap
ReDim Cars(0 To Car_CHUNK)
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If
Call ClearSetup
If flagNoRef = False Then
Call MsgBox("A .ref file already exists for this route, this option will now abort.", vbExclamation, App.Title)
Exit Sub
End If
FileCopy App.Path & "\stuffit.ref", App.Path & "\setupfiles\master.ref"
DoEvents
If FileExists(RoutePath & "\telepole.dat") Then
  Call CompactCheckForS(RoutePath & "\telepole.dat", booHaz)
 End If
If FileExists(RoutePath & "\speedpost.dat") Then
 Call CompactCheckForAce(RoutePath & "\speedpost.dat")
 Call CompactCheckForS(RoutePath & "\speedpost.dat", booHaz)
 
 End If
 
 If FileExists(RoutePath & "\sigcfg.dat") Then
 Call CompactCheckForAce(RoutePath & "\sigcfg.dat")
 
 
 End If

Call CompactRoute
Call CompactRef
If strReport <> vbNullString Then
strReport = "Unable to complete .REF file because of the following faulty .S files" & vbCrLf & "Fix with the 'Fix Bad .S File Format' option, then run Make .REF again" & vbCrLf & vbCrLf & strReport
 frmReport.Rich1.Text = strReport
 
     frmReport.Show 1
MousePointer = 0
Exit Sub
End If


flagway = 1
Call ConvertIt(App.Path & "\setupfiles\tempref.ref", flagway)
FileCopy App.Path & "\setupfiles\tempref.ref", OriginalRef

SB1.Panels(2) = "REF File Completed"

DoEvents
MousePointer = 0

Exit Sub
Errtrap:


End Sub



Private Sub Command64_Click()
Dim Filpath1$, strRP As String, x As Integer, strBatText As String


Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
MousePointer = 11
'Call UncompW

Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If
ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
 
   'strBatText = "java -Xmx256m TSUtil ichk " & chrw$(34) & RoutePath & chrw$(34) & " " & chrw$(34) & App.path & "\Reports\" & RouteName & ".log" & chrw$(34)
strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & ".log" & ChrW$(34) & " ichk -S " & ChrW$(34) & RoutePath & ChrW$(34)

  Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 
If Not FileExists(App.Path & "\Reports\" & RouteName & ".log") Then
Call MsgBox("As no .log file has been created, you can assume that TsUtils" _
            & vbCrLf & "did not complete the option you were attempting to run." _
            , vbCritical, App.Title)

MousePointer = 0
Exit Sub
End If
  MousePointer = 0
 frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & ".log"
 frmReport.Show 1
 Label12.Caption = "Finished."

End Sub


Private Sub Command65_Click(Index As Integer)
Dim booExists As Boolean, NewRouteName As String, Filpath1$, strTemp As String


cursouind = 0
Filpath1$ = App.Path & "\setupfiles"
If FileExists(App.Path & "\setupfiles\master.ref") Then
Kill App.Path & "\setupfiles\master.ref"
End If
If FileExists(App.Path & "\setupfiles\InstallMe.bat") Then
Kill App.Path & "\setupfiles\InstallMe.bat"
End If
Open App.Path & "\SetupFiles\Installme.bat" For Append As #12
Print #12, "@Echo Off"
Close #12
MousePointer = 11

RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath

x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If



Text1(0) = "*.trk"
If File1(cursouind).ListCount = 0 Then
Text1(0) = "*.tdb"
DoEvents
If File1(cursouind).ListCount = 0 Then
Call MsgBox(Lang(431) & vbCrLf & Lang(432), vbExclamation, Lang(407))
Text1(0) = "*.*"
MousePointer = 0
Exit Sub
Else
Call MsgBox(Lang(635) & vbCrLf & Lang(636), vbExclamation, App.Title)

Text1(0) = "*.*"
MousePointer = 0
Exit Sub
End If
End If
RouteName = File1(cursouind).List(i)
OldRouteName = RouteName

Call CheckForSMS(RoutePath & "\" & RouteName)

Call FindRouteName(RoutePath & "\" & RouteName, NewRouteName, booExists)
If booExists = False Then

'MsgBox Lang(339) & vbcr & Lang(340), 16, Lang(341)
MousePointer = 0

Exit Sub
Else
RouteName = NewRouteName
End If

Call FindRouteID(RoutePath & "\" & OldRouteName)
Call IsItElectric(RoutePath & "\" & OldRouteName, booElectric)

'End If
OriginalRef = RoutePath & "\" & RouteName & ".ref"
If FileExists(OriginalRef) Then
FileCopy RoutePath & "\" & RouteName & ".ref", App.Path & "\setupfiles\master.ref"
Else
flagNoRef = True
Call MsgBox(Lang(385) & vbCrLf & Lang(386), vbExclamation, App.Title)
FileCopy App.Path & "\stuffit.ref", App.Path & "\setupfiles\master.ref"
End If


RouteListed = True
TexturePath = RoutePath & "\Textures"
If Not DirExists(TexturePath) Then MkDir TexturePath
TexSnowPath = RoutePath & "\Textures\Snow"
If Not DirExists(TexSnowPath) Then MkDir TexSnowPath
TexNightPath = RoutePath & "\Textures\Night"
If Not DirExists(TexNightPath) Then MkDir TexNightPath
TexAutPath = RoutePath & "\Textures\Autumn"
If Not DirExists(TexAutPath) Then MkDir TexAutPath
TexAutSnowPath = RoutePath & "\Textures\AutumnSnow"
If Not DirExists(TexAutSnowPath) Then MkDir TexAutSnowPath
TexSprPath = RoutePath & "\Textures\Spring"
If Not DirExists(TexSprPath) Then MkDir TexSprPath
TexSprSnowPath = RoutePath & "\Textures\SpringSnow"
If Not DirExists(TexSprSnowPath) Then MkDir TexSprSnowPath
TexWinPath = RoutePath & "\Textures\Winter"
If Not DirExists(TexWinPath) Then MkDir TexWinPath
TexWinSnowPath = RoutePath & "\Textures\WinterSnow"
If Not DirExists(TexWinSnowPath) Then MkDir TexWinSnowPath
TilePath = RoutePath & "\Tiles"
ShapePath = RoutePath & "\Shapes"
SoundPath = RoutePath & "\Sound"
WorldPath = RoutePath & "\World"
EnvPath = RoutePath & "\Envfiles"
If Not DirExists(EnvPath) Then MkDir EnvPath
If Not DirExists(EnvPath & "\Textures") Then MkDir EnvPath & "\Textures"

If Not FileExists(RoutePath & "\" & RouteName & ".tdb") Then
strTemp = strTemp & vbCrLf & RoutePath & "\" & RouteName & ".tdb" & Lang(578)
End If
If Not FileExists(RoutePath & "\" & RouteName & ".tit") Then
strTemp = strTemp & vbCrLf & RoutePath & "\" & RouteName & ".tit" & Lang(578)
End If
If Not FileExists(RoutePath & "\tsection.dat") Then
strTemp = strTemp & vbCrLf & RoutePath & "\tsection.dat" & Lang(578) & Lang(580)
End If
If Not FileExists(TexturePath & "\acleantrack1.ace") Then
strTemp = strTemp & vbCrLf & TexturePath & "\acleantrack1.ace" & Lang(578) & Lang(581)
End If
If Not FileExists(TexturePath & "\acleantrack2.ace") Then
strTemp = strTemp & vbCrLf & TexturePath & "\acleantrack2.ace" & Lang(578) & Lang(581)
End If
If Not FileExists(TexSnowPath & "\acleantrack1.ace") Then
strTemp = strTemp & vbCrLf & TexSnowPath & "\acleantrack1.ace" & Lang(578) & Lang(582)
End If
If Not FileExists(TexSnowPath & "\acleantrack2.ace") Then
strTemp = strTemp & vbCrLf & TexSnowPath & "\acleantrack2.ace" & Lang(578) & Lang(582)
End If
If strTemp <> vbNullString Then
strTemp = Lang(583) & vbCrLf & vbCrLf & strTemp
 frmReport.Rich1.Text = strTemp
 
     frmReport.Show 1
     
     DoEvents


End If

Text1(0) = "*.*"
Rem ********** Delete any spare .w files ****

    cursouind = 1
    SparePath = App.Path & "\TempFiles"
    
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents

 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************
MousePointer = 0

End Sub


Private Sub Command66_Click()
Dim Filpath1$, strRP As String, x As Integer, strBatText As String
Dim strTemp As String
Dim NewFile As Integer, strResult As String, DX As Integer, DY As Integer


On Error GoTo Errtrap
Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If

If FileExists(RoutePath & "\CleanUp.bat") Then
Kill RoutePath & "\CleanUp.bat"
End If
If FileExists(RoutePath & "\CleanUpNew.bat") Then
Kill RoutePath & "\CleanUpNew.bat"
End If

MousePointer = 11

TryAgain:
strResult = InputBox("Enter number of squares to move route in east/west direction and north/south direction separated by space (West and South enter as negative)", "Move Route")

x = InStr(strResult, " ")
If x = 0 Then MousePointer = 0: Exit Sub
DX = Val(Left$(strResult, x - 1))
DY = Val(Mid$(strResult, x + 1))
If DX = 0 And DY = 0 Then GoTo Label2


Label2:
Rem ******** Uncompress files ***********



DoEvents

Rem ******* Delete Target Folder *******
If DirExists(RoutePath & "\newRoute") Then
NewFile = FreeFile
Open RoutePath & "\CleanUp.bat" For Append As #NewFile
Print #NewFile, "If not exist c:\windows\command\deltree.exe goto tryxp"
Print #NewFile, "Deltree /y " & ChrW$(34) & "newRoute" & ChrW$(34)
Print #NewFile, "GoTo CarryOn"
Print #NewFile, ":tryxp"
Print #NewFile, "RD /s /q " & ChrW$(34) & "newRoute" & ChrW$(34)
Print #NewFile, ":CarryOn"
Print #NewFile, "echo *"
Print #NewFile, "echo *"
Print #NewFile, "echo All done - Please close any open DOS windows"
Print #NewFile, "echo *"
Close #NewFile
ChDrive (Left$(RoutePath, 1))
ChDir RoutePath
 Call ShellAndWait(ChrW$(34) & RoutePath & "\CleanUp.bat" & ChrW$(34), True, vbNormalFocus)
End If
Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
'RoutePath = File1(cursouind).path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If
ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java -Xmx256m TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_Move.log" & ChrW$(34) & "  move -p " & ChrW$(34) & RoutePath & ChrW$(34) & " " & DX & " " & DY
  Call ShellAndWait(strBatText, True, vbNormalFocus)
 MousePointer = 0
 DoEvents
 Rem ******* Run bat to remove old files and copy new ones
 If Not FileExists(App.Path & "\Reports\" & RouteName & "_Move.log") Then
 
 Exit Sub
 End If
 frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_Move.log"
 frmReport.Show 1
 DoEvents
  NewFile = FreeFile
Open RoutePath & "\CleanUpNew.bat" For Append As #NewFile
Print #NewFile, "If not exist c:\windows\command\deltree.exe goto tryxp"
Print #NewFile, "Deltree /y td"
Print #NewFile, "Deltree /y world"
Print #NewFile, "Deltree /y tiles"
Print #NewFile, "Deltree /y lo_tiles"
Print #NewFile, "Deltree /y activities"
Print #NewFile, "Deltree /y paths"
Print #NewFile, "Deltree /y Services"
Print #NewFile, "Deltree /y Traffic"
Print #NewFile, "GoTo CarryOn"
Print #NewFile, ":tryxp"
Print #NewFile, "RD /s /q td"
Print #NewFile, "RD /s /q world"
Print #NewFile, "RD /s /q tiles"
Print #NewFile, "RD /s /q lo_Tiles"
Print #NewFile, "RD /s /q activities"
Print #NewFile, "RD /s /q paths"
Print #NewFile, "RD /s /q services"
Print #NewFile, "RD /s /q traffic"
Print #NewFile, ":CarryOn"
Print #NewFile, "echo *"
Print #NewFile, "echo *"
Print #NewFile, "echo All done - Please close any open DOS windows"
Print #NewFile, "echo *"
Close #NewFile
DoEvents
 Select Case MsgBox("New Files have been prepared in the Newroute folder of your route. If the report indicated an error in preparing the route then Click NO, otherwise it should be safe to click YES to complete the move.", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

ChDrive (Left$(RoutePath, 1))
ChDir RoutePath
 Call ShellAndWait(ChrW$(34) & RoutePath & "\CleanUpNew.bat" & ChrW$(34), True, vbNormalFocus)
DoEvents
strTemp = "xcopy " & ChrW$(34) & RoutePath & "\newroute\*.*" & ChrW$(34) & " " & ChrW$(34) & RoutePath & ChrW$(34) & " /s /y"
Call ShellAndWait(strTemp, True, vbNormalFocus)
 DoEvents
 
 Rem ********************************
 Call CheckMissFolders(RoutePath)
DoEvents
    Case vbNo

 End Select
 

 

Label12.Caption = "Finished."

  MousePointer = 0
 
 Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'Move a Route' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
     
 
End Sub
Private Sub Command67_Click()
Dim Filpath1$, strRP As String, x As Integer, strBatText As String
Dim booComp As Boolean

Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
MousePointer = 11
'Call UncompW

Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If
ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_tdb.log" & ChrW$(34) & "  rendb -a -w -m " & ChrW$(34) & RoutePath & ChrW$(34)

  Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 
If Not FileExists(App.Path & "\Reports\" & RouteName & "_tdb.log") Then
Call MsgBox("As no .log file has been created, you can assume that TsUtils" _
            & vbCrLf & "did not complete the option you were attempting to run." _
            , vbCritical, App.Title)

MousePointer = 0
Exit Sub
End If
Call MsgBox("Your route RENDB operation appears to have been successful. The new files are in the NewRoute folder within your first route. You may now replace the corresponding folders  and files in your original route with the equivalent folders and files from within NewRoute (but make a backup first).", vbExclamation, App.Title)
 
  MousePointer = 0
 frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_tdb.log"
 frmReport.Show 1
'
 If booComp = True Then
 booComp = False
 
 Rem *** compress .W files
 Call CompressWFiles
 
 End If
End Sub

Private Sub Command68_Click()
Dim Filpath1$, strRP As String, x As Integer, strBatText As String
Dim booComp As Boolean

Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
MousePointer = 11


Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If
ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java TSUtil chgdb " & ChrW$(34) & RoutePath & "\" & RouteName & ".tdb" & ChrW$(34) & " 40000 40000"

 Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 
 
  MousePointer = 0

 If booComp = True Then
 booComp = False
 
 Rem *** compress .W files
 Call CompressWFiles
 
 End If
End Sub


Private Sub Command69_Click()
frmTsUtil.Show

End Sub

Private Sub Command70_Click()
Dim q As Integer, strTempRoute As String

On Error GoTo Errtrap
Text1(0).Text = "*.*"

strTempRoute = RoutePath
For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then
If frmUtils.Controls(q).Caption <> Lang(30) And frmUtils.Controls(q).Caption <> Lang(53) Then
frmUtils.Controls(q).Enabled = False
ElseIf frmUtils.Controls(q).Caption = Lang(30) Then

frmUtils.Controls(q).Caption = Lang(637)
End If
End If
Next q
ReDim LocoPath(0 To CHUNK), LocoName(0 To CHUNK)
ReDim Wagpath(0 To CHUNK), WagonName(0 To CHUNK)
ReDim Locomotives(0 To CHUNK)
ReDim Wagons(0 To CHUNK)
ReDim Service(0 To CHUNK)
ReDim SrvPath(0 To CHUNK)
ReDim Activities(0 To CHUNK), ActPath(0 To CHUNK)
ReDim Traffic(0 To CHUNK), TfcPath(0 To CHUNK)
ReDim LocoCoup(0 To CHUNK), LocoFCoup(0 To CHUNK), LocoBrake(0 To CHUNK), LocoType(0 To CHUNK), LocoRigid(0 To CHUNK)
ReDim WagCoup(0 To CHUNK), WagFCoup(0 To CHUNK), WagBrake(0 To CHUNK), WagType(0 To CHUNK), WagRigid(0 To CHUNK)
ReDim Paths(0 To CHUNK), PathsPath(0 To CHUNK)
ReDim PathUsed(0 To CHUNK)



strbadbits = vbNullString
strReport = vbNullString
strForPrint = vbNullString
Set frmStock = Nothing
booStockOnly = False

ConEngNumber = 0: ConWagNumber = 0
lngAct = 0
lngSrv = 0
lngCon = 0
lngLoco = 0
lngWagons = 0
lngTfc = 0
lngPaths = 0
PathUsedNumb = 0

ActChecked = True
booActsChecked = True
For i = 0 To 5
Label7(i).Caption = vbNullString
Next i
Label3.Caption = vbNullString
SB2.Panels(2).Text = "Counting Rolling-stock"
Call CountStock

If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0

Exit Sub
End If

SB2.Panels(2).Text = "Counting Activities"
Call GetActivities2
If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0
Exit Sub
End If

Label7(2).Caption = Str(lngAct)
ReDim Preserve PathUsed(0 To PathUsedNumb)
SB2.Panels(2).Text = "Counting Services"
Call GetServices2
If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0
Exit Sub
End If

DoEvents
Label7(3).Caption = Str(lngSrv)
SB2.Panels(2).Text = "Counting Consists"
Call GetConsists
If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0
Exit Sub
End If

Label7(4).Caption = Str(lngCon)

SB2.Panels(2).Text = "Counting Traffic"

Call GetTraffic2
Label7(5).Caption = Str(lngTfc)

If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0
Exit Sub
End If


SB2.Panels(2).Text = "Counting Paths"

Call GetPaths2
If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0
Exit Sub
End If
Label3.Caption = lngPaths


DoEvents
booNoButtons = False

BooCheckAct = False
For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then frmUtils.Controls(q).Enabled = True
Next q
Command1(15).Caption = Lang(30)
DoEvents
SB2.Panels(2).Text = "Populating Grid"
frmUtils.Refresh
DoEvents
frmGrid.Show

     DoEvents
     
Command15.Visible = True

Dir1(0).Path = strTempRoute
Text1(0).Text = "*.s"
DoEvents
Text1(cursouind).Text = "*.*"
SB2.Panels(2).Text = vbNullString
strForPrint = vbNullString
strReport = vbNullString
strbadbits = vbNullString
Exit Sub
Errtrap:


Call MsgBox("An error " & Err & " occurred in subroutine 'CheckActivities' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub

Private Sub Command71_Click()
Dim Filpath1$, strRP As String, x As Integer, strBatText As String
Dim strTemp As String
Dim NewFile As Integer, strResult As String, DH As Integer, TilePath As String

On Error GoTo Errtrap
Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If

If FileExists(RoutePath & "\CleanUp.bat") Then
Kill RoutePath & "\CleanUp.bat"
End If
If FileExists(RoutePath & "\CleanUpNew.bat") Then
Kill RoutePath & "\CleanUpNew.bat"
End If

MousePointer = 11

TryAgain:
strResult = InputBox("Enter altitude adjustment for Route in metres.", "Adjust Route's Altitude")


DH = Val(strResult)
If DH = 0 Then Exit Sub
Rem ******** Uncompress files ***********
TilePath = RoutePath & "\Tiles"
'Call UncompT
Call DoDeCompFolder("t", TilePath, TilePath)
DoEvents
TilePath = RoutePath & "\L0_Tiles"
'Call UncompT
If DirExists(TilePath) Then
Call DoDeCompFolder("t", TilePath, TilePath)
End If
'Call UncompW
DoEvents
Rem ******* Delete Target Folder *******

If DirExists(RoutePath & "\newRoute") Then
NewFile = FreeFile
Open RoutePath & "\CleanUp.bat" For Append As #NewFile
Print #NewFile, "If not exist c:\windows\command\deltree.exe goto tryxp"
Print #NewFile, "Deltree /y " & ChrW$(34) & "newRoute" & ChrW$(34)
Print #NewFile, "GoTo CarryOn"
Print #NewFile, ":tryxp"
Print #NewFile, "RD /s /q " & ChrW$(34) & "newRoute" & ChrW$(34)
Print #NewFile, ":CarryOn"
Print #NewFile, "echo *"
Print #NewFile, "echo *"
Print #NewFile, "echo All done - Please close any open DOS windows"
Print #NewFile, "echo *"
Close #NewFile
ChDrive (Left$(RoutePath, 1))
ChDir RoutePath
 Call ShellAndWait(ChrW$(34) & RoutePath & "\CleanUp.bat" & ChrW$(34), True, vbNormalFocus)
End If
Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
'RoutePath = File1(cursouind).path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If
ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_Adjust.log" & ChrW$(34) & "  adjh " & ChrW$(34) & RoutePath & ChrW$(34) & " " & DH

  Call ShellAndWait(strBatText, True, vbNormalFocus)
 MousePointer = 0
 DoEvents
frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_Adjust.log"
 frmReport.Show 1
 DoEvents
Select Case MsgBox("New Files have been prepared in the Newroute folder of your route. If the report indicated an error in preparing the route then Click NO, otherwise it should be safe to click YES to complete the move.", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

strTemp = "xcopy " & ChrW$(34) & RoutePath & "\newroute\*.*" & ChrW$(34) & " " & ChrW$(34) & RoutePath & ChrW$(34) & " /S /Y"
Call ShellAndWait(strTemp, True, vbNormalFocus)
 DoEvents
  Call CheckMissFolders(RoutePath)
DoEvents
 Rem ********************************

    Case vbNo

 End Select
 
' If booComp = True Then
' booComp = False
' Label12.Caption = "Compressing .W files"
' Call CompressWFiles
' DoEvents
' End If
 
 Label12.Caption = "Compressing .T tiles."
 Call DoCompFolder("t", RoutePath & "\tiles", RoutePath & "\tiles")
 
 DoEvents
 Label12.Caption = "Compressing Lo .T tiles."
 If DirExists(RouteName & "\lo_tiles") Then
 Call DoCompFolder("t", RoutePath & "\lo_tiles", RoutePath & "\lo_tiles")
 End If
Label12.Caption = "Finished."

  MousePointer = 0
 
 Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'Change Route altitude' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
 Resume Next
 
 
End Sub

Private Sub Command72_Click()
Dim Filpath1$, x As Integer, strBatText As String
Dim strOrigTS As String

Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
MousePointer = 11
'Call UncompW(booComp)
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
RoutePath = File1(cursouind).Path
'frmUtils.caption = "Path=" & mstspath & "  Route=" & routepath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)

End If
Rem ************************************
DoEvents
CDL1.DialogTitle = "Select Original tsection.dat"
CDL1.Flags = cdlOFNExplorer
CDL1.InitDir = MSTSPath & "\Global"
CDL1.Filter = "Tsection Files (*.dat)|*.dat"
CDL1.ShowOpen

strOrigTS = CDL1.Filename



ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_chkup.log" & ChrW$(34) & "  chkup -v99:99 " & ChrW$(34) & strOrigTS & ChrW$(34) & " " & ChrW$(34)

 Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 If Not FileExists(App.Path & "\Reports\" & RouteName & "_Chkup.log") Then
Call MsgBox("As no .log file has been created, you can assume that TsUtils" _
            & vbCrLf & "did not complete the option you were attempting to run." _
            , vbCritical, App.Title)

MousePointer = 0
Exit Sub
End If
 

 
  MousePointer = 0
 frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_Chkup.log"
 frmReport.Show 1
 Label12.Caption = "Finished."

End Sub

Private Sub Command73_Click()
Dim Filpath1$, strRP As String, x As Integer, strBatText As String
Dim strOrigTS As String
Dim NewFile As Integer, strTemp As String

Label12.Caption = vbNullString
Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
MousePointer = 11
'Call UncompW
Rem ******* Delete Target Folder *******
If FileExists(RoutePath & "\CleanUp.bat") Then
Kill RoutePath & "\CleanUp.bat"
End If
DoEvents
If DirExists(RoutePath & "\newRoute") Then
NewFile = FreeFile
Open RoutePath & "\CleanUp.bat" For Append As #NewFile
Print #NewFile, "If not exist c:\windows\command\deltree.exe goto tryxp"
Print #NewFile, "Deltree /y " & ChrW$(34) & "newRoute" & ChrW$(34)
Print #NewFile, "GoTo CarryOn"
Print #NewFile, ":tryxp"
Print #NewFile, "RD /s /q " & ChrW$(34) & "newRoute" & ChrW$(34)
Print #NewFile, ":CarryOn"
Print #NewFile, "echo *"
Print #NewFile, "echo *"
Print #NewFile, "echo All done - Please close any open DOS windows"
Print #NewFile, "echo *"
Close #NewFile
ChDrive (Left$(RoutePath, 1))
ChDir RoutePath
 Call ShellAndWait(ChrW$(34) & RoutePath & "\CleanUp.bat" & ChrW$(34), True, vbNormalFocus)
End If
Rem ************************************
DoEvents
CDL1.DialogTitle = "Select Original tsection.dat"
CDL1.Flags = cdlOFNExplorer
CDL1.InitDir = MSTSPath & "\Global"
CDL1.Filter = "Tsection Files (*.dat)|*.dat"
CDL1.ShowOpen

strOrigTS = CDL1.Filename

Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If
ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_cvrt.log" & ChrW$(34) & "  cvrt -b" & ChrW$(34) & strOrigTS & ChrW$(34) & " -v99:99 " & ChrW$(34) & RoutePath & ChrW$(34)

  Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_cvrt.log"
 frmReport.Show 1
 DoEvents
 Select Case MsgBox("New Files have been prepared in the Newroute folder of your route. If the report indicated an error in preparing the route then Click NO, otherwise it should be safe to click YES to complete the move.", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes

 
 strTemp = "xcopy " & ChrW$(34) & RoutePath & "\newroute\*.*" & ChrW$(34) & " " & ChrW$(34) & RoutePath & ChrW$(34) & " /S /Y"
Call ShellAndWait(strTemp, True, vbNormalFocus)
 DoEvents
  Call CheckMissFolders(RoutePath)
DoEvents
 Case vbNo
 End Select

  MousePointer = 0
 
 Label12.Caption = "Finished."

End Sub

Private Sub Command74_Click()
Dim i As Integer, NewRouteName As String
Dim booExists As Boolean, OldRouteName As String

SparePath = App.Path & "\TempFiles"
MousePointer = 11
RoutePath = File1(0).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If
Text1(0) = "*.trk"
RouteName = File1(0).List(i)
OldRouteName = RouteName
Call FindRouteName(RoutePath & "\" & RouteName, NewRouteName, booExists)
If booExists = False Then
MsgBox Lang(339) & vbCr & Lang(340), 16, Lang(341)
Exit Sub
Else
RouteName = NewRouteName
End If
WorldPath = RoutePath & "\World"

RouteListed = True
Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
Close

cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents





If RouteName = vbNullString Then
Call MsgBox(Lang(463), vbExclamation, App.Title)

Exit Sub
End If
Call UncompressAllW(WorldPath)
 Drive1(1).Drive = Left$(RoutePath, 2)
Dir1(1).Path = RoutePath & "\textures"
Text1(1).Text = "*.ace"
flagChange = 2

frmRepShape.Show

     DoEvents
     
MousePointer = 0





End Sub


Private Sub Command75_Click()
Dim i As Integer

frmUtils.ZOrder


Trainspath = MSTSPath & "\Trains\"
strReport = vbNullString

'strFirst = Dir1(0).path
'Call CountStock
DoEvents
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = Trainspath
Text1(0) = "*.*"
frmUtils.Refresh
DoEvents

strReport = vbNullString
For i = 0 To 4
Label7(i).Caption = vbNullString
Next
Label3.Caption = vbNullString
lngAct = 0
lngSrv = 0
lngCon = 0
lngLoco = 0
lngWagons = 0
lngTfc = 0
cursouind = 0
Trainspath = Dir1(cursouind).Path
If Right$(Trainspath, 6) <> "Trains" Then
Call MsgBox("Please select the 'TRAINS' folder in the left hand folder window" _
            & vbCrLf & "which you wish to check." _
            & vbCrLf & "" _
            , vbExclamation, App.Title)

Exit Sub
End If
booStockOnly = False
SB2.Panels(2).Text = "Counting Stock"

Call CountStock2
SB2.Panels(2).Text = "Counting Consists"
Call GetConsists2
Call GetStock2
SB2.Panels(2).Text = "Populating Grid..."
frmStock.Show
Dir1(0).Path = Trainspath
     DoEvents
     SB2.Panels(2).Text = vbNullString
End Sub

Private Sub Command76_Click()
Dim strFirst As String, i As Integer
MousePointer = 11
For i = 0 To 4
Label7(i).Caption = vbNullString
Next i


Text1(0) = "*.*"
strReport = vbNullString
strFirst = Dir1(0).Path
Call CountStock

DoEvents
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = MSTSPath & "\Trains\Consists\"
Text1(0) = "*.con"
frmUtils.Refresh
DoEvents


booFixCon = True
frmSearch.Show
DoEvents
Dir1(0).Path = strFirst
booFixCon = False
Text1(0) = "*.*"
MousePointer = 0
End Sub



Private Sub Command77_Click()
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
strReport = vbNullString


cursouind = 0
Call ReadTrack


End Sub

Private Sub Command78_Click()
Text1(0).Text = "*.sms"
End Sub



Private Sub ReadRoute()

Dim i As Integer, j As Long, x As Integer
Dim booNoShape As Boolean, flagACE As Integer, strOrig As String

Dim GlobalPath As String, Lo_Tilepath As String, strS As String, strSpare As String

On Error GoTo Errtrap

ReDim Preserve strShp(0 To Shp_Chunk)
ReDim Preserve strGlobShp(0 To Shp_Chunk)
ReDim Ace1(0 To Shp_Chunk)
ReDim ForTex(0 To For_Chunk)
ReDim HazShp(0 To For_Chunk)
ReDim TerrTex(0 To For_Chunk)
ReDim Transfer(0 To For_Chunk)
x = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
MousePointer = 11
If numShp < 1 Then
numShp = 1
End If
numHaz = 0
numFor = 0
If numAce < 1 Then
numAce = 1
End If
numTerr = 0
numTrans = 0
WorldPath = RoutePath & "\world"
TertexPath = RoutePath & "\Terrtex"
TilePath = RoutePath & "\Tiles"
Lo_Tilepath = RoutePath & "\Lo_Tiles"
GlobalPath = MSTSPath & "\Global"
strSpare = App.Path & "\TempFiles"
frmUtils.Dir1(0).Path = WorldPath
frmUtils.Text1(0).Text = "*.w"
cursouind = 0
x = 2
booOKAll = False
booWriteFile = False
Call DoDeCompFolder("w", WorldPath, strSpare)
DoEvents
frmUtils.Dir1(0).Path = strSpare
frmUtils.Text1(0).Text = "*.w"
cursouind = 0
DoEvents
x = 3
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i
Rem copy to Spares **************************************
x = 4
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
frmUtils.Label12.Caption = "Reading: " & frmUtils.File1(cursouind).List(i)
'
  If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = strSpare & "\" & frmUtils.File1(cursouind).List(i)

Call ReadWorld(fullpath$)
'
   End If
'
   Next i
Call KillSpare("*.w")
DoEvents
   Rem ********** Read the Tiles *******************
x = 5
   frmUtils.Dir1(0).Path = TilePath
   SB1.Panels(2).Text = "Uncompressing .t files"
   DoEvents
Call DoDeCompFolder("t", TilePath, strSpare)
DoEvents
frmUtils.Dir1(0).Path = strSpare
frmUtils.Text1(0).Text = "*.t"
 cursouind = 0

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i
x = 6
For i = 0 To frmUtils.File1(cursouind).ListCount - 1

frmUtils.Label12.Caption = "Reading: " & frmUtils.File1(cursouind).List(i)
frmUtils.Label12.Refresh

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = strSpare & "\" & frmUtils.File1(cursouind).List(i)

Call ReadTerrain(fullpath$)

   End If

   Next i
Call KillSpare("*.t")
DoEvents
 Rem *********** Read any Lo_Tiles **********
 x = 7
 If DirExists(Lo_Tilepath) Then
     frmUtils.Dir1(0).Path = Lo_Tilepath
Call DoDeCompFolder("t", Lo_Tilepath, strSpare)
DoEvents
frmUtils.Dir1(0).Path = strSpare
frmUtils.Text1(0).Text = "*.t"
 cursouind = 0

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i
x = 8
For i = 0 To frmUtils.File1(cursouind).ListCount - 1

frmUtils.Label12.Caption = "Reading: " & frmUtils.File1(cursouind).List(i)
frmUtils.Label12.Refresh

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = strSpare & "\" & frmUtils.File1(cursouind).List(i)
  
Call ReadTerrain(fullpath$)

   End If

   Next i
Call KillSpare("*.t")
 End If
  x = 9
   Rem *********************************************
   If intCars > 0 Then
   QSort3 Cars(), 0, intCars - 1
   DoEvents
   RemD2 Cars(), Cars2()
   DoEvents
   For j = 0 To intCars - 1
   Cars(j) = vbNullString
   Next j
   
   intCars = UBound(Cars2)
   
   End If
   x = 10
    If numShp > 0 Then
    ReDim Preserve strShp(0 To numShp - 1)
   
    QSort3 strShp(), 0, numShp - 1
   
    DoEvents
    RemD2 strShp(), strShapes()
    DoEvents
    For j = 0 To numShp - 1
    strShp(j) = vbNullString
    Next j
    End If
    x = 11
    If numGlobShp > 0 Then
    ReDim Preserve strGlobShp(0 To numGlobShp - 1)
   
    QSort3 strGlobShp(), 0, numGlobShp - 1
    DoEvents
    RemD2 strGlobShp(), strGlobShp2()
    DoEvents
    For j = 0 To numGlobShp - 1
    strGlobShp(j) = vbNullString
    Next j
   End If
   x = 12
   If numFor > 0 Then
    ReDim Preserve ForTex(0 To numFor - 1)
    QSort3 ForTex(), 0, numFor - 1
    DoEvents
    RemD2 ForTex(), ForTex2()
    DoEvents
    For j = 0 To numFor - 1
    ForTex(j) = vbNullString
    Next j
   End If
    
 x = 13
    If numHaz > 0 Then
    ReDim Preserve HazShp(0 To numHaz - 1)
    QSort3 HazShp(), 0, numHaz - 1
    DoEvents
   
    RemD2 HazShp(), HazShp2()
    DoEvents
  
    numHaz = UBound(HazShp2) + 1
    End If
    If numTrans > 0 Then
    ReDim Preserve Transfer(0 To numTrans - 1)
    QSort3 Transfer(), 0, numTrans - 1
    DoEvents
    RemD2 Transfer(), Transfer2()
    DoEvents
  
    numTrans = UBound(Transfer2)
    End If
    x = 14
  If numTerr > 0 Then
    
    ReDim Preserve TerrTex(0 To numTerr - 1)
    QSort3 TerrTex(), 0, numTerr - 1
    DoEvents
    RemD2 TerrTex(), TerrTex2()
    DoEvents
    For j = 0 To numTerr - 1
    TerrTex(j) = vbNullString
    Next j
    End If
    
   
    numShp = UBound(strShapes)
    numGlobShp = UBound(strGlobShp2)
    numFor = UBound(ForTex2)
    numTerr = UBound(TerrTex2)
    Rem ********* Cars

 x = 15
    If intCars > 0 Then
    For j = 0 To intCars
   SB1.Panels(2) = Cars2(j)
    If FileExists(RoutePath & "\shapes\" & Cars2(j)) Then
  
   FileCopy RoutePath & "\shapes\" & Cars2(j), strSpare & "\" & Cars2(j)
    Else
    Call LookForShape(Cars2(j), booNoShape)
    End If
    Next j
    strOrig = RoutePath & "\shapes\"

    Call DoDeCompFolder("s", strSpare, strSpare)
    DoEvents

    For j = 0 To intCars
    Call ReadShape3(RoutePath & "\shapes", strSpare & "\" & Cars2(j), 3)
    Next j
    Call KillSpare("*.s")
 End If
    x = 16
  
    Rem **********************
    
    If numTrans > 0 Then
    For j = 0 To numTrans
    If Not FileExists(TexturePath & "\" & Transfer2(j)) Then
    flagACE = 1
    strS = " A World file "
   Call LookForACE2(Transfer2(j), flagACE, strS)
   End If
   If Not FileExists(TexturePath & "\Snow\" & Transfer2(j)) Then
    flagACE = 1
    strS = " A World file "
   Call LookForACESnow2(Transfer2(j), flagACE, strS)
   End If
   Next j
    End If
    Rem ******************************
    x = 17
    For j = 0 To numShp
    SB1.Panels(2) = strShapes(j)
   
    If strShapes(j) <> vbNullString Then
    If FileExists(RoutePath & "\shapes\" & strShapes(j)) Then
       FileCopy RoutePath & "\shapes\" & strShapes(j), strSpare & "\" & strShapes(j)
    Else
    Call LookForShape(strShapes(j), booNoShape)
    End If
    End If
    Next j
    SB1.Panels(2) = "Uncompressing Shape Files"
    DoEvents
    strOrig = RoutePath & "\shapes\"
    Call DoDeCompFolder("s", strSpare, strSpare)
    DoEvents
    x = 18
    For j = 0 To numShp

    Call ReadShape3(RoutePath & "\shapes", strSpare & "\" & strShapes(j), 1)
    Next j
    Call KillSpare("*.s")
      
    x = 19
    Rem *******************
    For j = 0 To numGlobShp
    SB1.Panels(2) = strGlobShp(j)
    If strGlobShp2(j) <> vbNullString Then
    If FileExists(GlobalPath & "\shapes\" & strGlobShp2(j)) Then
       FileCopy GlobalPath & "\shapes\" & strGlobShp2(j), strSpare & "\" & strGlobShp2(j)
    Else
    Call LookForGlobalShape(strGlobShp2(j), booNoShape)
    End If
    End If
    Next j
    SB1.Panels(2) = "Uncompressing Global Shape Files"
    DoEvents
    strOrig = GlobalPath & "\shapes\"
    Call DoDeCompFolder("s", strSpare, strSpare)
    DoEvents
    For j = 0 To numShp
    Call ReadShape3(GlobalPath & "\shapes", strSpare & "\" & strGlobShp2(j), 2)
    Next j
    Call KillSpare("*.s")
    
 x = 20
 If numFor > 0 Then
For j = 0 To numFor
SB1.Panels(2) = ForTex2(j)
Call CheckForACETree(ForTex2(j))
Next j
End If
x = 21
If numTerr > 0 Then
For j = 0 To numTerr
SB1.Panels(2) = TerrTex2(j)
If Not FileExists(TertexPath & "\" & TerrTex2(j)) Then

      Call LookForTerrtex(TerrTex2(j))
End If
If Not FileExists(TertexPath & "\Snow\" & TerrTex2(j)) Then

      Call LookForTerrtexSnow(TerrTex2(j))
End If
Next j
End If
x = 22
If booSpareTrack = True Then
booSpareTrack = False
Call MsgBox("At least one track section has been restored from Global\SpareTrack to" _
            & vbCrLf & "your Global\Shapes folder. If you are using a compacted tsection.dat file, you" _
            & vbCrLf & "must remove it and rename 'Master_tsection.dat' back to tsection.dat" _
            , vbExclamation, App.Title)

End If

Exit Sub
Errtrap:

If Err = -2147024893 Then
Resume Next
End If
If Err = 9 Then
Resume Next
End If

Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'Read route' please advise" _
                       & vbCrLf & "Mike that x = " & Str(x) _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
      ' Resume Next
    Exit Sub
    End Select
        
End Sub


Private Sub Command79_Click()
Dim strFirst As String, i As Integer

MousePointer = 11
For i = 0 To 4
Label7(i).Caption = vbNullString
Next i
DoEvents

Text1(0) = "*.*"
strReport = vbNullString
strFirst = Dir1(0).Path
Call CountStock
DoEvents
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = MSTSPath & "\Routes"
Text1(0) = "*.act"
frmUtils.Refresh
DoEvents


booFixAct = True
frmSearch.Show
DoEvents
Dir1(0).Path = strFirst
booFixAct = False
Text1(0) = "*.*"
MousePointer = 0
End Sub

Private Sub Command80_Click()
Dim i%
Dim strNew As String, SparePath As String
Dim x As Integer, Y As Integer
Dim j As Long, itExists As Boolean, q As Integer
Dim NewFile As Integer
Dim booGotIt As Boolean
Dim GlobalShapePath As String
Dim strEnvAce As String
'Dim tfh As TokenFileHandler

On Error GoTo Errtrap
numAce = 0
''Set tfh = New TokenFileHandler
Select Case MsgBox("If Soundsource errors are found, do you wish Route_Riter to patch them so that they appear on the current tile." _
                   & vbCrLf & "If you reply NO, then you can use TsUtils to correct the Soundsource correctly, where possible." _
                   , vbYesNo Or vbQuestion Or vbDefaultButton1, App.Title)

    Case vbYes
booFixSound = True
    Case vbNo
booFixSound = False
End Select
GlobalSparePath = MSTSPath & "\Global\SpareTrack\"
GlobalSoundPath = MSTSPath & "\Sound"

For j = 0 To 1000
Soundfile(j) = vbNullString
WavFile(j) = vbNullString
Next

SB1.Panels(4).Text = time
SB1.Panels(6).Text = vbNullString
numShp = 0
ReDim strShp(0 To Shp_Chunk)
ReDim strGlobShp(0 To Shp_Chunk)
flagNoTex = False

For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then
If frmUtils.Controls(q).Caption <> Lang(30) And frmUtils.Controls(q).Caption <> Lang(12) Then
frmUtils.Controls(q).Enabled = False
ElseIf frmUtils.Controls(q).Caption = Lang(30) Then

frmUtils.Controls(q).Caption = Lang(637)
End If
End If
Next q
ReDim Cars(0 To Car_CHUNK)
intCars = 0
intResponse = 0
intResponse2 = 0
frmUtils.Refresh

If envAceNumber > 0 Then
For j = 1 To envAceNumber
  EnvAceFile(j) = vbNullString
  Next j
  envAceNumber = 0
  End If

On Error GoTo Errtrap


SparePath = App.Path & "\TempFiles"
intTempFile = 0
strReport = vbNullString
 GlobalShapePath = MSTSPath & "\global\shapes\"
 Rem ********** Delete old .s files ************

If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
strReport = vbNullString


cursouind = 0

Rem *********** Check .trk file ***********
'Stop
Text1(0) = "*.trk"
RouteName = File1(cursouind).List(i)
OldRouteName = RouteName

Call CheckMissFolders(RoutePath)
DoEvents

If Not DirExists(RoutePath & "\Terrtex\Snow") Then
MkDir RoutePath & "\Terrtex\Snow"
End If
Call CheckForSMS(RoutePath & "\" & RouteName)


If Not DirExists(RoutePath & "\Envfiles") Then
Call MsgBox("The 'Envfiles' folder of your route is missing, no further" _
            & vbCrLf & "checking is possible until this error is corrected." _
            , vbExclamation, App.Title)
MousePointer = 0

Exit Sub
End If
If Not DirExists(RoutePath & "\Envfiles") Then
MkDir RoutePath & "\Envfiles"
DoEvents
End If
If Not DirExists(RoutePath & "\Envfiles\Textures") Then
MkDir RoutePath & "\Envfiles\Textures"
DoEvents
End If
If Not DirExists(RoutePath & "\Activities") Then
MkDir RoutePath & "\Activities"
DoEvents
End If
If Not DirExists(RoutePath & "\Paths") Then
MkDir RoutePath & "\Paths"
DoEvents
End If
If Not DirExists(RoutePath & "\Services") Then
MkDir RoutePath & "\Services"
DoEvents
End If
If Not DirExists(RoutePath & "\Traffic") Then
MkDir RoutePath & "\Traffic"
DoEvents
End If
If Not DirExists(RoutePath & "\Sound") Then
MkDir RoutePath & "\Sound"
DoEvents
End If
If Not DirExists(RoutePath & "\Terrtex\Snow") Then
MkDir RoutePath & "\Terrtex\Snow"
DoEvents
End If
If Not DirExists(RoutePath & "\Textures\Autumn") Then
MkDir RoutePath & "\Textures\Autumn"
DoEvents
End If
If Not DirExists(RoutePath & "\Textures\AutumnSnow") Then
MkDir RoutePath & "\Textures\AutumnSnow"
DoEvents
End If
If Not DirExists(RoutePath & "\Textures\Night") Then
MkDir RoutePath & "\Textures\Night"
DoEvents
End If
If Not DirExists(RoutePath & "\Textures\Snow") Then
MkDir RoutePath & "\Textures\Snow"
DoEvents
End If
If Not DirExists(RoutePath & "\Textures\Spring") Then
MkDir RoutePath & "\Textures\Spring"
DoEvents
End If
If Not DirExists(RoutePath & "\Textures\SpringSnow") Then
MkDir RoutePath & "\Textures\SpringSnow"
DoEvents
End If
If Not DirExists(RoutePath & "\Textures\Winter") Then
MkDir RoutePath & "\Textures\Winter"
DoEvents
End If
If Not DirExists(RoutePath & "\Textures\WinterSnow") Then
MkDir RoutePath & "\Textures\WinterSnow"
DoEvents
End If
Call CheckTrkForEnv(RoutePath & "\" & OldRouteName)

Rem *******************

DoEvents
cursouind = 0
If FileExists(RoutePath & "\carspawn.dat") Then

SB1.Panels(2).Text = "carspawn.dat"

 DoEvents
NewFile = FreeFile

     Open RoutePath & "\carspawn.dat" For Input As #NewFile
      Do While Not EOF(NewFile)
      Line Input #NewFile, strNew
     strNew = Trim$(strNew)
         If Left$(strNew, 14) = "CarSpawnerItem" Then
        
        x = InStr(strNew, ChrW$(34))
                If x > 0 Then
                Y = InStr(x + 1, strNew, ChrW$(34))
                strNew = Mid$(strNew, x + 1, Y - (x + 1))
                strNew = Trim$(strNew)
                Cars(intCars) = strNew
                intCars = intCars + 1
                        If intCars > UBound(Cars) Then
                          ReDim Preserve Cars(0 To intCars + Car_CHUNK)
                         End If
             
               End If
        End If
   strNew = vbNullString
  itExists = False
   Loop
   Close #NewFile
    ReDim Preserve Cars(0 To intCars - 1)
 End If

 
 Rem*************************************

 If FileExists(RoutePath & "\telepole.dat") Then
  Call CheckForS3(RoutePath & "\telepole.dat")
 End If
If FileExists(RoutePath & "\speedpost.dat") Then

 Call CheckForAce3(RoutePath & "\speedpost.dat")
 'Call CheckForS3(RoutePath & "\speedpost.dat")
 
 End If

 If FileExists(RoutePath & "\sigcfg.dat") Then
Call CheckForAce3(RoutePath & "\sigcfg.dat")


 End If

Rem ************ Check Envfiles ***************************
SB1.Panels(2).Text = "Environment Files"


If booAbort = True Then
booAbort = False
Command1(15).Caption = Lang(30)


MousePointer = 0
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = RoutePath
Text1(0).Text = "*.*"
strReport = vbNullString
strSeason = vbNullString
GoTo AbortNow
End If
 DoEvents
cursouind = 1
Drive1(cursouind).Drive = Left$(EnvPath, 2)
Dir1(cursouind).Path = EnvPath
Text1(cursouind) = "*.env"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
   Call CheckEnvForAce(EnvPath & "\" & File1(cursouind).List(i))
   
 End If
Next i

Drive1(cursouind).Drive = Left$(EnvPath, 2)
Dir1(cursouind).Path = EnvPath & "\Textures"
Text1(cursouind) = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For j = 1 To envAceNumber
 For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).List(i) = EnvAceFile(j) Then
   booGotIt = True
   GoTo GetAnother
   End If
   Next i
  If booGotIt = False Then

  strEnvAce = EnvAceFile(j)
  Call FindEnvAce(strEnvAce)
 
  End If
GetAnother:
booGotIt = False
   Next j
SB1.Panels(2).Text = "World Files"

Call ReadRoute
Call KillSpare("*.s")
MousePointer = 0

Call CheckACESMS
DoEvents


MousePointer = 0
 
 Text1(0) = "*.*"
  Text1(1) = "*.s"
 

  MousePointer = 0
  

  
 If strReport = vbNullString Then
 strReport = "No errors found..."
 Else
 strSeason = strSeason & vbCrLf & vbCrLf & Lang(498) & vbCrLf

 End If
frmReport.Rich1 = strMainReport & vbCrLf & vbCrLf & strReport & vbCrLf & vbCrLf & strSeason

 Close
 
 SB1.Panels(6).Text = time
 'frmReport.Show 1
 frmReport.Show 1
     DoEvents
     
 strReport = vbNullString
 strSeason = vbNullString

Text1(1).Text = "*.*"
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = RoutePath
Text1(0).Text = "*.*"

Call KillSpare("*.t")
DoEvents
AbortNow:
Command1(15).Caption = Lang(30)

For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then frmUtils.Controls(q).Enabled = True
Next q


Exit Sub
Errtrap:

Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'Command80_CLick' please advise" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
       
    Resume Next
        Case vbCancel
     'Resume Next
    Exit Sub
    End Select

End Sub

Private Sub Command81_Click()
Dim strFirst As String, i As Integer

MousePointer = 11
For i = 0 To 4
Label7(i).Caption = vbNullString
Next i
Text1(0) = "*.*"
strReport = vbNullString
strFirst = Dir1(0).Path
Call GetConsists

DoEvents
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = MSTSPath & "\Routes"
Text1(0) = "*.srv"
DoEvents
frmUtils.Refresh

booFixSrv = True
frmSearch.Show
DoEvents
Dir1(0).Path = strFirst
booFixSrv = False
MousePointer = 0
End Sub

Private Sub Command82_Click()
Dim strFirst As String, TrainsetPath As String

MousePointer = 11
TrainsetPath = MSTSPath & "\Trains\Trainset\"
strReport = vbNullString

strFirst = Dir1(0).Path
'Call CountStock
DoEvents
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = TrainsetPath
Text1(0) = "*.eng"
frmUtils.Refresh
DoEvents


booFixEng = True
frmSearch.Show
DoEvents
Dir1(0).Path = strFirst
booFixEng = False
MousePointer = 0
End Sub



Private Sub Command83_Click()
frmEngEdit.Show

End Sub

Private Sub Command84_Click()
Dim strFirst As String, strPath As String

MousePointer = 11
strPath = MSTSPath & "\"

strReport = vbNullString
strFirst = Dir1(0).Path
'Call CountStock
DoEvents
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = strPath
Text1(0) = "*.sd"
frmUtils.Refresh
DoEvents


booFixSD = True

frmSearch.Show
DoEvents
Dir1(0).Path = strFirst
booFixSD = False
MousePointer = 0
End Sub

Private Sub Command85_Click()
Dim strFirst As String, TrainsetPath As String

MousePointer = 11
TrainsetPath = MSTSPath & "\Trains\Trainset\"

strReport = vbNullString
strFirst = Dir1(0).Path
'Call CountStock
DoEvents
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = TrainsetPath
Text1(0) = "*.sms"
frmUtils.Refresh
DoEvents


booFixSMS = True

frmSearch.Show
DoEvents
Dir1(0).Path = strFirst
booFixSMS = False
MousePointer = 0
Select Case MsgBox("You must re-start Route_Riter after using the Fix .sms files option.", vbOKCancel Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbOK
Command1(15).value = True
    Case vbCancel
Exit Sub
End Select

End Sub





Private Sub Command86_Click()
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If
strReport = vbNullString
Call WorldCount
DoEvents

 MousePointer = 0
booWorldCount = True

DoEvents

MousePointer = 0
    
 frmReport.Rich1.Text = strReport
 frmReport.Show 1
 
End Sub

Private Sub Command87_Click()
frmMini.Show

End Sub

Private Sub Command88_Click()

If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If

Select Case MsgBox("Make sure you only run this option while you have a Mini-Route selected." _
                   & vbCrLf & "Otherwise you will delete many of the files in your Global\Shapes folder ....." _
                   , vbOKCancel Or vbExclamation Or vbDefaultButton1, "Warning!!!")

    Case vbOK
MousePointer = 11
    Case vbCancel
Exit Sub
End Select
strReport = vbNullString
Call MiniTrack
DoEvents

 MousePointer = 0
booWorldCount = True

'DoEvents
'Command39.value = True
DoEvents
MousePointer = 0

End Sub

Private Sub Command89_Click()
Call KillSpare("*.bat")
frmTrainset.Show
End Sub

Private Sub Command9_Click()
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If

Select Case MsgBox(Lang(613) & vbCrLf & Lang(614), vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
    strReport = vbNullString
Call CompressDXT
    Case vbNo
Exit Sub
End Select


End Sub


Sub AddDirSep(strPathName As String)
    If Right$(RTrim$(strPathName), Len("\")) <> "\" Then
        strPathName = RTrim$(strPathName) & "\"
    End If
End Sub
Private Sub CheckAnim(strAnim As String, LocoPath As String)

Dim x As Integer, strCorrEng As String, strTrains As String
Dim y1 As Integer, strAnimCorr As String
Dim strDir As String, strShape As String, strShapePath As String
On Error GoTo Errtrap
strTrains = MSTSPath & "\trains\trainset\"

x = InStrRev(LocoPath, "\")
strCorrEng = Mid$(LocoPath, x + 1)
strShapePath = Left$(LocoPath, x)
strAnimCorr = Replace(strAnim, "\\", "/")
strAnimCorr = Replace(strAnimCorr, "//", "/")

x = InStr(strAnimCorr, "common.crew")
If x > 0 Then
y1 = InStr(x + 12, strAnimCorr, "/")
strDir = Mid$(strAnimCorr, x + 12, y1 - (x + 12))
strShape = Mid$(strAnimCorr, y1 + 1)
If FileExists(strTrains & "common.crew\" & strDir & "\" & strShape) Then
Call CheckAnimForShape(strTrains & "common.crew\" & strDir & "\" & strShape, strTrains & "common.crew\" & strDir & "\", strCorrEng)
Else
strReport = strReport & "common.crew\" & strDir & "\" & strShape & Lang(584) & strCorrEng & vbCrLf
End If
ElseIf FileExists(strShapePath & strAnim) Then

Call CheckAnimForShape(strShapePath & strAnim, strShapePath, strCorrEng)
ElseIf Left$(strAnim, 2) <> ".." Then
strReport = strReport & strShapePath & strAnim & Lang(584) & strCorrEng & vbCrLf
End If
Exit Sub
Errtrap:
Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in  'CheckAnim' checking " & strAnim & " please send" _
                       & vbCrLf & "file to Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
    Exit Sub
    End Select
End Sub

Private Sub CheckAnimForShape(strShape As String, strShapePath As String, strCorrEng As String)
Dim x As Integer, strShapeName As String, Filpath1$
Dim A$, Y As Integer, yy As Integer, strShapePath2 As String
Dim strAce As String, y1 As Integer, y2 As Integer, booCom As Boolean, strCom As String

On Error GoTo Errtrap
   
DoEvents
x = InStrRev(strShapePath, "common.crew")
If x > 0 Then
booCom = True
strCom = Left$(strShapePath, x - 1)
End If
x = InStrRev(strShape, "\")
strShapeName = Mid$(strShape, x + 1)
strShapePath2 = Left(strShape, x - 1)
Filpath1$ = App.Path & "\TempFiles"
TokMode = 0
  booWriteFile = True
Call DoDeComp2(strShapeName, strShapePath2, Filpath1$)

 DoEvents

Open Filpath1$ & "\" & strShapeName For Input As #3
  Do While Not EOF(3)
  Line Input #3, A$
   Y = InStr(A$, "image (")
  If Y > 0 Then
  A$ = Replace(A$, "\\", "\")
  If booCom = True Then

  y2 = InStr(A$, "common.crew")
    If y2 > 0 Then
  yy = InStr(y2, A$, ")")
  strAce = Mid$(A$, y2, yy - (y2 + 1))
  strAce = Trim$(strAce)
  If Left(strAce, 1) = ChrW$(34) Then
  strAce = Mid(strAce, 2)
  End If
  If Right(strAce, 1) = ChrW$(34) Then
  strAce = Left(strAce, Len(strAce) - 1)
  End If
        End If
        If Not FileExists(strCom & strAce) Then
  strReport = strReport & strCom & strAce & Lang(585) & strCorrEng & vbCrLf
  
  
  End If
  Else
  
  
  yy = InStr(Y, A$, ")")
  strAce = Mid$(A$, 7, yy - (Y + 5))
  strAce = Trim$(strAce)
  y1 = InStrRev(strAce, "\")
  If y1 > 0 Then
    strAce = Mid$(strAce, y1 + 1)
  End If
  
  If Not FileExists(strShapePath & strAce) Then
  strReport = strReport & strShapePath & strAce & Lang(585) & strCorrEng & vbCrLf
  
  
  End If
  End If
  End If
Loop
Close #3
Exit Sub
Errtrap:
Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in  'CheckAnimForShape' checking " & strShapeName & " please send" _
                       & vbCrLf & "file to Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
    Exit Sub
    End Select
End Sub


Private Sub CheckForAce5(SFilepath As String, strSN As String)
Dim x As Long, Y As Long, Z As Long
Dim strS As String

'On Error GoTo errtrap
strReport = strSN & " requires the following .ace files:-" & vbCrLf

x = 1
 Do While x > 0
   x = InStr(x, SFilepath, ".ace")
   
   If x > 0 Then

   Y = InStrRev(SFilepath, "(", x)
  Z = InStrRev(SFilepath, ChrW$(34), x)
   If Z > Y Then
   strS = Mid$(SFilepath, Z + 1, (x + 4) - (Z + 1))
   Else
   strS = Mid$(SFilepath, Y + 1, (x + 4) - (Y + 1))
   
   End If
   
   strS = Trim$(strS)
   If Left$(strS, 1) = ChrW$(34) Then
  
   strS = Right$(strS, Len(strS) - 1)
   End If
   If Right$(strS, 1) = ChrW$(34) Then
   strS = Left$(strS, Len(strS) - 1)
   End If
   strReport = strReport & strS & vbCrLf
   x = x + 1
   End If
   
   Loop
   
   
   Exit Sub
   
End Sub

Private Sub CompressDXT()

Dim TerrPath As String
Dim TerrSnowPath As String, EnvPath As String
booCompressAll = True
On Error GoTo Errtrap

Call CompressDXT2(TexturePath)
DoEvents
Do While booAllCompressed = False
Loop

If DirExists(TexSnowPath) Then
Call CompressDXT2(TexSnowPath)
DoEvents
Do While booAllCompressed = False
Loop
End If

If DirExists(TexAutPath) Then
Call CompressDXT2(TexAutPath)
DoEvents
Do While booAllCompressed = False
Loop


End If
If DirExists(TexAutSnowPath) Then
Call CompressDXT2(TexAutSnowPath)
DoEvents
Do While booAllCompressed = False
Loop

End If
If DirExists(TexSprPath) Then
Call CompressDXT2(TexSprPath)
DoEvents
Do While booAllCompressed = False
Loop

End If
If DirExists(TexSprSnowPath) Then
Call CompressDXT2(TexSprSnowPath)
DoEvents
Do While booAllCompressed = False
Loop

End If
If DirExists(TexWinPath) Then
Call CompressDXT2(TexWinPath)
DoEvents
Do While booAllCompressed = False
Loop

End If
If DirExists(TexWinSnowPath) Then
Call CompressDXT2(TexWinSnowPath)
DoEvents
Do While booAllCompressed = False
Loop

End If
If DirExists(TexNightPath) Then
Call CompressDXT2(TexNightPath)
DoEvents
Do While booAllCompressed = False
Loop
End If

EnvPath = RoutePath & "\envfiles\textures"

If DirExists(EnvPath) Then
Call CompressDXT2(EnvPath)
DoEvents
Do While booAllCompressed = False
Loop
End If

TerrPath = RoutePath & "\terrtex"

If DirExists(TerrPath) Then
Call CompressDXT2(TerrPath)
DoEvents
Do While booAllCompressed = False
Loop
End If
TerrSnowPath = RoutePath & "\terrtex\snow"
If DirExists(TerrSnowPath) Then
Call CompressDXT2(TerrSnowPath)
DoEvents
Do While booAllCompressed = False
Loop
End If
Call CompressDXT2(RoutePath)
DoEvents

MousePointer = 11
Rem ************
 cursouind = 0
 Drive1(cursouind).Drive = Left$(TexturePath, 2)
Dir1(cursouind).Path = TexturePath
Text1(cursouind).Text = "*.*"

Text1(1).Text = "*.*"
  MousePointer = 0


  If strReport <> vbNullString Then
  frmReport.Rich1.Text = strReport
     frmReport.Show 1
     DoEvents
  End If
  booCompressAll = False
  
Exit Sub
Errtrap:

If Err = 53 Then
Resume Next
End If

End Sub

Private Sub CompressDXT2(strTexturePath As String)
Dim i As Integer, varBatText As String, Filpath1$
Dim shortTexPath As String
Dim strFound As String
Dim jj As Integer, x As Integer, strPath As String, strName As String
Dim strTGAF As String, booFileName As Boolean, strAceView As String

On Error GoTo Errtrap
SB1.Panels(2).Text = "Compressing"
If Not DirExists(App.Path & "\TempFiles") Then
MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents
FileCopy App.Path & "\AceIt.exe", App.Path & "\tempfiles\AceIt.exe"
MousePointer = 11
booAllCompressed = False
cursouind = 0
Drive1(cursouind).Drive = Left$(strTexturePath, 2)
Dir1(cursouind).Path = strTexturePath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

shortTexPath = File1(cursouind).Path

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
   strFound = shortTexPath & "\" & File1(cursouind).List(i)
  ' If Left$(File1(cursouind).list(i), 6) <> "aclean" And Left$(File1(cursouind).list(i), 6) <> "nclean" Then
x = InStrRev(strFound, "\")
 strPath = Left$(strFound, x - 1)
 strName = Mid$(strFound, x + 1)
   If Right$(strFound, 4) = ".ace" Then
      fullpath$ = strFound
   strOrigFile = strName
   Else
   MousePointer = 0
   Call MsgBox(Lang(393) & strOrigFile & Lang(405), vbExclamation, Lang(404))
    GoTo NextOne
    End If
      
Filpath1$ = App.Path & "\TempFiles"

   strAceView = strFound
  
   Call readFile(strAceView, bdata())
   If bdata(16) = 17 Then GoTo NextOne
   
   If bdata(8) <> 0 Or bdata(12) <> 0 Then GoTo NextOne
   If bdata(9) <> bdata(13) Then GoTo NextOne
   strTGASave = Filpath1$ & "\" & Left$(strOrigFile, Len(strOrigFile) - 3) & "tga"
   SB1.Panels(2).Text = strOrigFile & " to TGA"
   
   
   
  
  result = AceToTgaSquare(strAceView, strTGASave)
  DoEvents
  strTGAF = Left$(strOrigFile, Len(strOrigFile) - 3) & "tga"
  
  
   For jj = 1 To Len(strTGAF)
  
   If Mid$(strTGAF, jj, 1) = "&" Or Mid$(strTGAF, jj, 1) = " " Or Asc(Mid$(strTGAF, jj, 1)) > 122 Then
   booFileName = True
   Mid$(strTGAF, jj, 1) = "_"
     
    Name strTGASave As Filpath1$ & "\" & strTGAF
   End If
   Next jj
  
 ChDrive Left$(Filpath1$, 1)
 ChDir Filpath1$
  
    varBatText = "AceIt.exe " & strTGAF & " " & Left$(strTGAF, Len(strTGAF) - 4) & ".ace -dxt -q"
  Call ShellAndWait(varBatText, True, vbMinimizedFocus)
  
  DoEvents
        If booFileName = True Then
        strOldAce = Left$(strTGAF, Len(strTGAF) - 4)
        strOldAce = strOldAce & ".ace"
        
        Name Filpath1$ & "\" & strOldAce As Filpath1$ & "\" & strOrigFile
        
        End If
        
DoEvents
If FileExists(Filpath1$ & "\" & strOrigFile) Then
Kill strFound
DoEvents
FileCopy Filpath1$ & "\" & strOrigFile, strFound
Else
 strReport = strReport & strOrigFile & Lang(549) & vbCrLf
End If

   End If
  ' End If
  
NextOne:
   Next i
booAllCompressed = True
If booCompressAll = False Then
If strReport <> vbNullString Then
  frmReport.Rich1.Text = strReport
     frmReport.Show 1
     DoEvents
     
  End If
  End If
  SB1.Panels(2).Text = "Finished"
Exit Sub
Errtrap:

If Err = 53 Then
Resume Next
End If

End Sub


Private Sub SetLangMenu()
Dim i As Integer, strLang
Drive1(0).Drive = Left$(App.Path, 2)
Dir1(0).Path = App.Path
Text1(0) = "Lang_*.txt"
For i = 0 To File1(0).ListCount - 1
    File1(0).Selected(i) = True
Next i
For i = 0 To File1(0).ListCount - 1

   If File1(0).Selected(i) Then
   strLang = File1(0).List(i)
   strLang = Mid$(strLang, 6, Len(strLang) - 9)
   Language(i) = strLang
mnu1(i).Caption = strLang
mnu1(i).Visible = True
End If
Next i

End Sub

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Private Sub CheckEnvAce(EnvFile As String, Z As Integer)
Dim EnvName As String

MainRoutePath = MSTSPath & "\routes\"
Eur1Path = MainRoutePath & "Europe1\"
Eur2Path = MainRoutePath & "Europe2\"
Jap1Path = MainRoutePath & "Japan1\"
Jap2Path = MainRoutePath & "Japan2\"
USA1Path = MainRoutePath & "USA1\"
USA2Path = MainRoutePath & "USA2\"
TemplatePath = MSTSPath & "\Template\"

Open App.Path & "\SetupFiles\Installme.bat" For Append As #12

EnvName = EnvFile
If FileExists(TemplatePath & "Envfiles\Textures\" & EnvName) Then
If GetCRC(TemplatePath & "Envfiles\Textures\" & EnvName) = GetCRC(RoutePath & "\Envfiles\Textures\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\..\Template\Envfiles\Textures\" & EnvName & ChrW$(34) & " .\Envfiles\Textures\ /y"

Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\Textures\" & EnvName
End If
GoTo EndIt
End If

If FileExists(Eur1Path & "Envfiles\Textures\" & EnvName) Then
If GetCRC(Eur1Path & "Envfiles\Textures\" & EnvName) = GetCRC(RoutePath & "\Envfiles\Textures\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe1\Envfiles\Textures\" & EnvName & ChrW$(34) & " .\Envfiles\Textures\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\Textures\" & EnvName
End If
GoTo EndIt
End If
If FileExists(Eur2Path & "Envfiles\Textures\" & EnvName) Then
If GetCRC(Eur2Path & "Envfiles\Textures\" & EnvName) = GetCRC(RoutePath & "\Envfiles\Textures\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe2\Envfiles\Textures\" & EnvName & ChrW$(34) & " .\Envfiles\Textures\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\Textures\" & EnvName
End If
GoTo EndIt
End If
If FileExists(Jap1Path & "Envfiles\Textures\" & EnvName) Then
If GetCRC(Jap1Path & "Envfiles\Textures\" & EnvName) = GetCRC(RoutePath & "\Envfiles\Textures\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\Envfiles\Textures\" & EnvName & ChrW$(34) & " .\Envfiles\Textures\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\Textures\" & EnvName
End If
GoTo EndIt
End If
If FileExists(Jap2Path & "Envfiles\Textures\" & EnvName) Then
If GetCRC(Jap2Path & "Envfiles\Textures\" & EnvName) = GetCRC(RoutePath & "\Envfiles\Textures\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\Envfiles\Textures\" & EnvName & ChrW$(34) & " .\Envfiles\Textures\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\Textures\" & EnvName
End If
GoTo EndIt
End If
If FileExists(USA1Path & "Envfiles\Textures\" & EnvName) Then
If GetCRC(USA1Path & "Envfiles\Textures\" & EnvName) = GetCRC(RoutePath & "\Envfiles\Textures\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\Envfiles\Textures\" & EnvName & ChrW$(34) & " .\Envfiles\Textures\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\Textures\" & EnvName
End If
GoTo EndIt
End If
If FileExists(USA2Path & "Envfiles\Textures\" & EnvName) Then
If GetCRC(USA2Path & "Envfiles\Textures\" & EnvName) = GetCRC(RoutePath & "\Envfiles\Textures\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\Envfiles\Textures\" & EnvName & ChrW$(34) & " .\Envfiles\Textures\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\Textures\" & EnvName
End If
GoTo EndIt
End If
EndIt:
Close #12
End Sub

Private Sub CheckEnv(EnvFile As String, Z As Integer)
Dim EnvName As String

MainRoutePath = MSTSPath & "\routes\"
Eur1Path = MainRoutePath & "Europe1\"
Eur2Path = MainRoutePath & "Europe2\"
Jap1Path = MainRoutePath & "Japan1\"
Jap2Path = MainRoutePath & "Japan2\"
USA1Path = MainRoutePath & "USA1\"
USA2Path = MainRoutePath & "USA2\"

Open App.Path & "\SetupFiles\Installme.bat" For Append As #12

EnvName = EnvFile


If FileExists(Eur1Path & "Envfiles\" & EnvName) Then
If GetCRC(Eur1Path & "Envfiles\" & EnvName) = GetCRC(RoutePath & "\Envfiles\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe1\Envfiles\" & EnvName & ChrW$(34) & " .\Envfiles\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\" & EnvName
End If
GoTo EndIt
End If
If FileExists(Eur2Path & "Envfiles\" & EnvName) Then
If GetCRC(Eur2Path & "Envfiles\" & EnvName) = GetCRC(RoutePath & "\Envfiles\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe2\Envfiles\" & EnvName & ChrW$(34) & " .\Envfiles\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\" & EnvName
End If
GoTo EndIt
End If
If FileExists(Jap1Path & "Envfiles\" & EnvName) Then
If GetCRC(Jap1Path & "Envfiles\" & EnvName) = GetCRC(RoutePath & "\Envfiles\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\Envfiles\" & EnvName & ChrW$(34) & " .\Envfiles\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\" & EnvName
End If
GoTo EndIt
End If
If FileExists(Jap2Path & "Envfiles\" & EnvName) Then
If GetCRC(Jap2Path & "Envfiles\" & EnvName) = GetCRC(RoutePath & "\Envfiles\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\Envfiles\" & EnvName & ChrW$(34) & " .\Envfiles\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\" & EnvName
End If
GoTo EndIt
End If
If FileExists(USA1Path & "Envfiles\" & EnvName) Then
If GetCRC(USA1Path & "Envfiles\" & EnvName) = GetCRC(RoutePath & "\Envfiles\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\Envfiles\" & EnvName & ChrW$(34) & " .\Envfiles\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\" & EnvName
End If
GoTo EndIt
End If
If FileExists(USA2Path & "Envfiles\" & EnvName) Then
If GetCRC(USA2Path & "Envfiles\" & EnvName) = GetCRC(RoutePath & "\Envfiles\" & EnvName) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\Envfiles\" & EnvName & ChrW$(34) & " .\Envfiles\ /y"
Print #12, strBat
DoEvents
Z = Z + 1
KillEnvBat(Z) = RoutePath & "\Envfiles\" & EnvName
End If
GoTo EndIt
End If
EndIt:
Close #12
End Sub


Private Sub CheckForACETree(strAce As Variant)
Dim sFile As String, booTop As Boolean

sFile = " A tree texture "
booTop = False
  If Not FileExists(TexturePath & "\" & strAce) Then
  booTop = True
   Call LookForACESD(strAce, sFile, booTop)
    End If

   If Not FileExists(TexAutPath & "\" & strAce) Then
   Call LookForACESD("autumn\" & strAce, sFile, booTop)
  
   End If
   If Not FileExists(TexAutSnowPath & "\" & strAce) Then
    Call LookForACESD("autumnsnow\" & strAce, sFile, booTop)
   
   End If
   If Not FileExists(TexSprPath & "\" & strAce) Then
   Call LookForACESD("spring\" & strAce, sFile, booTop)
  
   End If
   If Not FileExists(TexSprSnowPath & "\" & strAce) Then
   Call LookForACESD("springsnow\" & strAce, sFile, booTop)
 
   End If
   If Not FileExists(TexWinPath & "\" & strAce) Then
   Call LookForACESD("winter\" & strAce, sFile, booTop)
 
   End If
   If Not FileExists(TexWinSnowPath & "\" & strAce) Then
   Call LookForACESD("wintersnow\" & strAce, sFile, booTop)
   
   End If
End Sub

Private Function ConvertESD(CompleteFilePath As String, flagway As Integer, strESD As String) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertESD = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer


'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then

xx = InStr(MyString, "ESD_A")


x = InStr(xx, MyString, vbCr)

strStart = Left$(MyString, xx - 1)
strEnd = Mid$(MyString, x)
MyString = strStart & "ESD_Alternative_Texture ( " & strESD & " )" & strEnd


End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertESD = True

ERRHANDLER:

  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function


Private Function FixCarPark(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
FixCarPark = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String
Dim xx As Integer

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If

If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then

xx = InStr(MyString, "ESD_Alternative_Teture")
If xx > 0 Then

strStart = Left$(MyString, xx - 1)
strEnd = Mid$(MyString, xx + 22)
MyString = strStart & "ESD_Alternative_Texture" & strEnd
End If

End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
FixCarPark = True

ERRHANDLER:

  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function

Private Function ConvertT(CompleteFilePath As String, strErrorBias As String, flagway As Integer) As Boolean

On Error GoTo ERRHANDLER
ConvertT = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim x As Long
Dim xx As Long, Y As Long

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If

If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
x = 1
TryAgain:
xx = InStr(x, MyString, "terrain_patchset_patch ( ")
    If xx > 0 Then
    Y = InStr(xx, MyString, ")")
        If Mid$(MyString, Y - 2, 1) <> strErrorBias Then
        Mid$(MyString, Y - 2, 1) = strErrorBias
        End If
    x = xx + 20
    GoTo TryAgain
    End If
End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertT = True

ERRHANDLER:

  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function




Public Sub ConvTGAtoACE(strOrigFilePath As String)
Dim i As Integer, varBatText As Variant
Dim strDrive As String
Dim booCompFound As Boolean
Dim strOrigFile As String, strTempF(1 To 50) As String, strOldF(1 To 50) As String
Dim intTemp As Integer, jj As Integer
Dim strOldAce As String, strNewAce As String

On Error GoTo Errtrap
Rem ********** Kill Textures in the temp directory
MousePointer = 11
    cursouind = 1
filepath1$ = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.tga"
Label11.Caption = "Converting .tga files back to .ace with DXT1 compression"
DoEvents

If FileExists(filepath1$ & "\do_makeace.bat") Then
Kill filepath1$ & "\do_makeace.bat"
End If
cursouind = 1
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   strOrigFile = File1(cursouind).List(i)
  
      If Right$(strOrigFile, 4) <> ".tga" Then
   MousePointer = 0
   Call MsgBox(Lang(393) & strOrigFile & Lang(489), vbExclamation, Lang(404))
GoTo NextOne
End If
   '***************
   
   For jj = 1 To Len(strOrigFile)
  
   If Mid$(strOrigFile, jj, 1) = "&" Or Mid$(strOrigFile, jj, 1) = " " Or Asc(Mid$(strOrigFile, jj, 1)) > 122 Then
   
   Mid$(strOrigFile, jj, 1) = "_"
   intTemp = intTemp + 1
    strTempF(intTemp) = strOrigFile
    strOldF(intTemp) = File1(cursouind).List(i)
    Name filepath1$ & "\" & strOldF(intTemp) As filepath1$ & "\" & strOrigFile
   End If
   Next jj
 
 
   booCompFound = True
  
    varBatText = "AceIt.exe " & strOrigFile & " " & Left$(strOrigFile, Len(strOrigFile) - 4) & ".ace -dxt -q"
     Open filepath1$ & "\do_makeace.bat" For Append As #6
   Print #6, varBatText
   Close 6
   End If
NextOne:
   Next i
   Close
  
   Text1(cursouind).Text = "*.*"
  strDrive = Left$(filepath1$, 1)
   ChDrive strDrive
ChDir filepath1$
mydir = CurDir

Call ShellAndWait("do_makeace.bat", True, vbNormalFocus)
DoEvents
Kill filepath1$ & "\*.tga"
MousePointer = 11

If intTemp > 0 Then
For i = 1 To intTemp
strOldAce = Left$(strOldF(i), Len(strOldF(i)) - 4)
strOldAce = strOldAce & ".ace"
strNewAce = Left$(strTempF(i), Len(strTempF(i)) - 4)
strNewAce = strNewAce & ".ace"
Name filepath1$ & "\" & strNewAce As filepath1$ & "\" & strOldAce
Next i
End If
Text1(1).Text = "*.*"
Rem ****************** Copying files back ****************

    cursouind = 1
filepath1$ = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.ace"
Label11.Caption = "Copying .ace files back to original folder"
DoEvents

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
   
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   strOrigFile = File1(cursouind).List(i)
   If FileExists(strOrigFilePath & "\" & strOrigFile) Then
   Kill strOrigFilePath & "\" & strOrigFile
   FileCopy fullpath$, strOrigFilePath & "\" & strOrigFile
   DoEvents
   Kill fullpath$
  End If
  End If
  Next i




Rem ******************************************************
'Call MsgBox("All selected Texture files have now been converted to DXT1 compression. The new files" _
'            & vbCrLf & "are in the Route_Riter\TempFiles folder, you may copy those required to your Route\Textures folder." _
'            , vbInformation, App.Title)

  MousePointer = 0
Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'ConvTGAToAce' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
'Resume Next
End Sub

Private Sub GetActDetails(strAct As String)
Dim NewFile As Integer, A$
Dim tempService As String
Dim i As Long, ii As Integer
Dim tempRoutePath As String, x As Integer
Dim svcExists As Boolean, Y As Integer, tfcExists As Boolean, j As Integer
Dim NameExists As Boolean, intEng As Integer, intWag As Integer
On Error GoTo Errtrap
lngAct = 1

ReDim Activities(0 To lngAct - 1)
 ReDim ActPath(0 To lngAct - 1)
 ReDim ActEng(0 To lngAct - 1, 0 To CHUNK)
 ReDim ActWag(0 To lngAct - 1, 0 To CHUNK)
 ReDim pPathName(0 To lngAct - 1, 0 To CHUNK)
 ReDim PConName(0 To lngAct - 1, 0 To CHUNK)
 ReDim PSvcName(0 To lngAct - 1, 0 To CHUNK)
 ReDim PTfcName(0 To lngAct - 1)

 Text1(0) = "*.*"
 Text1(0).Refresh
 
x = InStrRev(strAct, "\Activities")
tempRoutePath = Left$(strAct, x)
x = InStrRev(strAct, "\")
ActPath(0) = tempRoutePath & "Activities"
Activities(0) = Mid$(strAct, x + 1)

NewFile = FreeFile

 Open strAct For Input As #NewFile
  j = 0: i = 0
 Do While Not EOF(NewFile)
Line Input #NewFile, A$

 Rem ******** Get Players Consist ************
 x = InStr(A$, "Player_Service_Definition")
 If x > 0 Then
 tempService = Trim$(Mid$(A$, x + 27))
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If

tempService = tempService & ".srv"
 tempService = Trim$(tempService)

 Call CheckService(tempRoutePath & "Services\" & tempService, svcExists, i, j)

For ii = 1 To j
If PSvcName(i, ii) = tempService Then
NameExists = True
Exit For
End If
Next ii

If NameExists = False Then
 PSvcName(i, j) = tempService
 If j > UBound(PSvcName, 2) Then
     ReDim Preserve PSvcName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve pPathName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve PConName(0 To lngAct - 1, 0 To j + CHUNK)
      End If
 j = j + 1
 Else
 NameExists = False
 End If
 
 End If
 tempService = vbNullString
 Rem ************  Get Traffic Definition
 tempService = vbNullString

 x = InStr(A$, "Traffic_Definition")
 Y = InStr(A$, "Player_Traffic_Definition")

 If x > 0 And Y = 0 Then

 tempService = Trim$(Mid$(A$, x + 18))
 
 tempService = Trim$(tempService)
 
 If Left$(tempService, 1) = "(" Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ")" Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 tempService = Trim$(tempService)
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If

tempService = tempService & ".trf"
tempService = Trim$(tempService)
PTfcName(i) = tempService

 Call CheckTraffic(tempRoutePath & "Traffic\" & tempService, tfcExists, i)
 
 End If
 Rem *************** Get AI Traffic
 x = InStr(A$, "Service_Definition")
 Y = InStr(A$, "Player_Service")
 If x > 0 And Y = 0 Then
 booPlayer = False
 tempService = Trim$(Mid$(A$, x + 18))
 Y = InStrRev(tempService, " ")
 tempService = Left$(tempService, Y)
 tempService = Trim$(tempService)
 
 If Left$(tempService, 1) = "(" Then
 tempService = Mid$(tempService, 2)
 End If

 tempService = Trim$(tempService)
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 
tempService = tempService & ".srv"
tempService = Trim$(tempService)

 Call CheckService(tempRoutePath & "Services\" & tempService, svcExists, i, j)
 
 For ii = 1 To j
If PSvcName(i, ii) = tempService Then
NameExists = True
Exit For
End If
Next ii
If NameExists = False Then
 PSvcName(i, j) = tempService
 If j > UBound(PSvcName, 2) Then
     ReDim Preserve PSvcName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve pPathName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve PConName(0 To lngAct - 1, 0 To j + CHUNK)
      End If
 j = j + 1
 
 Else
 NameExists = False
 End If
 
 End If
 
 
 Loop
 Close #NewFile

 Call GetLooseActConsists(strAct, intEng, intWag)


frmActGrid.Show


     DoEvents
     
MousePointer = 0

Exit Sub
Errtrap:
Call MsgBox("Error in GetActDetails - " & Err.Description _
            & vbCrLf & "Processing " & strAct _
            , vbExclamation, App.Title)


Resume Next
End Sub


Private Sub CheckACESMS()
Dim tempPath As String
Dim GlobalPath As String, TertexPath As String, Ename As String
Dim MyString As String, AceTemp(1 To 100) As String
Dim strTemp As String, strTerr As String, strSpare As String

On Error GoTo selerr

MousePointer = 11
  Rem *********** Get the .ACE files ************
TertexPath = RoutePath & "\Terrtex"
GlobalPath = MSTSPath & "\Global\Shapes\"
  
 Rem ************** Find Sound Files ***************
 Call CheckDefaultSounds
 cursouind = 1
Drive1(cursouind).Drive = Left$(WorldPath, 2)
Dir1(cursouind).Path = WorldPath
Text1(cursouind).Text = "*.ws"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
   
   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   Rem ******************* New Bit *******************************
  
   Call CheckSoundSource(fullpath$)

   Rem ***********************************************
   Call CheckForSounds3(fullpath$)
   File1(cursouind).Selected(i) = False


   End If
   
   Next i
If FileExists(RoutePath & "\ttype.dat") Then
Call CheckForSounds3(RoutePath & "\ttype.dat")
End If

For i = 0 To SoundNumber
If Soundfile(i) <> vbNullString Then
If FileExists(SoundPath & "\" & Soundfile(i)) Then
  tempPath = SoundPath & "\" & Soundfile(i)
  
  ElseIf FileExists(GlobalSoundPath & "\" & Soundfile(i)) Then
  tempPath = GlobalSoundPath & "\" & Soundfile(i)
  Else
  Call LookForSound(Soundfile(i), "")
  tempPath = SoundPath & "\" & Soundfile(i)
  End If
  Call CheckForWav3(tempPath)
  End If
Next i

 
 

 
 Rem ************** Find terrain Textures ************


For i = 1 To numTerr


  SB1.Panels(2).Text = TerrTex2(i)
  
  If Not FileExists(TertexPath & "\" & TerrTex2(i)) Then
  
Call LookForTerrtex(TerrTex2(i))
End If
Next i


Label3:
Rem *********************** Check for Hazards


booHaz = True
If numHaz > 0 Then
For i = 0 To numHaz - 1
If HazShp2(i) = vbNullString Then
GoTo NextOne
End If
Call CompactCheckForS(RoutePath & "\" & HazShp2(i), booHaz)
DoEvents
If Hazard(i) <> vbNullString Then
fullpath$ = GlobalPath & Hazard(i)
If Not FileExists(fullpath$) Then
strReport = strReport & "Hazard " & Hazard(i) & " is missing from the Global\Shapes folder" & vbCrLf
GoTo NextOne
End If
If Right$(fullpath$, 1) = "\" Then
strReport = strReport & "Hazard " & Hazard(i) & " is missing from the Global\Shapes folder" & vbCrLf
GoTo NextOne
End If


Rem ***********
   strSpare = App.Path & "\TempFiles"
   NewFile2 = FreeFile
   x = InStrRev(fullpath$, "\")
   strTerr = Mid$(fullpath$, x + 1)
   
 Open fullpath$ For Binary As #NewFile2
    strTemp = String(2, " ")
    Get #NewFile2, , strTemp
 Close #NewFile2
 
 If Asc(Mid$(strTemp, 1, 1)) = 255 And Asc(Mid$(strTemp, 2, 1)) = 254 Then
 FileCopy fullpath$, strSpare & "\" & strTerr
 'MyString = ReadUniFile(fullpath$)
  Else
  MyMainString = vbNullString
  
   Ename = Left(GlobalPath, Len(GlobalPath) - 1)
   TokMode = 0
  Call DoDeComp2(strTerr, Ename, strSpare)

   'MyString = MyMainString
   End If
   
MyString = ReadUniFile(strSpare & "\" & strTerr)


      yy = 1
 Do
 
 yy = InStr(yy, MyString, "image (")
 If yy > 0 Then
 Z = InStr(yy, MyString, "(")
 zz = InStr(Z, MyString, ")")
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Mid$(strFName, 2)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
 intAce = intAce + 1
  AceTemp(intAce) = strFName
     yy = Z
     End If
    Loop While yy > 0
            For q = 1 To intAce
                    
              strNew = AceTemp(q)
               If Not FileExists(MSTSPath & "\Global\Textures\" & strNew) Then
strReport = strReport & "Hazard texture " & strNew & " is missing from the Global\Textures folder" & vbCrLf

End If
             Next q
        End If
 DoEvents


NextOne:
Next i
End If



MousePointer = 0

  
   Exit Sub
selerr:

Call MsgBox("Error " & Err & " " & Err.Description & " occurred while checking this route", vbExclamation, App.Title)


End Sub
Private Sub CheckSoundSource(strWS As String)
Dim Fnumber As Integer, strNew As String
Dim x As Long, Y As Long, zz As Long, xx As Long
Dim yy As Integer, Z As Long
Dim GlobalSoundPath As String, strTemp As String
Dim lx As Single, ly As Single, lz As Single
Dim strStart As String, strEnd As String
Dim R As Long, rr As Long, strFName As String, booChanged As Boolean

On Error GoTo Errtrap
GlobalSoundPath = MSTSPath & "\Sound\"
Fnumber = FreeFile
xx = 1


   strNew = ReadUniFile(strWS)
   Do
   x = InStr(xx, strNew, "Soundsource")
   If x > 0 Then
   Y = InStr(x, strNew, "Position")
   If Y > x And (Y - x) < 200 Then
   
   zz = InStr(Y, strNew, "(")
   xx = InStr(Y, strNew, ")")
   strTemp = Trim$(Mid$(strNew, zz + 1, xx - Z))
   yy = InStr(strTemp, " ")
   Z = InStr(yy + 1, strTemp, " ")
   
   lx = Val(Left$(strTemp, yy - 1))
   ly = Val(Mid$(strTemp, yy + 1, Z - yy))
   lz = Val(Mid$(strTemp, Z + 1))
        If lx <= -1024 Or lx >= 1024 Or lz <= -1024 Or lz >= 1024 Then
        If booFixSound = True Then
        booChanged = True
        R = InStr(xx, strNew, ".sms")
        rr = InStrRev(strNew, "(", R + 4)
        strFName = Mid$(strNew, rr + 1, (R + 4) - rr)
        strFName = Trim$(strFName)
        If lx <= -1024 Then lx = -1023
        If lx >= 1024 Then lx = 1023
        If lz <= -1024 Then lz = -1023
        If lz >= 1024 Then lz = 1023
        strTemp = "Position ( " & Str(lx) & " " & Str(ly) & " " & Str(lz) & " )"
        strStart = Left$(strNew, Y - 1)
        strEnd = Mid$(strNew, xx + 1)
        strNew = strStart & strTemp & strEnd
        strReport = strReport & "Soundsource " & strFName & " contained invalid position data in " & strWS & " the tile was modified." & vbCrLf
        GoTo NextBit
End If
        strReport = strReport & "Soundsource " & strFName & " contained invalid position data in " & strWS & " Run TsUtils option to fix Soundsources" & vbCrLf
NextBit:
        End If
   
   End If
   End If
   'xx = rr
   Loop While x > 0
   If booChanged = True Then
   If FileExists(strWS & ".bak") Then
   Kill strWS & ".bak"
   End If

   Name strWS As strWS & ".bak"
   DoEvents
   Call WriteUniFile(strWS, strNew)
   End If

 Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'CheckSoundSource' while checking " & strWS & " please advise" _
            & vbCrLf & "Support with details of operation being processed. Include screen print if possible" _
            , vbExclamation, App.Title)
            Resume Next
 
   
End Sub


Private Sub CheckTrk(strRoute As String)
Dim NewFile As Integer, strNew As String, x As Integer, Y As Integer, Z As Integer
Dim i As Integer, EnvTag(1 To 12), OldEnvDir As String, booExists As Boolean
Dim strRS As String

On Error GoTo Errtrap
EnvTag(1) = "springclear ("
EnvTag(2) = "springrain ("
EnvTag(3) = "springsnow ("
EnvTag(4) = "summerclear ("
EnvTag(5) = "summerrain ("
EnvTag(6) = "summersnow ("
EnvTag(7) = "autumnclear ("
EnvTag(8) = "autumnrain ("
EnvTag(9) = "autumnsnow ("
EnvTag(10) = "winterclear ("
EnvTag(11) = "winterrain ("
EnvTag(12) = "wintersnow ("

NewFile = FreeFile
Open strRoute For Input As #NewFile
Do While Not EOF(NewFile)
   
   Line Input #NewFile, strNew
  
   strNew = Trim$(strNew)
   For i = 1 To 12
   x = InStr(strNew, EnvTag(i))
   If x > 0 Then
    Y = InStr(x, strNew, "(")
    Z = InStr(Y, strNew, ")")
    OldEnv(i) = Trim$(Mid$(strNew, Y + 1, Z - Y - 1))
    
    End If
 Next i
 
 x = InStr(strNew, "RouteStart")
 If x > 0 Then
 Y = InStr(x, strNew, "(")
    Z = InStr(Y, strNew, ")")
    strRS = Trim$(Mid$(strNew, Y + 1, Z - Y - 1))
    x = InStr(strRS, " ")
    RStart(1) = Val(Left$(strRS, x - 1))
    Y = InStr(x + 1, strRS, " ")
    RStart(2) = Val(Mid$(strRS, x + 1, Y - x - 1))
    Z = InStr(Y + 1, strRS, " ")
    RStart(3) = Val(Mid$(strRS, Y + 1, Z - Y - 1))
    RStart(4) = Val(Mid$(strRS, Z + 1))
    
    
    End If
    Loop
  Close #NewFile
 
 OldEnvDir = EnvPath & "\OldEnvFiles"
If Not DirExists(OldEnvDir) Then
  MkDir (OldEnvDir)
  End If
  For i = 1 To 12
  FileCopy EnvPath & "\" & OldEnv(i), OldEnvDir & "\" & OldEnv(i)
    Next i
    FileCopy EnvPath & "\" & OldEnv(1), EnvPath & "\SpringClear.env"
    FileCopy EnvPath & "\" & OldEnv(2), EnvPath & "\SpringRain.env"
    FileCopy EnvPath & "\" & OldEnv(3), EnvPath & "\SpringSnow.env"
    FileCopy EnvPath & "\" & OldEnv(4), EnvPath & "\SummerClear.env"
    FileCopy EnvPath & "\" & OldEnv(5), EnvPath & "\SummerRain.env"
    FileCopy EnvPath & "\" & OldEnv(6), EnvPath & "\SummerSnow.env"
    FileCopy EnvPath & "\" & OldEnv(7), EnvPath & "\AutumnClear.env"
    FileCopy EnvPath & "\" & OldEnv(8), EnvPath & "\AutumnRain.env"
    FileCopy EnvPath & "\" & OldEnv(9), EnvPath & "\AutumnSnow.env"
    FileCopy EnvPath & "\" & OldEnv(10), EnvPath & "\WinterClear.env"
    FileCopy EnvPath & "\" & OldEnv(11), EnvPath & "\WinterRain.env"
    FileCopy EnvPath & "\" & OldEnv(12), EnvPath & "\WinterSnow.env"
    DoEvents
    If booExists = True Then GoTo CarryON
    For i = 1 To 12
    If FileExists(EnvPath & "\" & OldEnv(i)) Then
    Kill EnvPath & "\" & OldEnv(i)
    End If
    Next
CarryON:
   OldEnv(1) = "SpringClear.env"
   OldEnv(2) = "SpringRain.env"
   OldEnv(3) = "SpringSnow.env"
   OldEnv(4) = "SummerClear.env"
   OldEnv(5) = "SummerRain.env"
   OldEnv(6) = "SummerSnow.env"
   OldEnv(7) = "AutumnClear.env"
   OldEnv(8) = "AutumnRain.env"
   OldEnv(9) = "AutumnSnow.env"
   OldEnv(10) = "WinterClear.env"
   OldEnv(11) = "WinterRain.env"
   OldEnv(12) = "WinterSnow.env"
 
   Exit Sub
Errtrap:
   If Err = 70 Then
   booExists = True
   Resume Next
   
   End If
    
    
End Sub

Private Sub ClearSetup()
On Error GoTo Errtrap
Dim booExists As Boolean, NewRouteName As String, Filpath1$, OldRouteName As String

cursouind = 0
Filpath1$ = App.Path & "\setupfiles"
If FileExists(App.Path & "\setupfiles\master.ref") Then
Kill App.Path & "\setupfiles\master.ref"
End If
If FileExists(App.Path & "\setupfiles\InstallMe.bat") Then
Kill App.Path & "\setupfiles\InstallMe.bat"
End If

MousePointer = 11
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = RoutePath
DoEvents
'RoutePath = File1(cursouind).path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If

'Call MakeReadWrite(RoutePath)
'If Not FileExists(RoutePath & "\" & RouteName & ".ref") Then
Text1(0) = "*.trk"
RouteName = File1(cursouind).List(i)
OldRouteName = RouteName
Call CheckForSMS(RoutePath & "\" & RouteName)

Call FindRouteName(RoutePath & "\" & RouteName, NewRouteName, booExists)
If booExists = False Then

MsgBox Lang(339) & vbCr & Lang(340), 16, Lang(341)
'********************
Exit Sub
Else
RouteName = NewRouteName
End If

Call IsItElectric(RoutePath & "\" & OldRouteName, booElectric)

'End If
OriginalRef = RoutePath & "\" & RouteName & ".ref"
If FileExists(OriginalRef) Then
FileCopy RoutePath & "\" & RouteName & ".ref", App.Path & "\setupfiles\master.ref"
Else
flagNoRef = True
Call MsgBox(Lang(385) & vbCrLf & Lang(386), vbExclamation, "No .ref file")
FileCopy App.Path & "\stuffit.ref", App.Path & "\setupfiles\master.ref"
End If
RouteListed = True
TexturePath = RoutePath & "\Textures"
TexSnowPath = RoutePath & "\Textures\Snow"
TexNightPath = RoutePath & "\Textures\Night"
TexAutPath = RoutePath & "\Textures\Autumn"
TexAutSnowPath = RoutePath & "\Textures\AutumnSnow"
TexSprPath = RoutePath & "\Textures\Spring"
TexSprSnowPath = RoutePath & "\Textures\SpringSnow"
TexWinPath = RoutePath & "\Textures\Winter"
TexWinSnowPath = RoutePath & "\Textures\WinterSnow"
TilePath = RoutePath & "\Tiles"
ShapePath = RoutePath & "\Shapes"
SoundPath = RoutePath & "\Sound"
WorldPath = RoutePath & "\World"
EnvPath = RoutePath & "\Envfiles"

Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
    cursouind = 1

    Call KillSpare("*.t")
DoEvents
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************
MousePointer = 0
Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'ClearSetup' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub

Private Function ConvertAct(CompleteFilePath As String, flagway As Integer, strNew As String) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertAct = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer, Y As Integer

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
TryAgain:
x = InStr(MyString, "RouteID")
Y = InStr(x, MyString, vbCr)
xx = InStrRev(MyString, ")", Y)
If x = 0 Then GoTo CarryON
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xx + 1)
MyString = strStart & "RouteID ( " & ChrW$(34) & strNew & ChrW$(34) & " )" & strEnd

CarryON:

End If


The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertAct = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function

Private Function ConvertTrk(CompleteFilePath As String, flagway As Integer, strMid As String) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertTrk = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
x = InStr(MyString, "Environment (")

xx = InStr(x, MyString, "WinterSnow (")
xy = InStr(xx, MyString, ")")
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, xy + 1)
MyString = strStart & strMid & strEnd

CarryON:
End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertTrk = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function







Private Function ConvertSun(CompleteFilePath As String, flagway As Integer, Season As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertSun = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer, Y As Integer, strFadeIn As String, strFadeOut As String

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
x = InStr(MyString, "world_sky_satellite_rise_time")
xx = InStr(x, MyString, "(")
Y = InStr(xx, MyString, ")")
strStart = Left$(MyString, xx + 1)
strEnd = Mid$(MyString, Y - 1)
MyString = strStart & RiseSet(Season * 2 - 1) & strEnd
DoEvents
x = InStr(Y, MyString, "world_sky_satellite_set_time")
xx = InStr(x, MyString, "(")
Y = InStr(xx, MyString, ")")
strStart = Left$(MyString, xx + 1)
strEnd = Mid$(MyString, Y - 1)
MyString = strStart & RiseSet(Season * 2) & strEnd
Call GetFade(RiseSet(Season * 2 - 1), RiseSet(Season * 2), strFadeOut, strFadeIn)

DoEvents
x = InStr(Y, MyString, "world_sky_satellite_rise_time")
xx = InStr(x, MyString, "(")
Y = InStr(xx, MyString, ")")
strStart = Left$(MyString, xx + 1)
strEnd = Mid$(MyString, Y - 1)
MyString = strStart & MoonSet(Season * 2 - 1) & strEnd
DoEvents
x = InStr(Y, MyString, "world_sky_satellite_set_time")
xx = InStr(x, MyString, "(")
Y = InStr(xx, MyString, ")")
strStart = Left$(MyString, xx + 1)
strEnd = Mid$(MyString, Y - 1)
MyString = strStart & MoonSet(Season * 2) & strEnd
DoEvents
Rem **************** Fades
x = InStr(MyString, "world_sky_layer_fadein")
xx = InStr(x, MyString, "(")
Y = InStr(xx, MyString, ")")
strStart = Left$(MyString, xx + 1)
strEnd = Mid$(MyString, Y - 1)
MyString = strStart & strFadeIn & strEnd
DoEvents
x = InStr(Y, MyString, "world_sky_layer_fadeout")
xx = InStr(x, MyString, "(")
Y = InStr(xx, MyString, ")")
strStart = Left$(MyString, xx + 1)
strEnd = Mid$(MyString, Y - 1)
MyString = strStart & strFadeOut & strEnd
DoEvents

End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertSun = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function

Private Sub CountStock2()
Dim FirstPath As String, DirCount As Integer, TrainPath As String
On Error GoTo Errtrap

Trainspath = MSTSPath & "\Trains\"
If lngLoco > 0 Then
For i = 0 To lngLoco - 1
Locomotives(i) = vbNullString
LocoPath(i) = vbNullString
Next i
End If
If lngWagons > 0 Then
For i = 0 To lngWagons - 1
Wagons(i) = vbNullString
Wagpath(i) = vbNullString
Next i
End If
lngLoco = 0
lngWagons = 0
cursouind = 0
TrainPath = Trainspath & "trainset"
Dir1(cursouind).Path = TrainPath

If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
        Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If

    FirstPath = Dir1(cursouind).Path
    DirCount = Dir1(cursouind).ListCount
    
    result = DirDiverLoco2(FirstPath, DirCount, "")
    File1(cursouind).Path = Dir1(cursouind).Path
   
If booAbort = True Then


MousePointer = 0
Exit Sub
End If

If booAbort = True Then



MousePointer = 0
Exit Sub
End If
ReDim Preserve Locomotives(0 To lngLoco)
           ReDim Preserve LocoPath(0 To lngLoco - 1)
           ReDim Preserve LocoName(0 To lngLoco - 1)
           ReDim Preserve LocoCoup(0 To lngLoco - 1)
           ReDim Preserve LocoFCoup(0 To lngLoco - 1)
           ReDim Preserve LocoBrake(0 To lngLoco - 1)
           ReDim Preserve LocoType(0 To lngLoco - 1)
           ReDim Preserve Wagons(0 To lngWagons - 1)
           ReDim Preserve Wagpath(0 To lngWagons - 1)
           ReDim Preserve WagonName(0 To lngWagons - 1)
           ReDim Preserve WagCoup(0 To lngWagons - 1)
           ReDim Preserve WagFCoup(0 To lngWagons - 1)
           ReDim Preserve WagBrake(0 To lngWagons - 1)
           ReDim Preserve WagType(0 To lngWagons - 1)
           ReDim Preserve WagRigid(0 To lngWagons - 1)
           ReDim Preserve LocoRigid(0 To lngLoco - 1)
           ReDim Preserve WagFRigid(0 To lngWagons - 1)
           ReDim Preserve LocoFRigid(0 To lngLoco - 1)
           ReDim Preserve LocoSMS(0 To lngLoco - 1)
Label7(0).Caption = Str(lngLoco)
Label7(1).Caption = Str(lngWagons)
Exit Sub
Errtrap:
Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in  'CountStock2'  please advise" _
                       & vbCrLf & " Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
    Exit Sub
    End Select
End Sub

Private Sub CountStock()
Dim FirstPath As String, DirCount As Integer, TrainPath As String, i As Long

On Error GoTo Errtrap

If lngLoco > 0 Then
For i = 0 To lngLoco - 1
Locomotives(i) = vbNullString
LocoPath(i) = vbNullString
Next i
End If
If lngWagons > 0 Then
For i = 0 To lngWagons - 1
Wagons(i) = vbNullString
Wagpath(i) = vbNullString
Next i
End If
lngLoco = 0
lngWagons = 0
cursouind = 0
TrainPath = MSTSPath & "\trains\trainset"
Dir1(cursouind).Path = TrainPath

If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
        Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If

    FirstPath = Dir1(cursouind).Path
    DirCount = Dir1(cursouind).ListCount
    
    result = DirDiverLoco2(FirstPath, DirCount, "")
    File1(cursouind).Path = Dir1(cursouind).Path
  


If booAbort = True Then
MousePointer = 0
Exit Sub
End If
If lngLoco > 0 Then
           ReDim Preserve Locomotives(0 To lngLoco - 1)
           ReDim Preserve LocoPath(0 To lngLoco - 1)
           ReDim Preserve LocoName(0 To lngLoco - 1)
           ReDim Preserve LocoCoup(0 To lngLoco - 1)
           ReDim Preserve LocoFCoup(0 To lngLoco - 1)
           ReDim Preserve LocoBrake(0 To lngLoco - 1)
           ReDim Preserve LocoType(0 To lngLoco - 1)
           ReDim Preserve LocoRigid(0 To lngLoco - 1)
           ReDim Preserve LocoFRigid(0 To lngLoco - 1)
           ReDim Preserve LocoSMS(0 To lngLoco - 1)
 End If
 
 If lngWagons > 0 Then
           ReDim Preserve Wagons(0 To lngWagons - 1)
           ReDim Preserve Wagpath(0 To lngWagons - 1)
           ReDim Preserve WagonName(0 To lngWagons - 1)
           ReDim Preserve WagCoup(0 To lngWagons - 1)
           ReDim Preserve WagCoup(0 To lngWagons - 1)
           ReDim Preserve WagBrake(0 To lngWagons - 1)
           ReDim Preserve WagType(0 To lngWagons - 1)
           ReDim Preserve WagRigid(0 To lngWagons - 1)
           ReDim Preserve WagFRigid(0 To lngWagons - 1)
 End If
           
Label7(0).Caption = Str(lngLoco)
Label7(1).Caption = Str(lngWagons)
Exit Sub
Errtrap:


Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in  'CountStock'  lngLoco=" & lngLoco & " lngWagons=" & lngWagons _
                       & vbCrLf & " please advise Support with details" _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
      'Resume Next
        
    Exit Sub
    End Select
End Sub


Private Sub GetFade(strRise As String, strSet, strFadeOut As String, strFadeIn As String)
Dim strhr As String, hr2 As Integer, hr3 As Integer, strHr2 As String, strHr3 As String

strhr = Left$(strRise, 2)
hr2 = Val(strhr) - 1
hr3 = Val(strhr) + 1
strHr2 = Str(hr2)
strHr2 = Trim$(strHr2)
If Len(strHr2) = 1 Then
strHr2 = "0" & strHr2
End If
strHr3 = Str(hr3)
strHr3 = Trim$(strHr3)
If Len(strHr3) = 1 Then
strHr3 = "0" & strHr3
End If
strFadeOut = strHr2 & Mid$(strRise, 3) & " " & strHr3 & Mid$(strRise, 3)

strhr = Left$(strSet, 2)
hr2 = Val(strhr) - 1
hr3 = Val(strhr) + 1
strHr2 = Str(hr2)
strHr2 = Trim$(strHr2)
If Len(strHr2) = 1 Then
strHr2 = "0" & strHr2
End If
strHr3 = Str(hr3)
strHr3 = Trim$(strHr3)
If Len(strHr3) = 1 Then
strHr3 = "0" & strHr3
End If
strFadeIn = strHr2 & Mid$(strSet, 3) & " " & strHr3 & Mid$(strSet, 3)


End Sub



Private Sub ListLooseActConsists(CFilepath As String)
Dim strNew As String
Dim x As Long, j As Integer, Engpath As String, Engname As String
Dim Wagonpath As String, Wagname As String
Dim Z As Long, strnew3 As String, booEntry As Boolean
On Error GoTo Errtrap


j = 1
strNew = ReadUniFile(CFilepath)

strReport = strReport & CFilepath & vbCrLf
x = 1
Do
  itExists = False
  
x = InStr(x, strNew, "EngineData")
If x = 0 Then Exit Do
Z = InStr(x, strNew, vbLf)
strnew3 = Mid$(strNew, x, Z - x)
Call CheckEngineData(strnew3, Engname, Engpath, booEntry)


    Engname = Engname & ".eng"
 strReport = strReport & Engpath & "\" & Engname & vbCrLf
itExists = booEntry

   j = j + 1
   
   x = x + 1
   Loop
  
 
   
  x = 1
Do
  itExists = False
    x = InStr(x, strNew, "WagonData")
If x = 0 Then Exit Do
Z = InStr(x, strNew, vbLf)
strnew3 = Mid$(strNew, x, Z - x)
Call CheckWagonData(strnew3, Wagname, Wagonpath, booEntry)

   Wagname = Wagname & ".wag"
   strReport = strReport & Wagonpath & "\" & Wagname & vbCrLf


   x = x + 1
   Loop

Exit Sub
Errtrap:


Call MsgBox("An error " & Err.Description & " occurred in subroutine 'ListLooseActConsists' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
  


End Sub

Private Sub MiniLooseActConsists(CFilepath As String)
Dim strNew As String
Dim x As Long, j As Integer, Engpath As String, Engname As String
Dim Wagonpath As String, Wagname As String
Dim Z As Long, strnew3 As String, booEntry As Boolean
On Error GoTo Errtrap


j = 1
strNew = ReadUniFile(CFilepath)

x = 1
Do
  itExists = False
  
x = InStr(x, strNew, "EngineData")
If x = 0 Then Exit Do
Z = InStr(x, strNew, vbLf)
strnew3 = Mid$(strNew, x, Z - x)
Call CheckEngineData(strnew3, Engname, Engpath, booEntry)


    Engname = Engname & ".eng"
    
itExists = booEntry
If Not DirExists(MSTSPath & "\Trains\Trainset\" & Engpath) Then
MkDir MSTSPath & "\Trains\Trainset\" & Engpath
MkDir MSTSPath & "\Trains\Trainset\" & Engpath & "\CabView"
MkDir MSTSPath & "\Trains\Trainset\" & Engpath & "\Sound"


If DirExists(strTrainset & Engpath) Then
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & Engpath & "\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\" & Engpath & ChrW$(34) & " /S /I" & vbCrLf
Else
strReport = strReport & "Folder not found:- " & strTrainset & Engpath & vbCrLf
End If
End If
   j = j + 1
   
   x = x + 1
   Loop
  
 
   
  x = 1
Do
  itExists = False
    x = InStr(x, strNew, "WagonData")
If x = 0 Then Exit Do
Z = InStr(x, strNew, vbLf)
strnew3 = Mid$(strNew, x, Z - x)
Call CheckWagonData(strnew3, Wagname, Wagonpath, booEntry)

   Wagname = Wagname & ".wag"
   If Not DirExists(MSTSPath & "\Trains\Trainset\" & Wagonpath) Then
MkDir MSTSPath & "\Trains\Trainset\" & Wagonpath
If DirExists(strTrainset & Wagonpath) Then
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & strTrainset & Wagonpath & "\*.*" & ChrW$(34) & " " & ChrW$(34) & MSTSPath & "\Trains\Trainset\" & Wagonpath & ChrW$(34) & " /S /I" & vbCrLf
Else
strReport = strReport & "Folder not found:- " & strTrainset & Wagonpath & vbCrLf
End If
End If
  

   x = x + 1
   Loop
 
Exit Sub
Errtrap:


Call MsgBox("An error " & Err & " occurred in subroutine 'MiniLooseActConsists' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
  


End Sub


Private Sub GetLooseActConsists(CFilepath As String, k As Integer, k1 As Integer)
Dim strNew As String
Dim x As Long, j As Integer, Engpath As String, Engname As String
Dim Wagonpath As String, Wagname As String
Dim ActName As String, jj As Integer
Dim Z As Long, strnew3 As String, booEntry As Boolean
On Error GoTo Errtrap


x = InStrRev(CFilepath, "\")
ActName = Mid$(CFilepath, x + 1)



j = 1: jj = 1
strNew = ReadUniFile(CFilepath)

x = 1: k = 0
Do
  itExists = False
  
    x = InStr(x, strNew, "EngineData")
If x = 0 Then Exit Do
Z = InStr(x, strNew, vbLf)
strnew3 = Mid$(strNew, x, Z - x)
Call CheckEngineData(strnew3, Engname, Engpath, booEntry)
k = k + 1

    Engname = Engname & ".eng"
   
itExists = booEntry
   j = j + 1
   If itExists = False Then
  
   strReport = strReport & Lang(560) & Engpath & "\" & Engname & Lang(561) & ActName & vbCrLf
  ' Else
   frmStock.Grid3.AddItem ActName & vbTab & Engpath & "\" & Engname
   End If
   x = x + 1
   Loop
  
 
   
  x = 1: k1 = 0
Do
  itExists = False
    x = InStr(x, strNew, "WagonData")
If x = 0 Then Exit Do
Z = InStr(x, strNew, vbLf)
strnew3 = Mid$(strNew, x, Z - x)
Call CheckWagonData(strnew3, Wagname, Wagonpath, booEntry)
k1 = k1 + 1
    


   Wagname = Wagname & ".wag"
   
   jj = jj + 1
itExists = booEntry
  If itExists = False Then
 
   strReport = strReport & Lang(560) & Wagonpath & "\" & Wagname & Lang(561) & ActName & vbCrLf
   'Else
   frmStock.Grid3.AddItem ActName & vbTab & Wagonpath & "\" & Wagname
   End If
   x = x + 1
   Loop
 
Exit Sub
Errtrap:


Call MsgBox("An error " & Err & " occurred in subroutine 'GetLooseActConsists' please advise" _
            & vbCrLf & "Support, error in " & CFilepath & " ." _
            , vbExclamation, App.Title)
Resume Next


End Sub


Private Sub GetLooseConsists(CFilepath As String, intCon As Integer, k As Integer, k1 As Integer)
Dim strNew As String
Dim x As Long, j As Integer, Engpath As String, Engname As String
Dim Wagonpath As String, Wagname As String
Dim ActName As String
Dim Z As Long, strnew3 As String, booEntry As Boolean
On Error GoTo Errtrap


x = InStrRev(CFilepath, "\")
ActName = Mid$(CFilepath, x + 1)

strNew = ReadUniFile(CFilepath)

x = 1: k = 0
Do
  itExists = False
  
    x = InStr(x, strNew, "EngineData")
If x = 0 Then Exit Do
Z = InStr(x, strNew, vbLf)
strnew3 = Mid$(strNew, x, Z - x)
Call CheckEngineData(strnew3, Engname, Engpath, booEntry)
k = k + 1
If k > UBound(ActEng, 2) Then
     ReDim Preserve ActEng(0 To lngAct - 1, 0 To k + CHUNK)
           End If
    Engname = Engname & ".eng"
    For j = 0 To lngLoco - 1
   If Engname = Locomotives(j) Then
   ActEng(intCon, k) = j
   itExists = True
   Exit For
  End If
   Next j
   
   If itExists = False Then
  
   strReport = strReport & Lang(560) & Engpath & "\" & Engname & Lang(561) & ActName & vbCrLf
   'Else
   frmStock.Grid3.AddItem ActName & vbTab & Engpath & "\" & Engname
   End If
   x = x + 1
   Loop
  
 
   
  x = 1: k1 = 0
Do
  itExists = False
    x = InStr(x, strNew, "WagonData")
If x = 0 Then Exit Do
Z = InStr(x, strNew, vbLf)
strnew3 = Mid$(strNew, x, Z - x)
Call CheckWagonData(strnew3, Wagname, Wagonpath, booEntry)
k1 = k1 + 1
 If k1 > UBound(ActWag, 2) Then
     ReDim Preserve ActWag(0 To lngAct - 1, 0 To k1 + CHUNK)
           End If


   Wagname = Wagname & ".wag"
     For j = 0 To lngWagons - 1
   If Wagname = Wagons(j) Then
   itExists = True
   ActWag(intCon, k1) = j
   Exit For
    End If
  Next j
  
  If itExists = False Then
  
   strReport = strReport & Lang(560) & Wagonpath & "\" & Wagname & Lang(561) & ActName & vbCrLf
   'Else
   frmStock.Grid3.AddItem ActName & vbTab & Wagonpath & "\" & Wagname
   End If
   x = x + 1
   Loop
 
Exit Sub
Errtrap:



Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in 'GetLooseConsists' please advise" _
                       & vbCrLf & "checking " & ActName _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
       
    Resume Next
        Case vbCancel
       
    Exit Sub
    End Select

End Sub






Private Sub GetSeasons(CurrentSD As String, strESD As String)
Dim NewFile As Integer, x As Integer, xx As Integer, xy As Integer
Dim strSearch As String


strSearch = "ESD_A"
NewFile = FreeFile
  Open CurrentSD For Input As #NewFile
  Do While Not EOF(NewFile)
  Line Input #NewFile, A$
  
  x = InStr(A$, strSearch)
      If x = 0 Then GoTo TryAgain
      xx = InStr(x, A$, "(")
      xy = InStr(xx, A$, ")")
      strESD = Mid$(A$, xx + 1, xy - xx - 1)
      strESD = Trim$(strESD)
   Exit Do
   
TryAgain:
      Loop
      Close #NewFile
      
End Sub


Private Sub IsItMissing(sName As Variant, booMiss As Boolean)
Dim NewFile As Integer, strNew As String, x As Integer

x = InStrRev(sName, "\")
sName = Mid$(sName, x + 1)
NewFile = FreeFile
If FileExists(RoutePath & "\telepole.dat") Then
Open RoutePath & "\telepole.dat" For Input As NewFile
Do While Not EOF(NewFile)
   
   Line Input #NewFile, strNew
  
   strNew = Trim$(strNew)
   x = InStr(strNew, sName)
   If x > 0 Then
   booMiss = True
   Exit Sub
   End If
   Loop
   
Close NewFile
End If
NewFile = FreeFile
If FileExists(RoutePath & "\speedpost.dat") Then
Open RoutePath & "\speedpost.dat" For Input As NewFile
Do While Not EOF(NewFile)
   
   Line Input #NewFile, strNew
  
   strNew = Trim$(strNew)
   x = InStr(strNew, sName)
   If x > 0 Then
   booMiss = True
   Exit Sub
   End If
   Loop
   
Close NewFile
End If

End Sub

Private Sub KillRubbish()
Dim i As Integer


cursouind = 0
MousePointer = 11
Drive1(cursouind).Drive = Left$(TilePath, 2)

Dir1(cursouind).Path = TilePath
Text1(cursouind).Text = "*.bk"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i

'******************************


Dir1(cursouind).Path = RoutePath
Text1(cursouind).Text = "*.bk"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
Dir1(cursouind).Path = RoutePath
Text1(cursouind).Text = "*.*bk"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
Dir1(cursouind).Path = RoutePath
Text1(cursouind).Text = "*.*backup"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
Dir1(cursouind).Path = RoutePath
Text1(cursouind).Text = "*.bak"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
Dir1(cursouind).Path = ShapePath
Text1(cursouind).Text = "*.bak"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
Dir1(cursouind).Path = WorldPath
Text1(cursouind).Text = "*w.bk"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
Dir1(cursouind).Path = WorldPath
Text1(cursouind).Text = "*.bak"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
'************************
Text1(cursouind).Text = "*.*"
If FileExists(ShapePath & "\jp1shinjukust400.max") Then
Kill ShapePath & "\jp1shinjukust400.max"
End If
MousePointer = 0

Exit Sub

End Sub

Private Sub KillSpare2(strKill As String)
Dim SparePath As String, i As Integer

Close
cursouind = 1
SparePath = App.Path & "\SetupFiles"
Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = strKill
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1) = "*.*"

End Sub

Private Sub KillSpare(strKill As String)
Dim SparePath As String, i As Integer

Close
cursouind = 1
SparePath = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = strKill
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 DoEvents

 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1) = "*.*"

End Sub

Private Sub ListAce()
Dim i As Integer
Dim strOrigFile As String
Dim intType As Integer, intColor As Integer, strType As String
Dim intSize As Integer, intSizeY As Integer
On Error GoTo Errtrap

MousePointer = 11

 cursouind = 0
ACEPath = Dir1(0).Path

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
  
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   strOrigFile = File1(cursouind).List(i)
   If Right$(strOrigFile, 4) <> ".ace" Then
   MousePointer = 0
   Call MsgBox(Lang(393) & strOrigFile & Lang(405), vbExclamation, Lang(404))
    GoTo NextOne
    End If

   Call readFile(fullpath$, bdata())
  

intSize = bdata(8) + (bdata(9) * 256)
intSizeY = bdata(12) + (bdata(13) * 256)

 intType = bdata(16)
 intColor = bdata(20)
 
 
 Select Case intType
 Case 14, 16
 strType = "24 bit "
 If intColor = 4 Then
 strType = "24 bit + trans"
 End If
Case 17
strType = "32 bit "
If intColor = 5 Then
strType = "32 bit + 8 bit trans"
End If
Case 18
strType = "DXT1 "
If intColor = 3 Then
strType = "DXT1 opaque"
ElseIf intColor = 4 Then
strType = "DXT1 with 1 bit alpha"
End If
Case Else
strType = "Unknown "
End Select
Compressed:

   'strReport = strReport & strOrigFile & " - " & str(intSize) & " * " & str(intSizeY) & " - " & strType & vbtab & intType & vbtab & intColor & vbCrLf
   
   strReport = strReport & strOrigFile & " - " & Str(intSize) & " * " & Str(intSizeY) & " - " & strType & vbCrLf
   End If
   
NextOne:
   Next i
   'Close
 MousePointer = 0
frmReport.Rich1.Text = strReport
     frmReport.Show 1
     
     DoEvents
     
 cursouind = 0

Text1(0).Text = "*.*"

DoEvents
Text1(1).Text = "*.*"
 strReport = vbNullString
 
Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'ListAce' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub

Private Sub LookForSD(strShape As String, booNoShape As Boolean)
Dim booMiss As Boolean, strShape2 As String, strShape3 As String
Dim x As Integer, GlobalPath As String

strShape2 = Left$(strShape, Len(strShape) - 1)
x = InStrRev(strShape, "\")
strShape3 = Mid$(strShape, x + 1)
GlobalPath = MSTSPath & "\Global\Shapes\"
booNoShape = False

Call IsItMissing(strShape2, booMiss)
If booMiss = True Then
booMiss = False
Exit Sub
End If
If booMiss = False Then
'strTempPath = GlobalPath
'If FileExists(strTempPath & strShape3) Then
'GoTo ShowMessage
'End If
strTempPath = strComPath & "\shapes\"
If FileExists(strTempPath & strShape3) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\europe1\shapes\"
If FileExists(strTempPath & strShape3) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\europe2\shapes\"
If FileExists(strTempPath & strShape3) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\japan1\shapes\"
If FileExists(strTempPath & strShape3) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\japan2\shapes\"
If FileExists(strTempPath & strShape3) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\usa1\shapes\"
If FileExists(strTempPath & strShape3) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\usa2\shapes\"
If FileExists(strTempPath & strShape3) Then
GoTo ShowMessage
End If
End If
strReport = strReport & Lang(562) & strShape3 & Lang(563) & vbCrLf

booNoShape = True

Exit Sub
ShowMessage:

If intResponse = 4 Then
    strReport = strReport & Lang(562) & strShape3 & Lang(567) & vbCrLf
End If
If intResponse = 2 Then

FileCopy strTempPath & strShape3, ShapePath & "\" & strShape3
strReport = strReport & Lang(562) & strShape3 & Lang(566) & vbCrLf
End If
If intResponse = 0 Then
strResponse = Lang(562) & strShape3 & Lang(564) & vbCrLf & Lang(565)
frmDialog.Label1.Caption = strResponse
frmDialog.Show 1

     DoEvents
   

If intResponse = 1 Then
intResponse = 0

FileCopy strTempPath & strShape3, ShapePath & "\" & strShape3
strReport = strReport & Lang(562) & strShape3 & Lang(566) & vbCrLf


ElseIf intResponse = 2 Then

FileCopy strTempPath & strShape3, ShapePath & "\" & strShape3
strReport = strReport & Lang(562) & strShape3 & Lang(566) & vbCrLf

ElseIf intResponse = 3 Then
intResponse = 0
    strReport = strReport & Lang(562) & strShape3 & Lang(567) & vbCrLf
ElseIf intResponse = 4 Then
    strReport = strReport & Lang(562) & strShape3 & Lang(567) & vbCrLf
End If
End If
End Sub

Private Sub LookForSound(strSound As String, strCaller As String)
Dim strTempPath As String, i As Integer

strTempPath = strComPath & "\sound\"
If FileExists(strTempPath & strSound) Then
GoTo ShowMessage
End If
'strTempPath = MSTSPath & "\sound\"
'If FileExists(strTempPath & strSound) Then
'GoTo ShowMessage
'End If
For i = 0 To NumRoutes - 1
strTempPath = AllRoutes(i) & "\Sound\"
If FileExists(strTempPath & strSound) Then
GoTo ShowMessage
End If
Next i


strReport = strReport & Lang(568) & strSound & Lang(563) & " Called by " & strCaller & vbCrLf

Exit Sub
ShowMessage:

Rem*************************
If intResponse = 0 Then
strResponse = Lang(568) & strSound & Lang(564) & vbCrLf & Lang(565)
frmDialog.Label1.Caption = strResponse
frmDialog.Command1.Visible = False
frmDialog.Show 1

     DoEvents
     
End If
If intResponse = 1 Then
intResponse = 0
FileCopy strTempPath & strSound, SoundPath & "\" & strSound
strReport = strReport & Lang(568) & strSound & Lang(569) & strTempPath & strSound & vbCrLf
ElseIf intResponse = 2 Then
FileCopy strTempPath & strSound, SoundPath & "\" & strSound
strReport = strReport & Lang(568) & strSound & Lang(569) & strTempPath & strSound & vbCrLf
ElseIf intResponse = 3 Then
intResponse = 0
    strReport = strReport & Lang(568) & strSound & Lang(567) & vbCrLf
End If

End Sub
Private Sub LookForTerrtexSnow(strAce As Variant)
Dim strTempPath As String, i As Integer

On Error GoTo Errtrap
strTempPath = strComPath & "\Terrtex\Snow\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
For i = 0 To NumRoutes - 1
strTempPath = AllRoutes(i) & "\Terrtex\Snow\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
Next i

If intResponseSnow = 2 Then
FileCopy MSTSPath & "\routes\europe2\Terrtex\Snow\OEGrass2.ace", TerrPath & "\snow\" & strAce
strReport = strReport & "The Terrtex\Snow file " & strAce & Lang(564) & Lang(570) & vbCrLf


booTerrtexSnow = False
Exit Sub
End If
booTerrtexSnow = True
strResponse = "Terrtex file " & strAce & Lang(571) & vbCrLf & Lang(655)
frmDialogSnow.Command1.Visible = False
frmDialogSnow.Show 1



     DoEvents
  TerrPath = RoutePath & "\terrtex"

If intResponseSnow = 1 Then
FileCopy TerrPath & "\" & strAce, TerrPath & "\snow\" & strAce
strReport = strReport & "The Terrtex\Snow file " & strAce & Lang(564) & Lang(656) & vbCrLf
ElseIf intResponseSnow = 2 Then
FileCopy MSTSPath & "\routes\europe2\Terrtex\Snow\OEGrass2.ace", TerrPath & "\snow\" & strAce
strReport = strReport & "The Terrtex\Snow file " & strAce & Lang(564) & Lang(570) & vbCrLf

End If
booTerrtexSnow = False
Exit Sub
ShowMessage:
If intResponse = 0 Then
strResponse = "The Terrtex\Snow file " & strAce & Lang(564) & vbCrLf & Lang(565)
frmDialog.Label1.Caption = strResponse
frmDialog.Command1.Visible = False
frmDialog.Show 1

     DoEvents
     
End If
If intResponse = 1 Then
intResponse = 0
FileCopy strTempPath & strAce, RoutePath & "\Terrtex\Snow\" & strAce
strReport = strReport & "Terrtex\Snow file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponse = 2 Then
FileCopy strTempPath & strAce, RoutePath & "\Terrtex\Snow\" & strAce
strReport = strReport & "Terrtex\Snow file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponse = 3 Then
intResponse = 0
    strReport = strReport & "Terrtex\Snow file " & strAce & Lang(567) & vbCrLf

End If
Exit Sub

Errtrap:
Call MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'LookForTerrtexSnow' while" _
                       & vbCrLf & "checking " & strAce _
                       , vbExclamation, App.Title)
    

End Sub

Private Sub LookForTerrtex(strAce As Variant)
Dim strTempPath As String, i As Integer


strTempPath = strComPath & "\Terrtex\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
For i = 0 To NumRoutes - 1
strTempPath = AllRoutes(i) & "\Terrtex\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
Next i

'strTempPath = MSTSPath & "\routes\europe1\Terrtex\"
'If FileExists(strTempPath & strACE) Then
'GoTo ShowMessage
'End If
'strTempPath = MSTSPath & "\routes\europe2\Terrtex\"
'If FileExists(strTempPath & strACE) Then
'GoTo ShowMessage
'End If
'strTempPath = MSTSPath & "\routes\japan1\Terrtex\"
'If FileExists(strTempPath & strACE) Then
'GoTo ShowMessage
'End If
'strTempPath = MSTSPath & "\routes\japan2\Terrtex\"
'If FileExists(strTempPath & strACE) Then
'GoTo ShowMessage
'End If
'strTempPath = MSTSPath & "\routes\usa1\Terrtex\"
'If FileExists(strTempPath & strACE) Then
'GoTo ShowMessage
'End If
'strTempPath = MSTSPath & "\routes\usa2\Terrtex\"
'If FileExists(strTempPath & strACE) Then
'GoTo ShowMessage
'End If
strReport = strReport & "Terrtex file " & strAce & Lang(563) & vbCrLf

Exit Sub
ShowMessage:
If intResponse = 0 Then
strResponse = "The Terrtex file " & strAce & Lang(564) & vbCrLf & Lang(565) & vbCrLf
frmDialog.Command1.Visible = False
frmDialog.Label1.Caption = strResponse
frmDialog.Show 1

     DoEvents
     
End If
If intResponse = 1 Then
intResponse = 0
FileCopy strTempPath & strAce, RoutePath & "\Terrtex\" & strAce
strReport = strReport & "Terrtex file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponse = 2 Then
FileCopy strTempPath & strAce, RoutePath & "\Terrtex\" & strAce
strReport = strReport & "Terrtex file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponse = 3 Then
intResponse = 0
    strReport = strReport & "Terrtex file " & strAce & Lang(567) & vbCrLf

End If



End Sub


Private Sub LookForACESD(strAce As Variant, strOriginal As String, booTopLevel As Boolean)
Dim strTempPath As String, x As Integer, strTexture As String, flagACE As Integer, i As Integer
flagACE = 1

On Error GoTo Errtrap

If (strAce = "snow\cow.ace" Or strAce = "snow\femalehiker.ace" Or strAce = "snow\us2cow.ace" Or strAce = "snow\us1deer.ace" Or strAce = "snow\workman.ace") = True Then Exit Sub
x = InStrRev(strAce, "\")
strTexture = Mid$(strAce, x + 1)
strTempPath = strComPath & "\Textures\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
For i = 0 To NumRoutes - 1
strTempPath = AllRoutes(i) & "\Textures\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
Next i


If booTopLevel = True Then
strReport = strReport & "Missing texture file " & strAce & " referred by " & strOriginal & Lang(574) & vbCrLf
Exit Sub

Else
strResponse = "Seasonal texture file " & strAce & Lang(564) & vbCrLf & Lang(572)
If intResponse = 2 Then GoTo Label7
frmDialog.Label1.Caption = strResponse
frmDialog.Show 1

     DoEvents
Label7:


If intResponse = 1 Then
intResponse = 0
FileCopy TexturePath & "\" & strTexture, TexturePath & "\" & strAce
ElseIf intResponse = 2 Then
FileCopy TexturePath & "\" & strTexture, TexturePath & "\" & strAce
flagTerr = True

ElseIf intResponse = 3 Then
intResponse = 0
strReport = strReport & "Texture file " & strAce & " referred by " & strOriginal & Lang(573) & vbCrLf
ElseIf intResponse = 4 Then
intResponse = 0
flagNoTex = True
strReport = strReport & "Texture file " & strAce & " referred by " & strOriginal & Lang(573) & vbCrLf

End If
End If
'End If
Exit Sub
'******************
ShowMessage:

If intResponse = 0 Then
strResponse = Lang(393) & strAce & Lang(564) & vbCrLf & Lang(565)
frmDialog.Command1.Visible = False
frmDialog.Label1.Caption = strResponse
frmDialog.Show 1

     DoEvents
     
End If
If intResponse = 1 Then
intResponse = 0
FileCopy strTempPath & strAce, TexturePath & "\" & strAce
strReport = strReport & "Texture file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponse = 2 And flagACE = 1 Then
FileCopy strTempPath & strAce, TexturePath & "\" & strAce
strReport = strReport & "Texture file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponse = 3 And flagACE = 1 Then
intResponse = 0
    strReport = strReport & "The Texture file " & strAce & Lang(567) & vbCrLf
End If
Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'LookForAceSD' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub


Private Sub LookForACESnow2(strAce As Variant, flagACE As Integer, strOriginal As String)
Dim strTempPath As String, i As Integer

On Error GoTo Errtrap
strTempPath = strComPath & "\Textures\Snow\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
For i = 0 To NumRoutes - 1

strTempPath = AllRoutes(i) & "\Textures\Snow\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
Next i

If intResponseSnow = 0 Then
strResponse = "Transfer file " & strAce & Lang(571) & vbCrLf & Lang(572)
frmDialogSnow.Command1.Visible = False
frmDialogSnow.Show 1

     DoEvents
     


    If intResponseSnow = 1 Then
    intResponseSnow = 0
    FileCopy TexturePath & "\" & strAce, TexturePath & "\Snow\" & strAce
    strReport = strReport & "Transfer Snow file " & strAce & " was missing, the summer texture has been used instead" & vbCrLf
    ElseIf intResponseSnow = 2 Then
    FileCopy TexturePath & "\" & strAce, TexturePath & "\Snow\" & strAce
    strReport = strReport & "Transfer Snow file " & strAce & " was missing, the summer texture has been used instead" & vbCrLf
    ElseIf intResponseSnow = 3 Then
    intResponseSnow = 0
    strReport = strReport & "Transfer Snow file " & strAce & " referred by " & strOriginal & Lang(573) & vbCrLf
    ElseIf intResponseSnow = 4 Then
    strReport = strReport & "Transfer Snow file " & strAce & " referred by " & strOriginal & Lang(573) & vbCrLf
    End If
ElseIf intResponseSnow = 4 Then
strReport = strReport & "Transfer Snow file " & strAce & " referred by " & strOriginal & Lang(573) & vbCrLf
End If

Exit Sub
ShowMessage:

If intResponseSnow = 0 Then
strResponse = "The Transfer Snow file " & strAce & Lang(564) & vbCrLf & Lang(565)
frmDialogSnow.Command1.Visible = False
frmDialogSnow.Show 1

     DoEvents
     
End If
If intResponseSnow = 1 And flagACE = 1 Then
intResponseSnow = 0
FileCopy strTempPath & strAce, TexturePath & "\Snow\" & strAce
strReport = strReport & "Texture file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponseSnow = 2 And flagACE = 1 Then
FileCopy strTempPath & strAce, TexturePath & "\Snow\" & strAce
strReport = strReport & "Texture file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponseSnow = 3 And flagACE = 1 Then
intResponseSnow = 0
    strReport = strReport & "The Transfer Snow file " & strAce & Lang(567) & vbCrLf
End If
Exit Sub
Errtrap:
Select Case MsgBox("An error " & Err.Description & " occurred checking  " & strAce, vbOKCancel Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbOK
Resume Next
    Case vbCancel
    Resume Next
Exit Sub
End Select
End Sub


Private Sub LookForACE2(strAce As Variant, flagACE As Integer, strOriginal As String)
Dim strTempPath As String, i As Integer

strTempPath = strComPath & "\Textures\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
For i = 0 To NumRoutes - 1

strTempPath = AllRoutes(i) & "\textures\"

If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
Next i

strReport = strReport & "Texture file " & strAce & " referred by " & strOriginal & Lang(563) & vbCrLf

Exit Sub
ShowMessage:

If intResponse = 0 Then
strResponse = Lang(393) & strAce & Lang(564) & vbCrLf & Lang(565)
frmDialog.Command1.Visible = False
frmDialog.Label1.Caption = strResponse
frmDialog.Show 1

     DoEvents
     
End If
If intResponse = 1 And flagACE = 1 Then
intResponse = 0
FileCopy strTempPath & strAce, TexturePath & "\" & strAce
strReport = strReport & "Texture file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponse = 2 And flagACE = 1 Then
FileCopy strTempPath & strAce, TexturePath & "\" & strAce
strReport = strReport & "Texture file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponse = 3 And flagACE = 1 Then
intResponse = 0
    strReport = strReport & "The Texture file " & strAce & Lang(567) & vbCrLf
End If
Exit Sub
If intResponse = 1 And flagACE = 2 Then
intResponse = 0
FileCopy strTempPath & strAce, RoutePath & "\Terrtex\" & strAce
ElseIf intResponse = 2 And flagACE = 2 Then
FileCopy strTempPath & strAce, RoutePath & "\Terrtex\" & strAce
ElseIf intResponse = 3 And flagACE = 2 Then
intResponse = 0
    strReport = strReport & "The Terrtex file " & strAce & Lang(567) & vbCrLf
End If
Exit Sub
If intResponse = 1 And flagACE = 3 Then
intResponse = 0
FileCopy strTempPath & strAce, RoutePath & "\envfiles\textures\" & strAce
ElseIf intResponse = 2 And flagACE = 3 Then
FileCopy strTempPath & strAce, RoutePath & "\envfiles\textures\" & strAce
ElseIf intResponse = 3 And flagACE = 3 Then
intResponse = 0
    strReport = strReport & "The Environment Texture file " & strAce & Lang(567) & vbCrLf
End If
End Sub





Private Sub LookForACE(strAce As String, flagACE As Integer, strOriginal As String)
Dim strTempPath As String, i As Integer

strTempPath = strComPath & "\Textures\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
For i = 0 To NumRoutes - 1

strTempPath = AllRoutes(i) & "\textures\"

If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
Next i

strReport = strReport & "Texture file " & strAce & " referred by " & strOriginal & Lang(563) & vbCrLf

Exit Sub
ShowMessage:

If intResponse = 0 Then
strResponse = Lang(393) & strAce & Lang(564) & vbCrLf & Lang(565)
frmDialog.Label1.Caption = strResponse
frmDialog.Command1.Visible = False
frmDialog.Show 1

     DoEvents
     
End If
If intResponse = 1 And flagACE = 1 Then
intResponse = 0
FileCopy strTempPath & strAce, TexturePath & "\" & strAce
strReport = strReport & "Texture file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponse = 2 And flagACE = 1 Then
FileCopy strTempPath & strAce, TexturePath & "\" & strAce
strReport = strReport & "Texture file " & strAce & Lang(569) & strTempPath & strAce & vbCrLf
ElseIf intResponse = 3 And flagACE = 1 Then
intResponse = 0
    strReport = strReport & "The Texture file " & strAce & Lang(567) & vbCrLf
End If
Exit Sub
If intResponse = 1 And flagACE = 2 Then
intResponse = 0
FileCopy strTempPath & strAce, RoutePath & "\Terrtex\" & strAce
ElseIf intResponse = 2 And flagACE = 2 Then
FileCopy strTempPath & strAce, RoutePath & "\Terrtex\" & strAce
ElseIf intResponse = 3 And flagACE = 2 Then
intResponse = 0
    strReport = strReport & "The Terrtex file " & strAce & Lang(567) & vbCrLf
End If
Exit Sub
If intResponse = 1 And flagACE = 3 Then
intResponse = 0
FileCopy strTempPath & strAce, RoutePath & "\envfiles\textures\" & strAce
ElseIf intResponse = 2 And flagACE = 3 Then
FileCopy strTempPath & strAce, RoutePath & "\envfiles\textures\" & strAce
ElseIf intResponse = 3 And flagACE = 3 Then
intResponse = 0
    strReport = strReport & "The Environment Texture file " & strAce & Lang(567) & vbCrLf
End If
End Sub















Private Sub LookForShape(strShape As Variant, booNoShape As Boolean)
Dim strTempPath As String, booMiss As Boolean, FoundInSpare As Boolean
Dim strSpare As String

On Error GoTo Errtrap

GlobalPath = MSTSPath & "\Global\Shapes\"
strTempPath = strComPath & "\Shapes\"
strSpare = App.Path & "\TempFiles"

If FileExists(strTempPath & strShape) Then
GoTo ShowMessage
End If
If FileExists(GlobalPath & strShape) Then
strReport = strReport & strShape & " appears to be a Track or Road section being used as a Static shape" & vbCrLf
GoTo CarryON
End If
If FileExists(GlobalSparePath & strShape) Then
FoundInSpare = True
booSpareTrack = True
FileCopy GlobalSparePath & strShape, GlobalPath & strShape
FileCopy GlobalSparePath & strShape & "d", GlobalPath & strShape & "d"
FileCopy GlobalSparePath & strShape, strSpare & "\" & strShape
GoTo ShowMessage
End If

For i = 0 To NumRoutes - 1

strTempPath = AllRoutes(i) & "\shapes\"
If FileExists(strTempPath & strShape) And FileExists(strTempPath & strShape & "d") Then
GoTo ShowMessage
End If
Next i

Call IsItMissing(strShape, booMiss)
If booMiss = True Then
strTempPath = strComPath & "\shapes\"
If FileExists(strTempPath & strShape) Then
GoTo ShowMessage
End If
For i = 0 To NumRoutes - 1

strTempPath = AllRoutes(i) & "\shapes\"

If FileExists(strTempPath & strShape) Then
GoTo ShowMessage
End If
Next i
End If

strReport = strReport & "Shape file " & strShape & Lang(563) & vbCrLf

booNoShape = True

Exit Sub
ShowMessage:

Rem ***********************************
If FoundInSpare = True Then
FoundInSpare = False
FileCopy GlobalSparePath & strShape, ShapePath & "\" & strShape
FileCopy GlobalSparePath & strShape, strSpare & "\" & strShape
strReport = strReport & "Shape file " & strShape & " restored from Global\SpareTrack" & vbCrLf
DoEvents
If FileExists(GlobalSparePath & strShape & "d") Then
FileCopy GlobalSparePath & strShape & "d", ShapePath & "\" & strShape & "d"
strReport = strReport & "Shape definition file " & strShape & "d" & " restored from Global\SpareTrack" & vbCrLf
booMiss = False
DoEvents
End If
Exit Sub
End If
Rem ************************************
If intResponse = 0 Then
strResponse = "The Shape file " & strShape & Lang(564) & vbCrLf & Lang(565)
frmDialog.Label1.Caption = strResponse
frmDialog.Command1.Visible = False
frmDialog.Show 1

     DoEvents
     
End If
If intResponse = 1 Then
intResponse = 0
FileCopy strTempPath & strShape, ShapePath & "\" & strShape
FileCopy strTempPath & strShape, strSpare & "\" & strShape
strReport = strReport & "Shape file " & strShape & Lang(569) & strTempPath & strShape & vbCrLf
DoEvents
If FileExists(strTempPath & strShape & "d") Then
FileCopy strTempPath & strShape & "d", ShapePath & "\" & strShape & "d"
strReport = strReport & "Shape definition file " & strShape & "d" & Lang(569) & strTempPath & strShape & "d" & vbCrLf
booMiss = False
DoEvents
End If
'Call ReadShape(ShapePath & "\" & strShape, 1, strSpare)   '*****************************************
ElseIf intResponse = 2 Then
FileCopy strTempPath & strShape, ShapePath & "\" & strShape
FileCopy strTempPath & strShape, strSpare & "\" & strShape
strReport = strReport & "Shape file " & strShape & Lang(569) & strTempPath & strShape & vbCrLf
DoEvents
If FileExists(strTempPath & strShape & "d") Then
FileCopy strTempPath & strShape & "d", ShapePath & "\" & strShape & "d"
strReport = strReport & "Shape definition file " & strShape & "d" & Lang(569) & strTempPath & strShape & "d" & vbCrLf
booMiss = False
DoEvents
End If
'Call ReadShape(ShapePath & "\" & strShape, 1, strSpare)
ElseIf intResponse = 3 Then
intResponse = 0
    strReport = strReport & "Shape file " & strShape & Lang(567) & vbCrLf
End If
CarryON:
Exit Sub

Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'LookForShape' while looking" _
            & vbCrLf & "for " & strShape _
            , vbExclamation, App.Title)
Resume Next


End Sub

Private Sub LookForGlobalShape(strShape As Variant, booNoShape As Boolean)
Dim strTempPath As String, FoundInSpare As Boolean
Dim strGlobalSpare As String

On Error GoTo Errtrap

GlobalPath = MSTSPath & "\Global\Shapes\"
strGlobalSpare = MSTSPath & "\Global\SpareTrack\"
strTempPath = strComPath & "\Shapes\"
If FileExists(strGlobalSpare & strShape) Then
FoundInSpare = True
booSpareTrack = True
FileCopy GlobalSparePath & strShape, GlobalPath & strShape
FileCopy GlobalSparePath & strShape & "d", GlobalPath & strShape & "d"
GoTo ShowMessage
End If


strReport = strReport & "Global Shape file " & strShape & Lang(563) & vbCrLf

booNoShape = True

Exit Sub
ShowMessage:
Rem ***********************************
If FoundInSpare = True Then
FoundInSpare = False
strReport = strReport & "Global Shape file " & strShape & " restored from Global\SpareTrack" & vbCrLf
End If
Exit Sub

Exit Sub

Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'LookForGlobalShape' while looking" _
            & vbCrLf & "for " & strShape _
            , vbExclamation, App.Title)
Resume Next


End Sub

Private Sub SecondTextures()
Dim fullpath$, strBat As String

TexturePath = RoutePath & "\Textures"
TexSnowPath = RoutePath & "\Textures\Snow"
TexNightPath = RoutePath & "\Textures\Night"
TexAutPath = RoutePath & "\Textures\Autumn"
TexAutSnowPath = RoutePath & "\Textures\AutumnSnow"
TexSprPath = RoutePath & "\Textures\Spring"
TexSprSnowPath = RoutePath & "\Textures\SpringSnow"
TexWinPath = RoutePath & "\Textures\Winter"
TexWinSnowPath = RoutePath & "\Textures\WinterSnow"

Open App.Path & "\SetupFiles\Installme.bat" For Append As #12

cursouind = 0
Rem ************* Snow
Drive1(cursouind).Drive = Left$(TexSnowPath, 2)
Dir1(cursouind).Path = TexSnowPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexSnowPath & "\" & File1(cursouind).List(i)
If FileExists(TexturePath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexturePath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\Snow\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* Night
Drive1(cursouind).Drive = Left$(TexNightPath, 2)
Dir1(cursouind).Path = TexNightPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexNightPath & "\" & File1(cursouind).List(i)
If FileExists(TexturePath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexturePath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\Night\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* Autumn
Drive1(cursouind).Drive = Left$(TexAutPath, 2)
Dir1(cursouind).Path = TexAutPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexAutPath & "\" & File1(cursouind).List(i)
If FileExists(TexturePath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexturePath & "\" & File1(cursouind).List(i)) Then

strBat = "call Xcopy " & ChrW$(34) & ".\Textures\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\Autumn\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* AutumnSnow
Drive1(cursouind).Drive = Left$(TexAutSnowPath, 2)
Dir1(cursouind).Path = TexAutSnowPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexAutSnowPath & "\" & File1(cursouind).List(i)
If FileExists(TexturePath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexturePath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\AutumnSnow\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* Spring
Drive1(cursouind).Drive = Left$(TexSprPath, 2)
Dir1(cursouind).Path = TexSprPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexSprPath & "\" & File1(cursouind).List(i)
If FileExists(TexturePath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexturePath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\Spring\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* SpringSnow
Drive1(cursouind).Drive = Left$(TexSprSnowPath, 2)
Dir1(cursouind).Path = TexSprSnowPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexSprSnowPath & "\" & File1(cursouind).List(i)
If FileExists(TexturePath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexturePath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\SpringSnow\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* Winter
Drive1(cursouind).Drive = Left$(TexWinPath, 2)
Dir1(cursouind).Path = TexWinPath

Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexWinPath & "\" & File1(cursouind).List(i)
If FileExists(TexturePath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexturePath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy  " & ChrW$(34) & ".\Textures\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\Winter\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* WinterSnow
Drive1(cursouind).Drive = Left$(TexWinSnowPath, 2)
Dir1(cursouind).Path = TexWinSnowPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexWinSnowPath & "\" & File1(cursouind).List(i)
If FileExists(TexturePath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexturePath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\WinterSnow\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* WinterSnow against Snow
Drive1(cursouind).Drive = Left$(TexWinSnowPath, 2)
Dir1(cursouind).Path = TexWinSnowPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexWinSnowPath & "\" & File1(cursouind).List(i)
If FileExists(TexSnowPath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexSnowPath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\Snow\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\WinterSnow\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* AutumnSnow against Snow
Drive1(cursouind).Drive = Left$(TexAutSnowPath, 2)
Dir1(cursouind).Path = TexAutSnowPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexAutSnowPath & "\" & File1(cursouind).List(i)
If FileExists(TexSnowPath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexSnowPath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\Snow\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\AutumnSnow\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* SpringSnow against Snow
Drive1(cursouind).Drive = Left$(TexSprSnowPath, 2)
Dir1(cursouind).Path = TexSprSnowPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexSprSnowPath & "\" & File1(cursouind).List(i)
If FileExists(TexSnowPath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexSnowPath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\Snow\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\SpringSnow\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* Winter against Snow
Drive1(cursouind).Drive = Left$(TexWinPath, 2)
Dir1(cursouind).Path = TexWinPath

Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexWinPath & "\" & File1(cursouind).List(i)
If FileExists(TexSnowPath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexSnowPath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\Snow\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\Winter\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* AutumnSnow against WinterSnow
Drive1(cursouind).Drive = Left$(TexAutSnowPath, 2)
Dir1(cursouind).Path = TexAutSnowPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexAutSnowPath & "\" & File1(cursouind).List(i)
If FileExists(TexWinSnowPath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexWinSnowPath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\WinterSnow\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\AutumnSnow\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* SpringSnow against WinterSnow
Drive1(cursouind).Drive = Left$(TexSprSnowPath, 2)
Dir1(cursouind).Path = TexSprSnowPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexSprSnowPath & "\" & File1(cursouind).List(i)
If FileExists(TexWinSnowPath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexWinSnowPath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\WinterSnow\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\SpringSnow\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i
Rem ************* Winter against WinterSnow
Drive1(cursouind).Drive = Left$(TexWinPath, 2)
Dir1(cursouind).Path = TexWinPath

Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
      fullpath$ = TexWinPath & "\" & File1(cursouind).List(i)
If FileExists(TexWinSnowPath & "\" & File1(cursouind).List(i)) Then
If GetCRC(fullpath$) = GetCRC(TexWinSnowPath & "\" & File1(cursouind).List(i)) Then
strBat = "call Xcopy " & ChrW$(34) & ".\Textures\WinterSnow\" & File1(cursouind).List(i) & ChrW$(34) & " .\Textures\Winter\ /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill fullpath$
DoEvents
End If
End If
End If
Next i

Rem ************* Check for route graphic

If FileExists(RoutePath & "\" & strGraphic) Then
If GetCRC(MSTSPath & "\Template\template.ace") = GetCRC(RoutePath & "\" & strGraphic) Then

strBat = "call copy " & ChrW$(34) & "..\..\template\template.ace" & ChrW$(34) & " " & ChrW$(34) & ".\" & strGraphic & ChrW$(34) & " /y"
Print #12, strBat
strBat = vbNullString
DoEvents

Kill RoutePath & "\" & strGraphic
DoEvents
End If
End If
Close #12
End Sub

Private Sub WriteSnow(strSnowFile As String, varBatText As Variant)
Dim TertexPath As String, TerSnowPath As String
Dim x As Integer, strSnowName As String
Dim strSnowRoute As String

On Error GoTo Errtrap
x = InStrRev(strSnowFile, "\")
strSnowName = Mid$(strSnowFile, x + 1)
x = InStrRev(strSnowFile, "Routes\")
strSnowRoute = Mid$(strSnowFile, x + 7)

TertexPath = RoutePath & "\Terrtex"
TerSnowPath = TertexPath & "\Snow"
If Not FileExists(strSnowFile) Then
Call MsgBox("The Terrtex Snow file " & strSnowName & Lang(574) & vbCrLf & Lang(575), vbExclamation, App.Title)

Exit Sub
End If


MousePointer = 11
cursouind = 0
Drive1(cursouind).Drive = Left$(TerSnowPath, 2)
Dir1(cursouind).Path = TerSnowPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

SparePath = App.Path & "\TempFiles"

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) And File1(cursouind).List(i) <> "microtex.ace" Then
      fullpath$ = TerSnowPath & "\" & File1(cursouind).List(i)
      varBatText = varBatText & "copy " & ChrW$(34) & "..\" & strSnowRoute & ChrW$(34) & " " & ChrW$(34) & ".\Terrtex\Snow\" & File1(cursouind).List(i) & ChrW$(34) & vbCrLf
    DoEvents
    Kill fullpath$
    End If
Next i


 
 

Text1(1).Text = "*.*"
 
  
  Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'WriteSnow' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
  Resume Next
  
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
Public Sub FindMinMax2(ByRef dArray() As Double, ByRef dLowVal As Double, ByRef dHighVal As Double, MinIndex As Long, maxIndex As Long)

    Dim lIndex As Long
    Dim dFirstValIdx As Double
    Dim dLastValIdx As Double
    Dim dActVal As Double
    dFirstValIdx = LBound(dArray)
    dLastValIdx = UBound(dArray)
    dLowVal = dArray(dFirstValIdx) 'start value
    dHighVal = dArray(dFirstValIdx) 'start value


    For lIndex = dFirstValIdx To dLastValIdx
        dActVal = dArray(lIndex)


        If dActVal > dHighVal Then
            dHighVal = dActVal
            maxIndex = lIndex
        Else 'if value smaller Then high value


            If dActVal < dLowVal And dActVal > 0 Then
                dLowVal = dActVal
                MinIndex = lIndex
            End If

        End If

    Next lIndex

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
            

End Function

 Private Sub ReadWorld3(fullpath$)
   
   Dim MyString As String
   Dim x As Integer, strS As String
   Dim Y As Long, strTile As String
    
    On Error GoTo Errtrap
'    If Right$(fullpath$, 1) = "\" Then Exit Sub
'
   x = InStrRev(fullpath$, "\")
   strS = Mid$(fullpath$, x + 1)

MyString = ReadUniFile(fullpath$)

 Rem ******************* TrackObj
 Y = 1
 Do
 Y = InStr(Y, MyString, "TrackObj (")
 If Y > 0 Then

 numWorldTiles = numWorldTiles + 1
                strTile = Mid$(strS, 2)
                strTile = Left$(strTile, Len(strTile) - 2)
                If Len(strTile) = 14 Then
                strLeft = Left$(strTile, 7)
                strRight = Right$(strTile, 7)
                Call TileName(Val(strLeft), Val(strRight))
                strWorldTiles(numWorldTiles) = Text3(2) & ".t"
                Y = Y + 1
                If Y > UBound(strWorldTiles) Then
     ReDim Preserve strWorldTiles(1 To Y + CHUNK)
     
      End If
 
               
           End If
           End If
        Loop While Y > 0
   DoEvents


        Exit Sub
Errtrap:
If Err = -2146233086 Then
Exit Sub
End If

 Call MsgBox("An error " & Err & " occurred in subroutine 'ReadWorld3' please advise" _
            & vbCrLf & "reading " & fullpath$ _
            , vbExclamation, App.Title)
           
    Resume Next
End Sub
Private Sub ReadWorld2(fullpath$)
 
   Dim strGName As String, MyString As String
   Dim x As Integer, strS As String
   Dim Y As Long, yy As Long, Z As Long, zz As Long
    
    On Error GoTo Errtrap
    If Right$(fullpath$, 1) = "\" Then Exit Sub
    
   x = InStrRev(fullpath$, "\")
   strS = Mid$(fullpath$, x + 1)
   SB1.Panels(2).Text = strS

MyString = ReadUniFile(fullpath$)

 Rem ******************* TrackObj
 Y = 1
 Do
 Y = InStr(Y, MyString, "TrackObj (")
 If Y > 0 Then
 yy = InStr(Y, MyString, "FileName", 0)
 If yy = 0 Then GoTo CarryON
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 z1 = InStr(Z + 1, MyString, "(", 0)
 If z1 < zz Then
 zz = InStr(zz + 1, MyString, ")", 0)
 End If
 strGName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strGName = Trim$(strGName)

                If Right$(strGName, 2) = ".s" Then
                strGlobShp(numGlobShp) = strGName
               
                numGlobShp = numGlobShp + 1
                
                
                If numGlobShp > UBound(strGlobShp) Then
                    ReDim Preserve strGlobShp(0 To numGlobShp + Shp_Chunk)
                End If
                DoEvents
                End If
                    Y = yy
                    End If
CarryON:
    
    Loop While Y > 0
    DoEvents

        Exit Sub
Errtrap:
If Err = -2146233086 Then
Exit Sub
End If
If Err = -2147467262 Then
Call MsgBox(fullpath$ & " could not be uncompressed." _
            & vbCrLf & "This route can not be processed until this is fixed", vbExclamation, App.Title)
Exit Sub
End If
 Call MsgBox("An error " & Err & " occurred in subroutine 'ReadWorld2' please advise" _
            & vbCrLf & "reading " & fullpath$ _
            , vbExclamation, App.Title)
   ' Resume Next
End Sub

Private Sub ReadWorld(fullpath$)
   Dim strFName As String
   Dim strGName As String, MyString As String, strSpare As String
   Dim x As Integer, strS As String, strHaz As String
   Dim Y As Long, yy As Long, Z As Long, zz As Long, strPickup As String
   Dim strPath As String, z1 As Long
    
    On Error GoTo Errtrap

   x = InStrRev(fullpath$, "\")
   strS = Mid$(fullpath$, x + 1)
 strPath = Left$(fullpath$, x - 1)
   SB1.Panels(2).Text = strS

 MyString = ReadUniFile(fullpath$)
MyString = LCase$(MyString)
 Label9.Caption = "Processing:  " & strS

Rem *************** Static
 Y = 1
 Do
 Y = InStr(Y, MyString, "static (", 0)
 If Y > 0 Then
 yy = InStr(Y, MyString, "filename", 0)
 If yy = 0 Then GoTo CarryON
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 z1 = InStr(Z + 1, MyString, "(", 0)
 If z1 < zz Then
 zz = InStr(zz + 1, MyString, ")", 0)
 End If

 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)

                If Left$(strFName, 1) = ChrW$(34) Then
                strFName = Mid$(strFName, 2)
                End If
                If Right$(strFName, 1) = ChrW$(34) Then
                strFName = Left$(strFName, Len(strFName) - 1)
                End If
                If Left$(strFName, 1) = ChrW$(34) Then
                strFName = Mid$(strFName, 2)
                End If
                If Right$(strFName, 1) = ChrW$(34) Then
                strFName = Left$(strFName, Len(strFName) - 1)
                End If
                If Right$(strFName, 2) <> ".s" Then GoTo CarryON
                If Left$(strFName, 2) = ".." Then
              
                    x = InStrRev(strFName, "/")
                    If x = 0 Then
                    x = InStrRev(strFName, "\")
                    End If
                    strFName = Mid$(strFName, x + 1)
                End If
                strShp(numShp) = strFName
                  
                numShp = numShp + 1
                    If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(0 To numShp + Shp_Chunk)
                    End If
                    Y = yy
                    End If
CarryON:
    
    Loop While Y > 0
    Rem ******************* TrackObj
 Y = 1
 Do
 Y = InStr(Y, MyString, "trackobj (", 0)
 If Y > 0 Then
 yy = InStr(Y, MyString, "filename", 0)
 If yy = 0 Then GoTo CarryOn1
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 z1 = InStr(Z + 1, MyString, "(", 0)
 If z1 < zz Then
 zz = InStr(zz + 1, MyString, ")", 0)
 End If
 strGName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strGName = Trim$(strGName)

                
                If Left$(strGName, 2) = ".." Then
            
                    strGName = Replace(strGName, "\", "/")
                    x = InStrRev(strGName, "/")
                    strGName = Mid$(strGName, x + 1)
                    strShp(numShp) = strGName
                            
                     numShp = numShp + 1
                    If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(0 To numShp + Shp_Chunk)
                    End If
                   GoTo GetNextShape
                End If
                If Right$(strGName, 2) = ".s" Then
                strGlobShp(numGlobShp) = strGName
               
                numGlobShp = numGlobShp + 1
                
                
                If numGlobShp > UBound(strGlobShp) Then
                    ReDim Preserve strGlobShp(0 To numGlobShp + Shp_Chunk)
                End If
                DoEvents
                End If
GetNextShape:
                
                    Y = yy
                    End If
CarryOn1:
    
    Loop While Y > 0
    Rem ****************** Hazards
    Y = 1
 Do
 Y = InStr(Y, MyString, "hazard (", 0)
 If Y > 0 Then

 booHaz = True
 yy = InStr(Y, MyString, "filename", 0)
 If yy = 0 Then GoTo CarryOn2
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 z1 = InStr(Z + 1, MyString, "(", 0)
 If z1 < zz Then
 zz = InStr(zz + 1, MyString, ")", 0)
 End If
 strHaz = Mid$(MyString, Z + 1, zz - (Z + 1))
 strHaz = Trim$(strHaz)
                HazShp(numHaz) = strHaz
                numHaz = numHaz + 1
                If numHaz > UBound(HazShp) Then
                    ReDim Preserve HazShp(0 To numHaz + For_Chunk)
                    End If
            
                    Y = yy
                    End If

CarryOn2:
    Loop While Y > 0
    
    Rem ******************* Pickup Objects
        Y = 1
 Do
 Y = InStr(Y, MyString, "pickup (", 0)
 If Y > 0 Then
 yy = InStr(Y, MyString, "filename", 0)
 If yy = 0 Then GoTo CarryOn3
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 z1 = InStr(Z + 1, MyString, "(", 0)
 If z1 < zz Then
 zz = InStr(zz + 1, MyString, ")", 0)
 End If
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Right$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
 Z = InStr(Y, MyString, "pickuptype (", 0)
 Z = InStr(Z, MyString, "(", 0)
  zz = InStr(Z, MyString, ")", 0)
  strPickup = Mid$(MyString, Z + 1, zz - (Z + 1))
 strPickup = Trim$(strPickup)
 strPickup = Left$(strPickup, 1)
                 Select Case strPickup
                    Case "5"
                    booCoal = True
                    Case "6"
                    booWat = True
                    Case "7"
                    booDies = True
                End Select
                    
                strShp(numShp) = strFName
                
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(0 To numShp + Shp_Chunk)
                End If
                DoEvents
                
               ' End If
                         
                    Y = yy
                    End If

CarryOn3:
    Loop While Y > 0
Rem ******************** Signals
    Y = 1
 Do
 Y = InStr(Y, MyString, "signal (", 0)
 If Y > 0 Then
 booSig = True
 yy = InStr(Y, MyString, "filename", 0)
 If yy = 0 Then GoTo CarryOn4
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 z1 = InStr(Z + 1, MyString, "(", 0)
 If z1 < zz Then
 zz = InStr(zz + 1, MyString, ")", 0)
 End If
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Right$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
                strShp(numShp) = strFName
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(0 To numShp + Shp_Chunk)
                    End If
            
                    Y = yy
                    End If
CarryOn4:
    
    Loop While Y > 0
  Rem ******************** Gantries
    Y = 1
 Do
 Y = InStr(Y, MyString, "gantry (", 0)
 If Y > 0 Then
 
 yy = InStr(Y, MyString, "filename", 0)
 If yy = 0 Then GoTo CarryOn5
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 z1 = InStr(Z + 1, MyString, "(", 0)
 If z1 < zz Then
 zz = InStr(zz + 1, MyString, ")", 0)
 End If
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Right$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
                strShp(numShp) = strFName
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(0 To numShp + Shp_Chunk)
                    End If
            
                    Y = yy
                    End If

CarryOn5:
    Loop While Y > 0

     Rem ******************** Level Crossings
    Y = 1
 Do
 Y = InStr(Y, MyString, "levelcr (", 0)
 If Y > 0 Then
 booCross = True
 yy = InStr(Y, MyString, "filename", 0)
 If yy = 0 Then GoTo CarryOn6
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 z1 = InStr(Z + 1, MyString, "(", 0)
 If z1 < zz Then
 zz = InStr(zz + 1, MyString, ")", 0)
 End If
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Right$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
                strShp(numShp) = strFName
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(0 To numShp + Shp_Chunk)
                    End If
            
                    Y = yy
                    End If

CarryOn6:
    Loop While Y > 0
         Rem ******************** Collide Objects
    Y = 1
 Do
 Y = InStr(Y, MyString, "collideobject (", 0)
 If Y > 0 Then
 booCross = True
 yy = InStr(Y, MyString, "filename", 0)
 If yy = 0 Then GoTo CarryOn7
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 z1 = InStr(Z + 1, MyString, "(", 0)
 If z1 < zz Then
 zz = InStr(zz + 1, MyString, ")", 0)
 End If
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Right$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
                strShp(numShp) = strFName
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(0 To numShp + Shp_Chunk)
                    End If
                    Y = yy
                    End If
CarryOn7:
    
    Loop While Y > 0
    
         Rem ******************** Forest
    Y = 1
 Do
 Y = InStr(Y, MyString, "forest (", 0)
 If Y > 0 Then
 
 yy = InStr(Y, MyString, "treetexture", 0)
 If yy = 0 Then
 'strReport = strReport & "There is a Forest item in " & strS & " with no TreeTexture entry (not necessarily an error)" & vbCrLf
 Y = 0
 GoTo EndForest
 End If
 Rem
 
 
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 z1 = InStr(Z + 1, MyString, "(", 0)
 If z1 < zz Then
 zz = InStr(zz + 1, MyString, ")", 0)
 End If
 
 Rem
 
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Right$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 4) = ".ace" Then
                ForTex(numFor) = strFName
                numFor = numFor + 1
                If numFor > UBound(ForTex) Then
                    ReDim Preserve ForTex(0 To numFor + For_Chunk)
                    End If
            
                    Y = yy
                    End If
End If
EndForest:
    Loop While Y > 0
             Rem ******************** Transfer
    Y = 1
 Do
 Y = InStr(Y, MyString, "transfer (", 0)
 If Y > 0 Then

 yy = InStr(Y, MyString, "filename", 0)
 If yy = 0 Then GoTo CarryOn8
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 z1 = InStr(Z + 1, MyString, "(", 0)
 If z1 < zz Then
 zz = InStr(zz + 1, MyString, ")", 0)
 End If
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Right$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 4) = ".ace" Then
                Ace1(numAce) = strFName
                Transfer(numTrans) = strFName
                numTrans = numTrans + 1
                If numTrans > UBound(Transfer) Then
                    ReDim Preserve Transfer(0 To numTrans + For_Chunk)
                    End If
                numAce = numAce + 1
                If numAce > UBound(Ace1) Then
                    ReDim Preserve Ace1(0 To numAce + Shp_Chunk)
                    End If
            
                    Y = yy
                    End If
End If
CarryOn8:
    Loop While Y > 0
                 Rem ******************** SpeedPost
    Y = 1
 Do
 Y = InStr(Y, MyString, "speedpost (", 0)
 If Y > 0 Then
yy = InStr(Y, MyString, "speed_digit_tex", 0)
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Right$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 4) = ".ace" Then
                Ace1(numAce) = strFName
                numAce = numAce + 1
                If numAce > UBound(Ace1) Then
                    ReDim Preserve Ace1(0 To numAce + Shp_Chunk)
                    End If
 End If
 yy = InStr(Y, MyString, "filename", 0)
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Right$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
 If Right$(strFName, 2) = ".s" Then
 strShp(numShp) = strFName
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(0 To numShp + Shp_Chunk)
                    End If
                
                    Y = yy
                    End If
End If
    
    Loop While Y > 0
DoEvents

If FileExists(strSpare & "\" & strS) Then
Kill strSpare & "\" & strS
End If

        Exit Sub
Errtrap:

      Call MsgBox("An error " & Err & " occurred in subroutine 'ReadWorld' in " & fullpath$ _
            & vbCrLf & "Error Description = " & Err.Description _
            , vbExclamation, App.Title)
     
      Resume Next
   
End Sub
Private Sub ReadShape4(fullpath$, jx As Integer)
  Dim x As Integer, strS As String, AceTemp(1 To 100) As String, intAce As Integer
      Dim q As Integer, MyString As String, jj As Integer, strSpare As String
 
    
    On Error GoTo Errtrap
    
   x = InStrRev(fullpath$, "\")
   strS = Mid$(fullpath$, x + 1)

   SB1.Panels(2).Text = strS
   
      strSpare = App.Path & "\TempFiles"

  MyMainString = vbNullString
  Label9.Caption = "Processing:  " & strS

   DoEvents
   MyString = ReadUniFile(fullpath$)
   yy = 1
 Do
 
 yy = InStr(yy, MyString, "image (")
 If yy > 0 Then
 Z = InStr(yy, MyString, "(")
 zz = InStr(Z, MyString, ")")
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Mid$(strFName, 2)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
 intAce = intAce + 1
  AceTemp(intAce) = strFName
    
             
                           
                    yy = zz
                    End If

    
    Loop While yy > 0

    For q = 1 To intAce
           
            
              For jj = 1 To jx
              If AceTemp(intAce) = strGetAce(jj) Then
            
              strReport = strReport & strGetAce(jj) & " is used by " & vbTab & strS & vbCrLf
              End If
              Next jj
             Next q
   
 DoEvents


Exit Sub
Errtrap:

Call MsgBox("An error " & Err & " " & Err.Description & " reading file " & strS & " please send file to" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbExclamation, App.Title)
                      

    
 
End Sub

Private Sub ReadShape3(fullpath$, strSparePath As String, intPath As Integer)
   Dim x As Integer, strS As String, AceTemp(1 To 100) As String, intAce As Integer
   Dim i As Integer, intESD As Integer
   Dim MyString As String, strPath As String, GlobalPath As String
    
    On Error GoTo Errtrap
    
    GlobalPath = MSTSPath & "\Global\Shapes\"
   x = InStrRev(strSparePath, "\")
   strS = Mid$(strSparePath, x + 1)
    strPath = Left$(strSparePath, x - 1)
   SB1.Panels(2).Text = strS
   
  Label9.Caption = "Processing:  " & strS
 
' Call DoDeComp2(strS, strPath, strSpare)
   DoEvents
   MyString = ReadUniFile(strSparePath)
MyString = LCase$(MyString)
   
      yy = 1
 Do
 
 yy = InStr(yy, MyString, "image (", 0)
 If yy > 0 Then
 Z = InStr(yy, MyString, "(", 0)
 zz = InStr(Z, MyString, ")", 0)
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Mid$(strFName, 2)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If
 intAce = intAce + 1
  AceTemp(intAce) = strFName
    
             
                           
                    yy = zz
                    End If

    
    Loop While yy > 0
   

        Rem ********** Check .sd
        If FileExists(GlobalPath & strS & "d") Then
        intPath = 2

        End If
        
        If intPath = 1 Then
    
        Call ReadSD(fullpath$ & "\" & strS, intESD, intPath)
        ElseIf intPath = 2 Then
        intESD = 2
        ElseIf intPath = 3 Then
        intESD = 0
        End If
        
        
        For i = 1 To intAce
        Call ReadAce(AceTemp(i), intESD, strS)
        DoEvents
        AceTemp(i) = vbNullString
        Next i
        
      
Exit Sub
Errtrap:

Call MsgBox("An error " & Err & " " & Err.Description & " reading file " & strS & " please send file to" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbExclamation, App.Title)
                    
 Resume Next
End Sub

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
' read MSTS file into byte array
' works on if GZip compressed too
Public Function readFile(fName As String, ByRef bdata() As Byte) As Long
    Dim i As Integer
    Dim bufSize As Long
    Dim bHead(7) As Byte
   
    
    i = FreeFile
    On Error GoTo ErrEx
    
    Open fName For Binary Access Read As i
    Get #i, , bHead()
    If bHead(7) > 64 Then
        Get #i, , bufSize
        ReDim bdata(lOf(i) - 17)
        Get #i, 17, bdata()
        Set comp = New CompressZIt
        readFile = comp.DecompressData(bdata(), bufSize)
        Set comp = Nothing
    Else
        ReDim bdata(lOf(i) - 17)
        Get #i, 17, bdata()
    End If
    Close i
    
    Exit Function
ErrEx:
    readFile = Err.Number
End Function



Private Sub MakeSnow()
Dim TertexPath As String, TerSnowPath As String
Dim varBatText As Variant, strNewSnow As String


On Error GoTo Errtrap

CDL1.Filter = "Terrtex Files (*.ace)|*.ace"
CDL1.DialogTitle = "Select a SNOW .ace File"
CDL1.FilterIndex = 1
If strComPath <> vbNullString Then
CDL1.Filename = strComPath & "\Terrtex\Snow\US2TarGnd.ace"
Else
CDL1.Filename = MSTSPath & "\Routes\usa2\terrtex\snow\us2targnd.ace"
End If
CDL1.Action = 1
DoEvents
strNewSnow = CDL1.Filename


TertexPath = RoutePath & "\Terrtex"
TerSnowPath = TertexPath & "\Snow"

MousePointer = 11
cursouind = 0
Drive1(cursouind).Drive = Left$(TerSnowPath, 2)
Dir1(cursouind).Path = TerSnowPath
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

SparePath = App.Path & "\TempFiles"
varBatText = "md " & ChrW$(34) & TerSnowPath & "\backup" & ChrW$(34) & vbCrLf
varBatText = varBatText & "copy " & ChrW$(34) & TerSnowPath & "\*.*" & ChrW$(34) & " " & ChrW$(34) & TerSnowPath & "\backup" & ChrW$(34) & vbCrLf

If FileExists(SparePath & "\do_TerSnow.bat") Then
Kill SparePath & "\do_TerSnow.bat"
End If
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
      fullpath$ = TerSnowPath & "\" & File1(cursouind).List(i)
      
   
    varBatText = varBatText & "copy " & ChrW$(34) & strNewSnow & ChrW$(34) & " " & ChrW$(34) & fullpath$ & ChrW$(34) & vbCrLf
    
    End If
Next i
   varBatText = varBatText & "md " & ChrW$(34) & RoutePath & "\textures\winter\backup" & ChrW$(34) & vbCrLf
   varBatText = varBatText & "copy " & ChrW$(34) & RoutePath & "\Textures\Winter\*.*" & ChrW$(34) & " " & ChrW$(34) & RoutePath & "\Textures\Winter\backup" & ChrW$(34) & vbCrLf
varBatText = varBatText & "copy " & ChrW$(34) & RoutePath & "\Textures\WinterSnow\*.*" & ChrW$(34) & " " & ChrW$(34) & RoutePath & "\Textures\Winter" & ChrW$(34) & vbCrLf
   Open SparePath & "\do_TerSnow.bat" For Append As #6
   Print #6, varBatText
   Close 6
      
   Text1(cursouind).Text = "*.*"
  strDrive = Left$(SparePath, 1)
   ChDrive strDrive
ChDir SparePath
mydir = CurDir

  DoEvents

Call ShellAndWait("do_TerSnow.bat", True, vbNormalFocus)


Text1(1).Text = "*.*"
  MousePointer = 0
  
  Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'MakeSnow' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
  Resume Next
  
End Sub


Private Sub SetLanguage(strLang As String)
Dim intLan As Integer, i As Integer, NewFile As Integer, intNum As Integer, A$
On Error GoTo Errtrap

If Not FileExists(App.Path & "\" & strLang) Then

     Call MsgBox(App.Path & "\" & strLang & vbCrLf & Lang(479), vbCritical, App.Path)
Exit Sub
End If
 NewFile = FreeFile

Open App.Path & "\" & strLang For Input As #NewFile
Input #NewFile, intLan, A$
ReDim Lang(1 To intLan)
intNum = 1
For i = 1 To intLan
Input #NewFile, intNum

Input #NewFile, Lang(intNum)
Next i
Close #NewFile



For i = 0 To 6
SSTab1.TabCaption(i) = Lang(1 + i)
Next

Command7.Caption = Lang(8)
Command7.ToolTipText = Lang(9)
Command3.Caption = Lang(10)
Command3.ToolTipText = Lang(11)
Command80.Caption = Lang(12)
Command80.ToolTipText = Lang(13)
Command5.Caption = Lang(14)
Command5.ToolTipText = Lang(15)
Command9.Caption = Lang(16)
Command9.ToolTipText = Lang(17)
Command4.Caption = Lang(18)
Command4.ToolTipText = Lang(19)
Command8.Caption = Lang(20)
Command8.ToolTipText = Lang(21)
Command13.Caption = Lang(22)
Command13.ToolTipText = Lang(23)
Command12.Caption = Lang(24)
Command12.ToolTipText = Lang(25)
Command36.Caption = Lang(26)
Command36.ToolTipText = Lang(27)
Command2.Caption = Lang(28)
Command2.ToolTipText = Lang(29)

Command1(15).Caption = Lang(30)
Label1(0).Caption = Lang(31)
Label1(1).Caption = Lang(32)
Command41.Caption = Lang(33)
Command41.ToolTipText = Lang(34)
mnuFiles.Caption = Lang(35)
mnuPath.Caption = Lang(36)
mnuCommon.Caption = Lang(37)
mnuExit.Caption = Lang(38)
mnuLan.Caption = Lang(39)
'mnuEng.Caption = Lang(40)
'mnu1.Caption = Lang(41)
'mnu2.Caption = Lang(42)
'mnu3.Caption = Lang(43)
'mnu4.Caption = Lang(44)
'mnu5.Caption = Lang(45)
'mnu6.Caption = Lang(46)
mnuHelp.Caption = Lang(47)
mnuCont.Caption = Lang(48)
mnuAbout.Caption = Lang(49)
Frame2.Caption = Lang(50)
Frame5.Caption = Lang(51)

Frame4.Caption = Lang(52)
Command14(0).Caption = Lang(53)
Command14(0).ToolTipText = Lang(54)
Command15.Caption = Lang(55)
Command15.ToolTipText = Lang(56)
Command16.Caption = Lang(57)
Command16.ToolTipText = Lang(58)
Command35.Caption = Lang(59)
Command35.ToolTipText = Lang(60)
For i = 0 To 5
Label8(i).Caption = Lang(61 + i)
Next i

Command11(0).Caption = Lang(68)
Command11(0).ToolTipText = Lang(69)
Command26.Caption = Lang(72)
Command26.ToolTipText = Lang(73)
Command40.Caption = Lang(76)
Command40.ToolTipText = Lang(77)
Command6(0).Caption = Lang(78)
Command6(0).ToolTipText = Lang(79)
Command6(1).Caption = Lang(80)
Command6(1).ToolTipText = Lang(81)
Command21.Caption = Lang(82)
Command21.ToolTipText = Lang(83)
Command39.Caption = Lang(84)
Command39.ToolTipText = Lang(85)
Command45.Caption = Lang(86)
Command45.ToolTipText = Lang(87)
Command17.Caption = Lang(88)
Command17.ToolTipText = Lang(89)
Command18.Caption = Lang(90)
Command18.ToolTipText = Lang(91)
Command23.Caption = Lang(92)
Command23.ToolTipText = Lang(93)
Command29.Caption = Lang(94)
Command29.ToolTipText = Lang(95)
Command46.Caption = Lang(96)
Command46.ToolTipText = Lang(97)
Command30.Caption = Lang(98)
Command30.ToolTipText = Lang(99)
Command38.Caption = Lang(100)
Command38.ToolTipText = Lang(101)
Command44.Caption = Lang(364)
Command44.ToolTipText = Lang(365)
Command22.Caption = Lang(102)
Command22.ToolTipText = Lang(103)
Command55.Caption = Lang(104)
Command55.ToolTipText = Lang(105)
Command56(0).Caption = Lang(45)
Command56(0).ToolTipText = Lang(46)
Command56(1).Caption = Lang(615)
Command56(1).ToolTipText = Lang(616)
For i = 0 To 9
Command1(i).Caption = Lang(106 + i * 2)
Command1(i).ToolTipText = Lang(107 + i * 2)
Next i
For i = 11 To 14
Command1(i).Caption = Lang(104 + i * 2)
Command1(i).ToolTipText = Lang(105 + i * 2)
Next i
Command1(17).Caption = Lang(134)
Command1(17).ToolTipText = Lang(135)
Command34.Caption = Lang(136)
Command34.ToolTipText = Lang(137)
Command37.Caption = Lang(138)
Command37.ToolTipText = Lang(139)
Frame1.Caption = Lang(140)
Command19.Caption = Lang(141)
Command19.ToolTipText = Lang(142)
Command20.Caption = Lang(143)
Command20.ToolTipText = Lang(144)
Command28.Caption = Lang(145)
Command28.ToolTipText = Lang(146)
Command25.Caption = Lang(147)
Command25.ToolTipText = Lang(148)
Command31.Caption = Lang(149)
Command31.ToolTipText = Lang(150)
Command33.Caption = Lang(151)
Command33.ToolTipText = Lang(152)
Command32.Caption = Lang(153)
Command32.ToolTipText = Lang(154)
Command52.Caption = Lang(155)
Command52.ToolTipText = Lang(156)
Command11(3).Caption = Lang(157)
Command11(3).ToolTipText = Lang(158)
Command24.Caption = Lang(159)
Command24.ToolTipText = Lang(160)
Command54.Caption = Lang(179)
Command54.ToolTipText = Lang(180)
Check1.Caption = Lang(181)
Command49(0).Caption = Lang(161)
Command49(0).ToolTipText = Lang(162)
Command11(4).Caption = Lang(163)
Command11(4).ToolTipText = Lang(164)
Command47.Caption = Lang(165)
Command47.ToolTipText = Lang(166)
Command49(1).Caption = Lang(167)
Command49(1).ToolTipText = Lang(168)
'Command11(5).Caption = Lang(169)
Command11(5).ToolTipText = Lang(170)
Command51.Caption = Lang(171)
Command51.ToolTipText = Lang(172)
Frame9.Caption = Lang(173)
Label10.Caption = Lang(174)
Command50.Caption = Lang(175)
Command50.ToolTipText = Lang(176)
Command48.Caption = Lang(177)
Command48.ToolTipText = Lang(178)
Command10(0).ToolTipText = Lang(517)
Command10(1).ToolTipText = Lang(517)
Command42.ToolTipText = Lang(107)
Command53.ToolTipText = Lang(518)
Command43(0).ToolTipText = Lang(519)
Command43(1).ToolTipText = Lang(520)
Command43(2).ToolTipText = Lang(521)
SSTab1.ToolTipText = Lang(40)
Drive1(0).ToolTipText = Lang(41)
Drive1(1).ToolTipText = Lang(41)
Dir1(0).ToolTipText = Lang(42)
Dir1(1).ToolTipText = Lang(42)
File1(0).ToolTipText = Lang(43)
File1(1).ToolTipText = Lang(43)
Text1(0).ToolTipText = Lang(44)
Text1(1).ToolTipText = Lang(44)
'Command57.Caption = Lang(619)
'Command57.ToolTipText = Lang(619)
Command58.Caption = Lang(617)
Command58.ToolTipText = Lang(618)
Command59.ToolTipText = Lang(620)
'comAbort.Caption = Lang(637)
comAbortFil.Caption = Lang(637)
'comAbortAct.Caption = Lang(637)
Command60.Caption = Lang(642)
Command60.ToolTipText = Lang(643)
Command61.Caption = Lang(640)
Command61.ToolTipText = Lang(641)
Command62.Caption = Lang(638)
Command62.ToolTipText = Lang(639)
Command84.Caption = Lang(661)
Command84.ToolTipText = Lang(662)
Command82.Caption = Lang(663)
Command82.ToolTipText = Lang(664)
'Command83.Caption = Lang(665)
'Command83.ToolTipText = Lang(666)
Command85.Caption = Lang(667)
Command85.ToolTipText = Lang(668)
Command76.Caption = Lang(669)
Command76.ToolTipText = Lang(670)
Command79.Caption = Lang(671)
Command79.ToolTipText = Lang(672)
Command81.Caption = Lang(673)
Command81.ToolTipText = Lang(674)

Rem ************** Done to end of Graphics file utils.

Exit Sub
Errtrap:
Call MsgBox(Err.Description & " occurred in your Language file " & strLang, vbCritical, App.Title)

Resume Next

End Sub


Private Sub StripW2(strShape As String)
Dim i As Integer, filepath1$, fullpath$
Dim flagway As Integer


cursouind = 1
filepath1$ = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.W"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
  
   flagway = 0
    Call ConvertW2(fullpath$, strShape, flagway)
    flagway = 1
    Call ConvertW2(fullpath$, strShape, flagway)
    End If
    DoEvents
   Close
    Kill WorldPath & "\" & File1(cursouind).List(i)
    
DoEvents
FileCopy fullpath$, WorldPath & "\" & File1(cursouind).List(i)
DoEvents

    Next i
End Sub

Private Sub StripW(strShape As String)
Dim i As Integer, filepath1$, fullpath$
Dim flagway As Integer


cursouind = 1
filepath1$ = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.W"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
  
   flagway = 0
    Call ConvertW(fullpath$, strShape, flagway)
    flagway = 1
    Call ConvertW(fullpath$, strShape, flagway)
    End If
    DoEvents
   Close
    Kill WorldPath & "\" & File1(cursouind).List(i)
    
DoEvents
FileCopy fullpath$, WorldPath & "\" & File1(cursouind).List(i)
DoEvents

    Next i
End Sub
Private Sub UncompressSelectedW()

Dim strTemp As String, strTempObj As String, SparePath As String



'Set tfh = New TokenFileHandler

'

Rem *********Check World Files *********
SparePath = App.Path & "\TempFiles"

cursouind = 0
   
   For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) = True Then
   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   strTempObj = File1(cursouind).List(i)
   Open fullpath$ For Binary As #5
    strTemp = String(2, " ")
    Get #5, , strTemp
 Close #5

 If Asc(Mid$(strTemp, 1, 1)) = 255 And Asc(Mid$(strTemp, 2, 1)) = 254 Then
 FileCopy fullpath$, SparePath & "\" & strTempObj
Else
TokMode = 1
  booWriteFile = True
 Call DoDeComp2(strTempObj, File1(cursouind).Path, SparePath)
'result = tfh.decompress(fullpath$, SparePath & "\" & strtempobj)
DoEvents
End If
End If
GetAnother:
   Next i
     
   


End Sub

Private Sub UncompressAllW(strMyPath As String)
Dim SparePath As String

SparePath = App.Path & "\TempFiles"
Call DoDeCompFolder("w", strMyPath, SparePath)



End Sub


Private Function UnConvertSMS(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
UnConvertSMS = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
xx = 1
AnotherTry:
x = InStr(xx, MyString, "LoadAllWaves ( 1 )")
If x = 0 Then GoTo CarryON
strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, x + 18)
MyString = strStart & strEnd
xx = x + 19
GoTo AnotherTry
CarryON:
End If
'End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
UnConvertSMS = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function



         
         

         
         


Private Function ConvertW2(CompleteFilePath As String, strShape As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertW2 = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Long
Dim xx As Long, j As Integer
Dim X1(1 To 16) As Double, MinX As Double, MaxX As Double
'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If

If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll
xx = 1
If flagway = 0 Then
CarryON:

x = InStr(xx, MyString, strShape)
If x = 0 Then GoTo CarryOn2

X1(1) = InStr(x, MyString, "static (") - x
X1(2) = InStr(x, MyString, "gantry (") - x
X1(3) = InStr(x, MyString, "trackobj (") - x
X1(4) = InStr(x, MyString, "forest (") - x
X1(5) = InStr(x, MyString, "signal (") - x
X1(6) = InStr(x, MyString, "speedpost (") - x
X1(7) = InStr(x, MyString, "tr_watermark") - x
X1(8) = InStr(x, MyString, "levelcr (") - x
X1(9) = InStr(x, MyString, "pickup (") - x
X1(10) = InStr(x, MyString, "CollideObject (") - x
X1(11) = InStr(x, MyString, "Hazard (") - x
X1(12) = InStr(x, MyString, "Dyntrack (") - x
X1(13) = InStr(x, MyString, "Siding (") - x
X1(14) = InStr(x, MyString, "Platform (") - x
X1(15) = InStr(x, MyString, "CarSpawner (") - x
X1(16) = InStr(x, MyString, "Transfer (") - x
MinX = 0: MaxX = 0
Call FindMinMax(X1(), MinX, MaxX)

If MaxX = 0 Then

strEnd = vbCrLf & ")" & vbCrLf
Else
strEnd = Mid$(MyString, x + MinX)
End If
For j = 1 To 16
X1(j) = 0
Next j

X1(1) = x - InStrRev(MyString, "static (", x)
X1(2) = x - InStrRev(MyString, "gantry (", x)
X1(3) = x - InStrRev(MyString, "trackobj (", x)
X1(4) = x - InStrRev(MyString, "forest (", x)
X1(5) = x - InStrRev(MyString, "signal (", x)
X1(6) = x - InStrRev(MyString, "speedpost (", x)
X1(7) = x - InStrRev(MyString, "tr_watermark", x)
X1(8) = x - InStrRev(MyString, "levelcr (", x)
X1(9) = x - InStrRev(MyString, "pickup (", x)
X1(10) = x - InStrRev(MyString, "CollideObject (", x)
X1(11) = x - InStrRev(MyString, "Hazard (", x)
X1(12) = x - InStrRev(MyString, "Dyntrack (", x)
X1(13) = x - InStrRev(MyString, "Siding (", x)
X1(14) = x - InStrRev(MyString, "Platform (", x)
X1(15) = x - InStrRev(MyString, "CarSpawner (", x)
X1(16) = x - InStrRev(MyString, "Transfer (", x)
Call FindMinMax(X1(), MinX, MaxX)

If Mid$(MyString, MinX, 8) = "Gantry (" Then
strStart = Left$(MyString, x - MinX - 1)
MyString = strStart & strEnd
End If
For j = 1 To 15
X1(j) = 0
Next j
'MyString = strStart & strEnd

GoTo CarryON

CarryOn2:

End If

The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertW2 = True

ERRHANDLER:

  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function

Private Function ConvertW(CompleteFilePath As String, strShape As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertW = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Long
Dim xx As Long, j As Integer
Dim X1(1 To 16) As Double, MinX As Double, MaxX As Double
'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll
xx = 1
If flagway = 0 Then
CarryON:

x = InStr(xx, MyString, strShape)
If x = 0 Then GoTo CarryOn2

X1(1) = InStr(x, MyString, "static (") - x
X1(2) = InStr(x, MyString, "gantry (") - x
X1(3) = InStr(x, MyString, "trackobj (") - x
X1(4) = InStr(x, MyString, "forest (") - x
X1(5) = InStr(x, MyString, "signal (") - x
X1(6) = InStr(x, MyString, "speedpost (") - x
X1(7) = InStr(x, MyString, "tr_watermark") - x
X1(8) = InStr(x, MyString, "levelcr (") - x
X1(9) = InStr(x, MyString, "pickup (") - x
X1(10) = InStr(x, MyString, "CollideObject (") - x
X1(11) = InStr(x, MyString, "Hazard (") - x
X1(12) = InStr(x, MyString, "Dyntrack (") - x
X1(13) = InStr(x, MyString, "Siding (") - x
X1(14) = InStr(x, MyString, "Platform (") - x
X1(15) = InStr(x, MyString, "CarSpawner (") - x
X1(16) = InStr(x, MyString, "Transfer (") - x
MinX = 0: MaxX = 0
Call FindMinMax(X1(), MinX, MaxX)

If MaxX = 0 Then

strEnd = vbCrLf & ")" & vbCrLf
Else
strEnd = Mid$(MyString, x + MinX)
End If
For j = 1 To 16
X1(j) = 0
Next j

X1(1) = x - InStrRev(MyString, "static (", x)
X1(2) = x - InStrRev(MyString, "gantry (", x)
X1(3) = x - InStrRev(MyString, "trackobj (", x)
X1(4) = x - InStrRev(MyString, "forest (", x)
X1(5) = x - InStrRev(MyString, "signal (", x)
X1(6) = x - InStrRev(MyString, "speedpost (", x)
X1(7) = x - InStrRev(MyString, "tr_watermark", x)
X1(8) = x - InStrRev(MyString, "levelcr (", x)
X1(9) = x - InStrRev(MyString, "pickup (", x)
X1(10) = x - InStrRev(MyString, "CollideObject (", x)
X1(11) = x - InStrRev(MyString, "Hazard (", x)
X1(12) = x - InStrRev(MyString, "Dyntrack (", x)
X1(13) = x - InStrRev(MyString, "Siding (", x)
X1(14) = x - InStrRev(MyString, "Platform (", x)
X1(15) = x - InStrRev(MyString, "CarSpawner (", x)
X1(16) = x - InStrRev(MyString, "Transfer (", x)
Call FindMinMax(X1(), MinX, MaxX)

strStart = Left$(MyString, x - MinX - 1)
For j = 1 To 15
X1(j) = 0
Next j
MyString = strStart & strEnd

GoTo CarryON

CarryOn2:

End If

The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertW = True

ERRHANDLER:

  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function


Public Sub FindMinMax(ByRef dArray() As Double, ByRef dLowVal As Double, ByRef dHighVal As Double)
    Dim lIndex As Long
    Dim dFirstValIdx As Double
    Dim dLastValIdx As Double
    Dim dActVal As Double
    Dim booFirstTime As Boolean
    
    dFirstValIdx = LBound(dArray)
    dLastValIdx = UBound(dArray)
    

    booFirstTime = True

    For lIndex = dFirstValIdx To dLastValIdx
        dActVal = dArray(lIndex)

        If dActVal > 0 Then
        If booFirstTime = True Then
        booFirstTime = False
        dHighVal = dActVal
        dLowVal = dActVal
        End If
        If dActVal > dHighVal Then
            dHighVal = dActVal
        End If


            If dActVal < dLowVal And dActVal > 0 Then
                dLowVal = dActVal
            End If
        
        End If
    Next lIndex
   
    
    
End Sub


  Private Function ConvertENV(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertENV = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then

xx = InStr(MyString, "world_water")
CarryON:
x = InStr(xx, MyString, "BlendATexDiff")
If x = 0 Then GoTo CarryOn2
xx = x + 1
x = InStr(xx, MyString, "BlendATexDiff")
If x = 0 Then GoTo CarryOn2
strStart = Left$(MyString, x + 8)
strEnd = Mid$(MyString, x + 13)
MyString = strStart & strEnd
'xx = x + 7
'GoTo CarryOn

CarryOn2:

End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertENV = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function

         
       
Private Function ConvertSMS(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertSMS = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then
x = InStr(MyString, "LoadAllWaves")
If x > 0 Then GoTo CarryON
x = InStr(MyString, "scalabiltygroup( 5")
If x = 0 Then
x = InStr(MyString, "scalabiltygroup( 4")
Else
GoTo CarryOn3
End If
If x = 0 Then
x = InStr(MyString, "scalabiltygroup( 3")
Else
GoTo CarryOn3
End If
If x = 0 Then
x = InStr(MyString, "scalabiltygroup( 2")
Else
GoTo CarryOn3
End If
If x = 0 Then
x = InStr(MyString, "scalabiltygroup( 1")
Else
GoTo CarryOn3
End If
If x = 0 Then GoTo CarryOn2
CarryOn3:
xx = InStr(x, MyString, "Streams")
xy = InStrRev(MyString, ")", xx)
strStart = Left$(MyString, xy)
strEnd = Mid$(MyString, xy + 1)
MyString = strStart & vbCrLf & vbTab & vbTab & "LoadAllWaves ( 1 )" & strEnd
CarryOn2:
x = InStr(MyString, "scalabiltygroup( 0")
If x = 0 Then GoTo CarryON
xx = InStr(x, MyString, "Streams")
xy = InStrRev(MyString, ")", xx)
strStart = Left$(MyString, xy)
strEnd = Mid$(MyString, xy + 1)
MyString = strStart & vbCrLf & vbTab & vbTab & "LoadAllWaves ( 1 )" & strEnd

CarryON:
End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertSMS = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function

         

         
         

         
         
Private Sub CheckTraffic(tempPath As String, tfcExists As Boolean, i As Long)
Dim x As Integer, strTfc As String

If Not FileExists(tempPath) Then
tfcExists = False
x = InStrRev(tempPath, "\")
strTfc = Mid$(tempPath, x + 1)
strForPrint = strForPrint & Lang(593) & " " & Lang(344) & strTfc & Lang(345) & Activities(i) & vbCrLf
Else
tfcExists = True
End If
End Sub

Private Sub CheckPath(tempPath As String, itExists As Boolean)


If Not FileExists(tempPath) Then
itExists = False
Else
itExists = True
End If
End Sub

Private Sub ListCheckService(tempPath As String)
Dim tempRoutePath As String, ConsistPath As String

On Error GoTo Errtrap


actOnly = False

x = InStr(tempPath, "Services")
tempRoutePath = Left$(tempPath, x - 1)
If Not FileExists(tempPath) Then
svcExists = False
Exit Sub
Else
svcExists = True
NewFile = FreeFile
Open tempPath For Input As #NewFile
 Do While Not EOF(NewFile)
 Line Input #NewFile, A$
 
 x = InStr(A$, "Train_Config")
 If x > 0 Then
 
 tempService = Trim$(Mid$(A$, x + 12))
 
 If Left$(tempService, 1) = "(" Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ")" Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 tempService = Trim$(tempService)
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 
tempService = tempService & ".con"

ConsistPath = MSTSPath & "\Trains\Consists"

Call ListLooseActConsists(ConsistPath & "\" & tempService)

Label2:
 End If
 
 Loop
 Close #NewFile

End If
Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'MiniCheckService' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub


Private Sub MiniCheckService(tempPath As String)
Dim tempRoutePath As String, ConsistPath As String

On Error GoTo Errtrap


actOnly = False

x = InStr(tempPath, "Services")
tempRoutePath = Left$(tempPath, x - 1)
If Not FileExists(tempPath) Then
svcExists = False
Exit Sub
Else
svcExists = True
NewFile = FreeFile
Open tempPath For Input As #NewFile
 Do While Not EOF(NewFile)
 Line Input #NewFile, A$
 
 x = InStr(A$, "Train_Config")
 If x > 0 Then
 
 tempService = Trim$(Mid$(A$, x + 12))
 
 If Left$(tempService, 1) = "(" Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ")" Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 tempService = Trim$(tempService)
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 
tempService = tempService & ".con"

ConsistPath = MSTSPath & "\Trains\Consists"
 'Call CheckPath(ConsistPath & "\" & tempService, conExists)

 If Not FileExists(ConsistPath & "\" & tempService) Then
 If FileExists(strConPath & tempService) Then
 FileCopy strConPath & tempService, ConsistPath & "\" & tempService
 DoEvents
 End If
 End If
 
Label2:
 End If
 
 Loop
 Close #NewFile

End If
Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'MiniCheckService' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub



Private Sub CheckService(tempPath As String, svcExists As Boolean, ActNum As Long, j As Integer)
Dim tempSvcPath As String, pathExists As Boolean, x As Long, NewFile As Integer, A$
Dim tempRoutePath As String, ConsistPath As String, conExists As Boolean, tempService As String

On Error GoTo Errtrap

'Write #31, "CheckService - " & right$(tempPath, 40) & " " & Now
actOnly = False

x = InStr(tempPath, "Services")
tempRoutePath = Left$(tempPath, x - 1)

If Not FileExists(tempPath) Then
svcExists = False
Exit Sub
Else
svcExists = True
NewFile = FreeFile
Open tempPath For Input As #NewFile
 Do While Not EOF(NewFile)
 Line Input #NewFile, A$
 
 x = InStr(A$, "Train_Config")
 If x > 0 Then
 tempService = Trim$(Mid$(A$, x + 12))
 
 If Left$(tempService, 1) = "(" Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ")" Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 tempService = Trim$(tempService)
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 
tempService = tempService & ".con"

ConsistPath = MSTSPath & "\Trains\Consists"
 Call CheckPath(ConsistPath & "\" & tempService, conExists)
 
 PConName(ActNum, j) = tempService
 
 If actOnly = True Then GoTo Label1
 If conExists = True And flagFull = True Then
 strForPrint = strForPrint & "Consist Name = " & tempService & vbCrLf
 ElseIf conExists = False Then
 strForPrint = strForPrint & "Consist Name = " & tempService & "  " & Lang(530) & " " & vbCrLf

 End If
Label1:
 End If
 x = InStr(A$, "PathID")
 If x > 0 Then
 tempSvcPath = Trim$(Mid$(A$, x + 7))
 If Left$(tempSvcPath, 1) = "(" Then
 tempSvcPath = Mid$(tempSvcPath, 2)
 End If
 If Right$(tempSvcPath, 1) = ")" Then
 tempSvcPath = Left$(tempSvcPath, Len(tempSvcPath) - 1)
 End If
 tempSvcPath = Trim$(tempSvcPath)
 If Left$(tempSvcPath, 1) = ChrW$(34) Then
 tempSvcPath = Mid$(tempSvcPath, 2)
 End If
 If Right$(tempSvcPath, 1) = ChrW$(34) Then
 tempSvcPath = Left$(tempSvcPath, Len(tempSvcPath) - 1)
 End If
 'tempSvcPath = Trim$(tempSvcPath)

pPathName(ActNum, j) = tempSvcPath & ".pat"
PathUsed(PathUsedNumb) = tempSvcPath & ".pat"
PathUsedNumb = PathUsedNumb + 1
If PathUsedNumb > UBound(PathUsed) Then
           ReDim Preserve PathUsed(0 To PathUsedNumb + CHUNK)
           
           End If
If actOnly = True Then GoTo Label2
 Call CheckPath(tempRoutePath & "Paths\" & tempSvcPath & ".pat", pathExists)
 If pathExists = True And flagFull = True Then
 strForPrint = strForPrint & "Path Name = " & tempSvcPath & ".pat" & vbCrLf
 ElseIf pathExists = False Then

 strForPrint = strForPrint & "Path Name = " & tempSvcPath & ".pat  " & Lang(596) & Lang(345) & Activities(ActNum) & vbCrLf
 End If
Label2:
 End If
 
 Loop
 Close #NewFile

End If
Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in " & tempService & " subroutine 'CheckService' please advise" _
            & vbCrLf & "j=" & Str(j) & " ActNum=" & Str(ActNum) & " temppath=" & tempPath _
            , vbExclamation, App.Title)
'Resume Next
End Sub



Private Sub CountShapes()
Dim lngShapes As Long, x As Integer, strShapes As String


On Error GoTo Errtrap
Filpath1$ = App.Path & "\TempFiles"
MousePointer = 11
 cursouind = 1
 Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.w"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i


strSearch = "FileName"

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
   
   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
     
  lngShapes = 0
   Open fullpath$ For Input As #5
   Do While Not EOF(5)
    
   Line Input #5, strNew
 
    
     x = InStr(strNew, "FileName")
     If x > 0 Then
     lngShapes = lngShapes + 1
     GoTo GetNext
     End If
     x = InStr(strNew, "Forest")
     If x > 0 Then
     lngShapes = lngShapes + 1
     GoTo GetNext
     End If
     x = InStr(strNew, "Platform")
     If x > 0 Then
     lngShapes = lngShapes + 1
     GoTo GetNext
     End If
     x = InStr(strNew, "CarSpawner")
     If x > 0 Then
     lngShapes = lngShapes + 1
     GoTo GetNext
     End If
     x = InStr(strNew, "Dyntrack")
     If x > 0 Then
     lngShapes = lngShapes + 1
     GoTo GetNext
     End If
GetNext:
     Loop
    
     End If
     strShapes = strShapes & File1(cursouind).List(i) & "   " & Str(lngShapes) & vbCrLf
     
     Close #5
     Next i
     MousePointer = 0
     frmReport.Rich1.Text = strShapes
     frmReport.Show 1
     
     DoEvents
     
     
     Exit Sub
Errtrap:
     Resume Next
End Sub

Private Sub GetConsists()
Dim ConsistPath As String
On Error GoTo Errtrap
ConsistPath = MSTSPath & "\Trains\Consists"
cursouind = 0
Drive1(cursouind).Drive = Left$(ConsistPath, 2)
Dir1(cursouind).Path = ConsistPath
Text1(cursouind).Text = "*.con"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
lngCon = File1(cursouind).ListCount
ReDim Consists(0 To lngCon - 1)
ReDim ConIntName(0 To lngCon - 1)
ReDim ConIntWagName(0 To lngCon - 1)
DoEvents

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
   Consists(i) = File1(cursouind).List(i)

   End If
   
   Next i
   
Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'GetConsists' in  please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next



End Sub

Private Sub WorldACESMS(strThisTile As String, jj As Integer)

Dim GlobalPath As String

On Error GoTo Errtrap

GlobalPath = MSTSPath & "\global\shapes\"




 Rem ****************** Show Reports
   MousePointer = 0

DoEvents

strReport = strReport & vbCrLf & "Tile #" & Str(jj) & "  " & strThisTile & vbCrLf
strReport = strReport & "---------------------------" & vbCrLf
'Call QSort3(Ace1(), 0, numAce - 1)
For i = 0 To numShp - 1
strReport = strReport & strShp(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & "Unique shapes = " & Str(numShp) & vbCrLf
strReport = strReport & vbCrLf & "Total  shapes = " & Str(totShp) & vbCrLf & vbCrLf
gtotShp = gtotShp + totShp

For i = 0 To numGlobShp - 1
strReport = strReport & strGlobShp(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & "Unique Track Shapes = " & Str(numGlobShp) & vbCrLf
strReport = strReport & vbCrLf & "Total  shapes = " & Str(totGlobshp) & vbCrLf & vbCrLf
gtotGlobShp = gtotGlobShp + totGlobshp

For i = 0 To numAce - 1
If Len(Ace2(i)) > 1 Then
strReport = strReport & Left$(Ace2(i), Len(Ace2(i)) - 4) & vbCrLf
End If
Next i
strReport = strReport & vbCrLf & "Unique Texture Files:-  " & Str(numAce) & vbCrLf
strReport = strReport & vbCrLf & "Total Texture Files:-  " & Str(totAce) & vbCrLf & vbCrLf
gtotAce = gtotAce + totAce

DoEvents
For i = 0 To numTerr - 1
strReport = strReport & TerrTex2(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & "Unique Terrain Texture Files:-  " & Str(numTerr) & vbCrLf
strReport = strReport & vbCrLf & "Total Terrain Texture Files:-  " & Str(totTerr) & vbCrLf & vbCrLf
 gtotTerr = gtotTerr + totTerr

DoEvents
If numTrans > 0 Then
For i = 0 To numTrans - 1
strReport = strReport & Transfer2(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & "Unique Transfer Files:-  " & Str(numTrans) & vbCrLf
strReport = strReport & vbCrLf & "Total Transfer Files:-  " & Str(totTrans) & vbCrLf & vbCrLf
gtotTrans = gtotTrans + totTrans
End If
DoEvents
If numHaz > 0 Then
For i = 0 To numHaz - 1
strReport = strReport & HazShp2(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & "Unique Hazard Files:-  " & Str(numHaz) & vbCrLf
strReport = strReport & vbCrLf & "Total Hazard Files:-  " & Str(totHaz) & vbCrLf & vbCrLf
gtotHaz = gtotHaz + totHaz

End If
 DoEvents
 If numFor > 0 Then
For i = 0 To numFor - 1
strReport = strReport & ForTex2(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & "Unique Forest Texture Files:-  " & Str(numFor) & vbCrLf
strReport = strReport & vbCrLf & "Total Forest Texture Files:-  " & Str(totFor) & vbCrLf & vbCrLf
gtotFor = gtotFor + totFor
 End If
  

  
   Exit Sub
Errtrap:

Call MsgBox("An error " & Err & " occurred in subroutine 'WorldACESMS' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next


End Sub

Private Sub CompactACESMS()
Dim tempPath As String
Dim GlobalPath As String

On Error GoTo Errtrap

GlobalPath = MSTSPath & "\global\shapes\"


Call CompactRef

 
 Call CheckDefaultSounds
 cursouind = 1
Drive1(cursouind).Drive = Left$(WorldPath, 2)
Dir1(cursouind).Path = WorldPath
Text1(cursouind).Text = "*.ws"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
      SB1.Panels(2).Text = File1(cursouind).List(i)

   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   Call CheckForSounds(fullpath$)
   File1(cursouind).Selected(i) = False
   End If
   
   Next i
If FileExists(RoutePath & "\ttype.dat") Then
Call CheckForSounds(RoutePath & "\ttype.dat")
End If


For i = 0 To SoundNumber

If FileExists(SoundPath & "\" & Soundfile(i)) Then
  tempPath = SoundPath & "\" & Soundfile(i)
  
  ElseIf FileExists(GlobalSoundPath & "\" & Soundfile(i)) Then
  tempPath = GlobalSoundPath & "\" & Soundfile(i)
  
  End If
  Call CheckForWav(tempPath)
Next i

 

Label3:


 Rem ****************** Show Reports
   MousePointer = 0

DoEvents
'Call QSort3(Ace1(), 0, numAce - 1)
For i = 0 To numAce - 1
If Ace2(i) <> vbNullString And Ace2(i) <> "|" Then
strReport = strReport & Left$(Ace2(i), Len(Ace2(i)) - 4) & vbCrLf
End If
Next i
strReport = strReport & vbCrLf & vbCrLf & "Total Texture Files:-  " & Str(numAce) & vbCrLf & vbCrLf

 Call QSort(WavFile(), 1, WavNumber)
DoEvents
For i = 1 To WavNumber
strReport = strReport & WavFile(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & vbCrLf & "Total .WAV Files:-  " & Str(WavNumber) & vbCrLf & vbCrLf

Call QSort(Soundfile(), 0, SoundNumber)
DoEvents
For i = 0 To SoundNumber
strReport = strReport & Soundfile(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & vbCrLf & "Total Sound Files:-  " & Str(SoundNumber) & vbCrLf & vbCrLf

 

DoEvents
For i = 0 To numTerr - 1
strReport = strReport & TerrTex2(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & vbCrLf & "Total Terrain Texture Files:-  " & Str(numTerr) & vbCrLf & vbCrLf
 

DoEvents
If numTrans > 0 Then
For i = 0 To numTrans - 1
strReport = strReport & Transfer2(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & vbCrLf & "Total Transfer Files:-  " & Str(numTrans) & vbCrLf & vbCrLf
 

End If
DoEvents
If numHaz > 0 Then
For i = 0 To numHaz - 1
strReport = strReport & HazShp2(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & vbCrLf & "Total Hazard Files:-  " & Str(numHaz) & vbCrLf & vbCrLf
End If
 DoEvents
 If numFor > 0 Then
For i = 0 To numFor - 1
strReport = strReport & ForTex2(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & vbCrLf & "Total Forest Texture Files:-  " & Str(numFor) & vbCrLf & vbCrLf
 End If
  

  
   Exit Sub
Errtrap:

Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'CompactACESMS' please advise" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
        'Resume Next
    Exit Sub
    End Select

End Sub



Private Sub GetCoupling(LocoPath As String, flagCoup As Integer, intBrake As Integer, intType As Integer, strName As String, intCouplings As Integer, flagRigid As String, flagFCoup As Integer, strSMS As String)
Dim NewFile As Integer, A$, x As Long, xx As Long, strTender As String, Y As Long
Dim i As Integer, strBrake As String, xq As Long, booBrakes As Boolean, j As Long, jj As Long

On Error GoTo Errtrap
flagCoup = 0
NewFile = FreeFile
'Open LocoPath For Input As #NewFile
'Do While Not EOF(NewFile)
' Line Input #NewFile, A$

flagCoup = 0
flagFCoup = 0

A$ = ReadUniFile(LocoPath)
A$ = Replace(A$, "Type (Automatic", "Type ( Automatic")
A$ = Replace(A$, "Type (Chain", "Type ( Chain")
A$ = Replace(A$, "          ", " ")
A$ = Replace(A$, "         ", " ")
A$ = Replace(A$, "        ", " ")
A$ = Replace(A$, "       ", " ")
A$ = Replace(A$, "      ", " ")
A$ = Replace(A$, "     ", " ")
A$ = Replace(A$, "    ", " ")
A$ = Replace(A$, "   ", " ")
A$ = Replace(A$, "  ", " ")
A$ = Replace(A$, " ", " ")
 x = InStr(A$, "Type ( Engine )")
 If x > 0 Then intType = 1
 x = InStr(A$, "Type ( Freight )")
 If x > 0 Then intType = 2
 x = InStr(A$, "Type ( Tender )")
 If x > 0 Then intType = 4
 x = InStr(A$, "Type ( Carriage )")
 If x > 0 Then intType = 3
 xq = InStr(A$, "IsTenderRequired")
If xq > 0 Then
xx = InStr(xq, A$, "(")
xy = InStr(xx, A$, ")")
strTender = Trim$(Mid$(A$, xx + 1, xy - xx - 1))
If strTender = "1" Then
intType = 5
End If
End If

j = InStr(A$, "Coupling (")
If j = 0 Then intCouplings = 0
If j > 0 Then
jj = InStr(j + 5, A$, "Coupling (")
If jj > 0 Then
intCouplings = 2
Else
intCouplings = 1
End If

End If
If intCouplings = 1 Then
Y = InStr(A$, "Type ( Automatic")
If Y > 0 Then flagCoup = 1
If Y = 0 Then
Y = InStr(A$, "Type ( Chain")
If Y > 0 Then flagCoup = 2
End If
If Y = 0 Then
Y = InStr(A$, "Type ( Bar")
If Y > 0 Then flagCoup = 3
End If
ElseIf intCouplings = 2 Then
Y = InStr(A$, "Type ( Automatic")
If Y > 0 And Y < jj Then
flagCoup = 1
GoTo NextCoup
End If
Y = InStr(A$, "Type ( Chain")
If Y > 0 And Y < jj Then
flagCoup = 2
GoTo NextCoup
End If
Y = InStr(A$, "Type ( Bar")
If Y > 0 And Y < jj Then
flagCoup = 3
End If
NextCoup:
'End If
Rem *********** Second coupler
Y = InStr(jj, A$, "Type ( Automatic")
If Y > jj Then flagFCoup = 1
If Y < jj Then
Y = InStr(jj, A$, "Type ( Chain")
If Y > jj Then flagFCoup = 2
End If
If Y < jj Then
Y = InStr(jj, A$, "Type ( Bar")
If Y > jj Then flagFCoup = 3
End If
End If
x = InStr(A$, "brakesystemtype")
If x > 0 Then
booBrakes = True
xx = InStr(x, A$, "(")
xy = InStr(xx, A$, ")")
strBrake = Mid$(A$, xx + 1, xy - xx - 1)
strBrake = Trim$(strBrake)
If Left$(strBrake, 1) = ChrW$(34) Then
strBrake = Mid$(strBrake, 2, Len(strBrake) - 2)
strBrake = Trim$(strBrake)
End If
For i = 1 To 8
If strBrake = Brake(i) Then
intBrake = i
Exit For
End If
Next i
End If

x = InStr(A$, "Name (")
If x = 0 Then
x = InStr(A$, "Name  (")
End If
If x = 0 Then
x = InStr(A$, "Name   (")
End If
If x = 0 Then
x = InStr(A$, "Name(")
End If
If x > 0 Then
xx = InStr(x, A$, "(")
xy = InStr(xx, A$, ")")
strName = Mid$(A$, xx + 1, xy - xx - 1)
strName = Trim$(strName)
If Left$(strName, 1) = ChrW$(34) Then
strName = Mid$(strName, 2, Len(strName) - 2)
End If

End If
'Loop
'Close NewFile

If booBrakes = False Then
intBrake = 0
End If

x = InStr(A$, "Type ( Steam")
If x <> 0 Then
    If strTender = "1" Then
    intType = 5
    Else
    intType = 6
    End If
End If
x = InStr(A$, "Type ( Diesel")
If x <> 0 Then intType = 7
x = InStr(A$, "Type ( Electric")
If x <> 0 Then intType = 8

x = InStr(A$, "CouplingHasRigidConnection")
If x > 0 Then
xy = InStrRev(A$, "comment", x)
If xy > x - 15 Then
flagRigid = "1"
GoTo CarryON
End If
If Mid$(A$, x + 27, 2) = "()" Then
flagRigid = "2"
GoTo CarryON
End If
If Mid$(A$, x + 27, 5) = "( 1 )" Then
flagRigid = "3"
GoTo CarryON
End If
flagRigid = "0"
End If
CarryON:
x = InStr(x + 10, A$, "CouplingHasRigidConnection")
If x > 0 Then
xy = InStrRev(A$, "comment", x)
If xy > x - 15 Then
flagRigid = flagRigid & "1"
GoTo CarryON
End If
If Mid$(A$, x + 27, 2) = "()" Then
flagRigid = flagRigid & "2"
GoTo CarryON
End If
If Mid$(A$, x + 27, 5) = "( 1 )" Then
flagRigid = flagRigid & "3"
GoTo CarryON
End If
flagRigid = flagRigid & "0"
End If
Rem *************** Get Sound
If intType > 4 Then

x = InStr(A$, "Sound (")
If x = 0 Then
strSMS = ""
Else
xx = InStr(x, A$, ")")
strSMS = Mid(A$, x + 7, xx - (x + 7))
strSMS = Replace(strSMS, ChrW$(34), " ")
strSMS = Trim(strSMS)

End If
End If

Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'GetCoupling' please advise" _
            & vbCrLf & "while checking " & LocoPath _
            , vbExclamation, App.Title)
Resume Next

End Sub

Private Sub GetPaths2()
Dim strPathsPath As String, strPathsName As String
On Error GoTo Errtrap
MousePointer = 11
'RoutePath = MSTSPath & "\Routes"
cursouind = 0
Drive1(cursouind).Drive = Left$(MSTSPath, 2)
Dir1(cursouind).Path = RoutePath
Text1(cursouind) = "*.*"
DoEvents
'strForPrint = vbNullString

If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
        Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If


Dir1(cursouind).Path = RoutePath & "\Paths"
Text1(cursouind).Text = "*.pat"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
    If File1(cursouind).Selected(i) Then
    
    strPathsName = File1(cursouind).List(i)
    strPathsPath = RoutePath & "\Paths"
    If Not DirExists(strPathsPath) Then GoTo TryAnother
           
           If lngPaths > UBound(Paths) Then
           ReDim Preserve Paths(0 To lngPaths + CHUNK)
           ReDim Preserve PathsPath(0 To lngPaths + CHUNK)
           End If
            Paths(lngPaths) = strPathsName
            PathsPath(lngPaths) = strPathsPath
            lngPaths = lngPaths + 1
            End If
TryAnother:
        Next i

DoEvents
 ReDim Preserve Paths(0 To lngPaths - 1)
 ReDim Preserve PathsPath(0 To lngPaths - 1)
MousePointer = 0
Exit Sub
Errtrap:
If Err = 76 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'GetPaths' in " & strPathsName & " please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub
Private Sub GetPaths()
Dim jj As Integer, strPathsPath As String, strPathsName As String
On Error GoTo Errtrap
MousePointer = 11
RoutePath = MSTSPath & "\Routes"
cursouind = 0
Drive1(cursouind).Drive = Left$(MSTSPath, 2)
Dir1(cursouind).Path = RoutePath
Text1(cursouind) = "*.*"
DoEvents
'strForPrint = vbNullString

If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
        Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If

For jj = 0 To NumRoutes - 1
Dir1(cursouind).Path = AllRoutes(jj) & "\Paths"
Text1(cursouind).Text = "*.pat"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
    If File1(cursouind).Selected(i) Then
    'x = InStrRev(File1(cursouind).list(i), "\")
    strPathsName = File1(cursouind).List(i)
    strPathsPath = AllRoutes(jj) & "\Paths"
    If Not DirExists(strPathsPath) Then GoTo TryAnother
           
           If lngPaths > UBound(Paths) Then
           ReDim Preserve Paths(0 To lngPaths + CHUNK)
           ReDim Preserve PathsPath(0 To lngPaths + CHUNK)
           End If
            Paths(lngPaths) = strPathsName
            PathsPath(lngPaths) = strPathsPath
            lngPaths = lngPaths + 1
            End If
TryAnother:
        Next i
Next jj
   ' File1(cursouind).path = Dir1(cursouind).path
DoEvents
 ReDim Preserve Paths(0 To lngPaths - 1)
 ReDim Preserve PathsPath(0 To lngPaths - 1)
MousePointer = 0
Exit Sub
Errtrap:
If Err = 76 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'GetPaths' in " & strPathsName & " please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub

Private Sub GetTraffic2()
Dim strTfcPath As String, strTfcName As String
On Error GoTo Errtrap
MousePointer = 11
'RoutePath = MSTSPath & "\Routes"
cursouind = 0
Drive1(cursouind).Drive = Left$(MSTSPath, 2)
Dir1(cursouind).Path = RoutePath
Text1(cursouind) = "*.*"
DoEvents
'strForPrint = vbNullString

If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
        Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If


Dir1(cursouind).Path = RoutePath & "\Traffic"
Text1(cursouind).Text = "*.trf"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

    If File1(cursouind).Selected(i) Then
 
    strTfcName = File1(cursouind).List(i)
    strTfcPath = RoutePath & "\Traffic"
    If Not DirExists(strTfcPath) Then GoTo TryAnother
           
           If lngTfc > UBound(Traffic) Then
           ReDim Preserve Traffic(0 To lngTfc + CHUNK)
           ReDim Preserve TfcPath(0 To lngTfc + CHUNK)
           End If
            Traffic(lngTfc) = strTfcName
            TfcPath(lngTfc) = strTfcPath
            lngTfc = lngTfc + 1
            End If
TryAnother:
        Next i

    File1(cursouind).Path = Dir1(cursouind).Path
DoEvents
If lngTfc > 0 Then
 ReDim Preserve Traffic(0 To lngTfc - 1)
 ReDim Preserve TfcPath(0 To lngTfc - 1)
 End If
MousePointer = 0
Exit Sub
Errtrap:
If Err = 76 Then
Resume Next
End If
    Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'Get Traffic' please advise" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
    Exit Sub
    End Select
End Sub
Private Sub GetTraffic()
Dim jj As Integer, strTfcPath As String, strTfcName As String
On Error GoTo Errtrap
MousePointer = 11
RoutePath = MSTSPath & "\Routes"
cursouind = 0
Drive1(cursouind).Drive = Left$(MSTSPath, 2)
Dir1(cursouind).Path = RoutePath
Text1(cursouind) = "*.*"
DoEvents
'strForPrint = vbNullString

If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
        Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If

For jj = 0 To NumRoutes - 1
Dir1(cursouind).Path = AllRoutes(jj) & "\Traffic"
Text1(cursouind).Text = "*.trf"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

    If File1(cursouind).Selected(i) Then
  
    strTfcName = File1(cursouind).List(i)
    strTfcPath = AllRoutes(jj) & "\Traffic"
    If Not DirExists(strTfcPath) Then GoTo TryAnother
           
           If lngTfc > UBound(Traffic) Then
           ReDim Preserve Traffic(0 To lngTfc + CHUNK)
           ReDim Preserve TfcPath(0 To lngTfc + CHUNK)
           End If
            Traffic(lngTfc) = strTfcName
            TfcPath(lngTfc) = strTfcPath
            lngTfc = lngTfc + 1
            End If
TryAnother:
        Next i
Next jj
    File1(cursouind).Path = Dir1(cursouind).Path
DoEvents
If lngTfc > 0 Then
 ReDim Preserve Traffic(0 To lngTfc - 1)
 ReDim Preserve TfcPath(0 To lngTfc - 1)
 End If
MousePointer = 0
Exit Sub
Errtrap:
If Err = 76 Then
Resume Next
End If
    Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'Get Traffic' please advise" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
    Exit Sub
    End Select
End Sub

Private Sub GetServices2()

Dim strSrvName As String, strSrvPath As String

On Error GoTo Errtrap
MousePointer = 11
'RoutePath = MSTSPath & "\Routes"
cursouind = 0
Drive1(cursouind).Drive = Left$(MSTSPath, 2)
Dir1(cursouind).Path = RoutePath
Text1(cursouind) = "*.*"
DoEvents


If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
        Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
        Exit Sub
    End If
'For jj = 0 To NumRoutes - 1
Dir1(cursouind).Path = RoutePath & "\Services\"
Text1(cursouind).Text = "*.srv"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

    If File1(cursouind).Selected(i) Then
    'x = InStrRev(File1(cursouind).list(i), "\")
    strSrvName = File1(cursouind).List(i)
    strSrvPath = RoutePath & "\Services"
    If Not DirExists(strSrvPath) Then GoTo TryAnother
           
           If lngSrv > UBound(Service) Then
           ReDim Preserve Service(0 To lngSrv + CHUNK)
           ReDim Preserve SrvPath(0 To lngSrv + CHUNK)
           End If
            Service(lngSrv) = strSrvName
            SrvPath(lngSrv) = strSrvPath
            lngSrv = lngSrv + 1
            End If
TryAnother:
        Next i
'Next jj

   ' File1(cursouind).path = Dir1(cursouind).path
DoEvents
 ReDim Preserve Service(0 To lngSrv - 1)
 ReDim Preserve SrvPath(0 To lngSrv - 1)
MousePointer = 0
Exit Sub
Errtrap:
If Err = 76 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'GetServices' in " & strSrvName & " please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub

Private Sub GetServices()

Dim jj As Integer, strSrvName As String, strSrvPath As String

On Error GoTo Errtrap
MousePointer = 11
RoutePath = MSTSPath & "\Routes"
cursouind = 0
Drive1(cursouind).Drive = Left$(MSTSPath, 2)
Dir1(cursouind).Path = RoutePath
Text1(cursouind) = "*.*"
DoEvents


If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
        Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
        Exit Sub
    End If
For jj = 0 To NumRoutes - 1
Dir1(cursouind).Path = AllRoutes(jj) & "\Services\"
Text1(cursouind).Text = "*.srv"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

    If File1(cursouind).Selected(i) Then
    'x = InStrRev(File1(cursouind).list(i), "\")
    strSrvName = File1(cursouind).List(i)
    strSrvPath = AllRoutes(jj) & "\Services"
    If Not DirExists(strSrvPath) Then GoTo TryAnother
           
           If lngSrv > UBound(Service) Then
           ReDim Preserve Service(0 To lngSrv + CHUNK)
           ReDim Preserve SrvPath(0 To lngSrv + CHUNK)
           End If
            Service(lngSrv) = strSrvName
            SrvPath(lngSrv) = strSrvPath
            lngSrv = lngSrv + 1
            End If
TryAnother:
        Next i
Next jj

   ' File1(cursouind).path = Dir1(cursouind).path
DoEvents
 ReDim Preserve Service(0 To lngSrv - 1)
 ReDim Preserve SrvPath(0 To lngSrv - 1)
MousePointer = 0
Exit Sub
Errtrap:
If Err = 76 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'GetServices' in " & strSrvName & " please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub


Private Sub GetActivities2()
Dim i As Long, ii As Integer
Dim tempRoutePath As String, x As Integer, tempService As String
Dim svcExists As Boolean, Y As Integer, tfcExists As Boolean, j As Integer
Dim NameExists As Boolean, strActName As String, strActPath As String
On Error GoTo Errtrap
MousePointer = 11
'RoutePath = MSTSPath & "\Routes"
cursouind = 0
Drive1(cursouind).Drive = Left$(MSTSPath, 2)
Dir1(cursouind).Path = RoutePath
Text1(cursouind) = "*.*"
DoEvents


If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
        Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
        Exit Sub
    End If

'For jj = 0 To NumRoutes - 1

Drive1(cursouind).Drive = Left$(RoutePath, 2)

Dir1(cursouind).Path = RoutePath & "\Activities"
Text1(0).Text = "*.act"
DoEvents
If File1(0).ListCount = 0 Then
Exit Sub
Else

DoEvents

For i = 0 To File1(0).ListCount - 1
    
    strActName = File1(cursouind).List(i)
   
    strActPath = RoutePath & "\Activities"
'    If Not DirExists(strActPath) Then
'
'    GoTo TryAnother
'    End If
           
           If lngAct > UBound(Activities) Then
           ReDim Preserve Activities(0 To lngAct + CHUNK)
           ReDim Preserve ActPath(0 To lngAct + CHUNK)
           End If
            Activities(lngAct) = strActName
            ActPath(lngAct) = strActPath
            lngAct = lngAct + 1
           ' End If
TryAnother:
        Next i
        End If
TryMore:


    File1(cursouind).Path = Dir1(cursouind).Path
DoEvents

 ReDim Preserve Activities(0 To lngAct - 1)
 ReDim Preserve ActPath(0 To lngAct - 1)
 ReDim ActEng(0 To lngAct - 1, 0 To lngLoco + 500)
 ReDim ActWag(0 To lngAct - 1, 0 To lngWagons + 500)
 ReDim pPathName(0 To lngAct - 1, 0 To CHUNK)
 ReDim PConName(0 To lngAct - 1, 0 To CHUNK)
 ReDim PSvcName(0 To lngAct - 1, 0 To CHUNK)
 ReDim PTfcName(0 To lngAct - 1)
 
 
 


MousePointer = 11

For i = 0 To lngAct - 1

booPlayer = True
x = InStr(ActPath(i), "Activities")
tempRoutePath = Left$(ActPath(i), x - 1)

NewFile = FreeFile

 Open ActPath(i) & "\" & Activities(i) For Input As #NewFile

  j = 0
 Do While Not EOF(NewFile)
Line Input #NewFile, A$

 Rem ******** Get Players Consist ************
 x = InStr(A$, "Player_Service_Definition")
 If x > 0 Then
 tempService = Trim$(Mid$(A$, x + 27))
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If

tempService = tempService & ".srv"
 tempService = Trim$(tempService)

 Call CheckService(tempRoutePath & "Services\" & tempService, svcExists, i, j)
For ii = 0 To j
If PSvcName(i, ii) = tempService Then
NameExists = True
Exit For
End If
Next ii
If NameExists = False Then
 PSvcName(i, j) = tempService
 If j > UBound(PSvcName, 2) Then
     ReDim Preserve PSvcName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve pPathName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve PConName(0 To lngAct - 1, 0 To j + CHUNK)
      End If
 j = j + 1
 Else
 NameExists = False
 End If
 
 End If
 tempService = vbNullString
 Rem ************  Get Traffic Definition
 tempService = vbNullString

 x = InStr(A$, "Traffic_Definition")
 Y = InStr(A$, "Player")
 If x > 0 And Y = 0 Then

 tempService = Trim$(Mid$(A$, x + 18))
 
 tempService = Trim$(tempService)
 
 If Left$(tempService, 1) = "(" Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ")" Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 tempService = Trim$(tempService)
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If

tempService = tempService & ".trf"
tempService = Trim$(tempService)
PTfcName(i) = tempService

 Call CheckTraffic(tempRoutePath & "Traffic\" & tempService, tfcExists, i)
 
 End If
 Rem *************** Get AI Traffic
 x = InStr(A$, "Service_Definition")
 Y = InStr(A$, "Player_Service")
 If x > 0 And Y = 0 Then
 booPlayer = False
 tempService = Trim$(Mid$(A$, x + 18))
 Y = InStrRev(tempService, " ")
 tempService = Left$(tempService, Y)
 tempService = Trim$(tempService)
 
 If Left$(tempService, 1) = "(" Then
 tempService = Mid$(tempService, 2)
 End If

 tempService = Trim$(tempService)
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 
tempService = tempService & ".srv"
tempService = Trim$(tempService)

 Call CheckService(tempRoutePath & "Services\" & tempService, svcExists, i, j)
 
 For ii = 0 To j
If PSvcName(i, ii) = tempService Then
NameExists = True
Exit For
End If
Next ii
If NameExists = False Then
 PSvcName(i, j) = tempService
 If j > UBound(PSvcName, 2) Then
     ReDim Preserve PSvcName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve pPathName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve PConName(0 To lngAct - 1, 0 To j + CHUNK)
      End If
 j = j + 1
 
 Else
 NameExists = False
 End If
 
 End If
 
 
 Loop
 Close #NewFile

 
Next
MousePointer = 0




Exit Sub
Errtrap:

Call MsgBox("An error " & Err & " occurred in subroutine 'GetActivities' in " & strActName & " please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)


Resume Next
End Sub


Private Sub GetActivities()
Dim i As Long, ii As Integer
Dim tempRoutePath As String, x As Integer, tempService As String
Dim svcExists As Boolean, Y As Integer, tfcExists As Boolean, j As Integer
Dim NameExists As Boolean, jj As Integer, strActName As String, strActPath As String
On Error GoTo Errtrap
MousePointer = 11
RoutePath = MSTSPath & "\Routes"
cursouind = 0
Drive1(cursouind).Drive = Left$(MSTSPath, 2)
Dir1(cursouind).Path = RoutePath
Text1(cursouind) = "*.*"
DoEvents


If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
        Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
        Exit Sub
    End If

For jj = 0 To NumRoutes - 1

Drive1(cursouind).Drive = Left$(AllRoutes(jj), 2)
If Not DirExists(AllRoutes(jj) & "\Activities") Then GoTo TryMore
Dir1(cursouind).Path = AllRoutes(jj) & "\Activities"
Text1(0).Text = "*.act"
DoEvents
If File1(0).ListCount = 0 Then

Else

DoEvents

For i = 0 To File1(0).ListCount - 1
    
    strActName = File1(cursouind).List(i)
   
    strActPath = AllRoutes(jj) & "\Activities"
    If Not DirExists(strActPath) Then
    
    GoTo TryAnother
    End If
           
           If lngAct > UBound(Activities) Then
           ReDim Preserve Activities(0 To lngAct + CHUNK)
           ReDim Preserve ActPath(0 To lngAct + CHUNK)
           End If
            Activities(lngAct) = strActName
            ActPath(lngAct) = strActPath
            lngAct = lngAct + 1
           ' End If
TryAnother:
        Next i
        End If
TryMore:
Next jj

If lngAct = 0 Then
Call MsgBox("There are no Activities in this Route's Activities" _
            & vbCrLf & "folder, option aborted." _
            , vbCritical, App.Title)
booAbort = True
Exit Sub
End If

    File1(cursouind).Path = Dir1(cursouind).Path
DoEvents

 ReDim Preserve Activities(0 To lngAct - 1)
 ReDim Preserve ActPath(0 To lngAct - 1)
 ReDim ActEng(0 To lngAct - 1, 0 To lngLoco + 500)
 ReDim ActWag(0 To lngAct - 1, 0 To lngWagons + 500)
 ReDim pPathName(0 To lngAct - 1, 0 To CHUNK)
 ReDim PConName(0 To lngAct - 1, 0 To CHUNK)
 ReDim PSvcName(0 To lngAct - 1, 0 To CHUNK)
 ReDim PTfcName(0 To lngAct - 1)
 
 
 


MousePointer = 11

For i = 0 To lngAct - 1

booPlayer = True
x = InStr(ActPath(i), "Activities")
tempRoutePath = Left$(ActPath(i), x - 1)

NewFile = FreeFile

 Open ActPath(i) & "\" & Activities(i) For Input As #NewFile

  j = 0
 Do While Not EOF(NewFile)
Line Input #NewFile, A$

 Rem ******** Get Players Consist ************
 x = InStr(A$, "Player_Service_Definition")
 If x > 0 Then
 tempService = Trim$(Mid$(A$, x + 27))
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If

tempService = tempService & ".srv"
 tempService = Trim$(tempService)

 Call CheckService(tempRoutePath & "Services\" & tempService, svcExists, i, j)
For ii = 0 To j
If PSvcName(i, ii) = tempService Then
NameExists = True
Exit For
End If
Next ii
If NameExists = False Then
 PSvcName(i, j) = tempService
 If j > UBound(PSvcName, 2) Then
     ReDim Preserve PSvcName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve pPathName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve PConName(0 To lngAct - 1, 0 To j + CHUNK)
      End If
 j = j + 1
 Else
 NameExists = False
 End If
 
 End If
 tempService = vbNullString
 Rem ************  Get Traffic Definition
 tempService = vbNullString

 x = InStr(A$, "Traffic_Definition")
 Y = InStr(A$, "Player")
 If x > 0 And Y = 0 Then

 tempService = Trim$(Mid$(A$, x + 18))
 
 tempService = Trim$(tempService)
 
 If Left$(tempService, 1) = "(" Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ")" Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 tempService = Trim$(tempService)
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If

tempService = tempService & ".trf"
tempService = Trim$(tempService)
PTfcName(i) = tempService

 Call CheckTraffic(tempRoutePath & "Traffic\" & tempService, tfcExists, i)
 
 End If
 Rem *************** Get AI Traffic
 x = InStr(A$, "Service_Definition")
 Y = InStr(A$, "Player_Service")
 If x > 0 And Y = 0 Then
 booPlayer = False
 tempService = Trim$(Mid$(A$, x + 18))
 Y = InStrRev(tempService, " ")
 tempService = Left$(tempService, Y)
 tempService = Trim$(tempService)
 
 If Left$(tempService, 1) = "(" Then
 tempService = Mid$(tempService, 2)
 End If

 tempService = Trim$(tempService)
 If Left$(tempService, 1) = ChrW$(34) Then
 tempService = Mid$(tempService, 2)
 End If
 If Right$(tempService, 1) = ChrW$(34) Then
 tempService = Left$(tempService, Len(tempService) - 1)
 End If
 
tempService = tempService & ".srv"
tempService = Trim$(tempService)

 Call CheckService(tempRoutePath & "Services\" & tempService, svcExists, i, j)
 
 For ii = 0 To j
If PSvcName(i, ii) = tempService Then
NameExists = True
Exit For
End If
Next ii
If NameExists = False Then
 PSvcName(i, j) = tempService
 If j > UBound(PSvcName, 2) Then
     ReDim Preserve PSvcName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve pPathName(0 To lngAct - 1, 0 To j + CHUNK)
     ReDim Preserve PConName(0 To lngAct - 1, 0 To j + CHUNK)
      End If
 j = j + 1
 
 Else
 NameExists = False
 End If
 
 End If
 
 
 Loop
 Close #NewFile

 
Next
MousePointer = 0




Exit Sub
Errtrap:

Call MsgBox("An error " & Err & " occurred in subroutine 'GetActivities' in " & strActName & " please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)


Resume Next
End Sub


Private Sub GetStock2()
Dim TrainPath As String, i As Integer
Dim LocomotivePath As String
Dim ConsistPath As String, fullpath$, strName As String
Dim intEng As Integer, intWag As Integer, intBrake As Integer, intType As Integer, intCouplings As Integer
Dim LocoUsed() As Boolean, WagUsed() As Boolean, flagRigid As String, flagFCoup As Integer, strSMS As String
On Error GoTo Errtrap

MousePointer = 11

cursouind = 0

TrainPath = Trainspath & "\trainset"
Consistspath = TrainPath & "\Consists"

strForPrint2 = vbNullString
For i = 0 To lngLoco - 1
SB2.Panels(2).Text = Locomotives(i)
LocomotivePath = TrainPath & "\" & LocoPath(i) & "\" & Locomotives(i)

Call GetCoupling(LocomotivePath, flagCouple, intBrake, intType, strName, intCouplings, flagRigid, flagFCoup, strSMS)

LocoName(i) = strName
LocoCoup(i) = flagCouple
LocoFCoup(i) = flagFCoup
LocoBrake(i) = intBrake
LocoType(i) = intType
If Len(flagRigid) = 1 Then
LocoRigid(i) = Val(flagRigid)
LocoFRigid(i) = 0
ElseIf Len(flagRigid) = 2 Then
LocoRigid(i) = Val(Left$(flagRigid, 1))
LocoFRigid(i) = Val(Right$(flagRigid, 1))
End If
DoEvents
strName = vbNullString

Next i


For i = 0 To lngWagons - 1
SB2.Panels(2).Text = Wagons(i)
LocomotivePath = TrainPath & "\" & Wagpath(i) & "\" & Wagons(i)
Call GetCoupling(LocomotivePath, flagCouple, intBrake, intType, strName, intCouplings, flagRigid, flagFCoup, strSMS)
WagonName(i) = strName
WagCoup(i) = flagCouple
WagFCoup(i) = flagFCoup
WagBrake(i) = intBrake
WagType(i) = intType
If Len(flagRigid) = 1 Then
WagRigid(i) = Val(flagRigid)
WagFRigid(i) = 0
ElseIf Len(flagRigid) = 2 Then
WagRigid(i) = Val(Left$(flagRigid, 1))
WagFRigid(i) = Val(Right$(flagRigid, 1))
End If
DoEvents
strName = vbNullString


Next i
ReDim LocoUsed(0 To lngLoco - 1)
ReDim WagUsed(0 To lngWagons - 1)




  
ConsistPath = MSTSPath & "\Trains\Consists"
cursouind = 0
Drive1(cursouind).Drive = Left$(ConsistPath, 2)
Dir1(cursouind).Path = ConsistPath
Text1(cursouind).Text = "*.con"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
lngCon = File1(cursouind).ListCount
ReDim Consists(0 To lngCon - 1)
ReDim ConIntName(0 To lngCon - 1)
ReDim ConIntWagName(0 To lngCon - 1)
ReDim ConEng(0 To lngCon - 1, 0 To 80)
ReDim conWag(0 To lngCon - 1, 0 To 500)


For i = 0 To File1(cursouind).ListCount - 1
intEng = 0
intWag = 0
   If File1(cursouind).Selected(i) Then
      fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
SB2.Panels(2).Text = File1(cursouind).List(i)

      Call CheckForConsist(fullpath$, i, intEng, intWag)

      If intEng > 0 Then
      For j = 1 To intEng
      LocoUsed(ConEng(i, j)) = True
      frmStock.GridStock.AddItem File1(cursouind).List(i) & vbTab & ConIntName(i) & vbTab & LocoPath(ConEng(i, j)) & vbTab & Locomotives(ConEng(i, j)) & vbTab _
& LocoName(ConEng(i, j)) & vbTab & Coupling(LocoCoup(ConEng(i, j))) & vbTab & Brake(LocoBrake(ConEng(i, j))) & vbTab & StockType(LocoType(ConEng(i, j))) & vbTab & FCoupling(LocoFCoup(ConEng(i, j))) & vbTab & Rigid(LocoRigid(ConEng(i, j))) & vbTab & Rigid(LocoFRigid(ConEng(i, j)))
Next j
End If

If intWag > 0 Then
      For j = 1 To intWag
      If conWag(i, j) = 9999 Then

Else
      WagUsed(conWag(i, j)) = True
      frmStock.GridStock.AddItem File1(cursouind).List(i) & vbTab & ConIntName(i) & vbTab & Wagpath(conWag(i, j)) & vbTab & Wagons(conWag(i, j)) & vbTab _
& WagonName(conWag(i, j)) & vbTab & Coupling(WagCoup(conWag(i, j))) & vbTab & Brake(WagBrake(conWag(i, j))) & vbTab & StockType(WagType(conWag(i, j))) & vbTab & FCoupling(WagFCoup(conWag(i, j))) & vbTab & Rigid(WagRigid(conWag(i, j))) & Rigid(WagFRigid(conWag(i, j)))
End If
Next j
End If

     
   File1(cursouind).Selected(i) = False
   End If
Next i




For i = 0 To lngLoco - 1
If LocoUsed(i) = False Then
frmStock.GridStock.AddItem vbTab & vbTab & LocoPath(i) & vbTab & Locomotives(i) & vbTab _
& LocoName(i) & vbTab & Coupling(LocoCoup(i)) & vbTab & Brake(LocoBrake(i)) & vbTab & StockType(LocoType(i)) & vbTab & FCoupling(LocoFCoup(i)) & vbTab & Rigid(LocoRigid(i)) & vbTab & Rigid(LocoFRigid(i))
End If
If LocoPath(i) <> vbNullString Then
    frmStock.GridUnused.AddItem LocoPath(i) & vbTab & Locomotives(i) & vbTab & LocoName(i)
    End If
Next i

For i = 0 To lngWagons - 1
If WagUsed(i) = False Then
frmStock.GridStock.AddItem vbTab & vbTab & Wagpath(i) & vbTab & Wagons(i) & vbTab _
& WagonName(i) & vbTab & Coupling(WagCoup(i)) & vbTab & Brake(WagBrake(i)) & vbTab & StockType(WagType(i)) & vbTab & FCoupling(WagFCoup(i)) & vbTab & Rigid(WagRigid(i)) & vbTab & Rigid(WagFRigid(i))
End If
If Wagpath(i) <> vbNullString Then
    frmStock.GridUnused.AddItem Wagpath(i) & vbTab & Wagons(i) & vbTab & WagonName(i)
    End If
Next i



Rem *******************************
frmStock.GridStock.col = 0
frmStock.GridStock.Sort = flexSortStringAscending
frmStock.Grid3.col = 0
frmStock.Grid3.Sort = flexSortStringAscending
DoEvents
Text1(cursouind).Text = "*.*"
MousePointer = 0

frmReport.Rich1.Text = strReport
frmReport.Show 1

     DoEvents
     

Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'GetStock2' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub

Private Sub GetStock()
Dim TrainPath As String, i As Integer
Dim LocomotivePath As String
Dim ConsistPath As String, fullpath$, strName As String, strTempRoute As String
Dim intEng As Integer, intWag As Integer, intBrake As Integer, intType As Integer, flagFCoup As Integer
Dim LocoUsed() As Boolean, WagUsed() As Boolean, intCouplings As Integer, flagRigid As String, strSMS As String
On Error GoTo Errtrap

MousePointer = 11

cursouind = 0
TrainPath = MSTSPath & "\trains\trainset"

strForPrint2 = vbNullString
For i = 0 To lngLoco - 1
SB2.Panels(2).Text = Locomotives(i)
LocomotivePath = TrainPath & "\" & LocoPath(i) & "\" & Locomotives(i)

Call GetCoupling(LocomotivePath, flagCouple, intBrake, intType, strName, intCouplings, flagRigid, flagFCoup, strSMS)

LocoName(i) = strName
LocoCoup(i) = flagCouple
LocoFCoup(i) = flagFCoup
LocoBrake(i) = intBrake
LocoType(i) = intType
If Len(flagRigid) = 1 Then
LocoRigid(i) = Val(flagRigid)
LocoFRigid(i) = 0
ElseIf Len(flagRigid) = 2 Then
LocoRigid(i) = Val(Left$(flagRigid, 1))
LocoFRigid(i) = Val(Right$(flagRigid, 1))
End If
DoEvents
strName = vbNullString

Next i


For i = 0 To lngWagons - 1
SB2.Panels(2).Text = Wagons(i)
LocomotivePath = TrainPath & "\" & Wagpath(i) & "\" & Wagons(i)
Call GetCoupling(LocomotivePath, flagCouple, intBrake, intType, strName, intCouplings, flagRigid, flagFCoup, strSMS)
WagonName(i) = strName
WagCoup(i) = flagCouple
WagFCoup(i) = flagFCoup
WagBrake(i) = intBrake
WagType(i) = intType
If Len(flagRigid) = 1 Then
WagRigid(i) = Val(flagRigid)
WagFRigid(i) = 0
ElseIf Len(flagRigid) = 2 Then
WagRigid(i) = Val(Left$(flagRigid, 1))
WagFRigid(i) = Val(Right$(flagRigid, 1))
End If
DoEvents
strName = vbNullString


Next i
ReDim LocoUsed(0 To lngLoco - 1)
ReDim WagUsed(0 To lngWagons - 1)



  
ConsistPath = MSTSPath & "\Trains\Consists"
cursouind = 0
Drive1(cursouind).Drive = Left$(ConsistPath, 2)
Dir1(cursouind).Path = ConsistPath
Text1(cursouind).Text = "*.con"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
lngCon = File1(cursouind).ListCount
ReDim Consists(0 To lngCon - 1)
ReDim ConIntName(0 To lngCon - 1)
ReDim ConIntWagName(0 To lngCon - 1)
ReDim ConEng(0 To lngCon - 1, 0 To 80)
ReDim conWag(0 To lngCon - 1, 0 To 500)


For i = 0 To File1(cursouind).ListCount - 1
intEng = 0
intWag = 0
   If File1(cursouind).Selected(i) Then
      fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
SB2.Panels(2).Text = File1(cursouind).List(i)

      Call CheckForConsist(fullpath$, i, intEng, intWag)

      If intEng > 0 Then
      For j = 1 To intEng
      LocoUsed(ConEng(i, j)) = True
      frmStock.GridStock.AddItem File1(cursouind).List(i) & vbTab & ConIntName(i) & vbTab & LocoPath(ConEng(i, j)) & vbTab & Locomotives(ConEng(i, j)) & vbTab _
& LocoName(ConEng(i, j)) & vbTab & Coupling(LocoCoup(ConEng(i, j))) & vbTab & Brake(LocoBrake(ConEng(i, j))) & vbTab & StockType(LocoType(ConEng(i, j))) & vbTab & FCoupling(LocoFCoup(ConEng(i, j))) & vbTab & Rigid(LocoRigid(ConEng(i, j))) & vbTab & Rigid(LocoFRigid(ConEng(i, j)))

Next j
End If

If intWag > 0 Then
      For j = 1 To intWag
      
      Rem *******************************************
            If conWag(i, j) = 9999 Then

Else
      WagUsed(conWag(i, j)) = True
      frmStock.GridStock.AddItem File1(cursouind).List(i) & vbTab & ConIntName(i) & vbTab & Wagpath(conWag(i, j)) & vbTab & Wagons(conWag(i, j)) & vbTab _
& WagonName(conWag(i, j)) & vbTab & Coupling(WagCoup(conWag(i, j))) & vbTab & Brake(WagBrake(conWag(i, j))) & vbTab & StockType(WagType(conWag(i, j))) & vbTab & FCoupling(WagFCoup(conWag(i, j))) & vbTab & Rigid(WagRigid(conWag(i, j))) & vbTab & Rigid(WagFRigid(conWag(i, j)))
End If
Next j
End If

     
   File1(cursouind).Selected(i) = False
   End If
Next i

Rem ******************* Activities
For i = 0 To lngAct - 1
SB2.Panels(2).Text = Activities(i)
intEng = 0: intWag = 0
x = InStr(ActPath(i), "Activities")
tempRoutePath = Left$(ActPath(i), x - 1)
x = InStrRev(tempRoutePath, "\", Len(tempRoutePath) - 2)
strTempRoute = Mid(tempRoutePath, x + 1)
strTempRoute = Left(strTempRoute, Len(strTempRoute) - 1)

Call GetLooseConsists(tempRoutePath & "Activities\" & Activities(i), i, intEng, intWag)


      If intEng > 0 Then
      For j = 1 To intEng
      LocoUsed(ActEng(i, j)) = True
      
      frmStock.GridStock.AddItem Activities(i) & vbTab & vbTab & LocoPath(ActEng(i, j)) & vbTab & Locomotives(ActEng(i, j)) & vbTab _
& LocoName(ActEng(i, j)) & vbTab & Coupling(LocoCoup(ActEng(i, j))) & vbTab & Brake(LocoBrake(ActEng(i, j))) & vbTab & StockType(LocoType(ActEng(i, j))) & vbTab & FCoupling(LocoFCoup(ActEng(i, j))) & vbTab & Rigid(LocoRigid(ActEng(i, j))) & vbTab & Rigid(LocoFRigid(ActEng(i, j))) & vbTab & "" & vbTab & strTempRoute
Next j
End If
If intWag > 0 Then
      For j = 1 To intWag
      WagUsed(ActWag(i, j)) = True
      
      frmStock.GridStock.AddItem Activities(i) & vbTab & vbTab & Wagpath(ActWag(i, j)) & vbTab & Wagons(ActWag(i, j)) & vbTab _
& WagonName(ActWag(i, j)) & vbTab & Coupling(WagCoup(ActWag(i, j))) & vbTab & Brake(WagBrake(ActWag(i, j))) & vbTab & StockType(WagType(ActWag(i, j))) & vbTab & FCoupling(WagFCoup(ActWag(i, j))) & vbTab & Rigid(WagRigid(ActWag(i, j))) & vbTab & Rigid(WagFRigid(ActWag(i, j))) & vbTab & "" & vbTab & strTempRoute
Next j
End If




Next i


For i = 0 To lngLoco - 1
If LocoUsed(i) = False Then
frmStock.GridStock.AddItem vbTab & vbTab & LocoPath(i) & vbTab & Locomotives(i) & vbTab _
& LocoName(i) & vbTab & Coupling(LocoCoup(i)) & vbTab & Brake(LocoBrake(i)) & vbTab & StockType(LocoType(i)) & vbTab & FCoupling(LocoFCoup(i)) & vbTab & Rigid(LocoRigid(i))
    If LocoPath(i) <> vbNullString Then
    frmStock.GridUnused.AddItem LocoPath(i) & vbTab & Locomotives(i) & vbTab & LocoName(i)
    End If
End If
Next i

For i = 0 To lngWagons - 1
If WagUsed(i) = False Then
frmStock.GridStock.AddItem vbTab & vbTab & Wagpath(i) & vbTab & Wagons(i) & vbTab _
& WagonName(i) & vbTab & Coupling(WagCoup(i)) & vbTab & Brake(WagBrake(i)) & vbTab & StockType(WagType(i)) & vbTab & FCoupling(WagFCoup(i)) & vbTab & Rigid(WagRigid(i))
    If Wagpath(i) <> vbNullString Then
    frmStock.GridUnused.AddItem Wagpath(i) & vbTab & Wagons(i) & vbTab & WagonName(i)
    End If
End If
Next i



Rem *******************************
frmStock.GridStock.col = 0
frmStock.GridStock.Sort = flexSortStringAscending
frmStock.Grid3.col = 0
frmStock.Grid3.Sort = flexSortStringAscending
DoEvents
Text1(cursouind).Text = "*.*"
MousePointer = 0

frmReport.Rich1.Text = strReport
frmReport.Show 1

     DoEvents
     

Exit Sub
Errtrap:

Call MsgBox("An error " & Err & " occurred in subroutine 'GetStock' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub


Private Sub GetStock3()
Dim TrainPath As String, FirstPath As String, DirCount As Integer, i As Long
Dim LocomotivePath As String, intBrake As Integer, intType As Integer, intCouplings As Integer
Dim strName As String, flagRigid As String, flagFCoup As Integer, strSMS As String
Dim LN As Integer

On Error GoTo Errtrap

  MousePointer = 11
  TrainPath = Trainspath & "\trainset"
If booActsChecked = True Then GoTo CarryON
LN = 1
If lngLoco > 0 Then
For i = 0 To lngLoco - 1
Locomotives(i) = vbNullString
LocoPath(i) = vbNullString
Next i
End If
LN = 2
For i = 0 To 4
Label7(i).Caption = vbNullString
Next
LN = 3
Label3.Caption = vbNullString


cursouind = 0


If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
        Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If

LN = 4
    FirstPath = Dir1(cursouind).Path
    DirCount = Dir1(cursouind).ListCount
    result = DirDiverLoco2(FirstPath, DirCount, "")
    File1(cursouind).Path = Dir1(cursouind).Path


DoEvents
Label1:
LN = 5
Label7(0).Caption = Str(lngLoco)
Label7(1).Caption = Str(lngWagons)
strForPrint = vbNullString
strForPrint2 = vbNullString
           ReDim Preserve Locomotives(0 To lngLoco - 1)
           ReDim Preserve LocoPath(0 To lngLoco - 1)
           ReDim Preserve LocoName(0 To lngLoco - 1)
           ReDim Preserve LocoCoup(0 To lngLoco - 1)
           ReDim Preserve LocoFCoup(0 To lngLoco - 1)
           ReDim Preserve LocoBrake(0 To lngLoco - 1)
           ReDim Preserve LocoType(0 To lngLoco - 1)
           ReDim Preserve LocoRigid(0 To lngLoco - 1)
           ReDim Preserve LocoFRigid(0 To lngLoco - 1)
           ReDim Preserve LocoSMS(0 To lngLoco - 1)
CarryON:

ReDim Preserve LocoSMS(0 To lngLoco - 1)
LN = 6

For i = 0 To lngLoco - 1

LocomotivePath = TrainPath & "\" & LocoPath(i) & "\" & Locomotives(i)
If Not FileExists(LocomotivePath) Then
strReport = strReport & "File not found - " & LocomotivePath & vbCrLf
GoTo AnotherLoco
End If
LN = 7
If Locomotives(i) = vbNullString Or Right$(Locomotives(i), 3) <> "eng" Then GoTo AnotherLoco
DoEvents
LN = 8
Call CheckLoco(LocomotivePath)

LN = 9
Call GetCoupling(LocomotivePath, flagCouple, intBrake, intType, strName, intCouplings, flagRigid, flagFCoup, strSMS)
LN = 10
LocoName(i) = strName
LocoCoup(i) = flagCouple
LocoFCoup(i) = flagFCoup
LocoBrake(i) = intBrake
LocoType(i) = intType
LN = 11
If Len(flagRigid) = 1 Then
LocoRigid(i) = Val(flagRigid)
LocoFRigid(i) = 0
ElseIf Len(flagRigid) = 2 Then
LocoRigid(i) = Val(Left$(flagRigid, 1))
LocoFRigid(i) = Val(Right$(flagRigid, 1))
End If
LN = 12
LocoSMS(i) = strSMS
strName = vbNullString

frmStock.GridStock.AddItem vbTab & vbTab & LocoPath(i) & vbTab & Locomotives(i) & vbTab _
& LocoName(i) & vbTab & Coupling(LocoCoup(i)) & vbTab & Brake(LocoBrake(i)) & vbTab & StockType(LocoType(i)) & vbTab & FCoupling(LocoFCoup(i)) & vbTab & Rigid(LocoRigid(i)) & vbTab & Rigid(LocoFRigid(i)) & vbTab & LocoSMS(i)
AnotherLoco:
Next i
DoEvents
LN = 13
If booActsChecked = True Then GoTo CarryOn2
            ReDim Preserve Wagons(0 To lngWagons - 1)
           ReDim Preserve Wagpath(0 To lngWagons - 1)
           ReDim Preserve WagonName(0 To lngWagons - 1)
           ReDim Preserve WagCoup(0 To lngWagons - 1)
           ReDim Preserve WagFCoup(0 To lngWagons - 1)
           ReDim Preserve WagBrake(0 To lngWagons - 1)
           ReDim Preserve WagType(0 To lngWagons - 1)
           ReDim Preserve WagRigid(0 To lngWagons - 1)
           ReDim Preserve WagFRigid(0 To lngWagons - 1)
CarryOn2:
LN = 14
For i = 0 To lngWagons - 1

LocomotivePath = TrainPath & "\" & Wagpath(i) & "\" & Wagons(i)
If Not FileExists(LocomotivePath) Then
strReport = strReport & "File not found - " & LocomotivePath & vbCrLf
GoTo AnotherWagon
End If
If Wagons(i) = vbNullString Or Right$(Wagons(i), 3) <> "wag" Then GoTo AnotherWagon
DoEvents
'Call CheckWagon(LocomotivePath)
LN = 15
Call GetCoupling(LocomotivePath, flagCouple, intBrake, intType, strName, intCouplings, flagRigid, flagFCoup, strSMS)
WagonName(i) = strName
WagCoup(i) = flagCouple
WagFCoup(i) = flagFCoup
WagBrake(i) = intBrake
WagType(i) = intType
If Len(flagRigid) = 1 Then
WagRigid(i) = Val(flagRigid)
WagFRigid(i) = 0
ElseIf Len(flagRigid) = 2 Then
WagRigid(i) = Val(Left$(flagRigid, 1))
WagFRigid(i) = Val(Right$(flagRigid, 1))
End If
LN = 16
strName = vbNullString

frmStock.GridStock.AddItem vbTab & vbTab & Wagpath(i) & vbTab & Wagons(i) & vbTab _
& WagonName(i) & vbTab & Coupling(WagCoup(i)) & vbTab & Brake(WagBrake(i)) & vbTab & StockType(WagType(i)) & vbTab & FCoupling(WagFCoup(i)) & vbTab & Rigid(WagRigid(i)) & vbTab & Rigid(WagFRigid(i))
AnotherWagon:
Next i
DoEvents
LN = 17
frmStock.GridStock.col = 2
frmStock.GridStock.Sort = flexSortStringAscending
frmStock.Grid3.Visible = False
frmStock.Command5.Visible = False

DoEvents
Text1(cursouind).Text = "*.*"
MousePointer = 0
If strReport <> vbNullString Then
frmReport.Rich1.Text = strReport
frmReport.Show 1
End If
     DoEvents
     
frmStock.Show

     DoEvents
     
Exit Sub
Errtrap:

Call MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'GetSTock3' " _
            & vbCrLf & "LocomotivePath = " & LocomotivePath & vbCrLf & " Line=" & LN _
            , vbExclamation, App.Title)
            
Resume Next

End Sub
Private Sub CheckLoco(LocoPath As String)
Dim strCorrEng As String, flagway As Integer, x As Long
Dim NewFile As Integer, A$, ShapeOK As Boolean, Y As Long
Dim strCorrSD As String, strCorrS As String, strEngpath As String
Dim xx As Long, strAnim As String, j As Long
Dim Z As Long, NewFile2 As Integer, intAnim As Integer

On Error GoTo Errtrap

x = InStrRev(LocoPath, "\")
strCorrEng = Mid$(LocoPath, x + 1)

strEngpath = Left$(LocoPath, x)
Rem ************** Check .eng file
NewFile2 = FreeFile
  Open LocoPath For Input As #NewFile2
  Do While Not EOF(NewFile2)
  Line Input #NewFile2, A$
  
  Y = InStr(A$, "WagonShape")
  If Y > 0 Then
  x = InStr(A$, "(")
  xx = InStr(x, A$, ")")
  strCorrS = Mid$(A$, x + 1, xx - x - 1)
  
  End If
    Z = InStr(A$, "Freightanim")
   If Z > 0 Then
  
  x = InStr(A$, "common.crew")
   If x = 0 Then GoTo NextOne
   x = InStr(A$, "..")
   If x = 0 Then GoTo NextOne
  
  xx = InStr(x + 1, A$, ".s")
  strAnim = Mid$(A$, x, (xx + 2) - x)
  For j = 0 To intAnim
  If strAnim = strAnimUsed(j) Then
  GoTo NextOne
  End If
  Next j
  intAnim = intAnim + 1
  strAnimUsed(intAnim) = strAnim
  Call CheckAnim(strAnim, LocoPath)
  End If
NextOne:
  Loop
  strCorrS = Trim$(strCorrS)
  If Left$(strCorrS, 1) = ChrW$(34) Then
  strCorrS = Mid$(strCorrS, 2)
  End If
  If Right$(strCorrS, 1) = ChrW$(34) Then
  strCorrS = Left$(strCorrS, Len(strCorrS) - 1)
  End If
  strCorrS = Trim$(strCorrS)
  strCorrSD = strCorrS & "d"
  Close #NewFile2
  Rem ******************** Check for FreightAnim
 
If Not FileExists(strEngpath & strCorrS) Then

strReport = strReport & strCorrS & Lang(553) & LocoPath & Lang(554) & vbCrLf
End If
If Not FileExists(strEngpath & strCorrSD) Then

strReport = strReport & strCorrSD & Lang(553) & LocoPath & Lang(554) & vbCrLf
Exit Sub
End If

Rem ************** Check .sd file
If FileExists(strEngpath & strCorrSD) Then
  NewFile = FreeFile
  Open strEngpath & strCorrSD For Input As #NewFile
  Do While Not EOF(NewFile)
  Line Input #NewFile, A$
  Y = InStr(A$, strCorrS)
  If Y > 0 Then
  ShapeOK = True
  Exit Do
  End If
  Loop
Close NewFile
End If
If ShapeOK = False Then
      flagway = 0
      Call ConvertSD(strEngpath & strCorrSD, flagway, strCorrS)
      flagway = 1
      Call ConvertSD(strEngpath & strCorrSD, flagway, strCorrS)
      strReport = strReport & strCorrSD & Lang(555) & vbCrLf
      End If
Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'CheckLoco' " _
            & vbCrLf & "LocomotivePath = " & LocoPath _
                        , vbExclamation, App.Title)
Resume Next

End Sub

Private Function ConvertSD(CompleteFilePath As String, flagway As Integer, strCorrShape As String) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertSD = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean
Dim strStart As String, strEnd As String, x As Integer
Dim xx As Integer

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(401), vbInformation, frmUtils.Caption
  Exit Function
End If
'SET THE INFO LABEL
'lblInfo.Caption = "Converting " & chrw$(34) & _
        mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & chrw$(34)

'DETERMINE OPTION SELECTION
'TEXT2UNICODE
If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

'GET FIRST CHAR FROM THE FILE AND CHECK TO SEE IF IT IS ANSI 255 (ÿ)
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False

'MAKE SURE THE FILE IS UNICODE OR TEXT
If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, frmUtils.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, frmUtils.Caption
    Exit Function
  End If
End If

'OPEN THE FILE AND READ IT INTO MEMORY
Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll

If flagway = 0 Then

xx = InStr(MyString, "Shape")
If xx = 0 Then
strReport = strReport & strCorrShape & ".d does not have a valid 'Shape' entry and must be corrected" & vbCrLf
Exit Function
End If

x = InStr(xx, MyString, vbCr)

strStart = Left$(MyString, xx - 1)
strEnd = Mid$(MyString, x)
MyString = strStart & "Shape ( " & strCorrShape & strEnd


End If
The_obj.Close
fileflag = False

'SET THE TEMPFILE NAME AND PATH
tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"

'CREATE TEMP FILE IN TEMP FOLDER AND OVERWRITE A FILE WITH THE SAME NAME
Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True

'HANDLE ERROR FROM WRITE IF UNICODE FILE HAS BEEN ALTERED
On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, frmUtils.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False


FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertSD = True

ERRHANDLER:

  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, frmUtils.Caption
  
End Function






Private Sub IsItElectric(RouteName As String, booElectric As Boolean)
Dim NewFile As Integer, strNew As String, strTemp As String, x As Integer
Dim Y As Integer, yy As Integer, i As Integer, booExists As Boolean

NewFile = FreeFile
If Not FileExists(RouteName) Then
booExists = False
Exit Sub
End If
Open RouteName For Input As #NewFile
Do While Not EOF(NewFile)
   
   Line Input #NewFile, strNew
  
   strNew = Trim$(strNew)
   x = InStr(strNew, "Electrified (")
  
   If x > 0 Then

   Y = InStr(strNew, "(")
   yy = InStr(strNew, ")")
   strNew = Mid$(strNew, Y + 1, yy - (Y + 1))
   strNew = Trim$(strNew)
   If Left$(strNew, 1) = ChrW$(34) Then
    strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   If Val(strNew) = 1 Then
   booElectric = True
   Else
   booElectric = False
   End If
     End If
     x = InStr(strNew, ".env")
   If x > 0 Then
   i = i + 1
   Y = InStrRev(strNew, "(", x)
   strTemp = Mid$(strNew, Y + 1, (x + 3) - Y)
   strTemp = Trim$(strTemp)
   strEnv(i) = strTemp
   End If
   Loop
   strEnv(13) = "Editor.env"
   
   Close #NewFile
   
End Sub

Private Sub KillUnused()
Dim i As Integer, ii As Integer, AceFound As Boolean
Dim ShapePath As String, TempShape As String, strTempS As String
Dim TertexPath As String, Jumper As Boolean, flagway As Integer, intESD As Integer
Dim strAce As String, strMove As String

On Error GoTo Errtrap
If booKillMove = False Then
strMove = Lang(589)
Else
strMove = "Please confirm you wish to MOVE those files"
End If
Select Case MsgBox(strMove & vbCrLf & Lang(590), vbYesNo + vbExclamation + vbDefaultButton1, App.Title)

    Case vbYes
MousePointer = 11
cursouind = 1
ShapePath = RoutePath & "\shapes"
Drive1(cursouind).Drive = Left$(ShapePath, 2)
Dir1(cursouind).Path = ShapePath
Text1(cursouind).Text = "*.S"

DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numShp - 1

   If File1(cursouind).List(i) = strShp(ii) Then
     AceFound = True
   Exit For
   End If
   If Left(File1(cursouind).List(i), 8) = "DynaTrax" Then
     AceFound = True
   Exit For
   End If
   Next ii

 If AceFound = False Then

If booKillMove = False Then
  Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
  Else
  FileCopy File1(cursouind).Path & "\" & File1(cursouind).List(i), strKillPath & "\Shapes\" & File1(cursouind).List(i)
  DoEvents
  Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
  DoEvents
  End If
  
  strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
  TempShape = Left$(File1(cursouind).List(i), Len(File1(cursouind).List(i)) - 1) & "thm"
  If FileExists(ShapePath & "\" & TempShape) Then
  If booKillMove = False Then
 Kill ShapePath & "\" & TempShape
 Else
 FileCopy ShapePath & "\" & TempShape, strKillPath & "\Shapes\" & TempShape
 DoEvents
 Kill ShapePath & "\" & TempShape
 DoEvents
 End If
 
 strKillFiles = strKillFiles & ShapePath & "\" & TempShape & vbCrLf
 End If
  TempShape = File1(cursouind).List(i) & "d"
  If FileExists(ShapePath & "\" & TempShape) Then
  If booKillMove = False Then
 Kill ShapePath & "\" & TempShape
 Else
 FileCopy ShapePath & "\" & TempShape, strKillPath & "\Shapes\" & TempShape
 DoEvents
 Kill ShapePath & "\" & TempShape
 DoEvents
 End If
 strKillFiles = strKillFiles & ShapePath & "\" & TempShape & vbCrLf
 End If
  
 End If
 End If
 AceFound = False
 Next i
Rem ************ Kill Textures

Drive1(cursouind).Drive = Left$(TexturePath, 2)
Dir1(cursouind).Path = TexturePath
Text1(cursouind).Text = "*.ace"


DoEvents

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Left$(File1(cursouind).List(i), 11) = "acleantrack" Then
AceFound = True
GoTo GetAnother
End If
If Left$(File1(cursouind).List(i), 10) = "teleshadow" Then
AceFound = True
GoTo GetAnother
End If

If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
If File1(cursouind).List(i) = Left$(Ace2(ii), Len(Ace2(ii)) - 4) Then
   
   AceFound = True
   Exit For
   End If
   End If
 Next ii

 If AceFound = False Then

If booKillMove = False Then
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Else
 FileCopy File1(cursouind).Path & "\" & File1(cursouind).List(i), strKillPath & "\Textures\" & File1(cursouind).List(i)
 DoEvents
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 DoEvents
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 Rem ***************** Snow
 
 If FileExists(TexSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Snow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSnowPath & "\" & File1(cursouind).List(i)
 DoEvents
 End If
 strKillFiles = strKillFiles & TexSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexNightPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexNightPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexNightPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Night\" & File1(cursouind).List(i)
 DoEvents
 Kill TexNightPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexNightPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexWinPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexWinPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexWinPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Winter\" & File1(cursouind).List(i)
 DoEvents
 Kill TexWinPath & "\" & File1(cursouind).List(i)
 End If
  strKillFiles = strKillFiles & TexWinPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexWinSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexWinSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\WinterSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexWinSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexAutPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Autumn\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexAutSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\AutumnSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Spring\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\springSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 
 
 End If
 End If
GetAnother:
 AceFound = False
 Next i
 
Rem ************ Kill Snow Textures

Drive1(cursouind).Drive = Left$(TexSnowPath, 2)
Dir1(cursouind).Path = TexSnowPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
If Left$(File1(cursouind).List(i), 11) = "acleantrack" Then
AceFound = True
GoTo GetAnother2
End If
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
If File1(cursouind).List(i) = Left$(Ace2(ii), Len(Ace2(ii)) - 4) Then
   
   
   AceFound = True
   Exit For
   End If
   End If
   Next ii

 If AceFound = False Then
 
 If booKillMove = False Then
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Else
 FileCopy File1(cursouind).Path & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Snow\" & File1(cursouind).List(i)
 DoEvents
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 If FileExists(TexSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Snow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSnowPath & "\" & File1(cursouind).List(i)
 DoEvents
 End If
 strKillFiles = strKillFiles & TexSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexNightPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexNightPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexNightPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Night\" & File1(cursouind).List(i)
 DoEvents
 Kill TexNightPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexNightPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexWinPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexWinPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexWinPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Winter\" & File1(cursouind).List(i)
 DoEvents
 Kill TexWinPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexWinPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexWinSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexWinSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\wintersnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexWinSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexAutPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Autumn\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexAutSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\AutumnSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Spring\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\SpringSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 
 
 End If
 End If
GetAnother2:
 AceFound = False
 Next i
 Rem ************ Check Snow Textures for Unnecessary *********************  Errors from here......
 
 Drive1(cursouind).Drive = Left$(TexSnowPath, 2)
Dir1(cursouind).Path = TexSnowPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
If Left$(File1(cursouind).List(i), 11) = "acleantrack" Then
AceFound = True
GoTo GetAnother4
End If

For ii = 0 To numAce - 1

If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
strAce = Left$(Ace2(ii), Len(Ace2(ii)) - 4)
intESD = Val(Right$(Ace2(ii), 3))
End If

   If File1(cursouind).List(i) = strAce And (intESD = 1 Or intESD = 259 Or intESD = 257 Or intESD = 2) Then

   AceFound = True
   Exit For

   
   End If
 
GetAnother3:
   Next ii
   
 If AceFound = False Then

 If FileExists(TexSnowPath & "\" & File1(cursouind).List(i)) Then
 
If booKillMove = False Then
 Kill TexSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Snow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSnowPath & "\" & File1(cursouind).List(i)
 End If
  strKillFiles = strKillFiles & TexSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 End If
 AceFound = False
GetAnother4:
 Next i
 Rem ************ Kill Night Textures

Drive1(cursouind).Drive = Left$(TexNightPath, 2)
Dir1(cursouind).Path = TexNightPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
   If File1(cursouind).List(i) = Left$(Ace2(ii), Len(Ace2(ii)) - 4) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then

 If booKillMove = False Then
 Kill TexNightPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexNightPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Night\" & File1(cursouind).List(i)
 DoEvents
 Kill TexNightPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 If FileExists(TexWinPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexWinPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexWinPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\winter\" & File1(cursouind).List(i)
 DoEvents
 Kill TexWinPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexWinPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexWinSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexWinSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\winterSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexWinPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexAutPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Autumn\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexAutSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\AutumnSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\spring\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\SpringSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 
 
 End If
 End If
 AceFound = False
 Next i
  Rem ************ Check night Textures for Unnecessary *********************
 Drive1(cursouind).Drive = Left$(TexNightPath, 2)
Dir1(cursouind).Path = TexNightPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
strAce = Left$(Ace2(ii), Len(Ace2(ii)) - 4)
intESD = Val(Right$(Ace2(ii), 3))
   If File1(cursouind).List(i) = strAce And (intESD = 256 Or intESD = 257 Or intESD = 259) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then

 If FileExists(TexNightPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexNightPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexNightPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Night\" & File1(cursouind).List(i)
 DoEvents
 Kill TexNightPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexNightPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 End If
 Next i
 
 Rem ************ Kill Winter Textures

Drive1(cursouind).Drive = Left$(TexWinPath, 2)
Dir1(cursouind).Path = TexWinPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
   If File1(cursouind).List(i) = Left$(Ace2(ii), Len(Ace2(ii)) - 4) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then

 If booKillMove = False Then
 Kill TexWinPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexWinPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\winter\" & File1(cursouind).List(i)
 DoEvents
 Kill TexWinPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 ' strkillfiles=strkillfiles & x & vbcrlf
 If FileExists(TexWinSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexWinSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\winterSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexWinSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexAutPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Autumn\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexAutSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\AutumnSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Spring\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\SpringSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 
 
 End If
 End If
 AceFound = False
 Next i
   Rem ************ Check Winter Textures for Unnecessary *********************
 Drive1(cursouind).Drive = Left$(TexWinPath, 2)
Dir1(cursouind).Path = TexWinPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
strAce = Left$(Ace2(ii), Len(Ace2(ii)) - 4)
intESD = Val(Right$(Ace2(ii), 3))
   If File1(cursouind).List(i) = strAce And (intESD = 252 Or intESD = 259) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then

 If FileExists(TexWinPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexWinPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexWinPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\winter\" & File1(cursouind).List(i)
 DoEvents
 Kill TexWinPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexWinPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 End If
 Next i
  Rem ************ Kill Winter Snow Textures

Drive1(cursouind).Drive = Left$(TexWinSnowPath, 2)
Dir1(cursouind).Path = TexWinSnowPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
   If File1(cursouind).List(i) = Left$(Ace2(ii), Len(Ace2(ii)) - 4) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then

 If booKillMove = False Then
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexWinSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\winterSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 If FileExists(TexAutPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Autumn\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexAutSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\AutumnSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Spring\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\SpringSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 
 
 End If
 End If
 AceFound = False
 Next i
    Rem ************ Check Winter Snow Textures for Unnecessary *********************
 Drive1(cursouind).Drive = Left$(TexWinSnowPath, 2)
Dir1(cursouind).Path = TexWinSnowPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
strAce = Left$(Ace2(ii), Len(Ace2(ii)) - 4)
intESD = Val(Right$(Ace2(ii), 3))
   If File1(cursouind).List(i) = strAce And (intESD = 252 Or intESD = 259) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
  
   
 If AceFound = False Then
 
 If FileExists(TexWinSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexWinSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\wintersnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexWinSnowPath & "\" & File1(cursouind).List(i)
 End If
  strKillFiles = strKillFiles & TexWinSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 End If
 Next i
   Rem ************ Kill Autumn Textures

Drive1(cursouind).Drive = Left$(TexAutPath, 2)
Dir1(cursouind).Path = TexAutPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
   If File1(cursouind).List(i) = Left$(Ace2(ii), Len(Ace2(ii)) - 4) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then

 If booKillMove = False Then
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Autumn\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 If FileExists(TexAutSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\AutumnSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\spring\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\SpringSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 
 
 End If
 End If
 AceFound = False
 Next i
    Rem ************ Check Autumn Textures for Unnecessary *********************
 Drive1(cursouind).Drive = Left$(TexAutPath, 2)
Dir1(cursouind).Path = TexAutPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
strAce = Left$(Ace2(ii), Len(Ace2(ii)) - 4)
intESD = Val(Right$(Ace2(ii), 3))
   If File1(cursouind).List(i) = strAce And (intESD = 252 Or intESD = 259) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then

 If FileExists(TexAutPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Autumn\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexAutPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 End If
 Next i
Rem ************ Kill Autumn Snow Textures

Drive1(cursouind).Drive = Left$(TexAutSnowPath, 2)
Dir1(cursouind).Path = TexAutSnowPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
   If File1(cursouind).List(i) = Left$(Ace2(ii), Len(Ace2(ii)) - 4) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then

 If booKillMove = False Then
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\AutumnSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 If FileExists(TexSprPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\Spring\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 If FileExists(TexSprSnowPath & "\" & File1(cursouind).List(i)) Then
If booKillMove = False Then
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\SpringSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 End If
 AceFound = False
 Next i
     Rem ************ Check Autumn snow Textures for Unnecessary *********************
 Drive1(cursouind).Drive = Left$(TexAutSnowPath, 2)
Dir1(cursouind).Path = TexAutSnowPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
strAce = Left$(Ace2(ii), Len(Ace2(ii)) - 4)
intESD = Val(Right$(Ace2(ii), 3))
   If File1(cursouind).List(i) = strAce And (intESD = 252 Or intESD = 259) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then
 
 If FileExists(TexAutSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexAutSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\AutumnSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexAutSnowPath & "\" & File1(cursouind).List(i)
 End If
  strKillFiles = strKillFiles & TexAutSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 End If
 Next i
 Rem ************ Kill Spring Textures

Drive1(cursouind).Drive = Left$(TexSprPath, 2)
Dir1(cursouind).Path = TexSprPath
Text1(cursouind).Text = "*.ace"
DoEvents

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
   If File1(cursouind).List(i) = Left$(Ace2(ii), Len(Ace2(ii)) - 4) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then

 If booKillMove = False Then
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\spring\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 If FileExists(TexSprSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\SpringSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 End If
 AceFound = False
 Next i
     Rem ************ Check Spring Textures for Unnecessary *********************
 Drive1(cursouind).Drive = Left$(TexSprPath, 2)
Dir1(cursouind).Path = TexSprPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
strAce = Left$(Ace2(ii), Len(Ace2(ii)) - 4)
intESD = Val(Right$(Ace2(ii), 3))
   If File1(cursouind).List(i) = strAce And (intESD = 252 Or intESD = 259) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then
 If FileExists(TexSprPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\spring\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprPath & "\" & File1(cursouind).List(i)
 End If
  strKillFiles = strKillFiles & TexSprPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 End If
 Next i
  Rem ************ Kill Spring snow Textures

Drive1(cursouind).Drive = Left$(TexSprSnowPath, 2)
Dir1(cursouind).Path = TexSprSnowPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
   If File1(cursouind).List(i) = Left$(Ace2(ii), Len(Ace2(ii)) - 4) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then

 If booKillMove = False Then
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\SpringSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 AceFound = False
 Next i
     Rem ************ Check Spring Snow Textures for Unnecessary *********************
 Drive1(cursouind).Drive = Left$(TexSprSnowPath, 2)
Dir1(cursouind).Path = TexSprSnowPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numAce - 1
If Ace2(ii) <> vbNullString And Ace2(ii) <> "|" Then
strAce = Left$(Ace2(ii), Len(Ace2(ii)) - 4)
intESD = Val(Right$(Ace2(ii), 3))
   If File1(cursouind).List(i) = strAce And (intESD = 252 Or intESD = 259) Then
   AceFound = True
   Exit For
   End If
   End If
   Next ii
 If AceFound = False Then
 If FileExists(TexSprSnowPath & "\" & File1(cursouind).List(i)) Then
 If booKillMove = False Then
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TexSprSnowPath & "\" & File1(cursouind).List(i), strKillPath & "\Textures\SpringSnow\" & File1(cursouind).List(i)
 DoEvents
 Kill TexSprSnowPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & TexSprSnowPath & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 End If
 Next i
 Rem ******** Kill Terrain Textures


 
TertexPath = RoutePath & "\Terrtex"
Drive1(cursouind).Drive = Left$(TertexPath, 2)
Dir1(cursouind).Path = TertexPath

Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numTerr - 1

   If File1(cursouind).List(i) = TerrTex2(ii) Then
   AceFound = True

   Exit For
   End If
   Next ii
   
 If AceFound = False Then

 If booKillMove = False Then
 Kill TertexPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy TertexPath & "\" & File1(cursouind).List(i), strKillPath & "\Terrtex\" & File1(cursouind).List(i)
 DoEvents
 Kill TertexPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 DoEvents
 If booKillMove = False Then
 Kill TertexPath & "\Snow\" & File1(cursouind).List(i)
 Else
 FileCopy TertexPath & "\Snow\" & File1(cursouind).List(i), strKillPath & "\Terrtex\Snow\" & File1(cursouind).List(i)
 DoEvents
 Kill TertexPath & "\Snow\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\Snow\" & File1(cursouind).List(i) & vbCrLf
 DoEvents
  End If
 End If
 AceFound = False
 Next i
 Rem *********** Terrtex Snow ******************************

 Drive1(cursouind).Drive = Left$(TertexPath, 2)

 Dir1(cursouind).Path = TertexPath & "\Snow"
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
For ii = 0 To numTerr - 1

   If File1(cursouind).List(i) = TerrTex2(ii) Then
   AceFound = True
      Exit For
   End If
   Next ii
 If AceFound = False Then
 If booKillMove = False Then
 Kill TertexPath & "\Snow\" & File1(cursouind).List(i)
 Else
 FileCopy TertexPath & "\Snow\" & File1(cursouind).List(i), strKillPath & "\Terrtex\Snow\" & File1(cursouind).List(i)
 DoEvents
 Kill TertexPath & "\Snow\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
  End If
 End If
 AceFound = False
 Next i



Rem ************* Kill sounds
' strkillfiles=strkillfiles & x & vbcrlf
Drive1(cursouind).Drive = Left$(SoundPath, 2)
Dir1(cursouind).Path = SoundPath
Text1(cursouind).Text = "*.sms"

DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
AceFound = False
If File1(cursouind).Selected(i) = True Then
For ii = 0 To SoundNumber

   If File1(cursouind).List(i) = Soundfile(ii) Then
   AceFound = True
   
   Exit For
   End If
   Next ii
 If AceFound = False Then

 If booKillMove = False Then
 Kill SoundPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy SoundPath & "\" & File1(cursouind).List(i), strKillPath & "\Sound\" & File1(cursouind).List(i)
 DoEvents
 Kill SoundPath & "\" & File1(cursouind).List(i)
 End If
  strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 AceFound = False
 Next i

 Text1(cursouind).Text = "*.wav"
 DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
For ii = 0 To WavNumber

   If File1(cursouind).List(i) = WavFile(ii) Then
   AceFound = True
   
   Exit For
   End If
   Next ii
 If AceFound = False Then

 If booKillMove = False Then
 Kill SoundPath & "\" & File1(cursouind).List(i)
 Else
 FileCopy SoundPath & "\" & File1(cursouind).List(i), strKillPath & "\Sound\" & File1(cursouind).List(i)
 DoEvents
 Kill SoundPath & "\" & File1(cursouind).List(i)
 End If
 strKillFiles = strKillFiles & File1(cursouind).Path & "\" & File1(cursouind).List(i) & vbCrLf
 End If
 End If
 AceFound = False
 Next i
 Text1(cursouind).Text = "*.*"

Rem ***************** Copy over the new .ref file
flagway = 1
Call ConvertIt(App.Path & "\setupfiles\tempref.ref", flagway)
FileCopy App.Path & "\setupfiles\tempref.ref", OriginalRef

Rem ***************** Kill .S/t/w files

Call KillSpare("*.t")
DoEvents
 Rem *************** Remove .thm files


Drive1(cursouind).Drive = Left$(ShapePath, 2)
Dir1(cursouind).Path = ShapePath
Text1(cursouind).Text = "*.thm"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
  strTempSD = File1(cursouind).Path & "\" & File1(cursouind).List(i)
  
   Kill strTempSD
   
       
End If
Next i

Rem *************** Remove Orphaned .sd files


Drive1(cursouind).Drive = Left$(ShapePath, 2)
Dir1(cursouind).Path = ShapePath
Text1(cursouind).Text = "*.sd"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
  strTempSD = File1(cursouind).Path & "\" & File1(cursouind).List(i)
  strTempS = Left$(strTempSD, Len(strTempSD) - 1)
  If Not FileExists(strTempS) Then
   
   Kill strTempSD
   End If
       
End If
Next i
 
 Rem ************** Kill the rubbish files
If FileExists(TexturePath & "\tunnel1_noturrets.s") Then
 Kill TexturePath & "\tunnel1_noturrets.s"
 End If
 If FileExists(TexNightPath & "\us2powerplant.psd") Then
 Kill TexNightPath & "\us2powerplant.psd"
 End If
 If FileExists(TexSnowPath & "\jp2indirt") Then
 Kill TexSnowPath & "\jp2indirt"
 End If
 If FileExists(ShapePath & "\light.ace") Then
 Kill ShapePath & "\light.ace"
 End If
 If FileExists(ShapePath & "\lightmat.pal") Then
 Kill ShapePath & "\lightmat.pal"
 End If
 
 Rem ***************Copy .ref back to route*************************
 'FileCopy App.path & "\setupfiles\tempref.ref", OriginalRef
 Rem ******************************************************
 Drive1(1).Drive = Left$(App.Path, 2)
Dir1(1).Path = App.Path & "\setupfiles"
Text1(1).Text = "*.*"


MousePointer = 0
    Case vbNo
    Dim q As Integer
    For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then frmUtils.Controls(q).Enabled = True
Next q
Exit Sub
End Select


Exit Sub
Errtrap:

If Err = 76 Then
Jumper = True
Resume Next

ElseIf Err = 75 Then
MsgBox Lang(448) & vbCr & Lang(449), 48, Lang(450)
'********************
Resume Next
ElseIf Err = 53 Then
Resume Next
Else
Call MsgBox("An error " & Err & " occurred in subroutine 'KillUnused' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End If
End Sub


Public Function strGetShortFileName(ByVal strLongFileName As String) As String
 
    On Error GoTo errorHandler
 
    Dim lngRetVal As Long
    Dim strShortFileName As String
    Dim lngLen As Long
     
    'Set up buffer area for API function call return
    strShortFileName = Space(255)
     
    lngLen = Len(strShortFileName)
     
    'Call the function
    lngRetVal = GetShortPathName(strLongFileName, strShortFileName, lngLen)
     
    'Strip away unwanted characters.
    strGetShortFileName = Left$(strShortFileName, lngRetVal)
 
    Exit Function 'avoid executing the error handler
    
errorHandler:
'    If lngHandleError(Err.Number, mstrcModule, Err.Description, False) = tdcResume Then
'        Resume
'    Else
'        Err.Raise Err.Number, mstrcModule, Err.Description
'    End If
End Function


Private Sub CheckDefaultSounds2()

SoundNumber = 0

Soundfile(SoundNumber) = "ingame.sms"
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = DefWat
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = DefCross
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = DefCoal
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = DefDies
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = DefSig
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = "clear_ex.sms"
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = "clear_in.sms"
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = "rain_ex.sms"
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = "rain_in.sms"
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = "intro.sms"
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = "gui.sms"
SoundNumber = SoundNumber + 1
End Sub

Private Sub CheckDefaultSounds()
Dim strNew As String, j As Integer

'SoundNumber = 0
SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = "ingame.sms"
If booWat = True Then
strNew = DefWat
For j = 0 To SoundNumber
   If strNew = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   SoundNumber = SoundNumber + 1
   Soundfile(SoundNumber) = strNew
   End If
   strNew = vbNullString
  itExists = False
  End If
 If booCross = True Then
strNew = DefCross
For j = 0 To SoundNumber
   If strNew = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   SoundNumber = SoundNumber + 1
   Soundfile(SoundNumber) = strNew
   End If
   
   strNew = vbNullString
  itExists = False
  End If
  If booCoal = True Then
strNew = DefCoal
For j = 0 To SoundNumber
   If strNew = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   SoundNumber = SoundNumber + 1
   Soundfile(SoundNumber) = strNew
   End If
   
   strNew = vbNullString
  itExists = False
  End If
If booDies = True Then
strNew = DefDies
For j = 0 To SoundNumber
   If strNew = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   SoundNumber = SoundNumber + 1
   Soundfile(SoundNumber) = strNew
   End If
   
   strNew = vbNullString
  itExists = False
  End If
If booSig = True Then
strNew = DefSig
For j = 0 To SoundNumber
   If strNew = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   SoundNumber = SoundNumber + 1
   Soundfile(SoundNumber) = strNew
   End If
   
   strNew = vbNullString
  itExists = False
  End If

End Sub


Private Sub CheckForConsist(CFilepath As String, intCon As Integer, k As Integer, k1 As Integer)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, j As Integer, Engpath As String, Engname As String
Dim Wagonpath As String, Wagname As String, strCfg As String, strTemp As String
Dim TrainsetPath As String, ConName As String, ConsistPath As String
Dim xx As Integer, Z As Long, strnew3 As String, booEntry As Boolean
Dim conShort As String, strFoundPath As String, itExists As Boolean, flagway As Integer
On Error GoTo Errtrap

Fnumber = FreeFile
x = InStrRev(CFilepath, "\")
ConName = Mid$(CFilepath, x + 1)


conShort = Left$(ConName, Len(ConName) - 4)
TrainsetPath = MSTSPath & "\Trains\Trainset\"
ConsistPath = MSTSPath & "\Trains\Consists\"
Rem ******* Check Consist is Unicode ***************
   Open CFilepath For Binary As #5
    strTemp = String(2, " ")
    Get #5, , strTemp
 Close #5
If Asc(Mid$(strTemp, 1, 1)) <> 255 Then
If Asc(Mid$(strTemp, 2, 1)) <> 254 Then
 
   flagway = 1
   Call ConvertIt(CFilepath, flagway)
   End If
End If
   DoEvents
strNew = ReadUniFile(CFilepath)

Rem ************** Check traincfg

x = InStr(strNew, "TrainCfg")
If x > 0 Then
j = InStr(x, strNew, "(")
xx = InStr(j, strNew, vbCr)
strCfg = Mid$(strNew, j + 1, xx - (j + 1))
End If
strCfg = Trim$(strCfg)
strCfg = Replace(strCfg, ChrW$(34), "")
If strCfg <> conShort Then

strReport = strReport & "Consist " & conShort & Lang(487) & strCfg & vbCrLf & vbCrLf

End If
Rem ***************************

x = 1
Do
  
    x = InStr(x, strNew, "EngineData")
If x = 0 Then Exit Do
Z = InStr(x, strNew, vbLf)
strnew3 = Mid$(strNew, x, Z - x)

Call CheckEngineData(strnew3, Engname, Engpath, booEntry)

    Engname = Engname & ".eng"
    For j = 0 To lngLoco - 1
   If Engname = Locomotives(j) Then
   k = k + 1
   ConEng(intCon, k) = j
   itExists = True
   Exit For
  End If
   Next j

   If itExists = False Then

   FlagColRed = True
   strReport = strReport & Lang(560) & Engpath & "\" & Engname & Lang(561) & ConName & vbCrLf
   frmStock.Grid3.AddItem ConName & vbTab & Engpath & "\" & Engname
      frmStock.GridStock.AddItem ConName & vbTab & "" & vbTab & Engpath & vbTab & Engname & vbTab _
& "" & vbTab & "" & vbTab & "" & vbTab & ""
FlagColRed = False
   Else
   Rem *****************
   If Not FileExists(TrainsetPath & Engpath & "\" & Engname) Then
strFoundPath = vbNullString
Call LookForLoco(Engname, strFoundPath)

Rem *********************** Put option to alter consist here ********************************
If strFoundPath <> vbNullString Then
Select Case MsgBox(Engname & "  was not found in folder " & Engpath & " but was found in " & strFoundPath & " Do you wish to alter the Consist file?", vbOKCancel Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbOK
    
strReport = strReport & vbCrLf & Lang(533) & ConName & ", " & Lang(526) & Engpath & "\" & Engname & vbCrLf & Lang(631) & strFoundPath & " Consist corrected " & vbCrLf
    Call ChangeConPath(ConName, Engname, Engpath, strFoundPath, strNew)
    Case vbCancel
strReport = strReport & vbCrLf & Lang(533) & ConName & ", " & Lang(526) & Engpath & "\" & Engname & vbCrLf & Lang(631) & strFoundPath & vbCrLf
End Select

 End If
 End If
 End If
 Rem *************************
   x = x + 1
   Loop

 x = InStr(strNew, "Name (")
   If x > 0 Then
   xx = InStr(x, strNew, ")")
   strTemp = Mid$(strNew, x + 6, xx - (x + 7))
   strTemp = Trim$(strTemp)
   If Left$(strTemp, 1) = ChrW$(34) Then
   strTemp = Mid$(strTemp, 2, Len(strTemp) - 2)
   End If
   ConIntName(intCon) = strTemp
   End If
   
  x = 1: k1 = 0
Do
  
    x = InStr(x, strNew, "WagonData")
If x = 0 Then Exit Do

    Z = InStr(x, strNew, vbLf)
strnew3 = Mid$(strNew, x, Z - x)

Call CheckWagonData(strnew3, Wagname, Wagonpath, booEntry)

   Wagname = Wagname & ".wag"
   itExists = False
     For j = 0 To lngWagons - 1
   If Wagname = Wagons(j) Then
   
   itExists = True
   k1 = k1 + 1
   conWag(intCon, k1) = j
   Exit For
    End If
  Next j

  If itExists = False Then

 FlagColRed = True
   strReport = strReport & Lang(560) & Wagonpath & "\" & Wagname & Lang(561) & ConName & vbCrLf
   frmStock.Grid3.AddItem ConName & vbTab & Wagonpath & "\" & Wagname
   frmStock.GridStock.AddItem ConName & vbTab & "" & vbTab & Wagonpath & vbTab & Wagname & vbTab _
& "" & vbTab & "" & vbTab & "" & vbTab & ""
FlagColRed = False
   Else
      Rem *****************
   If Not FileExists(TrainsetPath & Wagonpath & "\" & Wagname) Then
strFoundPath = vbNullString
Call LookForWag(Wagname, strFoundPath)

If strFoundPath <> vbNullString Then
Select Case MsgBox(Wagname & "  was not found in folder " & Wagonpath & " but was found in " & strFoundPath & " Do you wish to alter the Consist file?", vbOKCancel Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbOK
 
strReport = strReport & vbCrLf & Lang(533) & ConName & ", " & Lang(527) & Wagonpath & "\" & Wagname & vbCrLf & Lang(632) & strFoundPath & " Consist corrected " & vbCrLf
    Call ChangeConPath(ConName, Wagname, Wagonpath, strFoundPath, strNew)
    Case vbCancel
strReport = strReport & vbCrLf & Lang(533) & ConName & ", " & Lang(527) & Wagonpath & "\" & Wagname & vbCrLf & Lang(632) & strFoundPath & vbCrLf
End Select

 End If
 End If
 End If
 Rem *************************
   x = x + 1
   Loop

Exit Sub
Errtrap:

   
    Call MsgBox("An error " & Err & " occurred in subroutine 'CheckForConsist' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
            Resume Next
'********************
  

End Sub

Private Sub LookForWag(strWag As String, strFoundPath As String)
Dim TrainsetPath As String, i As Integer

TrainsetPath = MSTSPath & "\Trains\Trainset\"
For i = 0 To lngWagons - 1
If strWag = Wagons(i) Then

If FileExists(TrainsetPath & Wagpath(i) & "\" & strWag) Then
strFoundPath = Wagpath(i) & "\" & strWag
End If
End If
Next i
End Sub

Private Sub LookForLoco(strEng As String, strFoundPath As String)
Dim TrainsetPath As String

TrainsetPath = MSTSPath & "\Trains\Trainset\"
For i = 0 To lngLoco - 1
If strEng = Locomotives(i) Then

If FileExists(TrainsetPath & LocoPath(i) & "\" & strEng) Then
strFoundPath = LocoPath(i) & "\" & strEng
Exit For
End If
End If
Next i

End Sub
Private Sub CompressACE()

On Error GoTo Errtrap

MousePointer = 11
cursouind = 0
If Right$(Dir1(cursouind).Path, 1) <> "\" Then Dir1(cursouind).Path = Dir1(cursouind).Path & "\"
frmMain.Show

     DoEvents
     


  MousePointer = 0
Exit Sub
Errtrap:

If Err = 53 Then
Resume Next
End If

End Sub



Private Sub UncompressTFiles()


Dim Filpath1$
Dim TilePath As String

Label9.Visible = True
On Error GoTo Errtrap
If Label2(0).Caption = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If

MousePointer = 11
cursouind = 0
TilePath = Dir1(0).Path & "\Tiles"
Filpath1$ = App.Path & "\TempFiles"
Label9.Caption = "Uncompressing T Files"
Call DoDeCompFolder("t", TilePath, Filpath1$)
DoEvents
Dir1(0).Path = TilePath
Text1(0) = "*.t"

Text1(1).Text = "*.t"
Text1(1).Text = "*.*"
  MousePointer = 0
  Label9.Visible = False
Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'UncompressTFIles' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub
Private Sub MakeACEDXTFile()

Dim i As Integer, strBatText As String
Dim Filpath1$, strDrive As String
Dim booCompFound As Boolean
Dim strOrigFile As String
Dim strNewAce As String


On Error GoTo Errtrap
Rem ********** Kill Textures in the temp directory
MousePointer = 11
    cursouind = 1
filepath1$ = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.bmp"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 DoEvents
 Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.tga"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.*"
ShapePath = Dir1(0).Path

 cursouind = 0


Filpath1$ = App.Path & "\tempfiles"

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    If Left$(File1(cursouind).List(i), 11) = "acleantrack" Then GoTo NextOne
   strOrigFile = File1(cursouind).List(i)
  strNewAce = Left$(strOrigFile, Len(strOrigFile) - 4) & ".ace"

  Call readFile(fullpath$, bdata())

   If bdata(16) = 18 Then GoTo NextOne
   If bdata(16) = 17 And bdata(20) = 5 Then GoTo NextOne
  
   booCompFound = True
   FileCopy fullpath$, Filpath1$ & "\" & strOrigFile
   DoEvents
     strDrive = Left$(Filpath1$, 1)
      ChDrive strDrive
    ChDir Filpath1$
   strBatText = App.Path & "\AceIt.exe " & strOrigFile & " " & strNewAce & " /dxt /q /u /filter:point"
   Call ShellAndWait(strBatText, True, vbNormalFocus)

DoEvents
Kill fullpath$
DoEvents
FileCopy Filpath1$ & "\" & strNewAce, File1(cursouind).Path & "\" & strNewAce
DoEvents
Kill Filpath1$ & "\" & strNewAce
DoEvents
strNewAce = vbNullString
End If
NextOne:
   Next i
   Close


 cursouind = 0
 Drive1(0).Drive = Left$(ShapePath, 2)
Dir1(0).Path = ShapePath
Text1(0).Text = "*.*"
 Drive1(1).Drive = Left$(Filpath1$, 2)
Dir1(1).Path = Filpath1$
Text1(1).Text = "*.ace"

  MousePointer = 0
Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'MakeACEDXTFile' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub




Private Sub MakeACECompFile(strSaveHere As String)

Dim i As Integer, strBatText As String
Dim Filpath1$, strDrive As String
Dim booCompFound As Boolean
Dim strOrigFile As String
Dim strNewAce As String


On Error GoTo Errtrap
Rem ********** Kill Textures in the temp directory
MousePointer = 11
    cursouind = 1
   
filepath1$ = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.bmp"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 DoEvents
 Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.tga"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.*"
ShapePath = Dir1(0).Path

 cursouind = 0


Filpath1$ = App.Path & "\TempFiles"

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   strOrigFile = File1(cursouind).List(i)
  strNewAce = Left$(strOrigFile, Len(strOrigFile) - 4) & ".ace"


   booCompFound = True
   FileCopy fullpath$, Filpath1$ & "\" & strOrigFile
   DoEvents
      strDrive = Left$(Filpath1$, 1)
      ChDrive strDrive
    ChDir Filpath1$
   
    
     strBatText = App.Path & "\AceIt.exe " & strOrigFile & " " & strNewAce & " -q"
  Call ShellAndWait(strBatText, True, vbNormalFocus)
DoEvents
'Kill fullpath$
DoEvents
FileCopy Filpath1$ & "\" & strNewAce, strSaveHere & "\" & strNewAce
DoEvents
Kill Filpath1$ & "\" & strNewAce
DoEvents
Kill Filpath1$ & "\" & strOrigFile
DoEvents
strNewAce = vbNullString
End If
NextOne:
   Next i
   Close


 cursouind = 0
 Drive1(0).Drive = Left$(ShapePath, 2)
Dir1(0).Path = ShapePath
Text1(0).Text = "*.*"
 Drive1(1).Drive = Left$(Filpath1$, 2)
Dir1(1).Path = strSaveHere
Text1(1).Text = "*.*"

  MousePointer = 0
Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'MakeACECompFile' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub




Private Sub MakeACEFile(strSaveHere As String)

Dim i As Integer, strBatText As String
Dim Filpath1$, strDrive As String
Dim booCompFound As Boolean
Dim strOrigFile As String
Dim strNewAce As String


On Error GoTo Errtrap
Rem ********** Kill Textures in the temp directory
MousePointer = 11
    cursouind = 1
filepath1$ = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.bmp"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 DoEvents
 Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.tga"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.*"
ShapePath = Dir1(0).Path

 cursouind = 0


Filpath1$ = App.Path & "\tempfiles"

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   strOrigFile = File1(cursouind).List(i)
  strNewAce = Left$(strOrigFile, Len(strOrigFile) - 4) & ".ace"

   booCompFound = True
   FileCopy fullpath$, Filpath1$ & "\" & strOrigFile
   DoEvents
     strDrive = Left$(Filpath1$, 1)
      ChDrive strDrive
    ChDir Filpath1$
    strBatText = App.Path & "\AceIt.exe " & strOrigFile & " " & strNewAce & " /u /q"
   
  
Call ShellAndWait(strBatText, True, vbNormalFocus)
DoEvents

DoEvents
FileCopy Filpath1$ & "\" & strNewAce, strSaveHere & "\" & strNewAce
DoEvents
Kill Filpath1$ & "\" & strNewAce
DoEvents
Kill Filpath1$ & "\" & strOrigFile
strNewAce = vbNullString
End If
NextOne:
   Next i
   Close


 cursouind = 0
 Drive1(0).Drive = Left$(ShapePath, 2)
Dir1(0).Path = ShapePath
Text1(0).Text = "*.*"
 Drive1(1).Drive = Left$(Filpath1$, 2)
Dir1(1).Path = strSaveHere
Text1(1).Text = "*.*"

  MousePointer = 0
Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If
Call MsgBox("An error " & Err.Description & " occurred in subroutine 'MakeACEFile' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub



Private Sub CompressFiles()
Dim i As Integer
Dim strOrigFile As String, strSpare As String, strSpare2 As String, lOf As Long


On Error GoTo Errtrap

MousePointer = 11

 cursouind = 0
ShapePath = Dir1(0).Path
strSpare = App.Path & "\Tempfiles"
strSpare2 = App.Path & "\Tempfiles2"
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
 
    fullpath$ = File1(cursouind).Path
   strOrigFile = File1(cursouind).List(i)
   FileCopy fullpath$ & "\" & strOrigFile, strSpare2 & "\" & strOrigFile
   lOf = Len(fullpath$ & "\" & strOrigFile)
   
   If Right$(strOrigFile, 2) <> ".s" And Right$(strOrigFile, 2) <> ".t" And Right$(strOrigFile, 2) <> ".w" Then
   MousePointer = 0
   Call MsgBox(Lang(393) & strOrigFile & Lang(410), vbExclamation, Lang(404))
    GoTo NextOne
    End If
   
    Call DoComp(strOrigFile, fullpath$, Dir1(1).Path)
   DoEvents

If Not FileExists(Dir1(1).Path & "\" & strOrigFile) Or FileLen(Dir1(1).Path & "\" & strOrigFile) = lOf Then
FileCopy strSpare2 & "\" & strOrigFile, fullpath$ & "\" & strOrigFile
strReport = strReport & strOrigFile & " Could not be compressed" & vbCrLf
DoEvents
Kill strSpare2 & "\" & strOrigFile
End If
   End If
NextOne:

   Next i
  
 cursouind = 0
 Drive1(0).Drive = Left$(ShapePath, 2)
Dir1(0).Path = ShapePath
Text1(0).Text = "*.*"
Text1(1).Text = "*.s"
DoEvents
Text1(1).Text = "*.*"
  MousePointer = 0
Exit Sub

Errtrap:


If Err = 53 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'CompressFiles' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub

Private Sub CompressSFiles()


'Set tfh = New TokenFileHandler
On Error GoTo Errtrap

MousePointer = 11

 cursouind = 0
 Drive1(0).Drive = Left$(ShapePath, 2)
Dir1(0).Path = ShapePath
Text1(0).Text = "*.s"


ShapePath = Dir1(0).Path
SB1.Panels(2).Text = "Compressing .s Files"
Call DoCompFolder("s", ShapePath, ShapePath)


 cursouind = 0
 Drive1(0).Drive = Left$(ShapePath, 2)
Dir1(0).Path = ShapePath
Text1(0).Text = "*.*"
Text1(1).Text = "*.s"
DoEvents
Text1(1).Text = "*.*"
  MousePointer = 0
Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'CompressSFiles' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next


End Sub


Private Sub CompressWFiles()


On Error GoTo Errtrap

MousePointer = 11



SB1.Panels(2).Text = "Compressing .w Files"
Call DoCompFolder("w", WorldPath, WorldPath)



 cursouind = 0
 Drive1(0).Drive = Left$(WorldPath, 2)
Dir1(0).Path = WorldPath
Text1(0).Text = "*.*"
Text1(1).Text = "*.w"
DoEvents
Text1(1).Text = "*.*"
  MousePointer = 0
Exit Sub
Errtrap:

If Err = 53 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'CompressWFiles' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next

End Sub
Private Sub FindFolder(ThePath As String)
On Error GoTo Errtrap
Dim x As Integer
If Right$(ThePath, 1) = "\" Then
ThePath = Left$(ThePath, Len(ThePath) - 1)
End If
x = InStrRev(ThePath, "\")
ThePath = Mid$(ThePath, x + 1)
Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'FindFolder' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub

Private Sub KillEnv()
Dim i As Integer, strMid As String, flagway As Integer, SparePath As String

SparePath = App.Path & "\TempFiles"
Rem ************** change .trk file back to default
strMid = "Environment (" & vbCrLf
strMid = strMid & vbTab & "SpringClear (Sun.env)" & vbCrLf
strMid = strMid & vbTab & "SpringRain (Rain.env)" & vbCrLf
strMid = strMid & vbTab & "SpringSnow (Snow.env)" & vbCrLf
strMid = strMid & vbTab & "SummerClear (Sun.env)" & vbCrLf
strMid = strMid & vbTab & "SummerRain (Rain.env)" & vbCrLf
strMid = strMid & vbTab & "SummerSnow (Snow.env)" & vbCrLf
strMid = strMid & vbTab & "AutumnClear (Sun.env)" & vbCrLf
strMid = strMid & vbTab & "AutumnRain (Rain.env)" & vbCrLf
strMid = strMid & vbTab & "AutumnSnow (Snow.env)" & vbCrLf
strMid = strMid & vbTab & "WinterClear (Sun.env)" & vbCrLf
strMid = strMid & vbTab & "WinterRain (Rain.env)" & vbCrLf
strMid = strMid & vbTab & "WinterSnow (Snow.env)" & vbCrLf

FileCopy RoutePath & "\" & OldRouteName, SparePath & "\" & OldRouteName
flagway = 0
Call ConvertTrk(SparePath & "\" & OldRouteName, flagway, strMid)
DoEvents
flagway = 1
Call ConvertTrk(SparePath & "\" & OldRouteName, flagway, strMid)
DoEvents
Close

FileCopy SparePath & "\" & OldRouteName, RoutePath & "\" & OldRouteName
DoEvents
If FileExists(SparePath & "\" & OldRouteName) Then
Kill SparePath & "\" & OldRouteName
End If



Rem *********************************
cursouind = 0
Drive1(cursouind).Drive = Left$(RoutePath, 2)
Dir1(cursouind).Path = RoutePath & "\Envfiles"
Text1(cursouind).Text = "*.env"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 End If
  Next i
  
 
Dir1(cursouind).Path = RoutePath & "\Envfiles\textures"
Text1(cursouind).Text = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 End If
  Next i


End Sub

Private Sub MakeReadWrite(strRoute As String)
Dim strBat(0 To 23) As String, strDrive As String, Filpath1$, i As Integer

Filpath1$ = App.Path & "\TempFiles"
'strRoute = strGetShortFileName(strRoute)
strBat(0) = "ATTRIB -R " & ChrW$(34) & strRoute & "\*.*" & ChrW$(34)
strBat(1) = "ATTRIB -R " & ChrW$(34) & strRoute & "\shapes\*.*" & ChrW$(34)
strBat(2) = "ATTRIB -R " & ChrW$(34) & strRoute & "\textures\*.*" & ChrW$(34)
strBat(3) = "ATTRIB -R " & ChrW$(34) & strRoute & "\terrtex\*.*" & ChrW$(34)
strBat(4) = "ATTRIB -R " & ChrW$(34) & strRoute & "\sound\*.*" & ChrW$(34)
strBat(5) = "ATTRIB -R " & ChrW$(34) & strRoute & "\tiles\*.*" & ChrW$(34)
strBat(6) = "ATTRIB -R " & ChrW$(34) & strRoute & "\world\*.*" & ChrW$(34)
strBat(7) = "ATTRIB -R " & ChrW$(34) & strRoute & "\terrtex\snow\*.*" & ChrW$(34)
strBat(8) = "ATTRIB -R " & ChrW$(34) & strRoute & "\textures\night\*.*" & ChrW$(34)
strBat(9) = "ATTRIB -R " & ChrW$(34) & strRoute & "\textures\winter\*.*" & ChrW$(34)
strBat(10) = "ATTRIB -R " & ChrW$(34) & strRoute & "\textures\spring\*.*" & ChrW$(34)
strBat(11) = "ATTRIB -R " & ChrW$(34) & strRoute & "\textures\autumn\*.*" & ChrW$(34)
strBat(12) = "ATTRIB -R " & ChrW$(34) & strRoute & "\textures\wintersnow\*.*" & ChrW$(34)
strBat(13) = "ATTRIB -R " & ChrW$(34) & strRoute & "\textures\springsnow\*.*" & ChrW$(34)
strBat(14) = "ATTRIB -R " & ChrW$(34) & strRoute & "\textures\autumnsnow\*.*" & ChrW$(34)
strBat(15) = "ATTRIB -R " & ChrW$(34) & strRoute & "\textures\snow\*.*" & ChrW$(34)
strBat(16) = "ATTRIB -R " & ChrW$(34) & strRoute & "\envfiles\*.*" & ChrW$(34)
strBat(17) = "ATTRIB -R " & ChrW$(34) & strRoute & "\envfiles\textures\*.*" & ChrW$(34)
strBat(18) = "ATTRIB -R " & ChrW$(34) & strRoute & "\activities\*.*" & ChrW$(34)
strBat(19) = "ATTRIB -R " & ChrW$(34) & strRoute & "\paths\*.*" & ChrW$(34)
strBat(20) = "ATTRIB -R " & ChrW$(34) & strRoute & "\services\*.*" & ChrW$(34)
strBat(21) = "ATTRIB -R " & ChrW$(34) & strRoute & "\traffic\*.*" & ChrW$(34)
strBat(22) = "ATTRIB -R " & ChrW$(34) & strRoute & "\Td\*.*" & ChrW$(34)
strBat(23) = "ATTRIB -R " & ChrW$(34) & strRoute & "\Lo_tiles\*.*" & ChrW$(34)




Newfile3 = FreeFile

Open Filpath1$ & "\do_read.bat" For Output As #Newfile3
For i = 0 To 21
   Print #Newfile3, strBat(i)
   Next i
   Close Newfile3
  
   strDrive = Left$(Filpath1$, 1)
      ChDrive strDrive
ChDir Filpath1$
mydir = CurDir
  DoEvents
'result = Shell(Environ$("comspec") & " /c do_read.bat", vbNormalFocus)
Call ShellAndWait("do_read.bat", True, vbNormalFocus)
 MousePointer = 0
End Sub

Private Sub SmsAlreadyExists(Shapefile As String, ShapeFound As Boolean)
Dim strBat As String
MainRoutePath = MSTSPath & "\routes\"
Eur1Path = MainRoutePath & "Europe1\"
Eur2Path = MainRoutePath & "Europe2\"
Jap1Path = MainRoutePath & "Japan1\"
Jap2Path = MainRoutePath & "Japan2\"
USA1Path = MainRoutePath & "USA1\"
USA2Path = MainRoutePath & "USA2\"

Open App.Path & "\SetupFiles\Installme.bat" For Append As #12



If FileExists(Eur1Path & "Sound\" & Shapefile) Then
If GetCRC(Eur1Path & "Sound\" & Shapefile) = GetCRC(RoutePath & "\Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\europe1\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo Label1
End If
End If
If FileExists(Eur2Path & "Sound\" & Shapefile) Then
If GetCRC(Eur2Path & "Sound\" & Shapefile) = GetCRC(RoutePath & "\Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\europe2\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString
      
ShapeFound = True
GoTo Label1
End If
End If
If FileExists(Jap1Path & "Sound\" & Shapefile) Then
If GetCRC(Jap1Path & "Sound\" & Shapefile) = GetCRC(RoutePath & "\Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\Japan1\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString
    
    
ShapeFound = True
GoTo Label1
End If
End If
If FileExists(Jap2Path & "Sound\" & Shapefile) Then
If GetCRC(Jap2Path & "Sound\" & Shapefile) = GetCRC(RoutePath & "\Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\Japan2\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString

ShapeFound = True
GoTo Label1
End If
End If
If FileExists(USA1Path & "Sound\" & Shapefile) Then
If GetCRC(USA1Path & "Sound\" & Shapefile) = GetCRC(RoutePath & "\Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\USA1\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString

ShapeFound = True
GoTo Label1
End If
End If
If FileExists(USA2Path & "Sound\" & Shapefile) Then
If GetCRC(USA2Path & "Sound\" & Shapefile) = GetCRC(RoutePath & "\Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\USA2\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString

    
ShapeFound = True
GoTo Label1
End If
End If
ShapeFound = False
Label1:
Close #12
End Sub
Private Sub WavAlreadyExists(Shapefile As String, ShapeFound As Boolean)
Dim strBat As String
MainRoutePath = MSTSPath & "\routes\"
Eur1Path = MainRoutePath & "Europe1\"
Eur2Path = MainRoutePath & "Europe2\"
Jap1Path = MainRoutePath & "Japan1\"
Jap2Path = MainRoutePath & "Japan2\"
USA1Path = MainRoutePath & "USA1\"
USA2Path = MainRoutePath & "USA2\"

Open App.Path & "\SetupFiles\Installme.bat" For Append As #12




If FileExists(Eur1Path & "Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\europe1\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString
    
    ShapeFound = True
GoTo Label1
End If
If FileExists(Eur2Path & "Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\europe2\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString
      
ShapeFound = True
GoTo Label1
End If
If FileExists(Jap1Path & "Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\Japan1\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString
    
    
ShapeFound = True
GoTo Label1
End If
If FileExists(Jap2Path & "Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\Japan2\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString

ShapeFound = True
GoTo Label1
End If
If FileExists(USA1Path & "Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\USA1\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString

ShapeFound = True
GoTo Label1
End If
If FileExists(USA2Path & "Sound\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\USA2\Sound\" & Shapefile & ChrW$(34) & " .\Sound\ /s /y"
Print #12, strBat
strBat = vbNullString

    
ShapeFound = True
GoTo Label1
End If
ShapeFound = False
Label1:
Close #12
End Sub



Private Function ReadASCIIFile(CompleteFilePath As String) As String

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
mytristate = 0
DoEvents 'DISPLAY CATCHES UP WITH PROGGIE

Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll
The_obj.Close
fileflag = False
ReadASCIIFile = MyString
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


Private Function ConvertIt(CompleteFilePath As String, flagway As Integer) As Boolean
'CONVERTS A FILE FROM UNICODE TO TEXT OR VICE VERSA
'RETURNS TRUE IF THE PROCESS IS SUCCESSFUL, OTHERWISE IT RETURNS FALSE
'THE FILE MUST BE < (2^31 - 1) BYTES (2,147,483,647 BYTES)

On Error GoTo ERRHANDLER
ConvertIt = False

Dim length As Long, mytristate As Integer, outfiletype As Boolean
Dim MyString As String, tempfile As String, tempfolder As String
Dim File_obj As Object, The_obj As Object, fileflag As Boolean

'GET TEMP FOLDER
tempfolder = Environ("TEMP")
If tempfolder = vbNullString Then
  MkDir (Left$(App.Path, 3) & "Temp")
  tempfolder = Left$(App.Path, 3) & "Temp"
End If
If Right$(tempfolder, 1) <> "\" Then tempfolder = tempfolder & "\"

Set File_obj = CreateObject("Scripting.FileSystemObject")
'MAKE SURE THE FILE EXISTS
If Not File_obj.FileExists(CompleteFilePath) Then Exit Function
'MAKE SURE THAT THE FILE LENGTH WILL FIT INTO MYSTRING VARIABLE
length = FileLen(CompleteFilePath)
If length >= 2000000000 Then
  MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & " is too large to convert!", vbInformation, Me.Caption
  Exit Function
End If

If flagway = 1 Then
  
  mytristate = 0 'INFILE IS ANSI
  outfiletype = True 'OUTFILE WILL BE UNICODE
Else 'UNICODE2TEXT
  
  mytristate = -1 'INFILE IS UNICODE
  outfiletype = False 'OUTFILE WILL BE ANSI
End If
DoEvents


Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, 0)
fileflag = True
MyString = The_obj.Read(1)
The_obj.Close
fileflag = False


If flagway = 1 Then
  If Asc(MyString) > 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(402), vbInformation, Me.Caption
    Exit Function
  End If
Else
  If Asc(MyString) < 254 Then
    MsgBox ChrW$(34) & Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ChrW$(34) & Lang(403), vbInformation, Me.Caption
    Exit Function
  End If
End If


Set The_obj = File_obj.OpenTextFile(CompleteFilePath, 1, False, mytristate)
fileflag = True
MyString = The_obj.ReadAll
The_obj.Close
fileflag = False


tempfile = tempfolder & _
      Mid$(CompleteFilePath, InStrRev(CompleteFilePath, "\") + 1) & ".TEMP"


Set The_obj = File_obj.CreateTextFile(tempfile, True, outfiletype)
fileflag = True


On Error Resume Next
  The_obj.Write (MyString)
  If Err.Number <> 0 Then
    If fileflag Then The_obj.Close
    MyString = vbNullString
    MsgBox Lang(404), vbExclamation, Me.Caption
    Kill tempfile
    Exit Function
  End If
On Error GoTo ERRHANDLER

The_obj.Close
fileflag = False



FileCopy tempfile, CompleteFilePath
Kill tempfile
ConvertIt = True

ERRHANDLER:
  If fileflag Then The_obj.Close
  MyString = vbNullString
  If Err.Number <> 0 Then MsgBox "Windows Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, Me.Caption
  
End Function





Private Sub TransferExists(Shapefile As String, ShapeFound As Boolean)
Dim strBat As String
MainRoutePath = MSTSPath & "\routes\"
Eur1Path = MainRoutePath & "Europe1\"
Eur2Path = MainRoutePath & "Europe2\"
Jap1Path = MainRoutePath & "Japan1\"
Jap2Path = MainRoutePath & "Japan2\"
USA1Path = MainRoutePath & "USA1\"
USA2Path = MainRoutePath & "USA2\"
ShapeFound = False
'If ShapeFile = "terrain.ace" Then Exit Sub
Open App.Path & "\SetupFiles\Installme.bat" For Append As #12




If FileExists(Eur1Path & "TerrTex\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\europe1\TerrTex\" & Shapefile & ChrW$(34) & " .\Textures\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo Snow
End If
If FileExists(Eur2Path & "TerrTex\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\europe2\TerrTex\" & Shapefile & ChrW$(34) & " .\Textures\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo Snow
End If
If FileExists(Jap1Path & "TerrTex\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\Japan1\TerrTex\" & Shapefile & ChrW$(34) & " .\Textures\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo Snow
End If
If FileExists(Jap2Path & "TerrTex\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\Japan2\TerrTex\" & Shapefile & ChrW$(34) & " .\Textures\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo Snow
End If
If FileExists(USA1Path & "TerrTex\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\USA1\TerrTex\" & Shapefile & ChrW$(34) & " .\Textures\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo Snow
End If
If FileExists(USA2Path & "TerrTex\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\USA2\TerrTex\" & Shapefile & ChrW$(34) & " .\Textures\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
End If

Snow:


If FileExists(Eur1Path & "TerrTex\Snow\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\europe1\TerrTex\Snow\" & Shapefile & ChrW$(34) & " .\Textures\Snow\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
If FileExists(Eur2Path & "TerrTex\Snow\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\europe2\TerrTex\Snow\" & Shapefile & ChrW$(34) & " .\Textures\Snow\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
If FileExists(Jap1Path & "TerrTex\Snow\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\Japan1\TerrTex\Snow\" & Shapefile & ChrW$(34) & " .\Textures\Snow\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
If FileExists(Jap2Path & "TerrTex\Snow\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\Japan2\TerrTex\Snow\" & Shapefile & ChrW$(34) & " .\Textures\Snow\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
If FileExists(USA1Path & "TerrTex\Snow\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\USA1\TerrTex\Snow\" & Shapefile & ChrW$(34) & " .\Textures\Snow\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
If FileExists(USA2Path & "TerrTex\Snow\" & Shapefile) Then
strBat = "call Xcopy " & ChrW$(34) & "..\USA2\TerrTex\Snow\" & Shapefile & ChrW$(34) & " .\Textures\Snow\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
End If
    
GotIt:

Close #12
End Sub

Private Sub TerrAlreadyExists(Shapefile As String, ShapeFound As Boolean, SnowShapeFound As Boolean)
Dim strBat As String
MainRoutePath = MSTSPath & "\routes\"
Eur1Path = MainRoutePath & "Europe1\"
Eur2Path = MainRoutePath & "Europe2\"
Jap1Path = MainRoutePath & "Japan1\"
Jap2Path = MainRoutePath & "Japan2\"
USA1Path = MainRoutePath & "USA1\"
USA2Path = MainRoutePath & "USA2\"
'TemplatePath = MSTSPath & "\Template\"
'If ShapeFile = "terrain.ace" Then Exit Sub
Open App.Path & "\SetupFiles\Installme.bat" For Append As #12

'************************************
Rem ****************************
ShapeFound = False
SnowShapeFound = False


If booSnow = False Then

If FileExists(Eur1Path & "TerrTex\" & Shapefile) Then
If GetCRC(Eur1Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\europe1\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If
End If
If FileExists(Eur1Path & "TerrTex\snow\" & Shapefile) Then
If GetCRC(Eur1Path & "Terrtex\snow\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\europe1\TerrTex\snow\" & Shapefile & ChrW$(34) & " .\TerrTex\Snow\ /y"
Print #12, strBat
strBat = vbNullString
SnowShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If

End If
If FileExists(Eur2Path & "TerrTex\" & Shapefile) And ShapeFound = False Then
If GetCRC(Eur2Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\europe2\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If
End If
If FileExists(Eur2Path & "TerrTex\snow\" & Shapefile) Then
If GetCRC(Eur2Path & "Terrtex\snow\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\europe2\TerrTex\snow\" & Shapefile & ChrW$(34) & " .\TerrTex\snow\ /y"
Print #12, strBat
strBat = vbNullString
SnowShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If
End If
If FileExists(Jap1Path & "TerrTex\" & Shapefile) And ShapeFound = False Then
If GetCRC(Jap1Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If
End If
If FileExists(Jap1Path & "TerrTex\snow\" & Shapefile) Then
If GetCRC(Jap1Path & "Terrtex\snow\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\TerrTex\snow\" & Shapefile & ChrW$(34) & " .\TerrTex\snow\ /y"
Print #12, strBat
strBat = vbNullString
SnowShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If
End If
If FileExists(Jap2Path & "TerrTex\" & Shapefile) And ShapeFound = False Then
If GetCRC(Jap2Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If
End If
If FileExists(Jap2Path & "TerrTex\snow\" & Shapefile) Then
If GetCRC(Jap2Path & "Terrtex\snow\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\TerrTex\snow\" & Shapefile & ChrW$(34) & " .\TerrTex\snow\ /y"
Print #12, strBat
strBat = vbNullString
SnowShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If
End If
If FileExists(USA1Path & "TerrTex\" & Shapefile) And ShapeFound = False Then
If GetCRC(USA1Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If
End If
If FileExists(USA1Path & "TerrTex\snow\" & Shapefile) Then
If GetCRC(USA1Path & "Terrtex\snow\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\TerrTex\snow\" & Shapefile & ChrW$(34) & " .\TerrTex\snow\ /y"
Print #12, strBat
strBat = vbNullString
SnowShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If
End If
If FileExists(USA2Path & "TerrTex\" & Shapefile) And ShapeFound = False Then
If GetCRC(USA2Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If
End If
If FileExists(USA2Path & "TerrTex\snow\" & Shapefile) Then
If GetCRC(USA2Path & "Terrtex\snow\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\TerrTex\snow\" & Shapefile & ChrW$(34) & " .\TerrTex\snow\ /y"
Print #12, strBat
strBat = vbNullString
SnowShapeFound = True
If ShapeFound = True And SnowShapeFound = True Then GoTo GotIt
End If
End If
ElseIf booSnow = True Then
If FileExists(Eur1Path & "TerrTex\" & Shapefile) Then
If GetCRC(Eur1Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\europe1\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
If FileExists(Eur2Path & "TerrTex\" & Shapefile) Then
If GetCRC(Eur2Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\europe2\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
If FileExists(Jap1Path & "TerrTex\" & Shapefile) Then
If GetCRC(Jap1Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
If FileExists(Jap2Path & "TerrTex\" & Shapefile) Then
If GetCRC(Jap2Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
If FileExists(USA1Path & "TerrTex\" & Shapefile) Then
If GetCRC(USA1Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
If FileExists(USA2Path & "TerrTex\" & Shapefile) Then
If GetCRC(USA2Path & "Terrtex\" & Shapefile) = GetCRC(RoutePath & "\Terrtex\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\TerrTex\" & Shapefile & ChrW$(34) & " .\TerrTex\ /s /y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If

End If
GotIt:

Close #12

End Sub


Private Sub ShapeAlreadyExists(Shapefile As String, ShapeFound As Boolean)
Dim strBat As String
Dim booE1 As Boolean, booE2 As Boolean, booJ1 As Boolean, booJ2 As Boolean
Dim booU1 As Boolean, booU2 As Boolean

MainRoutePath = MSTSPath & "\routes\"
Eur1Path = MainRoutePath & "Europe1\shapes\"
Eur2Path = MainRoutePath & "Europe2\shapes\"
Jap1Path = MainRoutePath & "Japan1\shapes\"
Jap2Path = MainRoutePath & "Japan2\shapes\"
USA1Path = MainRoutePath & "USA1\shapes\"
USA2Path = MainRoutePath & "USA2\shapes\"

Close

Open App.Path & "\SetupFiles\Installme.bat" For Append As #12



Rem ****************************

If FileExists(Eur1Path & Shapefile) Then
booE1 = True
'If Right$(Shapefile, 2) = "sd" Then
'strBat = "call xcopy " & chrw$(34) & "..\europe1\shapes\" & Shapefile & chrw$(34) & " .\Shapes\" & " /Y"
'Print #12, strBat
'strBat = vbNullString
'ShapeFound = True
'GoTo GotIt
'End If
If GetCRC(Eur1Path & Shapefile) = GetCRC(RoutePath & "\shapes\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\europe1\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
If FileExists(Eur2Path & Shapefile) Then
booE2 = True
'If Right$(Shapefile, 2) = "sd" Then
'strBat = "call xcopy " & chrw$(34) & "..\europe2\shapes\" & Shapefile & chrw$(34) & " .\Shapes\" & " /Y"
'Print #12, strBat
'strBat = vbNullString
'ShapeFound = True
'GoTo GotIt
'End If
If GetCRC(Eur2Path & Shapefile) = GetCRC(RoutePath & "\shapes\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\europe2\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
If FileExists(Jap1Path & Shapefile) Then
booJ1 = True
'If Right$(Shapefile, 2) = "sd" Then
'strBat = "call xcopy " & chrw$(34) & "..\japan1\shapes\" & Shapefile & chrw$(34) & " .\Shapes\" & " /Y"
'Print #12, strBat
'strBat = vbNullString
'ShapeFound = True
'GoTo GotIt
'End If
If GetCRC(Jap1Path & Shapefile) = GetCRC(RoutePath & "\shapes\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\japan1\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
If FileExists(Jap2Path & Shapefile) Then
booJ2 = True
'If Right$(Shapefile, 2) = "sd" Then
'strBat = "call xcopy " & chrw$(34) & "..\japan2\shapes\" & Shapefile & chrw$(34) & " .\Shapes\" & " /Y"
'Print #12, strBat
'strBat = vbNullString
'ShapeFound = True
'GoTo GotIt
'End If
If GetCRC(Jap2Path & Shapefile) = GetCRC(RoutePath & "\shapes\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\japan2\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
If FileExists(USA1Path & Shapefile) Then
booU1 = True
'If Right$(Shapefile, 2) = "sd" Then
'strBat = "call xcopy " & chrw$(34) & "..\usa1\shapes\" & Shapefile & chrw$(34) & " .\Shapes\" & " /Y"
'Print #12, strBat
'strBat = vbNullString
'ShapeFound = True
'GoTo GotIt
'End If
If GetCRC(USA1Path & Shapefile) = GetCRC(RoutePath & "\shapes\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\usa1\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
If FileExists(USA2Path & Shapefile) Then
booU2 = True
'If Right$(Shapefile, 2) = "sd" Then
'strBat = "call xcopy " & chrw$(34) & "..\usa2\shapes\" & Shapefile & chrw$(34) & " .\Shapes\" & " /Y"
'Print #12, strBat
'strBat = vbNullString
'ShapeFound = True
'GoTo GotIt
'End If
If GetCRC(USA2Path & Shapefile) = GetCRC(RoutePath & "\shapes\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\usa2\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
If booE1 = True And intResponse <> 3 And intResponse <> 4 Then
    booExact = True
   strResponse = Shapefile & Lang(412) & vbCrLf & Lang(413) & vbCrLf & Lang(421)
   frmDialog.Show 1
   
     DoEvents
     
   booExact = False
   
   If intResponse = 2 Or intResponse = 4 Then
strBat = "call xcopy " & ChrW$(34) & "..\europe1\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
ElseIf booE2 = True And intResponse <> 3 And intResponse <> 4 Then
    booExact = True
   strResponse = Shapefile & Lang(412) & vbCrLf & Lang(414) & vbCrLf & Lang(421)
  
   frmDialog.Show 1
   
     DoEvents
     
   booExact = False
   
   If intResponse = 2 Or intResponse = 4 Then
strBat = "call xcopy " & ChrW$(34) & "..\europe2\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
ElseIf booJ1 = True And intResponse <> 3 And intResponse <> 4 Then
    booExact = True
   strResponse = Shapefile & Lang(412) & vbCrLf & Lang(415) & vbCrLf & Lang(421)
   
   frmDialog.Show 1
   
     DoEvents
     
   booExact = False
   
   If intResponse = 2 Or intResponse = 4 Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
 End If
ElseIf booJ2 = True And intResponse <> 3 And intResponse <> 4 Then
    booExact = True
   strResponse = Shapefile & Lang(412) & vbCrLf & Lang(416) & vbCrLf & Lang(421)
   
   frmDialog.Show 1
   
     DoEvents
     
   booExact = False
   
   If intResponse = 2 Or intResponse = 4 Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If

ElseIf booU1 = True And intResponse <> 3 And intResponse <> 4 Then
    booExact = True
   strResponse = Shapefile & Lang(412) & vbCrLf & Lang(417) & vbCrLf & Lang(421)
  
   frmDialog.Show 1
   
     DoEvents
     
   booExact = False
   
   If intResponse = 2 Or intResponse = 4 Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If

ElseIf booU2 = True And intResponse <> 3 And intResponse <> 4 Then
    booExact = True
   strResponse = Shapefile & Lang(412) & vbCrLf & Lang(418) & vbCrLf & Lang(421)
   
   frmDialog.Show 1
   
     DoEvents
     
   booExact = False
   
If intResponse = 2 Or intResponse = 4 Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\shapes\" & Shapefile & ChrW$(34) & " .\Shapes\" & " /Y"
Print #12, strBat
strBat = vbNullString
ShapeFound = True
GoTo GotIt
End If
End If
GotIt:
Close #12
End Sub



Private Sub AceAlreadyExists2(Shapefile As String)
Dim strBat As String, TemplatePath As String, SF1 As Boolean, SF2 As Boolean
Dim SF3 As Boolean, SF4 As Boolean, SF5 As Boolean, SF6 As Boolean, SF7 As Boolean
Dim SF8 As Boolean, SF9 As Boolean, i As Integer
Dim booSum(1 To 6) As Boolean, booSnow(1 To 6) As Boolean, booNight(1 To 6) As Boolean
Dim booSpr(1 To 6) As Boolean, booSprS(1 To 6) As Boolean, booAut(1 To 6) As Boolean
Dim booAutS(1 To 6) As Boolean, booWin(1 To 6) As Boolean, booWinS(1 To 6) As Boolean
Dim strRoute(1 To 6)
On Error GoTo Errtrap

strRoute(1) = "USA2"
strRoute(2) = "Europe1"
strRoute(3) = "Europe2"
strRoute(4) = "Japan1"
strRoute(5) = "Japan2"
strRoute(6) = "USA1"


MainRoutePath = MSTSPath & "\routes\"
Eur1Path = MainRoutePath & "Europe1\"
Eur2Path = MainRoutePath & "Europe2\"
Jap1Path = MainRoutePath & "Japan1\"
Jap2Path = MainRoutePath & "Japan2\"
USA1Path = MainRoutePath & "USA1\"
USA2Path = MainRoutePath & "USA2\"
TemplatePath = MSTSPath & "\Template\"


Open App.Path & "\SetupFiles\Installme.bat" For Append As #12



Rem ********** Europe1 **************************

If FileExists(Eur1Path & "Textures\" & Shapefile) And FileExists(RoutePath & "\textures\" & Shapefile) Then
booSum(2) = True
If SF1 = False Then
If GetCRC(Eur1Path & "Textures\" & Shapefile) = GetCRC(RoutePath & "\textures\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe1\Textures\" & Shapefile & ChrW$(34) & " .\Textures\ /y"
Print #12, strBat
strBat = vbNullString
SF1 = True
KillEm(KE) = RoutePath & "\textures\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If

If FileExists(Eur1Path & "Textures\Snow\" & Shapefile) And FileExists(RoutePath & "\textures\snow\" & Shapefile) And SF2 = False Then
booSnow(2) = True
If GetCRC(Eur1Path & "Textures\Snow\" & Shapefile) = GetCRC(RoutePath & "\textures\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe1\Textures\snow\" & Shapefile & ChrW$(34) & " .\Textures\snow\ /y"
Print #12, strBat
strBat = vbNullString
SF2 = True
KillEm(KE) = RoutePath & "\textures\snow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If

If FileExists(Eur1Path & "Textures\Spring\" & Shapefile) And FileExists(RoutePath & "\textures\spring\" & Shapefile) And SF3 = False Then
booSpr(2) = True
If GetCRC(Eur1Path & "Textures\Spring\" & Shapefile) = GetCRC(RoutePath & "\textures\spring\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe1\Textures\spring\" & Shapefile & ChrW$(34) & " .\Textures\spring\ /y"
Print #12, strBat
strBat = vbNullString
SF3 = True
KillEm(KE) = RoutePath & "\textures\spring\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur1Path & "Textures\springSnow\" & Shapefile) And FileExists(RoutePath & "\textures\springsnow\" & Shapefile) And SF4 = False Then
booSprS(2) = True
If GetCRC(Eur1Path & "Textures\springSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\springsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe1\Textures\springsnow\" & Shapefile & ChrW$(34) & " .\Textures\springsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF4 = True
KillEm(KE) = RoutePath & "\textures\springsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur1Path & "Textures\Winter\" & Shapefile) And FileExists(RoutePath & "\textures\Winter\" & Shapefile) And SF5 = False Then
booWin(2) = True
If GetCRC(Eur1Path & "Textures\Winter\" & Shapefile) = GetCRC(RoutePath & "\textures\Winter\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe1\Textures\Winter\" & Shapefile & ChrW$(34) & " .\Textures\Winter\ /y"
Print #12, strBat
strBat = vbNullString
SF5 = True
KillEm(KE) = RoutePath & "\textures\winter\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur1Path & "Textures\WinterSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Wintersnow\" & Shapefile) And SF6 = False Then
booWinS(2) = True
If GetCRC(Eur1Path & "Textures\WinterSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Wintersnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe1\Textures\Wintersnow\" & Shapefile & ChrW$(34) & " .\Textures\Wintersnow\ /y"
Print #12, strBat
strBat = vbNullString
SF6 = True
KillEm(KE) = RoutePath & "\textures\wintersnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If

End If
End If
If FileExists(Eur1Path & "Textures\Autumn\" & Shapefile) And FileExists(RoutePath & "\textures\Autumn\" & Shapefile) And SF7 = False Then
booAut(2) = True
If GetCRC(Eur1Path & "Textures\Autumn\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumn\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe1\Textures\Autumn\" & Shapefile & ChrW$(34) & " .\Textures\Autumn\ /y"
Print #12, strBat
strBat = vbNullString
SF7 = True
KillEm(KE) = RoutePath & "\textures\Autumn\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur1Path & "Textures\AutumnSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Autumnsnow\" & Shapefile) And SF8 = False Then
booAutS(2) = True
If GetCRC(Eur1Path & "Textures\AutumnSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumnsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe1\Textures\Autumnsnow\" & Shapefile & ChrW$(34) & " .\Textures\Autumnsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF8 = True
KillEm(KE) = RoutePath & "\textures\Autumnsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur1Path & "Textures\Night\" & Shapefile) And FileExists(RoutePath & "\textures\Night\" & Shapefile) And SF9 = False Then
booNight(2) = True
If GetCRC(Eur1Path & "Textures\Night\" & Shapefile) = GetCRC(RoutePath & "\textures\Night\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe1\Textures\Night\" & Shapefile & ChrW$(34) & " .\Textures\Night\ /y"
Print #12, strBat
strBat = vbNullString
SF9 = True
KillEm(KE) = RoutePath & "\textures\night\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
End If
Rem ****************************
Rem ********** Europe2 **************************
If FileExists(Eur2Path & "Textures\" & Shapefile) And FileExists(RoutePath & "\textures\" & Shapefile) Then
booSum(3) = True
If SF1 = False Then
If GetCRC(Eur2Path & "Textures\" & Shapefile) = GetCRC(RoutePath & "\textures\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe2\Textures\" & Shapefile & ChrW$(34) & " .\Textures\ /y"
Print #12, strBat
strBat = vbNullString
SF1 = True
KillEm(KE) = RoutePath & "\textures\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur2Path & "Textures\Snow\" & Shapefile) And FileExists(RoutePath & "\textures\snow\" & Shapefile) And SF2 = False Then
booSnow(3) = True
If GetCRC(Eur2Path & "Textures\Snow\" & Shapefile) = GetCRC(RoutePath & "\textures\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe2\Textures\snow\" & Shapefile & ChrW$(34) & " .\Textures\snow\ /y"
Print #12, strBat
strBat = vbNullString
SF2 = True
KillEm(KE) = RoutePath & "\textures\snow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur2Path & "Textures\Spring\" & Shapefile) And FileExists(RoutePath & "\textures\spring\" & Shapefile) And SF3 = False Then
booSpr(3) = True
If GetCRC(Eur2Path & "Textures\Spring\" & Shapefile) = GetCRC(RoutePath & "\textures\spring\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe2\Textures\spring\" & Shapefile & ChrW$(34) & " .\Textures\spring\ /y"
Print #12, strBat
strBat = vbNullString
SF3 = True
KillEm(KE) = RoutePath & "\textures\spring\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur2Path & "Textures\springSnow\" & Shapefile) And FileExists(RoutePath & "\textures\springsnow\" & Shapefile) And SF4 = False Then
booSprS(3) = True
If GetCRC(Eur2Path & "Textures\springSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\springsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe2\Textures\springsnow\" & Shapefile & ChrW$(34) & " .\Textures\springsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF4 = True
KillEm(KE) = RoutePath & "\textures\springsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur2Path & "Textures\Winter\" & Shapefile) And FileExists(RoutePath & "\textures\Winter\" & Shapefile) And SF5 = False Then
booWin(3) = True
If GetCRC(Eur2Path & "Textures\Winter\" & Shapefile) = GetCRC(RoutePath & "\textures\Winter\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe2\Textures\Winter\" & Shapefile & ChrW$(34) & " .\Textures\Winter\ /y"
Print #12, strBat
strBat = vbNullString
SF5 = True
KillEm(KE) = RoutePath & "\textures\winter\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur2Path & "Textures\WinterSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Wintersnow\" & Shapefile) And SF6 = False Then
booWinS(3) = True
If GetCRC(Eur2Path & "Textures\WinterSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Wintersnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe2\Textures\Wintersnow\" & Shapefile & ChrW$(34) & " .\Textures\Wintersnow\ /y"
Print #12, strBat
strBat = vbNullString
SF6 = True
KillEm(KE) = RoutePath & "\textures\wintersnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur2Path & "Textures\Autumn\" & Shapefile) And FileExists(RoutePath & "\textures\Autumn\" & Shapefile) And SF7 = False Then
booAut(3) = True
If GetCRC(Eur2Path & "Textures\Autumn\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumn\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe2\Textures\Autumn\" & Shapefile & ChrW$(34) & " .\Textures\Autumn\ /y"
Print #12, strBat
strBat = vbNullString
SF7 = True
KillEm(KE) = RoutePath & "\textures\Autumn\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur2Path & "Textures\AutumnSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Autumnsnow\" & Shapefile) And SF8 = False Then
booAutS(3) = True
If GetCRC(Eur2Path & "Textures\AutumnSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumnsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe2\Textures\Autumnsnow\" & Shapefile & ChrW$(34) & " .\Textures\Autumnsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF8 = True
KillEm(KE) = RoutePath & "\textures\Autumnsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Eur2Path & "Textures\Night\" & Shapefile) And FileExists(RoutePath & "\textures\Night\" & Shapefile) And SF9 = False Then
booNight(3) = True
If GetCRC(Eur2Path & "Textures\Night\" & Shapefile) = GetCRC(RoutePath & "\textures\Night\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Europe2\Textures\Night\" & Shapefile & ChrW$(34) & " .\Textures\Night\ /y"
Print #12, strBat
strBat = vbNullString
SF9 = True
KillEm(KE) = RoutePath & "\textures\night\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
End If
Rem ****************************
Rem ********** Japan1 **************************
If FileExists(Jap1Path & "Textures\" & Shapefile) And FileExists(RoutePath & "\textures\" & Shapefile) Then
booSum(4) = True
If SF1 = False Then
If GetCRC(Jap1Path & "Textures\" & Shapefile) = GetCRC(RoutePath & "\textures\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\Textures\" & Shapefile & ChrW$(34) & " .\Textures\ /y"
Print #12, strBat
strBat = vbNullString
SF1 = True
KillEm(KE) = RoutePath & "\textures\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap1Path & "Textures\Snow\" & Shapefile) And FileExists(RoutePath & "\textures\snow\" & Shapefile) And SF2 = False Then
booSnow(4) = True
If GetCRC(Jap1Path & "Textures\Snow\" & Shapefile) = GetCRC(RoutePath & "\textures\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\Textures\snow\" & Shapefile & ChrW$(34) & " .\Textures\snow\ /y"
Print #12, strBat
strBat = vbNullString
SF2 = True
KillEm(KE) = RoutePath & "\textures\snow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap1Path & "Textures\Spring\" & Shapefile) And FileExists(RoutePath & "\textures\spring\" & Shapefile) And SF3 = False Then
booSpr(4) = True
If GetCRC(Jap1Path & "Textures\Spring\" & Shapefile) = GetCRC(RoutePath & "\textures\spring\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\Textures\spring\" & Shapefile & ChrW$(34) & " .\Textures\spring\ /y"
Print #12, strBat
strBat = vbNullString
SF3 = True
KillEm(KE) = RoutePath & "\textures\spring\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap1Path & "Textures\springSnow\" & Shapefile) And FileExists(RoutePath & "\textures\springsnow\" & Shapefile) And SF4 = False Then
booSprS(4) = True
If GetCRC(Jap1Path & "Textures\springSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\springsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\Textures\springsnow\" & Shapefile & ChrW$(34) & " .\Textures\springsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF4 = True
KillEm(KE) = RoutePath & "\textures\springsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap1Path & "Textures\Winter\" & Shapefile) And FileExists(RoutePath & "\textures\Winter\" & Shapefile) And SF5 = False Then
booWin(4) = True
If GetCRC(Jap1Path & "Textures\Winter\" & Shapefile) = GetCRC(RoutePath & "\textures\Winter\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\Textures\Winter\" & Shapefile & ChrW$(34) & " .\Textures\Winter\ /y"
Print #12, strBat
strBat = vbNullString
SF5 = True
KillEm(KE) = RoutePath & "\textures\winter\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap1Path & "Textures\WinterSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Wintersnow\" & Shapefile) And SF6 = False Then
booWinS(4) = True
If GetCRC(Jap1Path & "Textures\WinterSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Wintersnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\Textures\Wintersnow\" & Shapefile & ChrW$(34) & " .\Textures\Wintersnow\ /y"
Print #12, strBat
strBat = vbNullString
SF6 = True
KillEm(KE) = RoutePath & "\textures\wintersnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap1Path & "Textures\Autumn\" & Shapefile) And FileExists(RoutePath & "\textures\Autumn\" & Shapefile) And SF7 = False Then
booAut(4) = True
If GetCRC(Jap1Path & "Textures\Autumn\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumn\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\Textures\Autumn\" & Shapefile & ChrW$(34) & " .\Textures\Autumn\ /y"
Print #12, strBat
strBat = vbNullString
SF7 = True
KillEm(KE) = RoutePath & "\textures\Autumn\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap1Path & "Textures\AutumnSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Autumnsnow\" & Shapefile) And SF8 = False Then
booAutS(4) = True
If GetCRC(Jap1Path & "Textures\AutumnSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumnsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\Textures\Autumnsnow\" & Shapefile & ChrW$(34) & " .\Textures\Autumnsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF8 = True
KillEm(KE) = RoutePath & "\textures\Autumnsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap1Path & "Textures\Night\" & Shapefile) And FileExists(RoutePath & "\textures\Night\" & Shapefile) And SF9 = False Then
booNight(4) = True
If GetCRC(Jap1Path & "Textures\Night\" & Shapefile) = GetCRC(RoutePath & "\textures\Night\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan1\Textures\Night\" & Shapefile & ChrW$(34) & " .\Textures\Night\ /y"
Print #12, strBat
strBat = vbNullString
SF9 = True
KillEm(KE) = RoutePath & "\textures\night\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
End If
Rem ****************************
Rem ********** Japan2 **************************
If FileExists(Jap2Path & "Textures\" & Shapefile) And FileExists(RoutePath & "\textures\" & Shapefile) Then
booSum(5) = True
If SF1 = False Then
If GetCRC(Jap2Path & "Textures\" & Shapefile) = GetCRC(RoutePath & "\textures\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\Textures\" & Shapefile & ChrW$(34) & " .\Textures\ /y"
Print #12, strBat
strBat = vbNullString
SF1 = True
KillEm(KE) = RoutePath & "\textures\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap2Path & "Textures\Snow\" & Shapefile) And FileExists(RoutePath & "\textures\snow\" & Shapefile) And SF2 = False Then
booSnow(5) = True
If GetCRC(Jap2Path & "Textures\Snow\" & Shapefile) = GetCRC(RoutePath & "\textures\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\Textures\snow\" & Shapefile & ChrW$(34) & " .\Textures\snow\ /y"
Print #12, strBat
strBat = vbNullString
SF2 = True
KillEm(KE) = RoutePath & "\textures\snow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap2Path & "Textures\Spring\" & Shapefile) And FileExists(RoutePath & "\textures\spring\" & Shapefile) And SF3 = False Then
booSpr(5) = True
If GetCRC(Jap2Path & "Textures\Spring\" & Shapefile) = GetCRC(RoutePath & "\textures\spring\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\Textures\spring\" & Shapefile & ChrW$(34) & " .\Textures\spring\ /y"
Print #12, strBat
strBat = vbNullString
SF3 = True
KillEm(KE) = RoutePath & "\textures\spring\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap2Path & "Textures\springSnow\" & Shapefile) And FileExists(RoutePath & "\textures\springsnow\" & Shapefile) And SF4 = False Then
booSprS(5) = True
If GetCRC(Jap2Path & "Textures\springSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\springsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\Textures\springsnow\" & Shapefile & ChrW$(34) & " .\Textures\springsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF4 = True
KillEm(KE) = RoutePath & "\textures\springsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap2Path & "Textures\Winter\" & Shapefile) And FileExists(RoutePath & "\textures\Winter\" & Shapefile) And SF5 = False Then
booWin(5) = True
If GetCRC(Jap2Path & "Textures\Winter\" & Shapefile) = GetCRC(RoutePath & "\textures\Winter\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\Textures\Winter\" & Shapefile & ChrW$(34) & " .\Textures\Winter\ /y"
Print #12, strBat
strBat = vbNullString
SF5 = True
KillEm(KE) = RoutePath & "\textures\winter\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap2Path & "Textures\WinterSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Wintersnow\" & Shapefile) And SF6 = False Then
booWinS(5) = True
If GetCRC(Jap2Path & "Textures\WinterSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Wintersnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\Textures\Wintersnow\" & Shapefile & ChrW$(34) & " .\Textures\Wintersnow\ /y"
Print #12, strBat
strBat = vbNullString
SF6 = True
KillEm(KE) = RoutePath & "\textures\wintersnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap2Path & "Textures\Autumn\" & Shapefile) And FileExists(RoutePath & "\textures\Autumn\" & Shapefile) And SF7 = False Then
booAut(5) = True
If GetCRC(Jap2Path & "Textures\Autumn\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumn\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\Textures\Autumn\" & Shapefile & ChrW$(34) & " .\Textures\Autumn\ /y"
Print #12, strBat
strBat = vbNullString
SF7 = True
KillEm(KE) = RoutePath & "\textures\Autumn\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap2Path & "Textures\AutumnSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Autumnsnow\" & Shapefile) And SF8 = False Then
booAutS(5) = True
If GetCRC(Jap2Path & "Textures\AutumnSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumnsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\Textures\Autumnsnow\" & Shapefile & ChrW$(34) & " .\Textures\Autumnsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF8 = True
KillEm(KE) = RoutePath & "\textures\Autumnsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(Jap2Path & "Textures\Night\" & Shapefile) And FileExists(RoutePath & "\textures\Night\" & Shapefile) And SF9 = False Then
booNight(5) = True
If GetCRC(Jap2Path & "Textures\Night\" & Shapefile) = GetCRC(RoutePath & "\textures\Night\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\Japan2\Textures\Night\" & Shapefile & ChrW$(34) & " .\Textures\Night\ /y"
Print #12, strBat
strBat = vbNullString
SF9 = True
KillEm(KE) = RoutePath & "\textures\night\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
End If
Rem ****************************
Rem ********** USA1 **************************
If FileExists(USA1Path & "Textures\" & Shapefile) And FileExists(RoutePath & "\textures\" & Shapefile) Then
booSum(6) = True
If SF1 = False Then
If GetCRC(USA1Path & "Textures\" & Shapefile) = GetCRC(RoutePath & "\textures\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\Textures\" & Shapefile & ChrW$(34) & " .\Textures\ /y"
Print #12, strBat
strBat = vbNullString
SF1 = True
KillEm(KE) = RoutePath & "\textures\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA1Path & "Textures\Snow\" & Shapefile) And FileExists(RoutePath & "\textures\snow\" & Shapefile) And SF2 = False Then
booSnow(6) = True
If GetCRC(USA1Path & "Textures\Snow\" & Shapefile) = GetCRC(RoutePath & "\textures\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\Textures\snow\" & Shapefile & ChrW$(34) & " .\Textures\snow\ /y"
Print #12, strBat
strBat = vbNullString
SF2 = True
KillEm(KE) = RoutePath & "\textures\snow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA1Path & "Textures\Spring\" & Shapefile) And FileExists(RoutePath & "\textures\spring\" & Shapefile) And SF3 = False Then
booSpr(6) = True
If GetCRC(USA1Path & "Textures\Spring\" & Shapefile) = GetCRC(RoutePath & "\textures\spring\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\Textures\spring\" & Shapefile & ChrW$(34) & " .\Textures\spring\ /y"
Print #12, strBat
strBat = vbNullString
SF3 = True
KillEm(KE) = RoutePath & "\textures\spring\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA1Path & "Textures\springSnow\" & Shapefile) And FileExists(RoutePath & "\textures\springsnow\" & Shapefile) And SF4 = False Then
booSprS(6) = True
If GetCRC(USA1Path & "Textures\springSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\springsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\Textures\springsnow\" & Shapefile & ChrW$(34) & " .\Textures\springsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF4 = True
KillEm(KE) = RoutePath & "\textures\springsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA1Path & "Textures\Winter\" & Shapefile) And FileExists(RoutePath & "\textures\Winter\" & Shapefile) And SF5 = False Then
booWin(6) = True
If GetCRC(USA1Path & "Textures\Winter\" & Shapefile) = GetCRC(RoutePath & "\textures\Winter\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\Textures\Winter\" & Shapefile & ChrW$(34) & " .\Textures\Winter\ /y"
Print #12, strBat
strBat = vbNullString
SF5 = True
KillEm(KE) = RoutePath & "\textures\winter\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA1Path & "Textures\WinterSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Wintersnow\" & Shapefile) And SF6 = False Then
booWinS(6) = True
If GetCRC(USA1Path & "Textures\WinterSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Wintersnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\Textures\Wintersnow\" & Shapefile & ChrW$(34) & " .\Textures\Wintersnow\ /y"
Print #12, strBat
strBat = vbNullString
SF6 = True
KillEm(KE) = RoutePath & "\textures\wintersnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA1Path & "Textures\Autumn\" & Shapefile) And FileExists(RoutePath & "\textures\Autumn\" & Shapefile) And SF7 = False Then
booAut(6) = True
If GetCRC(USA1Path & "Textures\Autumn\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumn\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\Textures\Autumn\" & Shapefile & ChrW$(34) & " .\Textures\Autumn\ /y"
Print #12, strBat
strBat = vbNullString
SF7 = True
KillEm(KE) = RoutePath & "\textures\Autumn\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA1Path & "Textures\AutumnSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Autumnsnow\" & Shapefile) And SF8 = False Then
booAutS(6) = True
If GetCRC(USA1Path & "Textures\AutumnSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumnsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\Textures\Autumnsnow\" & Shapefile & ChrW$(34) & " .\Textures\Autumnsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF8 = True
KillEm(KE) = RoutePath & "\textures\Autumnsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA1Path & "Textures\Night\" & Shapefile) And FileExists(RoutePath & "\textures\Night\" & Shapefile) And SF9 = False Then
booNight(6) = True
If GetCRC(USA1Path & "Textures\Night\" & Shapefile) = GetCRC(RoutePath & "\textures\Night\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA1\Textures\Night\" & Shapefile & ChrW$(34) & " .\Textures\Night\ /y"
Print #12, strBat
strBat = vbNullString
SF9 = True
KillEm(KE) = RoutePath & "\textures\night\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
End If
Rem ****************************
Rem ********** USA2 **************************
If FileExists(USA2Path & "Textures\" & Shapefile) And FileExists(RoutePath & "\textures\" & Shapefile) Then
booSum(1) = True
If SF1 = False Then
If GetCRC(USA2Path & "Textures\" & Shapefile) = GetCRC(RoutePath & "\textures\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\Textures\" & Shapefile & ChrW$(34) & " .\Textures\ /y"
Print #12, strBat
strBat = vbNullString
SF1 = True
KillEm(KE) = RoutePath & "\textures\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA2Path & "Textures\Snow\" & Shapefile) And FileExists(RoutePath & "\textures\snow\" & Shapefile) And SF2 = False Then
booSnow(1) = True
If GetCRC(USA2Path & "Textures\Snow\" & Shapefile) = GetCRC(RoutePath & "\textures\snow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\Textures\snow\" & Shapefile & ChrW$(34) & " .\Textures\snow\ /y"
Print #12, strBat
strBat = vbNullString
SF2 = True
KillEm(KE) = RoutePath & "\textures\snow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA2Path & "Textures\Spring\" & Shapefile) And FileExists(RoutePath & "\textures\spring\" & Shapefile) And SF3 = False Then
booSpr(1) = True
If GetCRC(USA2Path & "Textures\Spring\" & Shapefile) = GetCRC(RoutePath & "\textures\spring\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\Textures\spring\" & Shapefile & ChrW$(34) & " .\Textures\spring\ /y"
Print #12, strBat
strBat = vbNullString
SF3 = True
KillEm(KE) = RoutePath & "\textures\spring\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA2Path & "Textures\springSnow\" & Shapefile) And FileExists(RoutePath & "\textures\springsnow\" & Shapefile) And SF4 = False Then
booSprS(1) = True
If GetCRC(USA2Path & "Textures\springSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\springsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\Textures\springsnow\" & Shapefile & ChrW$(34) & " .\Textures\springsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF4 = True
KillEm(KE) = RoutePath & "\textures\springsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA2Path & "Textures\Winter\" & Shapefile) And FileExists(RoutePath & "\textures\Winter\" & Shapefile) And SF5 = False Then
booWin(1) = True
If GetCRC(USA2Path & "Textures\Winter\" & Shapefile) = GetCRC(RoutePath & "\textures\Winter\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\Textures\Winter\" & Shapefile & ChrW$(34) & " .\Textures\Winter\ /y"
Print #12, strBat
strBat = vbNullString
SF5 = True
KillEm(KE) = RoutePath & "\textures\winter\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA2Path & "Textures\WinterSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Wintersnow\" & Shapefile) And SF6 = False Then
booWinS(1) = True
If GetCRC(USA2Path & "Textures\WinterSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Wintersnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\Textures\Wintersnow\" & Shapefile & ChrW$(34) & " .\Textures\Wintersnow\ /y"
Print #12, strBat
strBat = vbNullString
SF6 = True
KillEm(KE) = RoutePath & "\textures\wintersnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA2Path & "Textures\Autumn\" & Shapefile) And FileExists(RoutePath & "\textures\Autumn\" & Shapefile) And SF7 = False Then
booAut(1) = True
If GetCRC(USA2Path & "Textures\Autumn\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumn\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\Textures\Autumn\" & Shapefile & ChrW$(34) & " .\Textures\Autumn\ /y"
Print #12, strBat
strBat = vbNullString
SF7 = True
KillEm(KE) = RoutePath & "\textures\Autumn\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA2Path & "Textures\AutumnSnow\" & Shapefile) And FileExists(RoutePath & "\textures\Autumnsnow\" & Shapefile) And SF8 = False Then
booAutS(1) = True
If GetCRC(USA2Path & "Textures\AutumnSnow\" & Shapefile) = GetCRC(RoutePath & "\textures\Autumnsnow\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\Textures\Autumnsnow\" & Shapefile & ChrW$(34) & " .\Textures\Autumnsnow\ /y"
Print #12, strBat
strBat = vbNullString
SF8 = True
KillEm(KE) = RoutePath & "\textures\Autumnsnow\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
If FileExists(USA2Path & "Textures\Night\" & Shapefile) And FileExists(RoutePath & "\textures\Night\" & Shapefile) And SF9 = False Then
booNight(1) = True
If GetCRC(USA2Path & "Textures\Night\" & Shapefile) = GetCRC(RoutePath & "\textures\Night\" & Shapefile) Then
strBat = "call xcopy " & ChrW$(34) & "..\USA2\Textures\Night\" & Shapefile & ChrW$(34) & " .\Textures\Night\ /y"
Print #12, strBat
strBat = vbNullString
SF9 = True
KillEm(KE) = RoutePath & "\textures\night\" & Shapefile
KE = KE + 1
If KE > UBound(KillEm) Then
ReDim Preserve KillEm(0 To KE + REF_CHUNK)
End If
End If
End If
End If
Rem ****************************
If SF1 = False And intAceResponse <> 3 And intAceResponse <> 4 Then
  For i = 1 To 6
    
    If booSum(i) = True Then
        booExact = True
        strResponse = Shapefile & Lang(412) & vbCrLf & Lang(420) & strRoute(i) & vbCrLf & Lang(421)
        frmDialog3.Label1 = strResponse
        'frmDialog.Command1.Visible = False
        If intAceResponse <> 4 Then
        frmDialog3.Show 1
        End If
     DoEvents
     
        booExact = False
        'Select Case MsgBox(Lang(419) & ShapeFile & Lang(412) & vbCrLf & Lang(420) & strRoute(i) & Lang(421), vbYesNo + vbExclamation + vbDefaultButton1, App.Title)
    'Case vbYes
   If intAceResponse = 2 Or intAceResponse = 4 Then
    If i = 1 Then
    Rem ..\..\
    strBat = "call xcopy " & ChrW$(34) & "..\" & strRoute(i) & "\Textures\" & Shapefile & ChrW$(34) & " .\Textures\ /s /y"
    Else
    strBat = "call xcopy " & ChrW$(34) & "..\" & strRoute(i) & "\Textures\" & Shapefile & ChrW$(34) & " .\Textures\ /s /y"
    End If
    Print #12, strBat
    strBat = vbNullString
    ShapeFound = True
    KillEm(KE) = RoutePath & "\textures\" & Shapefile
    KE = KE + 1
        If KE > UBound(KillEm) Then
        ReDim Preserve KillEm(0 To KE + REF_CHUNK)
        End If
        Exit For
    Else
    Exit For
    End If
   End If
 Next i
End If







EndIt:
Close #12

Rem ****************************************

Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'AceAlreadyExists2' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub

Function GetCRC(strFileName As String) As String
Dim strTemp As String

 Set m_CRC = New clsCRC
 m_CRC.Algorithm = 1
 strTemp = Hex(m_CRC.CalculateFile(strFileName))
 GetCRC = strTemp
 
End Function


Public Sub CheckForWav4(SFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Y As Integer, Z As Integer
Dim GlobalSoundPath As String, itExists As Boolean
Dim j As Integer

On Error GoTo Errtrap
If Not FileExists(SFilepath) Then
Call MsgBox(SFilepath & " Was not found.", vbExclamation, App.Title)
Exit Sub
End If
GlobalSoundPath = MSTSPath & "\Sound\"
Fnumber = FreeFile


Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 
   x = InStr(strNew, ".wav")
   
   If x > 0 Then

   Y = InStrRev(strNew, " ", x)
   Z = InStrRev(strNew, ChrW$(34), x)
    If Z > 0 Then
   strNew = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   Else
   strNew = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   
   End If
   
   strNew = Trim$(strNew)
    If Left$(strNew, 1) = ChrW$(34) Then
     strNew = Right$(strNew, Len(strNew) - 1)
    End If
    If Right$(strNew, 1) = ChrW$(34) Then
    strNew = Left$(strNew, Len(strNew) - 1)
    End If

         For j = 0 To numWave
   If strNew = strWave(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   numWave = numWave + 1
   strWave(numWave) = strNew
   End If
   itExists = False
   
   
  
   strNew = vbNullString
    End If
   Loop
   
   Close #Fnumber

   Exit Sub
Errtrap:


End Sub

Public Sub CheckForWav3(SFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Y As Integer, Z As Integer
Dim GlobalSoundPath As String


On Error GoTo Errtrap
If Not FileExists(SFilepath) Then
Call MsgBox(SFilepath & " Was not found.", vbExclamation, App.Title)
Exit Sub
End If
GlobalSoundPath = MSTSPath & "\Sound\"
Fnumber = FreeFile


Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 
   x = InStr(strNew, ".wav")
   
   If x > 0 Then

   Y = InStrRev(strNew, " ", x)
   Z = InStrRev(strNew, ChrW$(34), x)
    If Z > 0 Then
   strNew = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   Else
   strNew = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   
   End If
   
   strNew = Trim$(strNew)
    If Left$(strNew, 1) = ChrW$(34) Then
     strNew = Right$(strNew, Len(strNew) - 1)
    End If
    If Right$(strNew, 1) = ChrW$(34) Then
    strNew = Left$(strNew, Len(strNew) - 1)
    End If
   
     If Not FileExists(SoundPath & "\" & strNew) And Not FileExists(GlobalSoundPath & strNew) Then
    Call LookForSound(strNew, SFilepath)
    End If
  
   strNew = vbNullString
    End If
   Loop
   
   Close #Fnumber

   Exit Sub
Errtrap:

End Sub


Public Sub CheckForWav(SFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Y As Integer, j As Integer, Z As Integer
On Error GoTo Errtrap


Fnumber = FreeFile


Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)

   x = InStr(strNew, ".wav")
   
   If x > 0 Then

   Y = InStrRev(strNew, " ", x)
   Z = InStrRev(strNew, ChrW$(34), x)
   If Z > 0 Then
   strNew = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   Else
   strNew = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   
   End If
   
   strNew = Trim$(strNew)
   If Left$(strNew, 1) = ChrW$(34) Then
  
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   
   
   For j = 0 To WavNumber
   If strNew = WavFile(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   WavNumber = WavNumber + 1
   WavFile(WavNumber) = strNew
   End If
   End If
   strNew = vbNullString
  itExists = False
   Loop
   
   Close #Fnumber
   Exit Sub
Errtrap:

   If Err = 53 Then
   
   MsgBox Lang(342) & SFilepath & Lang(343) & vbCr & Lang(344), 48, Lang(345)
'********************
   End If

End Sub


Public Sub CheckForSMS(SFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Y As Integer, Z As Integer
Dim strLoad As String, itExists As Boolean
Dim j As Integer, strSound As String

On Error GoTo Errtrap
Fnumber = FreeFile

j = InStrRev(SFilepath, "\")
strSound = Left$(SFilepath, j - 1)
If FileExists(strSound & "\Sound\clear_ex.sms") Then
  For j = 0 To SoundNumber
   If "clear_ex.sms" = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
    If itExists = False Then
        SoundNumber = SoundNumber + 1
        Soundfile(SoundNumber) = "clear_ex.sms"
    End If
End If
If FileExists(strSound & "\Sound\clear_in.sms") Then
  For j = 0 To SoundNumber
   If "clear_in.sms" = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
    If itExists = False Then
        SoundNumber = SoundNumber + 1
        Soundfile(SoundNumber) = "clear_in.sms"
    End If
End If
If FileExists(strSound & "\Sound\rain_ex.sms") Then
  For j = 0 To SoundNumber
   If "rain_ex.sms" = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
    If itExists = False Then
        SoundNumber = SoundNumber + 1
        Soundfile(SoundNumber) = "rain_ex.sms"
    End If
End If
If FileExists(strSound & "\Sound\rain_in.sms") Then
  For j = 0 To SoundNumber
   If "rain_in.sms" = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
    If itExists = False Then
        SoundNumber = SoundNumber + 1
        Soundfile(SoundNumber) = "rain_in.sms"
    End If
End If
If FileExists(strSound & "\Sound\intro.sms") Then

  For j = 0 To SoundNumber
   If "intro.sms" = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
    If itExists = False Then
        SoundNumber = SoundNumber + 1
        Soundfile(SoundNumber) = "intro.sms"
    End If
End If

Open SFilepath For Input As Fnumber

   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
   
   x = InStr(strNew, "defaultcrossingsms")
   If x > 0 Then
   x = InStr(strNew, ".sms")
   If x > 0 Then
   Y = InStrRev(strNew, " ", x)
   Z = InStrRev(strNew, ChrW$(34), x)
        If Y > Z Then
        strNew = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
        Else
        strNew = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
        End If
   strNew = Trim$(strNew)
        If Left$(strNew, 1) = ChrW$(34) Then
        strNew = Right$(strNew, Len(strNew) - 1)
        End If
        If Right$(strNew, 1) = ChrW$(34) Then
        strNew = Left$(strNew, Len(strNew) - 1)
        End If
   DefCross = strNew
   End If
   End If
   x = InStr(strNew, "defaultsignalsms")
   If x > 0 Then
   x = InStr(strNew, ".sms")
   If x > 0 Then
   Y = InStrRev(strNew, " ", x)
   Z = InStrRev(strNew, ChrW$(34), x)
   If Y > Z Then
   strNew = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   Else
   strNew = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   End If
   strNew = Trim$(strNew)
   If Left$(strNew, 1) = ChrW$(34) Then
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   DefSig = strNew
   End If
   End If
   x = InStr(strNew, "defaultwatertowersms")
   If x > 0 Then
   x = InStr(strNew, ".sms")
   If x > 0 Then
   Y = InStrRev(strNew, " ", x)
   Z = InStrRev(strNew, ChrW$(34), x)
   If Y > Z Then
   strNew = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   Else
   strNew = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   End If
   strNew = Trim$(strNew)
   If Left$(strNew, 1) = ChrW$(34) Then
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   DefWat = strNew
   End If
   End If
  x = InStr(strNew, "defaultcoaltowersms")
   If x > 0 Then
   x = InStr(strNew, ".sms")
   If x > 0 Then
   Y = InStrRev(strNew, " ", x)
   Z = InStrRev(strNew, ChrW$(34), x)
   If Y > Z Then
   strNew = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   Else
   strNew = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   End If
   strNew = Trim$(strNew)
   If Left$(strNew, 1) = ChrW$(34) Then
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   DefCoal = strNew
   End If
   End If
   x = InStr(strNew, "defaultdieseltowersms")
   If x > 0 Then
   x = InStr(strNew, ".sms")
   If x > 0 Then
   Y = InStrRev(strNew, " ", x)
   Z = InStrRev(strNew, ChrW$(34), x)
   If Y > Z Then
   strNew = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   Else
   strNew = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   End If
   strNew = Trim$(strNew)
   If Left$(strNew, 1) = ChrW$(34) Then
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   DefDies = strNew
   End If
   End If
   x = InStr(strNew, "Graphic (")
   If x > 0 Then
   Y = InStrRev(strNew, ")")
   strGraphic = Mid$(strNew, x + 9, Y - (x + 9))
   strGraphic = Trim$(strGraphic)
   If Left$(strGraphic, 1) = ChrW$(34) Then
   strGraphic = Mid$(strGraphic, 2)
   End If
   If Right$(strGraphic, 1) = ChrW$(34) Then
   strGraphic = Left$(strGraphic, Len(strGraphic) - 1)
   End If
   strGraphic = Trim$(strGraphic)
  
   If Not FileExists(RoutePath & "\" & strGraphic) Then
   
   Call MsgBox(Lang(422) & strGraphic & Lang(423) & vbCrLf & Lang(424), vbExclamation, App.Title)
   
   End If
   End If
   
   x = InStr(strNew, "LoadingScreen (")
   If x > 0 Then
   Y = InStr(x, strNew, ")")
   strLoad = Mid$(strNew, x + 15, Y - (x + 15))
   strLoad = Trim$(strLoad)
   If Left$(strLoad, 1) = ChrW$(34) Then
   strLoad = Mid$(strLoad, 2)
   End If
   If Right$(strLoad, 1) = ChrW$(34) Then
   strLoad = Left$(strLoad, Len(strLoad) - 1)
   End If
   strLoad = Trim$(strLoad)
   If Not FileExists(RoutePath & "\" & strLoad) And Not FileExists(MSTSPath & "\gui\screens\load.ace") Then
   
   Call MsgBox(Lang(424) & strLoad & Lang(423) & vbCrLf & Lang(426), vbExclamation, App.Title)
   
   End If
   End If
   
   
   
   
   
   
   Loop
 
   Close #Fnumber
   Exit Sub
Errtrap:

   If Err = 53 Then
      
     MsgBox Lang(342) & SFilepath & Lang(343) & vbCr & Lang(344), 48, Lang(345)
'********************
   End If

   If Err = 76 Then
   Exit Sub
   End If

End Sub


Public Sub CheckForSounds3(SFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Y As Integer, Z As Integer, xx As Integer
Dim GlobalSoundPath As String, strTemp As String

On Error GoTo Errtrap
GlobalSoundPath = MSTSPath & "\Sound\"
Fnumber = FreeFile
xx = 1

Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
   strTemp = strNew
   
   
SecondTry:
   x = InStr(xx, strTemp, ".sms")
   
   If x > 0 Then
 
   Y = InStrRev(strTemp, " ", x)
   Z = InStrRev(strTemp, ChrW$(34), x)
   If Z > 0 Then
   strNew = Mid$(strTemp, Z + 1, (x + 4) - (Z + 1))
   Else
   strNew = Mid$(strTemp, Y + 1, (x + 4) - (Y + 1))
   
   End If
   
   strNew = Trim$(strNew)
   
   If Left$(strNew, 1) = ChrW$(34) Then
  
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   
   Rem **********************
      For j = 0 To SoundNumber
   If strNew = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   SoundNumber = SoundNumber + 1
   Soundfile(SoundNumber) = strNew
   End If
   
   Rem ***************************

 If Not FileExists(SoundPath & "\" & strNew) And Not FileExists(GlobalSoundPath & strNew) Then

  Call LookForSound(strNew, SFilepath)
  End If
  itExists = False
  strNew = vbNullString
  xx = x + 1
 GoTo SecondTry
  End If
   strNew = vbNullString
   strTemp = vbNullString
  itExists = False
 xx = 1
 
   Loop
   
   Close #Fnumber
   

   
   
   Exit Sub
Errtrap:

   If Err = 53 Then
      
    MsgBox Lang(342) & SFilepath & Lang(343) & vbCr & Lang(344), 48, Lang(345)
'********************
   End If

   If Err = 76 Then
   Exit Sub
   End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : CheckForSounds
' DateTime  : 30/11/2006 14:25
' Author    : Mike
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CheckForSounds(SFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Y As Integer, j As Integer, Z As Integer
Dim strTemp As String, xx As Integer

On Error GoTo Errtrap
Fnumber = FreeFile
xx = 1

Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
   strTemp = strNew
   
   
SecondTry:
   x = InStr(xx, strTemp, ".sms")
   
   If x > 0 Then
   Y = InStrRev(strTemp, " ", x)
   Z = InStrRev(strTemp, ChrW$(34), x)
   If Z > 0 Then
   strNew = Mid$(strTemp, Z + 1, (x + 4) - (Z + 1))
   Else
   strNew = Mid$(strTemp, Y + 1, (x + 4) - (Y + 1))
   
   End If
   
   strNew = Trim$(strNew)
   
   If Left$(strNew, 1) = ChrW$(34) Then
  
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   
   
   For j = 0 To SoundNumber
   If strNew = Soundfile(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   SoundNumber = SoundNumber + 1
   Soundfile(SoundNumber) = strNew
   End If
   
    itExists = False
  strNew = vbNullString
  xx = x + 1
 GoTo SecondTry
  End If
   strNew = vbNullString
   strTemp = vbNullString
  itExists = False
 xx = 1
 
   Loop
   
   Close #Fnumber
   

   
   Exit Sub
Errtrap:

   If Err = 53 Then
      
    MsgBox Lang(342) & SFilepath & Lang(343) & vbCr & Lang(344), 48, Lang(345)
'********************
   End If

   If Err = 76 Then
   Exit Sub
   End If

End Sub


Public Sub CheckForHazAce(SFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Y As Integer, j As Integer, Z As Integer

On Error GoTo Errtrap

Fnumber = FreeFile


Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 
   x = InStr(strNew, ".ace")
   
   If x > 0 Then

   Y = InStrRev(strNew, "(", x)
  Z = InStrRev(strNew, ChrW$(34), x)
   If Z > Y Then
   strNew = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   Else
   strNew = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   
   End If
   
   strNew = Trim$(strNew)
   If Left$(strNew, 1) = ChrW$(34) Then
  
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   
   
   For j = 0 To hazAceNumber
   If strNew = HazAce(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   hazAceNumber = hazAceNumber + 1
   HazAce(hazAceNumber) = strNew
   End If
   End If
   strNew = vbNullString
  itExists = False
   Loop
   
   Close #Fnumber

   Exit Sub
Errtrap:

      If Err = 53 Then
         
     MsgBox Lang(342) & SFilepath & Lang(343) & vbCr & Lang(344), 48, Lang(345)
'********************
   End If


End Sub


Public Sub CheckEnvForAce(SFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Y As Integer, j As Integer, Z As Integer

On Error GoTo Errtrap

Fnumber = FreeFile


Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 
   x = InStr(strNew, ".ace")
   
   If x > 0 Then
 
   Y = InStrRev(strNew, "(", x)
  Z = InStrRev(strNew, ChrW$(34), x)
   If Z > Y Then
   strNew = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   Else
   strNew = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   
   End If
   
   strNew = Trim$(strNew)
   If Left$(strNew, 1) = ChrW$(34) Then
  
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   
   
   For j = 1 To envAceNumber
   If strNew = EnvAceFile(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   envAceNumber = envAceNumber + 1
   EnvAceFile(envAceNumber) = strNew
   End If
   End If
   strNew = vbNullString
  itExists = False
   Loop
   
   Close #Fnumber
  
   Exit Sub
Errtrap:

      If Err = 53 Then
         
     MsgBox Lang(342) & SFilepath & Lang(343) & vbCr & Lang(344), 48, Lang(345)
'********************
   End If


End Sub



Public Sub CheckForAce3(SFilepath As String)
Dim Fnumber As Integer, strNew As String, strTemp As String, j As Integer

Dim x As Integer, Y As Integer, Z As Integer
Dim flagACE As Integer, strS As String

On Error GoTo Errtrap

Fnumber = FreeFile
x = InStrRev(SFilepath, "\")
strS = Mid$(SFilepath, x + 1)
flagACE = 1
Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
CarryON:
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 j = 1
TryAgain:
   x = InStr(j, strNew, ".ace")
   
   If x > 0 Then

   Y = InStrRev(strNew, "(", x)
  Z = InStrRev(strNew, ChrW$(34), x)
   If Z > Y Then
   strTemp = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   Else
   strTemp = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   
   End If
   
   strTemp = Trim$(strTemp)
   If Left$(strTemp, 1) = ChrW$(34) Then
  
   strTemp = Right$(strTemp, Len(strTemp) - 1)
   End If
   If Right$(strTemp, 1) = ChrW$(34) Then
   strTemp = Left$(strTemp, Len(strTemp) - 1)
   End If
   
   If Not FileExists(TexturePath & "\" & strTemp) Then
      Call LookForACE(strTemp, flagACE, strS)
   End If
   End If
   
     strTemp = vbNullString
  itExists = False
  j = x + 5
  If j > Len(strNew) Or x = 0 Then GoTo CarryON
  GoTo TryAgain
   Loop
   
   Close #Fnumber
   Exit Sub
Errtrap:

      If Err = 53 Then
         
     MsgBox Lang(342) & SFilepath & Lang(343) & vbCr & Lang(344), 48, Lang(345)
'********************
   End If


End Sub




Public Sub CompactAceSeasons(strSpare As String, booCar As Boolean)
Dim strNew As String
Dim x As Integer
Dim strSDFile As String, GlobalPath As String
Dim strS As String
Dim strSPath As String
 

On Error GoTo Errtrap
TryAgain:
   
   x = InStrRev(strSpare, "\")
   strSPath = Left$(strSpare, x - 1)
strSDFile = Mid$(strSpare, x + 1)
strSDFile = strSDFile & "d"
GlobalPath = MSTSPath & "\Global\Shapes\"
   
   strS = Mid$(strSpare, x + 1)
   SB1.Panels(2).Text = strS
   
   
 MyString = ReadUniFile(strSpare)

      yy = 1
 Do
 
 yy = InStr(yy, MyString, "image (")
 If yy > 0 Then
 Z = InStr(yy, MyString, "(")
 zz = InStr(Z, MyString, ")")
 strFName = Mid$(MyString, Z + 1, zz - (Z + 1))
 strFName = Trim$(strFName)
 If Left$(strFName, 1) = ChrW$(34) Then
 strFName = Mid$(strFName, 2)
 End If
 If Right$(strFName, 1) = ChrW$(34) Then
 strFName = Left$(strFName, Len(strFName) - 1)
 End If

    
   strNew = strFName
             
              Call CompactAllAce(strNew, strSDFile, booCar)
                           
                    yy = zz
                    End If

    
    Loop While yy > 0
   
 DoEvents
 
   Exit Sub
Errtrap:

 
      If Err = 53 Then
         
     MsgBox Lang(342) & SFilepath & Lang(343) & vbCr & Lang(344), 48, Lang(345)
'********************

Exit Sub
 End If
 
Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'CompactACESeasons' please advise" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
    'Resume Next
   
    Exit Sub
    End Select

End Sub




Public Sub CompactCheckForAce(SFilepath As String)
Dim Fnumber As Integer, strNew As String, strTemp As String
Dim x As Integer, Y As Integer, Z As Integer, j As Integer

On Error GoTo Errtrap

Fnumber = FreeFile


Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
CarryON:
If EOF(Fnumber) = True Then Exit Do

   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 j = 1
TryAgain:

   x = InStr(j, strNew, ".ace")
   
   If x > 0 Then

   Y = InStrRev(strNew, "(", x)
  Z = InStrRev(strNew, ChrW$(34), x)
   If Z > Y Then
   strTemp = Mid$(strNew, Z + 1, (x + 4) - (Z + 1))
   Else
   strTemp = Mid$(strNew, Y + 1, (x + 4) - (Y + 1))
   
   End If
   
   strTemp = Trim$(strTemp)
   If Left$(strTemp, 1) = ChrW$(34) Then
  
   strTemp = Right$(strTemp, Len(strTemp) - 1)
   End If
   If Right$(strTemp, 1) = ChrW$(34) Then
   strNew = Left$(strTemp, Len(strTemp) - 1)
   End If
 For Y = 0 To numMisc
 
   If strTemp = MiscAce(Y) Then
   itExists = True
   Exit For
   End If
   Next Y
   If itExists = False Then
   MiscAce(numMisc) = strTemp
   numMisc = numMisc + 1
   End If

   End If
   strTemp = vbNullString
  itExists = False
  j = x + 5
  If j > Len(strNew) Or x = 0 Then GoTo CarryON
  GoTo TryAgain
  
   Loop
   
   Close #Fnumber
   
   Exit Sub
Errtrap:


Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'CompactCheckForAce' please advise" _
                       & vbCrLf & "Support with details of operation being processed." _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        
        Case vbRetry
    Resume Next
        Case vbCancel
       'Resume Next
    Exit Sub
    End Select


End Sub































Private Sub CheckForS3(SFilepath As String)
Dim Fnumber As Integer, strNew As String
Dim x As Integer, Y As Integer, strS As String

On Error GoTo Errtrap

Fnumber = FreeFile

x = InStrRev(SFilepath, "\")
strS = Mid$(SFilepath, x + 1)
Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 
   x = InStr(strNew, ".s")
   
   If x > 0 Then
 
   Y = InStrRev(strNew, "(", x)
   Z = InStrRev(strNew, ChrW$(34), x)
   If Y > Z Then
   
   strNew = Mid$(strNew, Y + 1, (x + 2) - (Y + 1))
   Else
   strNew = Mid$(strNew, Z + 1, (x + 2) - (Z + 1))
   End If
   
  
   strNew = Trim$(strNew)
   If Left$(strNew, 1) = ChrW$(34) Then
  
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   strShp(numShp) = strNew
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(1 To numShp + Shp_Chunk)
  
   End If
 End If
   strNew = vbNullString
  
   Loop
   Close #Fnumber
Exit Sub
Errtrap:

   If Err = 53 Then
     
    MsgBox Lang(342) & SFilepath & Lang(343) & vbCr & Lang(344), 48, Lang(345)
'********************
   End If

End Sub

Private Sub CompactCheckForS(SFilepath As String, booHaz As Boolean)
Dim Fnumber As Integer, strNew As String, strTemp As String
Dim x As Integer, Y As Integer, j As Integer

On Error GoTo Errtrap

Fnumber = FreeFile
x = InStrRev(SFilepath, "\")
strTemp = Mid$(SFilepath, x + 1)

Open SFilepath For Input As Fnumber
   
   Do While Not EOF(Fnumber)
   
   Line Input #Fnumber, strNew
   strNew = Trim$(strNew)
 
   x = InStr(strNew, ".s")
   
   If x > 0 Then
 
   Y = InStrRev(strNew, "(", x)
   Z = InStrRev(strNew, ChrW$(34), x)
   If Y > Z Then
   
   strNew = Mid$(strNew, Y + 1, (x + 2) - (Y + 1))
   Else
   strNew = Mid$(strNew, Z + 1, (x + 2) - (Z + 1))
   End If
   
  
   strNew = Trim$(strNew)
   If Left$(strNew, 1) = ChrW$(34) Then
  
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
  
   If booHaz = True Then
   
    For j = 0 To hazNumber
   If strNew = Hazard(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
   
   Hazard(hazNumber) = strNew
   HazName(hazNumber) = strTemp
   hazNumber = hazNumber + 1
   End If
   Else
   For j = 0 To intCars
   If strNew = Cars(j) Then
   itExists = True
   Exit For
   End If
  Next j
   If itExists = False Then
      Cars(intCars) = strNew
   intCars = intCars + 1
   If intCars > UBound(Cars) Then
     ReDim Preserve Cars(0 To intCars + Car_CHUNK)
    End If
   
   End If
   End If
   End If
   strNew = vbNullString
  itExists = False
   Loop
   Close #Fnumber

Exit Sub
Errtrap:

Select Case MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'CompactCheckforS' while checking" _
                       & vbCrLf & SFilepath _
                       , vbRetryCancel Or vbExclamation Or vbDefaultButton1, App.Title)
    
        Case vbRetry
    Resume Next
        Case vbCancel
      ' Resume Next
    Exit Sub
    End Select

End Sub

Private Sub FindRouteID(RouteName As String)
Dim NewFile As Integer, strNew As String, x As Integer, strFolderName As String, booExists As Boolean
Dim Y As Integer, yy As Integer

On Error GoTo Errtrap
NewFile = FreeFile
If Not FileExists(RouteName) Then
booExists = False
Exit Sub
End If
x = InStrRev(RoutePath, "\")
strFolderName = Mid$(RoutePath, x + 1)


Open RouteName For Input As #NewFile
Do While Not EOF(NewFile)
  
   Line Input #NewFile, strNew
  
   strNew = Trim$(strNew)
   x = InStr(strNew, "RouteID")
  
   If x > 0 Then

   Y = InStr(strNew, "(")
   yy = InStrRev(strNew, ")")
   strNew = Mid$(strNew, Y + 1, yy - (Y + 1))
   strNew = Trim$(strNew)
   
 
   If Left$(strNew, 1) = ChrW$(34) Then
  
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
'   If strFolderName <> strNew Then
'   Call MsgBox(Lang(427) & strNew & Lang(428) & vbCrLf & strFolderName & Lang(429), vbExclamation, Lang(430))
'
'   End If
   Exit Sub
   End If
   Loop
   Close #NewFile
   
   booExists = False
   Exit Sub
   
Errtrap:
   If Err = 76 Then
   Call MsgBox(Lang(431) & vbCrLf & Lang(432), vbExclamation, Lang(407))
   End If
End Sub


Private Sub FindRouteInternalName(RouteName As String, NewRoute As String, booExists As Boolean)
Dim NewFile As Integer, strNew As String
On Error GoTo Errtrap
NewFile = FreeFile
If Not FileExists(RouteName) Then
booExists = False
Exit Sub
End If
Open RouteName For Input As #NewFile
Do While Not EOF(NewFile)
   
   Line Input #NewFile, strNew
  
   strNew = Trim$(strNew)
   x = InStr(strNew, "Name")
  
   If x > 0 Then

   Y = InStr(strNew, "(")
   yy = InStr(strNew, ")")
   strNew = Mid$(strNew, Y + 1, yy - (Y + 1))
   strNew = Trim$(strNew)
   
 
   If Left$(strNew, 1) = ChrW$(34) Then
  
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   NewRoute = strNew
   booExists = True
   Exit Sub
   End If
   Loop
   Close #NewFile
   
   booExists = False
   
   Exit Sub
   
Errtrap:
   If Err = 76 Then
   Call MsgBox(Lang(431) & vbCrLf & Lang(432), vbExclamation, Lang(407))
   
   End If
End Sub




Private Sub FindRouteName(RouteName As String, NewRoute As String, booExists As Boolean)
Dim NewFile As Integer, strNew As String, x As Integer, Y As Integer, yy As Integer
On Error GoTo Errtrap
NewFile = FreeFile
If Not FileExists(RouteName) Then
booExists = False
Exit Sub
End If
Open RouteName For Input As #NewFile
Do While Not EOF(NewFile)
   
   Line Input #NewFile, strNew
  
   strNew = Trim$(strNew)
   x = InStr(strNew, "FileName")
  
   If x > 0 Then

   Y = InStr(strNew, "(")
   yy = InStrRev(strNew, ")")
   strNew = Mid$(strNew, Y + 1, yy - (Y + 1))
   strNew = Trim$(strNew)
   
 
   If Left$(strNew, 1) = ChrW$(34) Then
  
   strNew = Right$(strNew, Len(strNew) - 1)
   End If
   If Right$(strNew, 1) = ChrW$(34) Then
   strNew = Left$(strNew, Len(strNew) - 1)
   End If
   NewRoute = strNew
   booExists = True
   Exit Sub
   End If
   Loop
   Close #NewFile
   
   booExists = False
   
   Exit Sub
   
Errtrap:
   If Err = 76 Then
   Call MsgBox(Lang(431) & vbCrLf & Lang(432), vbExclamation, Lang(407))
   
   End If
End Sub








Private Function DirDiverLoco2(NewPath As String, DirCount As Integer, Backup As String) As Integer
'  Recursively search directories from NewPath down...
'  NewPath is searched on this recursion.
'  BackUp is origin of this recursion.
'  DirCount is number of subdirectories in this directory.

Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, x As Integer, xx As Integer
Dim retval As Integer

    SearchFlag = True           ' Set flag so the user can interrupt.
    DirDiverLoco2 = False            ' Set to True if there is an error.
    retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    If SearchFlag = False Then
        DirDiverLoco2 = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = Dir1(0).ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = Dir1(0).Path                    ' Save old path for next recursion.
        Dir1(0).Path = NewPath
        If Dir1(0).ListCount > 0 Then
        
            ' Get to the node bottom.
            Dir1(0).Path = Dir1(0).List(DirsToPeek - 1)
            AbandonSearch = DirDiverLoco2((Dir1(0).Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    
    ' Call function to enumerate files.
    If File1(0).ListCount Then
        If Len(Dir1(0).Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = Dir1(0).Path                  ' If at root level, leave as is...
        Else
            ThePath = Dir1(0).Path + "\"            ' Otherwise put "\" before the filename.
        End If
        

        x = InStrRev(ThePath, "\", Len(ThePath) - 1)
        xx = InStrRev(ThePath, "\", x - 1)
        If Mid$(ThePath, xx + 1, 8) <> "Trainset" Then
        
        GoTo SkipThis
        End If
        Call FindFolder(ThePath)
        'Text1(cursouind).Text = "*.eng"
       
    File1(cursouind).Pattern = "*.eng;*.wag"
    DoEvents
        For ind = 0 To File1(0).ListCount - 1
       If Right$(File1(0).List(ind), 3) = "eng" Then
        
        If lngLoco > UBound(Locomotives) Then
           ReDim Preserve Locomotives(0 To lngLoco + CHUNK)
           ReDim Preserve LocoPath(0 To lngLoco + CHUNK)
           ReDim Preserve LocoName(0 To lngLoco + CHUNK)
           ReDim Preserve LocoCoup(0 To lngLoco + CHUNK)
           ReDim Preserve LocoFCoup(0 To lngLoco + CHUNK)
           ReDim Preserve LocoBrake(0 To lngLoco + CHUNK)
           ReDim Preserve LocoType(0 To lngLoco + CHUNK)
           ReDim Preserve LocoRigid(0 To lngLoco + CHUNK)
           ReDim Preserve LocoFRigid(0 To lngLoco + CHUNK)
           ReDim Preserve LocoSMS(0 To lngLoco + CHUNK)
           End If
        
            Locomotives(lngLoco) = File1(0).List(ind)
            LocoPath(lngLoco) = ThePath
            lngLoco = lngLoco + 1
    ElseIf Right$(File1(0).List(ind), 3) = "wag" Then
       
        
        If lngWagons > UBound(Wagons) Then
           ReDim Preserve Wagons(0 To lngWagons + CHUNK)
           ReDim Preserve Wagpath(0 To lngWagons + CHUNK)
           ReDim Preserve WagonName(0 To lngWagons + CHUNK)
           ReDim Preserve WagCoup(0 To lngWagons + CHUNK)
           ReDim Preserve WagFCoup(0 To lngWagons + CHUNK)
           ReDim Preserve WagBrake(0 To lngWagons + CHUNK)
           ReDim Preserve WagType(0 To lngWagons + CHUNK)
           ReDim Preserve WagRigid(0 To lngWagons + CHUNK)
           ReDim Preserve WagFRigid(0 To lngWagons + CHUNK)
         
           End If
        ' Add conforming files in this directory to the list box.
            Wagons(lngWagons) = File1(0).List(ind)
            Wagpath(lngWagons) = ThePath
            lngWagons = lngWagons + 1
            End If
            Next ind
  
SkipThis:
    End If
    If Backup <> vbNullString Then        ' If there is a superior directory, move it.
        Dir1(0).Path = Backup
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiverLoco2 = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    ElseIf Err = 9 Then                      ' Otherwise display error message and quit.
     
     Resume Next
     ElseIf Err = 76 Then
     GoTo SkipThis
'Call MsgBox("File " & NewPath _
'            & vbCrLf & "In subroutine dirdiverloco2 Error=" & Err _
'            , vbExclamation, "An error occurred while processing")
'
'Resume Next
     Else
        MsgBox Error
        End
    End If
End Function

Private Function DirDiver(NewPath As String, DirCount As Integer, Backup As String) As Integer
'  Recursively search directories from NewPath down...
'  NewPath is searched on this recursion.
'  BackUp is origin of this recursion.
'  DirCount is number of subdirectories in this directory.
Dim fhi As BY_HANDLE_FILE_INFORMATION, Links As Long, hFile As Long
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, Entry As String
Dim retval As Integer, x As Integer
Dim buff As OFSTRUCT

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
            AbandonSearch = DirDiver((frmUtils.Dir1(cursouind).Path), DirCount%, OldPath)
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
        x = InStrRev(ThePath, "\", Len(ThePath) - 1)
        'xx = InStrRev(ThePath, "\", X - 1)
        
        If Mid$(ThePath, x + 1, 5) = "World" Or Mid$(ThePath, x + 1, 5) = "Tiles" Or Mid$(ThePath, x + 1, 5) = "Paths" Then
       
        GoTo SkipThis
        End If
        If Mid$(ThePath, x + 1, 10) = "Activities" Or Mid$(ThePath, x + 1, 8) = "Services" Or Mid$(ThePath, x + 1, 7) = "Traffic" Then
       
        GoTo SkipThis
        End If
        For ind = 0 To frmUtils.File1(cursouind).ListCount - 1        ' Add conforming files in this directory to the list box.
            Entry = ThePath + frmUtils.File1(cursouind).List(ind)
            hFile = OpenFile(Entry, buff, OF_READ)
            GetFileInformationByHandle hFile, fhi
            Links = fhi.nNumberOfLinks
            If booLink = True Then
            'If Links = 1 Then
            frmLinks.GridLinks.AddItem Str(Links) & vbTab & Entry
            lLink = lLink + 1
            'End If
            ElseIf booLink = False Then
            frmLinks.GridLinks.AddItem Str(Links) & vbTab & Entry
           lLink = lLink + 1
            End If
            
           ' lblCount(0).Caption = str(Val(lblCount(0).Caption) + 1)
            DoEvents
            If booAbort = True Then
            SearchFlag = False
            Exit Function
            End If
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




Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
                                         '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left$(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left$(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid$(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = vbNullString                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function















Private Sub comAbortFil_Click()
booAbort = True
End Sub

Private Sub Command1_Click(Index As Integer)
Dim MSG$, Response%, i%, Filpath$, Filpath1$
Dim FileCounter As Integer
Dim strFileName As String, x As Integer
Dim strPathName As String, strDir As String

On Error GoTo selerr
Select Case Index
Case 2
Rem delete
If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
MSG = Lang(110) & tit$
Response = MsgBox(MSG, 36, Lang(110))
Select Case Response
Case vbYes
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   
   
   End If
       
End If
Next i
File1(cursouind).Refresh


tit$ = vbNullString
Case vbNo
tit$ = vbNullString
End Select

File1(cursouind).Refresh
Case 7
Rem rename

For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
FileCounter = FileCounter + 1
End If
Next i

If FileCounter = 1 Then
tit4$ = InputBox("Rename " & tit$ & " as ", "Rename ", tit$)
If tit4$ = vbNullString Then Exit Sub
Name tit$ As tit4$
 tit$ = vbNullString: tit4$ = vbNullString
File1(cursouind).Refresh
Exit Sub
End If
If FileCounter > 1 Then
Select Case MsgBox(Lang(433) & vbCrLf & Lang(434), vbOKCancel + vbQuestion + vbDefaultButton1, Lang(435))

    Case vbOK
    x = InStrRev(tit$, "\")
    strFileName = Mid$(tit$, x + 1)
    strPathName = Left$(tit$, x)
    frmRename.Show 1
   
     DoEvents
     
    If strWith = vbNullString And strPrefix = vbNullString And strReplace = vbNullString Then Exit Sub
    If strPrefix <> vbNullString Then

For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   Name File1(cursouind).Path & "\" & File1(cursouind).List(i) As File1(cursouind).Path & "\" & strPrefix & File1(cursouind).List(i)
End If
Next i
End If
If strReplace <> vbNullString Then

For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
 
   Name File1(cursouind).Path & "\" & File1(cursouind).List(i) As Left$(File1(cursouind).Path & "\" & File1(cursouind).List(i), Len(File1(cursouind).Path & "\" & File1(cursouind).List(i)) - Len(strReplace)) & strWith

End If
Next i
End If

    Case vbCancel
Exit Sub
End Select
End If
File1(cursouind).Refresh
Case 8
Rem MakeDir
tit$ = Label2(cursouind).Caption
tit4$ = InputBox("Make new Directory " & tit$, "Make New Directory", tit$)
If tit4$ = vbNullString Then Exit Sub
MkDir tit4$
File1(cursouind).Refresh
Dir1(cursouind).Refresh
Case 9
Rem Remove Dir
tit$ = Dir1(cursouind).Path
tit4$ = InputBox("Remove Directory " & tit$, "Remove Directory", tit$)
If tit4$ = vbNullString Then Exit Sub
RmDir tit4$
File1(cursouind).Refresh
Dir1(cursouind).Refresh
Case 0
Rem All
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True

 
   Next i
   
Case 1
Rem None
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = False
Next i
Label2(cursouind).Caption = vbNullString
tit$ = vbNullString
Case 3
Rem Copy
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
    Filpath$ = File1(cursouind).Path
   If Right$(Filpath$, 1) = "\" Then
   Filpath$ = Left$(Filpath$, Len(Filpath$) - 1)
   End If
Filpath1$ = File1(curtarind).Path
   If Right$(Filpath1$, 1) = "\" Then
   Filpath1$ = Left$(Filpath1$, Len(Filpath1$) - 1)
   End If
  
   If FileExists(Filpath1$ & "\" & File1(cursouind).List(i)) Then
   MSG = "File " & Filpath1$ & "\" & File1(cursouind).List(i) & " already exists" & nl
   MSG = MSG & "Overwrite ?"
   Response = MsgBox(MSG, vbOKCancel)
   End If
   FileCopy Filpath$ & "\" & File1(cursouind).List(i), Filpath1$ & "\" & File1(cursouind).List(i)
   End If
   Next i
File1(cursouind).Refresh
File1(curtarind).Refresh


Case 15
Rem Quit
Call KillSpare2("*.s")
DoEvents

If Command1(15).Caption = Lang(637) Then
booAbort = True
Dim q As Integer
For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then frmUtils.Controls(q).Enabled = True
Next q

Command1(15).Caption = Lang(30)
Exit Sub
Else
Unload Me
End If
Case 16
Rem Reset List


' lstFoundFiles.Clear
   ' lblCount.Caption = 0
    SearchFlag = False                  ' Flag indicating search in progress.
    Dir1(cursouind).Path = CurDir
    Drive1(cursouind).Drive = Dir1(cursouind).Path ' Reset the path.
    Text1(cursouind).Text = "*.*"

Case 11
Rem Parent
Dir1(cursouind).Path = ".."
Dir1(cursouind).Refresh
File1(cursouind).Refresh
Case 12
Rem Root
Dir1(cursouind).Path = "\"
Dir1(cursouind).Refresh
File1(cursouind).Refresh
Case 4
Rem Move
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
    Filpath$ = File1(cursouind).Path
   If Right$(Filpath$, 1) = "\" Then
   Filpath$ = Left$(Filpath$, Len(Filpath$) - 1)
   End If
Filpath1$ = File1(curtarind).Path
   If Right$(Filpath1$, 1) = "\" Then
   Filpath1$ = Left$(Filpath1$, Len(Filpath1$) - 1)
   End If
   If FileExists(Filpath1$ & "\" & File1(cursouind).List(i)) Then
   MSG = "File " & Filpath1$ & "\" & File1(cursouind).List(i) & " already exists" & nl
   MSG = MSG & "Overwrite ?"
   Response = MsgBox(MSG, vbOKCancel)
   End If
   FileCopy Filpath$ & "\" & File1(cursouind).List(i), Filpath1$ & "\" & File1(cursouind).List(i)
   DoEvents
   Kill Filpath$ & "\" & File1(cursouind).List(i)
   End If
   Next i
File1(cursouind).Refresh
File1(curtarind).Refresh




Case 5
'Edit
frmReport.Show 1

     DoEvents
     
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
   numselect = numselect + 1
   End If
Next i
  If numselect > 1 Then
   MsgBox Lang(436), 48, "Viewer Warning"
'********************
   End If
For i = 0 To File1(cursouind).ListCount - 1
   
  
   If File1(cursouind).Selected(i) Then
    If Right$(File1(cursouind).Path, 1) = "\" Then

   fullpath$ = File1(cursouind).Path & File1(cursouind).List(i)
   Else
   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
   
   frmReport.Rich1.LoadFile fullpath$
 
    End If
    File1(cursouind).Selected(i) = False
Next i
Case 6
Rem Properties

For i = 0 To File1(cursouind).ListCount - 1
    
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   fullpath$ = File1(cursouind).Path & File1(cursouind).List(i)
   Else
   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
  
 frmProp.Text1.Text = fullpath$
 frmProp.Label2(1).Caption = FileLen(fullpath$)
 frmProp.Label2(2).Caption = FileDateTime(fullpath$)
 myattr = GetAttr(fullpath$)

 resultr = myattr And vbReadOnly
 resulth = myattr And vbHidden
 results = myattr And vbSystem
 resulta = myattr And vbArchive
 If resultr Then
 frmProp.Check1(0).value = 1
 End If
 If resulth Then
 frmProp.Check1(1).value = 1
 End If
 If results Then
 frmProp.Check1(2).value = 1
 End If
 If resulta Then
 frmProp.Check1(3).value = 1
 End If
 
 frmProp.Label2(3).Caption = Str$(myattr)
 frmProp.Show 1
 
     DoEvents
     
 End If
Next i
Case 13
Rem Execute
tit$ = Label2(cursouind).Caption
retval = Shell(tit$, 1)
Case 14
Rem Hidden/System
If File1(cursouind).Hidden = False Then
File1(cursouind).Hidden = True
File1(cursouind).System = True
File1(cursouind).BackColor = &HFF&
Else
File1(cursouind).Hidden = False
File1(cursouind).System = False
File1(cursouind).BackColor = &H80000005
End If
Case 17
Rem Dir to Text
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
   Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
  strDir = strDir & File1(cursouind).List(i) & vbCrLf
  End If
  Next i
frmReport.Rich1.Text = strDir
frmReport.Show 1

     DoEvents
     
End Select
Exit Sub
selerr:

MSG = "Error " & Error$(Err) & " (" & Str$(Err) & ") occured " & nl
MSG = MSG & "Unable to complete requested procedure."
Response = MsgBox(MSG)
Resume Next
End Sub



Private Sub Command10_Click(Index As Integer)
Dim result As String

'result = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", lang(293))
result = MSTSPath
Text1(Index) = "*.*"
Drive1(Index).Drive = Left$(result, 2)
Dir1(Index).Path = result & "\Routes"
End Sub

Private Sub Command11_Click(Index As Integer)
Dim MSG$, Response%, strSaveFolder As String

On Error GoTo selerr

Select Case Index
Case 0
Rem compress .s/t/w
Select Case MsgBox(Lang(437) & vbCrLf & Lang(438), vbOKCancel + vbInformation + vbDefaultButton1, Lang(439))

    Case vbOK
    Label9.Visible = True
Call CompressFiles
File1(cursouind).Refresh
tit$ = vbNullString
Label9.Visible = False
Close

    Case vbCancel
Exit Sub
End Select
File1(cursouind).Refresh
If strReport = vbNullString Then
strReport = vbCrLf & "Operation concluded successfully, no errors found" & vbCrLf
End If
   
   frmReport.Rich1 = strReport
   frmReport.Show 1
   strReport = vbNullString


Case 3
Rem Make ACE
If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If

strSaveFolder = Label2(1).Caption
MSG = "Make uncompressed ACE file from selected files?"
Response = MsgBox(MSG, 36, "Make ACE")
Select Case Response
Case vbYes
Call MsgBox("ACE file(s) will be saved in " & strSaveFolder, vbExclamation, App.Title)


Call MakeACEFile(strSaveFolder)
File1(cursouind).Refresh
tit$ = vbNullString
Case vbNo
tit$ = vbNullString
End Select
File1(cursouind).Refresh

Case 4
Rem Make Compressed ACE - Zlib
If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
strSaveFolder = Label2(1).Caption
MSG = "Make Compressed ACE file(s) from selected files"
Response = MsgBox(MSG, 36, "Make Compressed ACE")
Select Case Response
Case vbYes
Call MsgBox("ACE file(s) will be saved in " & strSaveFolder, vbExclamation, App.Title)
Call MakeACECompFile(strSaveFolder)
File1(cursouind).Refresh
tit$ = vbNullString
Case vbNo
tit$ = vbNullString
End Select
File1(cursouind).Refresh
Case 5
Rem Make Compressed ACE - DXT
If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
strSaveFolder = Label2(1).Caption
MSG = "Convert Selected .ace files to DXT1 files"
Response = MsgBox(MSG, 36, "Make DXT1")
Select Case Response
Case vbYes

Call MakeACEDXTFile
File1(cursouind).Refresh
tit$ = vbNullString
Case vbNo
tit$ = vbNullString
End Select
File1(cursouind).Refresh
End Select
Exit Sub
selerr:
Call MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'frmUtils-Command11-" & Index & "' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub

Private Sub Command12_Click()
Dim Filpath1$, strBat As String, varBatText As Variant, EnvFile As String, EnvAceFile As String
Dim ShapeFound As Boolean, Shapefile As String, BooEnv As Boolean, Z As Integer, q As Integer
Dim SnowShapeFound As Boolean

For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then
If frmUtils.Controls(q).Caption <> Lang(30) And frmUtils.Controls(q).Caption <> Lang(24) Then
frmUtils.Controls(q).Enabled = False
End If
End If
Next q
Command7.value = True

ReDim KillEm(0 To REF_CHUNK)
intResponse = 0
intResponse2 = 0
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)

GoTo CleanUP
End If
cursouind = 0
MousePointer = 11
Drive1(cursouind).Drive = Left$(RoutePath, 2)
Dir1(cursouind).Path = RoutePath
Text1(cursouind).Text = "*.*"

Dim iMsg As Integer
iMsg = MsgBox(Lang(333) & vbCr & Lang(334), 49, Lang(335))
Select Case iMsg

    Case vbOK
    
MousePointer = 11
    Case vbCancel
    
GoTo CleanUP

End Select


iMsg = MsgBox(Lang(336) & vbCr & Lang(337), 36, Lang(338))
Select Case iMsg

    Case vbYes
    
BooEnv = False
    Case vbNo
    
BooEnv = True
End Select
intResponse = 0
Filpath1$ = App.Path & "\SetupFiles"
cursouind = 0
Drive1(cursouind).Drive = Left$(ShapePath, 2)
Dir1(cursouind).Path = ShapePath
Text1(cursouind).Text = "*.S"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1

If File1(cursouind).Selected(i) = True Then
   If File1(cursouind).List(i) <> vbNullString Then
   Shapefile = File1(cursouind).List(i)
   
   Call ShapeAlreadyExists(Shapefile, ShapeFound)
   End If
   
   
  If ShapeFound = True Then
        If FileExists(File1(cursouind).Path & "\" & File1(cursouind).List(i)) Then
        Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
        End If
  
  TempShape = Left$(File1(cursouind).List(i), Len(File1(cursouind).List(i)) - 1) & "thm"
  If FileExists(ShapePath & "\" & TempShape) Then
 Kill ShapePath & "\" & TempShape
 End If
 ShapeFound = False
 
 If FileExists(ShapePath & "\" & Shapefile & "d") Then
 Call ShapeAlreadyExists(Shapefile & "d", ShapeFound)
 If ShapeFound = True Then
  If FileExists(ShapePath & "\" & Shapefile & "d") Then
 Kill ShapePath & "\" & Shapefile & "d"
 End If
 End If
 End If
 End If
End If
 ShapeFound = False
 Next i
Rem *************Textures*************
Drive1(cursouind).Drive = Left$(TexturePath, 2)
Dir1(cursouind).Path = TexturePath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
  If File1(cursouind).List(i) <> vbNullString Then
   
  Shapefile = File1(cursouind).List(i)
  

 Call AceAlreadyExists2(Shapefile)
     


 End If
 End If
 ShapeFound = False
 Next i
If KE > 0 Then
   ReDim Preserve KillEm(0 To KE - 1)
For i = 0 To KE - 1

If FileExists(KillEm(i)) And KillEm(i) <> vbNullString Then

 Kill KillEm(i)
 DoEvents
 End If
 Next i
 End If
 Rem ***********Terrain Textures ******************
 
 TertexPath = RoutePath & "\Terrtex"
 frmSnow.Show 1
 
     DoEvents
     

 If booSnow = True Then
 Call WriteSnow(strSnowName, varBatText)
 End If

Drive1(cursouind).Drive = Left$(TertexPath, 2)
Dir1(cursouind).Path = TertexPath
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

If File1(cursouind).Selected(i) = True Then
   If File1(cursouind).List(i) <> vbNullString Then
   
   Shapefile = File1(cursouind).List(i)
 
   Call TerrAlreadyExists(Shapefile, ShapeFound, SnowShapeFound)
   
   End If
  
    If ShapeFound = True Then
    
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 If FileExists(TertexPath & "\Snow\" & File1(cursouind).List(i)) And booSnow = False And SnowShapeFound = True Then
 Kill TertexPath & "\Snow\" & File1(cursouind).List(i)
 End If
 End If
 End If
 ShapeFound = False
 Next i
 If booSnow = True Then
 Open App.Path & "\SetupFiles\Installme.bat" For Append As #12
 Print #12, varBatText
 Close #12
 End If
 Rem ************* Transfer Files ******************
  TertexPath = RoutePath & "\Terrtex"
For i = 0 To TransferNumber - 1
   Shapefile = TransferFile(i)
'   If ShapeFile = "terrtex.ace" Then
'   ShapeFound = False
'   Else
   Call TransferExists(Shapefile, ShapeFound)
   'End If
   
    If ShapeFound = True Then
     If FileExists(TexturePath & "\" & Shapefile) Then

 Kill TexturePath & "\" & Shapefile
 End If
 If FileExists(TexturePath & "\Snow\" & Shapefile) Then
 Kill TexturePath & "\Snow\" & Shapefile
 End If
 End If
 
 ShapeFound = False
 Next i
 
 Rem ************* Sound Files *******************
 Drive1(cursouind).Drive = Left$(SoundPath, 2)
Dir1(cursouind).Path = SoundPath
Text1(cursouind).Text = "*.sms"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
   If File1(cursouind).List(i) <> vbNullString Then
   Shapefile = File1(cursouind).List(i)
   

   Call SmsAlreadyExists(Shapefile, ShapeFound)
  
   End If
    If ShapeFound = True Then
Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 End If
 End If
 ShapeFound = False

 Next i
Rem ************* .wav files

 Text1(cursouind).Text = "*.wav"
 DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
   If File1(cursouind).List(i) <> vbNullString Then
   Shapefile = File1(cursouind).List(i)
   Call WavAlreadyExists(Shapefile, ShapeFound)
   End If
    If ShapeFound = True Then
    Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
    End If
End If
 ShapeFound = False

Next i
Rem ********** Hazards
If DirExists(RoutePath & "\Global") Then
Open App.Path & "\SetupFiles\Installme.bat" For Append As #12
strBat = "call xcopy .\Global\*.* ..\..\Global\*.* /s /y"
Print #12, strBat
strBat = vbNullString
Close #12
DoEvents
End If
Rem ************ Env files *******

If BooEnv = True Then
Open App.Path & "\SetupFiles\Installme.bat" For Append As #12
strBat = "call xcopy ..\..\template\envfiles\*.* .\Envfiles\ /s /y"
Print #12, strBat
strBat = vbNullString
Close #12
DoEvents

Call KillEnv
ElseIf BooEnv = False Then    '****************************************************Env...
 Drive1(cursouind).Drive = Left$(EnvPath, 2)
Dir1(cursouind).Path = EnvPath
Text1(cursouind).Text = "*.env"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
Z = 0
For i = 1 To 100
KillEnvBat(i) = vbNullString
Next i
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
   If File1(cursouind).List(i) <> vbNullString Then
   EnvFile = File1(cursouind).List(i)
   
   Call CheckEnv(EnvFile, Z)
    End If
   End If
Next i
If Z > 0 Then
For i = 1 To Z
Kill KillEnvBat(i)
DoEvents
Next i
End If
 Drive1(cursouind).Drive = Left$(EnvPath, 2)
Dir1(cursouind).Path = EnvPath & "\Textures"
Text1(cursouind).Text = "*.ace"
DoEvents
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
Z = 0
For i = 1 To 100
KillEnvBat(i) = vbNullString
Next i
For i = 0 To File1(cursouind).ListCount - 1
If File1(cursouind).Selected(i) = True Then
   If File1(cursouind).List(i) <> vbNullString Then
   EnvAceFile = File1(cursouind).List(i)
   If Left$(EnvAceFile, 4) <> "hitw" Then
   Call CheckEnvAce(EnvAceFile, Z)
   End If
    End If
   End If
Next i
If Z > 0 Then
For i = 1 To Z
Kill KillEnvBat(i)
DoEvents
Next i
End If
End If
 Text1(0).Text = "*.*"
 Text1(1).Text = "*.*"
DoEvents
Rem **************** Second Pass *********************

Call SecondTextures

'*****************************************************

FileCopy App.Path & "\setupfiles\Installme.bat", RoutePath & "\Installme.bat"

CleanUP:
MousePointer = 0
cursouind = 0
Dir1(cursouind).Path = RoutePath


For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then frmUtils.Controls(q).Enabled = True
Next q
Call MsgBox("Your route is now ready for packaging. As some .zip programs have problems with empty folders, suggest" _
            & vbCrLf & "that you check the route for empty folders, and if you find any, place a short empty text file in them, e.g. 0.txt" _
            , vbInformation, App.Title)

End Sub


Private Sub Command13_Click()
Dim i As Integer, strMoveTiles As String

If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If
frmUnwanted.Show 1

If booKillMove = True Then
strKillPath = RoutePath & "\RRBackups"
strMoveTiles = strKillPath & "\Tiles"
If Not DirExists(strKillPath) Then
MkDir strKillPath
End If
If Not DirExists(strKillPath & "\Tiles") Then
MkDir strKillPath & "\Tiles"
End If
End If
cursouind = 0
MousePointer = 11
Drive1(cursouind).Drive = Left$(TilePath, 2)
Dir1(cursouind).Path = TilePath
Text1(cursouind).Text = "*e.raw"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If booKillMove = True Then
   FileCopy File1(cursouind).Path & "\" & File1(cursouind).List(i), strMoveTiles & "\" & File1(cursouind).List(i)
   DoEvents
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
Dir1(cursouind).Path = TilePath
Text1(cursouind).Text = "*n.raw"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If booKillMove = True Then
   FileCopy File1(cursouind).Path & "\" & File1(cursouind).List(i), strMoveTiles & "\" & File1(cursouind).List(i)
   DoEvents
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
Dir1(cursouind).Path = TilePath
Text1(cursouind).Text = "*.bk"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i

'******************************


Dir1(cursouind).Path = RoutePath
Text1(cursouind).Text = "*.bk"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
Dir1(cursouind).Path = RoutePath
Text1(cursouind).Text = "*.bak"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
Dir1(cursouind).Path = WorldPath
Text1(cursouind).Text = "*w.bk"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
Dir1(cursouind).Path = WorldPath
Text1(cursouind).Text = "*.bak"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
   If Right$(File1(cursouind).Path, 1) = "\" Then

   Kill File1(cursouind).Path & File1(cursouind).List(i)
   Else
   Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
   End If
       
End If
Next i
'************************

Dir1(0).Path = RoutePath
Text1(0).Text = "*.*"
DoEvents
Text1(cursouind).Text = "*.*"
MousePointer = 0
Exit Sub

End Sub




Private Sub Command14_Click(Index As Integer)
Dim q As Integer, strTempRoute As String

On Error GoTo Errtrap
Text1(0).Text = "*.*"

strTempRoute = RoutePath
For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then
If frmUtils.Controls(q).Caption <> Lang(30) And frmUtils.Controls(q).Caption <> Lang(53) Then
frmUtils.Controls(q).Enabled = False
ElseIf frmUtils.Controls(q).Caption = Lang(30) Then

frmUtils.Controls(q).Caption = Lang(637)
End If
End If
Next q
ReDim LocoPath(0 To CHUNK), LocoName(0 To CHUNK)
ReDim Wagpath(0 To CHUNK), WagonName(0 To CHUNK)
ReDim Locomotives(0 To CHUNK)
ReDim Wagons(0 To CHUNK)
ReDim Service(0 To CHUNK)
ReDim SrvPath(0 To CHUNK)
ReDim Activities(0 To CHUNK), ActPath(0 To CHUNK)
ReDim Traffic(0 To CHUNK), TfcPath(0 To CHUNK)
ReDim LocoCoup(0 To CHUNK), LocoFCoup(0 To CHUNK), LocoBrake(0 To CHUNK), LocoType(0 To CHUNK), LocoRigid(0 To CHUNK)
ReDim WagCoup(0 To CHUNK), WagFCoup(0 To CHUNK), WagBrake(0 To CHUNK), WagType(0 To CHUNK), WagRigid(0 To CHUNK)
ReDim Paths(0 To CHUNK), PathsPath(0 To CHUNK)
ReDim PathUsed(0 To CHUNK)

strbadbits = vbNullString
strReport = vbNullString
strForPrint = vbNullString
Set frmStock = Nothing
booStockOnly = False

ConEngNumber = 0: ConWagNumber = 0
lngAct = 0
lngSrv = 0
lngCon = 0
lngLoco = 0
lngWagons = 0
lngTfc = 0
lngPaths = 0
PathUsedNumb = 0

ActChecked = True
booActsChecked = True
For i = 0 To 5
Label7(i).Caption = vbNullString
Next i
Label3.Caption = vbNullString
SB2.Panels(2).Text = "Counting Rolling-stock"
Call CountStock

If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0

Exit Sub
End If

SB2.Panels(2).Text = "Counting Activities"
Call GetActivities
If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0
Exit Sub
End If

Label7(2).Caption = Str(lngAct)
ReDim Preserve PathUsed(0 To PathUsedNumb)
SB2.Panels(2).Text = "Counting Services"
Call GetServices
If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0
Exit Sub
End If

DoEvents
Label7(3).Caption = Str(lngSrv)
SB2.Panels(2).Text = "Counting Consists"
Call GetConsists
If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0
Exit Sub
End If

Label7(4).Caption = Str(lngCon)

SB2.Panels(2).Text = "Counting Traffic"

Call GetTraffic
Label7(5).Caption = Str(lngTfc)

If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0
Exit Sub
End If


SB2.Panels(2).Text = "Counting Paths"

Call GetPaths
If booAbort = True Then
booAbort = False


Text1(0).Text = "*.*"

MousePointer = 0
Exit Sub
End If
Label3.Caption = lngPaths


DoEvents
booNoButtons = False

BooCheckAct = False
For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then frmUtils.Controls(q).Enabled = True
Next q
Command1(15).Caption = Lang(30)
DoEvents
SB2.Panels(2).Text = "Populating Grid"
frmUtils.Refresh
DoEvents
frmGrid.Show

     DoEvents
     
Command15.Visible = True

Dir1(0).Path = strTempRoute
Text1(0).Text = "*.s"
DoEvents
Text1(cursouind).Text = "*.*"
SB2.Panels(2).Text = vbNullString
strForPrint = vbNullString
strReport = vbNullString
strbadbits = vbNullString
Exit Sub
Errtrap:


Call MsgBox("An error " & Err & " occurred in subroutine 'CheckActivities' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub


Private Sub Command15_Click()


frmUtils.ZOrder

strReport = vbNullString
strbadbits = vbNullString
Call GetStock
frmStock.Show

     DoEvents
     
End Sub

Private Sub Command16_Click()
Dim TrainsetPath As String

TrainsetPath = MSTSPath & "\Trains\"
If booActsChecked = True Then GoTo CarryON
ReDim LocoPath(0 To CHUNK), LocoName(0 To CHUNK)
ReDim Wagpath(0 To CHUNK), WagonName(0 To CHUNK)
ReDim Locomotives(0 To CHUNK)
ReDim Wagons(0 To CHUNK)
ReDim LocoCoup(0 To CHUNK), LocoFCoup(0 To CHUNK), LocoBrake(0 To CHUNK), LocoType(0 To CHUNK), LocoRigid(0 To CHUNK)
ReDim WagCoup(0 To CHUNK), WagFCoup(0 To CHUNK), WagBrake(0 To CHUNK), WagType(0 To CHUNK), WagRigid(0 To CHUNK)

ConEngNumber = 0: ConWagNumber = 0
lngAct = 0
lngSrv = 0
lngCon = 0
lngLoco = 0
lngWagons = 0
lngTfc = 0
CarryON:
ReDim strAnimUsed(0 To 50), LocoSMS(0 To CHUNK)

booStockOnly = True
    Call KillSpare("*.s")
DoEvents

strReport = vbNullString
cursouind = 0
DoEvents
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = TrainsetPath
Text1(0) = "*.*"
frmUtils.Refresh
Trainspath = Dir1(cursouind).Path
If Right$(Trainspath, 6) <> "Trains" Then
Call MsgBox("Please select the 'TRAINS' folder in the left hand folder window" _
            & vbCrLf & "which you wish to check." _
            & vbCrLf & "" _
            , vbExclamation, App.Title)

Exit Sub
End If
Call GetStock3
DoEvents
Text1(0) = "*.*"
frmUtils.Refresh


End Sub



Private Sub Command17_Click()
Dim booExists As Boolean, NewRouteName As String, Filpath1$, OldRouteName As String
Dim SparePath As String

'Set tfh = New TokenFileHandler
strReport = vbNullString
Filpath1$ = App.Path & "\setupfiles"
If FileExists(App.Path & "\setupfiles\master.ref") Then
Kill App.Path & "\setupfiles\master.ref"
End If
If FileExists(App.Path & "\setupfiles\InstallMe.bat") Then
Kill App.Path & "\setupfiles\InstallMe.bat"
End If

RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If


'If Not FileExists(RoutePath & "\" & RouteName & ".ref") Then
Text1(0) = "*.trk"
RouteName = File1(cursouind).List(i)
OldRouteName = RouteName
Call CheckForSMS(RoutePath & "\" & RouteName)

Call FindRouteName(RoutePath & "\" & RouteName, NewRouteName, booExists)
If booExists = False Then

MsgBox Lang(339) & vbCr & Lang(340), 16, Lang(341)
'********************
Exit Sub
Else
RouteName = NewRouteName
End If

Call IsItElectric(RoutePath & "\" & OldRouteName, booElectric)

'End If
OriginalRef = RoutePath & "\" & RouteName & ".ref"
If FileExists(OriginalRef) Then
FileCopy RoutePath & "\" & RouteName & ".ref", App.Path & "\setupfiles\master.ref"
Else
flagNoRef = True
Call MsgBox(Lang(385) & vbCrLf & Lang(386), vbExclamation, "No .ref file")
FileCopy App.Path & "\stuffit.ref", App.Path & "\setupfiles\master.ref"
End If
RouteListed = True
TexturePath = RoutePath & "\Textures"
TexSnowPath = RoutePath & "\Textures\Snow"
TexNightPath = RoutePath & "\Textures\Night"
TexAutPath = RoutePath & "\Textures\Autumn"
TexAutSnowPath = RoutePath & "\Textures\AutumnSnow"
TexSprPath = RoutePath & "\Textures\Spring"
TexSprSnowPath = RoutePath & "\Textures\SpringSnow"
TexWinPath = RoutePath & "\Textures\Winter"
TexWinSnowPath = RoutePath & "\Textures\WinterSnow"
TilePath = RoutePath & "\Tiles"
ShapePath = RoutePath & "\Shapes"
SoundPath = RoutePath & "\Sound"
WorldPath = RoutePath & "\World"
Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
    cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************
MousePointer = 11
Filpath1$ = App.Path & "\setupfiles"
cursouind = 0

Rem *********Check World Files *********
WorldPath = RoutePath & "\World"
Call DoDeCompFolder("w", WorldPath, SparePath)


DoEvents
Call CountShapes
    cursouind = 1
SparePath = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.W"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
'Rem *******************
MousePointer = 0
End Sub


Private Sub Command18_Click()
Dim strTempSMS As String, flagway As Integer, tempSoundPath As String
On Error GoTo Errtrap
Select Case MsgBox(Lang(444) & vbCrLf & Lang(445), vbOKCancel + vbExclamation + vbDefaultButton1, Lang(446))

    Case vbOK
Rem Go Ahead
    Case vbCancel
Exit Sub
End Select
If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
SparePath = App.Path & "\TempFiles\"

cursouind = 1

Drive1(cursouind).Drive = Left$(SparePath, 2)
Dir1(cursouind).Path = SparePath
Text1(cursouind).Text = "*.sms"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 Drive1(cursouind).Drive = Left$(SparePath, 2)
Dir1(cursouind).Path = SparePath
Text1(cursouind).Text = "*.*"


cursouind = 0

tempSoundPath = Dir1(cursouind).Path


Text1(cursouind).Text = "*.sms"
If File1(cursouind).ListCount = 0 Then
Call MsgBox(Lang(447), vbExclamation, App.Title)

Exit Sub
End If

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   strTempSMS = File1(cursouind).List(i)
   FileCopy fullpath$, SparePath & strTempSMS
   Next i

   cursouind = 1
Drive1(cursouind).Drive = Left$(SparePath, 2)
Dir1(cursouind).Path = SparePath
Text1(cursouind).Text = "*.sms"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
flagway = 0
Call ConvertSMS(fullpath$, flagway)
flagway = 1
Call ConvertSMS(fullpath$, flagway)
FileCopy fullpath$, tempSoundPath & "\" & File1(cursouind).List(i)
DoEvents
Kill fullpath$
Next i
Text1(cursouind).Text = "*.*"
Exit Sub
Errtrap:
If Err = 75 Then
MsgBox Lang(448) & vbCr & Lang(449), 48, Lang(450)
'********************
Resume Next
Else
Call MsgBox("An error " & Err & " occurred in subroutine 'Add LoadAllWaves' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End If

End Sub

Private Sub Command19_Click()
Dim booExists As Boolean, NewRouteName As String, SparePath, OldRouteName As String
Dim flagway As Integer, strMid As String
On Error GoTo Errtrap
SparePath = App.Path & "\TempFiles"



RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If



Text1(0) = "*.trk"
RouteName = File1(cursouind).List(i)
OldRouteName = RouteName

Call FindRouteName(RoutePath & "\" & RouteName, NewRouteName, booExists)
If booExists = False Then

MsgBox Lang(339) & vbCr & Lang(340), 16, Lang(341)

Exit Sub
Else
RouteName = NewRouteName
End If
booEnvFound = True

RouteListed = True
TexturePath = RoutePath & "\Textures"
TexSnowPath = RoutePath & "\Textures\Snow"
TexNightPath = RoutePath & "\Textures\Night"
TexAutPath = RoutePath & "\Textures\Autumn"
TexAutSnowPath = RoutePath & "\Textures\AutumnSnow"
TexSprPath = RoutePath & "\Textures\Spring"
TexSprSnowPath = RoutePath & "\Textures\SpringSnow"
TexWinPath = RoutePath & "\Textures\Winter"
TexWinSnowPath = RoutePath & "\Textures\WinterSnow"
TilePath = RoutePath & "\Tiles"
ShapePath = RoutePath & "\Shapes"
SoundPath = RoutePath & "\Sound"
WorldPath = RoutePath & "\World"
EnvPath = RoutePath & "\Envfiles"
Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************

Call CheckTrk(RoutePath & "\" & OldRouteName)

strMid = "Environment (" & vbCrLf
strMid = strMid & vbTab & "SpringClear (SpringClear.env)" & vbCrLf
strMid = strMid & vbTab & "SpringRain (SpringRain.env)" & vbCrLf
strMid = strMid & vbTab & "SpringSnow (SpringSnow.env)" & vbCrLf
strMid = strMid & vbTab & "SummerClear (SummerClear.env)" & vbCrLf
strMid = strMid & vbTab & "SummerRain (SummerRain.env)" & vbCrLf
strMid = strMid & vbTab & "SummerSnow (SummerSnow.env)" & vbCrLf
strMid = strMid & vbTab & "AutumnClear (AutumnClear.env)" & vbCrLf
strMid = strMid & vbTab & "AutumnRain (AutumnRain.env)" & vbCrLf
strMid = strMid & vbTab & "AutumnSnow (AutumnSnow.env)" & vbCrLf
strMid = strMid & vbTab & "WinterClear (WinterClear.env)" & vbCrLf
strMid = strMid & vbTab & "WinterRain (WinterRain.env)" & vbCrLf
strMid = strMid & vbTab & "WinterSnow (WinterSnow.env)" & vbCrLf

FileCopy RoutePath & "\" & OldRouteName, SparePath & "\" & OldRouteName
flagway = 0
Call ConvertTrk(SparePath & "\" & OldRouteName, flagway, strMid)
DoEvents
flagway = 1
Call ConvertTrk(SparePath & "\" & OldRouteName, flagway, strMid)
DoEvents
Close

FileCopy SparePath & "\" & OldRouteName, RoutePath & "\" & OldRouteName
DoEvents
If FileExists(SparePath & "\" & OldRouteName) Then
Kill SparePath & "\" & OldRouteName
End If
Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " occurred in subroutine 'Set up new environment' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub
Function GetApplicationPath(ByVal ExeName As String) As String
    GetApplicationPath = GetRegistryValue(HKEY_LOCAL_MACHINE, _
        "Software\Microsoft\Microsoft Games\" & ExeName, "")
End Function

Private Sub Command2_Click()
Dim i As Integer, strTemp As String, strTempZ As String


    If RouteListed = False Then
Select Case MsgBox(Lang(451) & vbCrLf & Dir1(cursouind).Path, vbOKCancel + vbExclamation + vbDefaultButton1, "Route Selection")

    Case vbOK

    Case vbCancel
Exit Sub
End Select

End If

''    'Open an archive  exist will create a new archive
'



    On Error Resume Next
    FromZip = 1

NewZipPath = RoutePath
ZipName = RouteName

frmNewZip.Show 1

     DoEvents
     
i = 0
If booMulti = True Then
If FileExists(strZipName) Then
Name strZipName As strZipName & ".zip"
End If
Do
strTemp = Trim$(Str(i))
strTemp = String(3 - Len(strTemp), "0") & strTemp
strTempZ = Trim$(Str(i + 1))
strTempZ = String(2 - Len(strTempZ), "0") & strTempZ
strTempZ = ".z" & strTempZ
strTemp = "." & strTemp
If FileExists(strZipName & strTemp) Then
Name strZipName & strTemp As strZipName & strTempZ
i = i + 1
Else
Exit Do
End If
Loop
End If
End Sub

Sub CountFiles(ByVal Path As String)
    Dim names() As String, i As Long
    ' Ensure that there is a trailing backslash.
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
  
        names() = GetFiles(Path & "*.*")          ' & Choose(j, "exe", "bat", "com"))
        ' Load partial results in the ListBox lst.
        For i = 1 To UBound(names)
            lngZipSize = lngZipSize + FileLen(Path & names(i))
           lngZipFiles = lngZipFiles + 1
        Next
   
    names() = GetDirectories(Path, vbHidden)
    For i = 1 To UBound(names)
        CountFiles Path & names(i)
    Next
End Sub

Function GetFiles(filespec As String, Optional Attributes As _
    VbFileAttribute) As String()
    Dim result() As String
    Dim Filename As String, Count As Long
    Const ALLOC_CHUNK = 50
    ReDim result(0 To ALLOC_CHUNK) As String
    Filename = Dir$(filespec, Attributes)
    Do While Len(Filename)
        Count = Count + 1
        If Count > UBound(result) Then
            ' Resize the result array if necessary.
            ReDim Preserve result(0 To Count + ALLOC_CHUNK) As String
        End If
        result(Count) = Filename
        ' Get ready for the next iteration.
        Filename = Dir$
    Loop
    ' Trim the result array.
    ReDim Preserve result(0 To Count) As String
    GetFiles = result
End Function

Function GetDirectories(Path As String, Optional Attributes As _
    VbFileAttribute, Optional IncludePath As Boolean) As String()
    Dim result() As String
    Dim dirName As String, Count As Long, path2 As String
    Const ALLOC_CHUNK = 50
    ReDim result(ALLOC_CHUNK) As String
    ' Build the path name + backslash.
    path2 = Path
    If Right$(path2, 1) <> "\" Then path2 = path2 & "\"
    dirName = Dir$(path2 & "*.*", vbDirectory Or Attributes)
    Do While Len(dirName)
        If dirName = "." Or dirName = ".." Then
            ' Exclude the "." and ".." entries.
        ElseIf (GetAttr(path2 & dirName) And vbDirectory) = 0 Then
            ' This is a regular file.
        Else
            ' This is a directory.
            Count = Count + 1
            If Count > UBound(result) Then
                ' Resize the result array if necessary.
                ReDim Preserve result(Count + ALLOC_CHUNK) As String
            End If
            ' Include the path if requested.
            If IncludePath Then dirName = path2 & dirName
            result(Count) = dirName
        End If
        dirName = Dir$
    Loop
    ' Trim the result array.
    ReDim Preserve result(Count) As String
    GetDirectories = result
End Function


Private Sub Command20_Click()
Dim i As Integer, flagway As Integer, Season As Integer
If booEnvFound = False Then
Call MsgBox(Lang(452) & vbCrLf & Lang(453), vbExclamation, Lang(454))

Exit Sub
End If
SparePath = App.Path & "\TempFiles"

frmGetSun.Show 1

     DoEvents
     
If booCancel = True Then
booCancel = False
Exit Sub
End If

If booSouth = False Then
For i = 1 To 12
Select Case i
Case 1 To 3
Season = 1
Case 4 To 6
Season = 2
Case 7 To 9
Season = 3
Case 10 To 12
Season = 4
End Select

FileCopy EnvPath & "\" & OldEnv(i), SparePath & "\" & OldEnv(i)
flagway = 0
Call ConvertSun(SparePath & "\" & OldEnv(i), flagway, Season)
flagway = 1
Call ConvertSun(SparePath & "\" & OldEnv(i), flagway, Season)
FileCopy SparePath & "\" & OldEnv(i), EnvPath & "\" & OldEnv(i)
Kill SparePath & "\" & OldEnv(i)
Next i
ElseIf booSouth = True Then
For i = 1 To 12
Select Case i
Case 1 To 3
Season = 1
Case 4 To 6
Season = 4
Case 7 To 9
Season = 3
Case 10 To 12
Season = 1
End Select

FileCopy EnvPath & "\" & OldEnv(i), SparePath & "\" & OldEnv(i)
flagway = 0
Call ConvertSun(SparePath & "\" & OldEnv(i), flagway, Season)
flagway = 1
Call ConvertSun(SparePath & "\" & OldEnv(i), flagway, Season)
FileCopy SparePath & "\" & OldEnv(i), EnvPath & "\" & OldEnv(i)
Kill SparePath & "\" & OldEnv(i)
Next i
End If
End Sub

Private Sub Command21_Click()
Dim flagway As Integer, NewFile2 As Integer, strTemp As String
'frmReport.Show 1
Close
For i = 0 To File1(cursouind).ListCount - 1
   
  
   If File1(cursouind).Selected(i) Then
        If Right$(File1(cursouind).Path, 1) = "\" Then
        fullpath$ = File1(cursouind).Path & File1(cursouind).List(i)
        Else
        fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
        End If
   NewFile2 = FreeFile
   Open fullpath$ For Binary As #NewFile2
    strTemp = String(2, " ")
    Get #NewFile2, , strTemp
 Close #NewFile2
 
 If Asc(Mid$(strTemp, 1, 1)) = 255 And Asc(Mid$(strTemp, 2, 1)) = 254 Then
 
   flagway = 0
   Call ConvertIt(fullpath$, flagway)
   DoEvents
   booUniEdit = True
   strUniName = fullpath$
   frmReport.Rich1.LoadFile fullpath$
   If Right$(fullpath$, 3) = "eng" Then
   frmReport.Text1 = "sound"
   End If
   frmReport.Show 1
   
     DoEvents
     
 DoEvents

 flagway = 1
   Call ConvertIt(fullpath$, flagway)
    File1(cursouind).Selected(i) = False
    Else
    Call MsgBox(Lang(344) & fullpath$ & vbCrLf & Lang(455), vbExclamation, App.Title)
                File1(cursouind).Selected(i) = False
    GoTo Another
    End If
    End If
Another:
Next i
End Sub

Private Sub Command22_Click()
Dim strPath As String, flagway As Integer, strAI As String, strEng As String

On Error GoTo Errtrap
If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
SparePath = App.Path & "\TempFiles"
frmAI.Show vbModal

     DoEvents
     

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
   strPath = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   If Right$(strPath, 4) <> ".eng" Then
   Call MsgBox(Lang(393) & File1(cursouind).List(i) & vbCrLf & Lang(456), vbExclamation, App.Title)
   
   GoTo CarryON
   End If
   If booMU = False Then
   FileCopy strPath, SparePath & "\#" & File1(cursouind).List(i)
   strAI = SparePath & "\#" & File1(cursouind).List(i)
   strEng = "#" & File1(cursouind).List(i)
   ElseIf booMU = True Then
   FileCopy strPath, SparePath & "\MU_" & File1(cursouind).List(i)
   strAI = SparePath & "\MU_" & File1(cursouind).List(i)
   strEng = "MU_" & File1(cursouind).List(i)
   End If
   flagway = 0
   Call ConvertAI(strAI, flagway)
   DoEvents
   flagway = 1
   Call ConvertAI(strAI, flagway)
   DoEvents
  If booMU = False Then
   FileCopy SparePath & "\#" & File1(cursouind).List(i), File1(cursouind).Path & "\#" & File1(cursouind).List(i)
   DoEvents
   Kill SparePath & "\#" & File1(cursouind).List(i)
   ElseIf booMU = True Then
   FileCopy SparePath & "\MU_" & File1(cursouind).List(i), File1(cursouind).Path & "\MU_" & File1(cursouind).List(i)
   DoEvents
   Kill SparePath & "\MU_" & File1(cursouind).List(i)
   End If
   End If
CarryON:
   Next i
   
   Text1(cursouind) = "*.eng"
   DoEvents
   Text1(cursouind) = "*.*"
Exit Sub
Errtrap:


End Sub

Private Sub Command23_Click()
Dim strTempSMS As String, flagway As Integer, tempSoundPath As String

Select Case MsgBox(Lang(457) & vbCrLf & Lang(445), vbOKCancel + vbExclamation + vbDefaultButton1, Lang(446))

    Case vbOK
Rem Go Ahead
    Case vbCancel
Exit Sub
End Select
If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
SparePath = App.Path & "\TempFiles\"

cursouind = 1

Drive1(cursouind).Drive = Left$(SparePath, 2)
Dir1(cursouind).Path = SparePath
Text1(cursouind).Text = "*.sms"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 Drive1(cursouind).Drive = Left$(SparePath, 2)
Dir1(cursouind).Path = SparePath
Text1(cursouind).Text = "*.*"


cursouind = 0

tempSoundPath = Dir1(cursouind).Path

Text1(cursouind).Text = "*.sms"
If File1(cursouind).ListCount = 0 Then
Call MsgBox(Lang(447), vbExclamation, App.Title)

Exit Sub
End If

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   strTempSMS = File1(cursouind).List(i)
   FileCopy fullpath$, SparePath & strTempSMS
   Next i

   cursouind = 1
Drive1(cursouind).Drive = Left$(SparePath, 2)
Dir1(cursouind).Path = SparePath
Text1(cursouind).Text = "*.sms"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
flagway = 0
Call UnConvertSMS(fullpath$, flagway)
flagway = 1
Call UnConvertSMS(fullpath$, flagway)
FileCopy fullpath$, tempSoundPath & "\" & File1(cursouind).List(i)
DoEvents
Kill fullpath$
Next i
Text1(cursouind).Text = "*.*"
End Sub

Private Sub Command24_Click()


On Error GoTo Errtrap



Call CompressACE
File1(cursouind).Refresh
tit$ = vbNullString


Errtrap:
Resume Next
End Sub


Private Sub Command25_Click()
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If
Select Case MsgBox(Lang(458) & vbCrLf & Lang(459), vbOKCancel + vbExclamation + vbDefaultButton1, "Warning")

    Case vbOK
Call MakeSnow
    Case vbCancel
Exit Sub
End Select

End Sub

Private Sub Command26_Click()
Dim i As Integer
Dim Filpath1$
Dim strOrigFile As String, Ename As String
'Call KillSpare2("*.s")

If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
Select Case MsgBox(Lang(460) & vbCrLf & Lang(438), vbOKCancel + vbInformation + vbDefaultButton1, App.Title)
    Case vbOK


Label9.Visible = True
On Error GoTo Errtrap

MousePointer = 11

ShapePath = Dir1(0).Path

 cursouind = 0


Filpath1$ = App.Path & "\TempFiles"

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   strOrigFile = File1(cursouind).List(i)
      If Right$(strOrigFile, 2) <> ".s" And Right$(strOrigFile, 2) <> ".t" And Right$(strOrigFile, 2) <> ".w" Then
   MousePointer = 0
   Call MsgBox(Lang(393) & strOrigFile & Lang(410), vbExclamation, Lang(404))
GoTo GetNext

 Else
 Ename = Dir1(1).Path
 Label9.Caption = "Processing:  " & strOrigFile
 DoEvents
   
   Call DoDeComp2(strOrigFile, File1(cursouind).Path, Ename)
   End If
   End If
GetNext:
   Next i


 Call KillSpare("*.s")
 DoEvents
 cursouind = 0
 Drive1(0).Drive = Left$(ShapePath, 2)
Dir1(0).Path = ShapePath
Text1(0).Text = "*.*"
Drive1(1).Drive = Left$(Ename, 2)
Dir1(1).Path = Ename
'Text1(1).Text = "*.s"
'Text1(1).Text = "*.*"
  MousePointer = 0
  Label9.Visible = False
 
Exit Sub
Errtrap:


If Err = 53 Then
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'Command26' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
File1(cursouind).Refresh
tit$ = vbNullString

    Case vbCancel
Exit Sub
End Select

End Sub

Private Sub Command28_Click()
Dim flagway As Integer

SparePath = App.Path & "\TempFiles"
MousePointer = 11

cursouind = 0
RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If
EnvPath = RoutePath & "\envfiles"
If Not DirExists(EnvPath) Then
Call MsgBox(Lang(461) & vbCrLf & Lang(462), vbCritical, App.Title)

Exit Sub
End If
Dir1(cursouind).Path = EnvPath
Text1(cursouind) = "*.env"

For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1
   fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   flagway = 0
   Call ConvertENV(fullpath$, flagway)
   flagway = 1
   Call ConvertENV(fullpath$, flagway)
Next i
MousePointer = 0
Text1(0).Text = "*.*"

End Sub

Private Sub Command29_Click()
Dim i As Integer, NewRouteName As String
Dim booExists As Boolean, OldRouteName As String

cursouind = 0
SparePath = App.Path & "\TempFiles"

RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If
Text1(0) = "*.trk"
RouteName = File1(cursouind).List(i)
OldRouteName = RouteName
Call FindRouteName(RoutePath & "\" & RouteName, NewRouteName, booExists)
If booExists = False Then
MsgBox Lang(339) & vbCr & Lang(340), 16, Lang(341)
Exit Sub
Else
RouteName = NewRouteName
End If
WorldPath = RoutePath & "\World"

RouteListed = True
Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************


If RouteName = vbNullString Then
Call MsgBox(Lang(463), vbExclamation, App.Title)

Exit Sub
End If
MousePointer = 11
Call UncompressAllW(WorldPath)

GetAnother:
Close

frmDelShape.Show 1

     DoEvents
     
cursouind = 0

 If strDelShape <> vbNullString Then
 
MSG = Lang(110) & strDelShape & Lang(465) & RouteName
Response = MsgBox(MSG, 36, Lang(464))
Select Case Response
Case vbYes

Call StripW(strDelShape)

strDelShape = vbNullString
Case vbNo
strDelShape = vbNullString
End Select
End If
Select Case MsgBox("Do you wish to delete any more shapes from this Route?", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbYes
GoTo GetAnother
    Case vbNo

End Select
MousePointer = 0
Text1(1).Text = "*.*"
End Sub


Private Sub Command3_Click()

Call MakeReadWrite(RoutePath)
'Call Command10_Click(0)
End Sub

Private Sub Command30_Click()
cursouind = 1
SparePath = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.s"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
booCopy = True

Command10(1).value = True
DoEvents
Command10(0).value = True
DoEvents
frmCopy.Show

     DoEvents
     


End Sub

Private Sub Command31_Click()
 On Error GoTo Errtrap
 strComPath = MSTSPath & "\Common"
 
 If Not DirExists(strComPath) Then
   MkDir (strComPath)
 End If
 MkDir strComPath & "\Envfiles"
 MkDir strComPath & "\Shapes"
 MkDir strComPath & "\Sound"
 MkDir strComPath & "\Terrtex"
 MkDir strComPath & "\Textures"
 MkDir strComPath & "\Stuffed"
 If FileExists(App.Path & "\StuffitPack\EZstuff3.bat") Then
   FileCopy App.Path & "\StuffitPack\EZstuff3.bat", strComPath & "\EZstuff3.bat"
 Else
   Call MsgBox(Lang(466) & vbCrLf & Lang(467), vbCritical, Lang(347))
   Exit Sub
 End If
 Exit Sub
Errtrap:
 If Err = 75 Then Resume Next
 
End Sub


Private Sub Command32_Click()
 On Error Resume Next
 
 booCommon = True
 Command10(1).value = True
DoEvents
Command10(0).value = True
DoEvents
 frmCommon.Show
 
     DoEvents
     
 
End Sub

Private Sub Command33_Click()
 On Error Resume Next
 Dim strDrive As String, strBat As String
 
 Eur1Path = MainRoutePath & "Europe1\"
 Eur2Path = MainRoutePath & "Europe2\"
 Jap1Path = MainRoutePath & "Japan1\"
 Jap2Path = MainRoutePath & "Japan2\"
 USA1Path = MainRoutePath & "USA1\"
 USA2Path = MainRoutePath & "USA2\"
 
 Rem *******************************
 strBat = "xcopy " & ChrW$(34) & App.Path & "\StuffitPack\stuffed\*.*" & ChrW$(34) & " " & ChrW$(34) & strComPath & "\stuffed\*.*" & ChrW$(34) & " /s /y"

Open App.Path & "\TempFiles\do_stuff.bat" For Output As #12

Print #12, strBat
Close #12

ChDrive (Left$(App.Path, 1))
ChDir App.Path & "\TempFiles"

Call ShellAndWait("do_stuff.bat", True, vbNormalFocus)
DoEvents

 strDrive = Left$(strComPath, 1)
 ChDrive strDrive
 ChDir strComPath
 mydir = CurDir
 DoEvents
 Call ShellAndWait(ChrW$(34) & strComPath & "\ezstuff3.bat" & ChrW$(34), True, vbNormalFocus)
 DoEvents
 
 Call MsgBox(Lang(468) & vbCrLf & Lang(469), vbExclamation, Lang(470))
 
End Sub




Private Sub Command34_Click()
Dim tit$, tit2$, tit3$, x As Integer
On Error GoTo Errtrap

Select Case MsgBox(Lang(471) & vbCrLf & Lang(472), vbOKCancel + vbInformation + vbDefaultButton1, App.Title)

    Case vbOK
   
    
tit3$ = Dir1(1).List(Dir1(1).ListIndex)

   tit2$ = Dir1(0).List(Dir1(0).ListIndex)
    If Drive1(0).Drive <> Drive1(1).Drive Then
    Call MsgBox(Lang(473), vbCritical + vbDefaultButton1, App.Title)
    
    Exit Sub
    End If
   x = InStrRev(tit2$, "\")
   tit$ = Mid$(tit2$, x + 1)
   Name tit2$ As tit3$ & "\" & tit$
 Dir1(0).Refresh
 Dir1(1).Refresh
    Case vbCancel
Exit Sub
End Select

Exit Sub
Errtrap:
If Err = 58 Then
Call MsgBox(Lang(474) & tit$ & Lang(475) & vbCrLf & Lang(476), vbExclamation + vbDefaultButton1, App.Title)


End If

 

End Sub

Private Sub Command35_Click()
Dim i As Integer
ReDim PathUsed(0 To CHUNK)
For i = 0 To File1(cursouind).ListCount - 1

  If File1(cursouind).Selected(i) = True Then
    Filpath$ = File1(cursouind).Path
   
   Exit For
    End If
    Next i
   If Right$(Filpath$, 1) = "\" Then
   Filpath$ = Left$(Filpath$, Len(Filpath$) - 1)
   End If

tit$ = Filpath$ & "\" & File1(cursouind).List(i)
If Right$(tit$, 4) <> ".act" Then
Call MsgBox(Lang(477) & vbCrLf & Lang(478), vbExclamation, Lang(480))
Exit Sub
End If


strForPrint = vbNullString
strReport = vbNullString
strbadbits = vbNullString
PathUsedNumb = 0
BooCheckAct = True

Call GetActDetails(tit$)
DoEvents


strbadbits = vbNullString
'End If
'Next i
If strActReport <> "" Then
frmReport.Rich1.Text = strActReport
frmReport.Show 1

End If
DoEvents
If strForPrint <> vbNullString Then
 frmReport.Rich1.Text = strForPrint
     frmReport.Show 1
     
     DoEvents
     
End If
strForPrint = vbNullString
strReport = vbNullString
End Sub

Private Sub Command36_Click()
Dim result As String, i As Integer

For i = 0 To 1
result = MSTSPath
Text1(i) = "*.*"
Drive1(i).Drive = Left$(result, 2)
Dir1(i).Path = result & "\Routes"
Next i
booUpdate = True
frmUpdate.Show

     DoEvents
     

End Sub

Private Sub Command37_Click()
Dim i As Integer, strTemp As String, strTempZ As String

frmNewZip.Show 1

     DoEvents
     

i = 0
If booMulti = True Then
If FileExists(strZipName) Then
Name strZipName As strZipName & ".zip"
End If
Do
strTemp = Trim$(Str(i))
strTemp = String(3 - Len(strTemp), "0") & strTemp
strTempZ = Trim$(Str(i + 1))
strTempZ = String(2 - Len(strTempZ), "0") & strTempZ
strTempZ = ".z" & strTempZ
strTemp = "." & strTemp
If FileExists(strZipName & strTemp) Then
Name strZipName & strTemp As strZipName & strTempZ
i = i + 1
Else
Exit Do
End If
Loop
End If

End Sub

Private Sub Command38_Click()
Dim i As Integer, NewRouteName As String
Dim booExists As Boolean, OldRouteName As String

SparePath = App.Path & "\TempFiles"
strReport = vbNullString
RoutePath = File1(0).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If
Text1(0) = "*.trk"
RouteName = File1(0).List(i)
OldRouteName = RouteName
Call FindRouteName(RoutePath & "\" & RouteName, NewRouteName, booExists)
If booExists = False Then
MsgBox Lang(339) & vbCr & Lang(340), 16, Lang(341)
Exit Sub
Else
RouteName = NewRouteName
End If
WorldPath = RoutePath & "\World"

RouteListed = True
Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
Close

cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************


If RouteName = vbNullString Then
Call MsgBox(Lang(463), vbExclamation, App.Title)

Exit Sub
End If

Call UncompressAllW(WorldPath)


booGetW = True
frmDelShape.Show 1

     DoEvents
     
cursouind = 0

 If strDelShape = vbNullString Then Exit Sub
 

Call ListInW(strDelShape)

'Call StripW(strDelShape)

strDelShape = vbNullString

booGetW = False

End Sub


Private Sub ListInW(strShape As String)
Dim i As Integer, filepath1$, fullpath$, x As Long, Z As Long, Y As Long
Dim NewFile As Integer, strNew As String, strPart As String, strPos As String

Rem
cursouind = 1
filepath1$ = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(filepath1$, 2)
Dir1(1).Path = filepath1$
Text1(1).Text = "*.W"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
    NewFile = FreeFile
    
  Open fullpath$ For Binary As #NewFile
  
strNew = String(lOf(NewFile), vbNullChar)
Get NewFile, , strNew
strPart = Replace(strNew, Chr$(0), "")
  Close #NewFile
  Z = 1
  Do
  x = InStr(Z, strPart, strShape)
  If x > 0 Then
  x = InStr(x, strPart, "Position")
  Y = InStr(x, strPart, ")")
  strPos = Mid$(strPart, x, Y + 1 - x)
  strReport = strReport & File1(cursouind).List(i) & vbTab & strShape & vbTab & strPos & vbCrLf
  Z = x + 1
  End If
  Loop Until x = 0
    End If
   

    
    Next i
    Close
    
     frmReport.Rich1.Text = strReport
     frmReport.Show 1
     
     DoEvents
     strReport = vbNullString
End Sub

Private Sub Command39_Click()
Dim GlobalShapePath As String, i As Integer
Dim GlobalPath As String

On Error GoTo Errtrap
 GlobalShapePath = MSTSPath & "\global\shapes\"
 GlobalPath = MSTSPath & "\global\"
 
 cursouind = 1
 Text1(cursouind) = "*.s"
Drive1(cursouind).Drive = Left$(MSTSPath, 2)
Dir1(cursouind).Path = GlobalShapePath
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
 
Select Case MsgBox(Lang(481) & vbCrLf & Lang(482), vbOKCancel + vbExclamation + vbDefaultButton1, App.Title)

    Case vbOK
    booTsection = True

If FileExists(GlobalPath & "Master_tsection.dat") Then
Name GlobalPath & "Master_tsection.dat" As GlobalPath & "Master_tsection.dat.old"
DoEvents
FileCopy GlobalPath & "tsection.dat", GlobalPath & "Master_tsection.dat"
Else
FileCopy GlobalPath & "tsection.dat", GlobalPath & "Master_tsection.dat"
End If
Call ParseTsection(GlobalPath & "Master_tsection.dat", GlobalPath & "tsection.dat")
    Case vbCancel
Exit Sub
End Select
Exit Sub
Errtrap:
Call MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'Command39' " _
            & vbCrLf & "GlobalPath=" & GlobalPath _
            , vbExclamation, App.Title)


Resume Next

End Sub
Private Sub ParseTsection2(pathOld As String, pathNew As String)
Dim NewFile As Integer, strNew As String
Dim x As Integer
Dim MaxSec As Long, MaxShape As Long, Tsec As Long, Tshp As Long
Dim xy As Integer, Y As Integer, yy As Integer, TSName As String
Dim booFound As Boolean, TSShape As String
Dim LastSec As Long, strTsec As String

Call ConvertIt(pathOld, 0)
MousePointer = 11
NewFile = FreeFile
Open pathOld For Input As NewFile

Do While Not EOF(NewFile)
   Line Input #NewFile, strNew
    x = InStr(strNew, "Tracksections (")
    If x > 0 And Len(strNew) < 46 Then
MaxSec = Val(Trim$(Mid$(strNew, x + 15)))

PartOne = "SIMISA@@@@@@@@@@JINX0F0t______"
PartOne = PartOne & vbCrLf & vbCrLf & "Tracksections (" & Str(MaxSec) & vbCrLf & vbCrLf
   Exit Do
    End If
    Loop
 Do While Not EOF(NewFile)
Label1:
   Line Input #NewFile, strNew
    strNew = Trim$(strNew)
    
Label2:
    If Left$(strNew, 14) = "Tracksection (" And Len(strNew) < 44 Then
      
    strTsec = strNew & vbCrLf
    Do
    Line Input #NewFile, strNew
    strNew = Trim$(strNew)
    If Left$(strNew, 5) <> "_skip" And Left$(strNew, 12) <> "Tracksection" And Left$(strNew, 11) <> "trackshapes" And Left$(strNew, 5) <> "_Info" Then
    'If Left$(strNew, 5) <> "_skip" And Left$(strNew, 12) <> "Tracksection" And Left$(strNew, 11) <> "trackshapes" Then
      strTsec = strTsec & strNew & vbCrLf
    Else
    
    PartOne = PartOne & strTsec
    strTsec = vbNullString

    Exit Do
    End If
    Loop
    
    If Left$(strNew, 12) = "Tracksection" Then GoTo Label2
    If Left$(strNew, 11) = "trackshapes" Then Exit Do
    If Left$(strNew, 5) = "_Info" Then GoTo Label1
    If Left$(strNew, 5) = "_skip" Then
    Do
    Line Input #NewFile, strNew
   
        strNew = Trim$(strNew)
        Loop Until strNew = ")"
    End If
    End If
    
   Loop
    Close #NewFile
 LastSec = Tsec - 1
 
   
   
    NewFile = FreeFile
Open pathOld For Input As NewFile
Do While Not EOF(NewFile)
   Line Input #NewFile, strNew
    x = InStr(strNew, "Trackshapes (")
    If x > 0 And Len(strNew) < 44 Then
   
    MaxShape = Val(Trim$(Mid$(strNew, x + 13)))
    PartTwo = "Trackshapes ( " & Str(MaxShape) & vbCrLf & vbCrLf
    
    Exit Do
    End If
    Loop
     Do While Not EOF(NewFile)
Label3:
   Line Input #NewFile, strNew
    strNew = Trim$(strNew)
    
Label4:
    If Left$(strNew, 12) = "Trackshape (" And Len(strNew) < 44 Then
    TSShape = strNew & vbCrLf

        Do While Not EOF(NewFile)
        Line Input #NewFile, strNew
        strNew = Trim$(strNew)
        If Left$(strNew, 10) <> "Trackshape" And Left$(strNew, 5) <> "_SKIP" And Left$(strNew, 5) <> "shape" And Left$(strNew, 5) <> "_Info" Then
            TSShape = TSShape & strNew & vbCrLf
            Y = InStr(strNew, "Filename")
            If Y > 0 Then
            xy = InStr(Y, strNew, "(")
            yy = InStr(xy, strNew, ")")
            TSName = Trim$(Mid$(strNew, xy + 1, yy - xy - 1))
            booFound = False
           Call FileInUse2(TSName, booFound)
           End If

        Else
        If booFound = True Then
        booFound = False
        PartTwo = PartTwo & TSShape & vbCrLf
        TSShape = vbNullString
        End If
            Tshp = Tshp + 1

        Exit Do
        End If
        Loop
    If Left$(strNew, 5) = "shape" Then GoTo Label3
    If Left$(strNew, 5) = "_Info" Then GoTo Label3
    If Left$(strNew, 10) = "Trackshape" Then GoTo Label4
    If Left$(strNew, 5) = "_skip" Then
    Do
    Line Input #NewFile, strNew
   
        strNew = Trim$(strNew)
        Loop Until strNew = ")"
    End If
    End If
    
    
    Loop
    If booFound = True Then
        booFound = False
        PartTwo = PartTwo & TSShape & vbCrLf
        TSShape = vbNullString
        End If
    Close #NewFile

    PartOne = PartOne & vbCrLf & PartTwo & vbCrLf & ")"
    Open pathNew For Output As #5
    Print #5, PartOne
    Close #5
    Call ConvertIt(pathNew, 1)
    DoEvents
    Call ConvertIt(pathOld, 1)
MousePointer = 0
End Sub

Private Sub ParseTsection(pathOld As String, pathNew As String)
Dim NewFile As Integer, strNew As String
Dim x As Integer
Dim MaxSec As Long, MaxShape As Long, Tsec As Long, Tshp As Long
Dim xy As Integer, Y As Integer, yy As Integer, TSName As String
Dim booFound As Boolean, TSShape As String
Dim LastSec As Long, strTsec As String

Call ConvertIt(pathOld, 0)
MousePointer = 11
NewFile = FreeFile
Open pathOld For Input As NewFile

Do While Not EOF(NewFile)
   Line Input #NewFile, strNew
    x = InStr(strNew, "Tracksections (")
    If x > 0 And Len(strNew) < 46 Then
MaxSec = Val(Trim$(Mid$(strNew, x + 15)))

PartOne = "SIMISA@@@@@@@@@@JINX0F0t______"
PartOne = PartOne & vbCrLf & vbCrLf & "Tracksections (" & Str(MaxSec) & vbCrLf & vbCrLf
   Exit Do
    End If
    Loop
 Do While Not EOF(NewFile)
Label1:
   Line Input #NewFile, strNew
    strNew = Trim$(strNew)
    
Label2:
    If Left$(strNew, 14) = "Tracksection (" And Len(strNew) < 44 Then
     
    strTsec = strNew & vbCrLf
    Do
    Line Input #NewFile, strNew
    strNew = Trim$(strNew)
    If Left$(strNew, 5) <> "_skip" And Left$(strNew, 12) <> "Tracksection" And Left$(strNew, 11) <> "trackshapes" And Left$(strNew, 5) <> "_Info" Then
  
    strTsec = strTsec & strNew & vbCrLf
    Else
    
    PartOne = PartOne & strTsec
    strTsec = vbNullString

    Exit Do
    End If
    Loop
    
    If Left$(strNew, 12) = "Tracksection" Then GoTo Label2
    If Left$(strNew, 11) = "trackshapes" Then Exit Do
    If Left$(strNew, 5) = "_Info" Then GoTo Label1
    If Left$(strNew, 5) = "_skip" Then
    Do
    Line Input #NewFile, strNew
   
        strNew = Trim$(strNew)
        Loop Until strNew = ")"
    End If
    End If
    
   Loop
    Close #NewFile
 LastSec = Tsec - 1
 
   
   
    NewFile = FreeFile
Open pathOld For Input As NewFile
Do While Not EOF(NewFile)
   Line Input #NewFile, strNew
    x = InStr(strNew, "Trackshapes (")
    If x > 0 And Len(strNew) < 44 Then
   
    MaxShape = Val(Trim$(Mid$(strNew, x + 13)))
    PartTwo = "Trackshapes ( " & Str(MaxShape) & vbCrLf & vbCrLf
    
    Exit Do
    End If
    Loop
     Do While Not EOF(NewFile)
Label3:
   Line Input #NewFile, strNew
    strNew = Trim$(strNew)
    
Label4:
    If Left$(strNew, 12) = "Trackshape (" And Len(strNew) < 44 Then
    TSShape = strNew & vbCrLf

        Do While Not EOF(NewFile)
        Line Input #NewFile, strNew
        strNew = Trim$(strNew)
        If Left$(strNew, 10) <> "Trackshape" And Left$(strNew, 5) <> "_SKIP" And Left$(strNew, 5) <> "shape" And Left$(strNew, 5) <> "_Info" Then
        
        
        TSShape = TSShape & strNew & vbCrLf
            Y = InStr(strNew, "Filename")
            If Y > 0 Then
            xy = InStr(Y, strNew, "(")
            yy = InStr(xy, strNew, ")")
            TSName = Trim$(Mid$(strNew, xy + 1, yy - xy - 1))
            booFound = False
           Call FileInUse(TSName, booFound)
           End If

        Else
        If booFound = True Then
        booFound = False
        PartTwo = PartTwo & TSShape & vbCrLf
        TSShape = vbNullString
        End If
            Tshp = Tshp + 1

        Exit Do
        End If
        Loop
    If Left$(strNew, 5) = "shape" Then GoTo Label3
    If Left$(strNew, 5) = "_info" Then GoTo Label3
    If Left$(strNew, 10) = "Trackshape" Then GoTo Label4
    If Left$(strNew, 5) = "_skip" Then
    Do
    Line Input #NewFile, strNew
   
        strNew = Trim$(strNew)
        Loop Until strNew = ")"
    End If
    End If
    
    
    Loop
    If booFound = True Then
        booFound = False
        PartTwo = PartTwo & TSShape & vbCrLf
        TSShape = vbNullString
        End If
    Close #NewFile

    PartOne = PartOne & vbCrLf & PartTwo & vbCrLf & ")"
    Open pathNew For Output As #5
    Print #5, PartOne
    Close #5
    Call ConvertIt(pathNew, 1)
    DoEvents
    Call ConvertIt(pathOld, 1)
MousePointer = 0
End Sub


Private Sub FileInUse2(strFile As String, booFound As Boolean)
Dim i As Integer


For i = 1 To UBound(strGlobShp2)
If strFile = strGlobShp2(i) Then
booFound = True
Exit For
End If
Next i

End Sub


Private Sub FileInUse(strFile As String, booFound As Boolean)
Dim i As Integer

cursouind = 1
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
If strFile = frmUtils.File1(cursouind).List(i) Then
booFound = True
End If
Next i

End Sub
Private Sub FindEnvAce(strAce As String)
Dim strTempPath As String

strTempPath = strComPath & "\envfiles\textures\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\europe1\envfiles\textures\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\europe2\envfiles\textures\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\japan1\envfiles\textures\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\japan2\envfiles\textures\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\usa1\envfiles\textures\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
strTempPath = MSTSPath & "\routes\usa2\envfiles\textures\"
If FileExists(strTempPath & strAce) Then
GoTo ShowMessage
End If
strReport = strReport & "Environment texture file " & strAce & Lang(563) & vbCrLf

Exit Sub
ShowMessage:
If intResponse = 0 Then
strResponse = "The Environment texture file " & strAce & Lang(564) & vbCrLf & Lang(565)
frmDialog.Label1 = strResponse
frmDialog.Command1.Visible = False

frmDialog.Show 1

     DoEvents
     
End If
If intResponse = 1 Then
intResponse = 0
FileCopy strTempPath & strAce, RoutePath & "\envfiles\textures\" & strAce
ElseIf intResponse = 2 Then
FileCopy strTempPath & strAce, RoutePath & "\envfiles\textures\" & strAce
ElseIf intResponse = 3 Then
intResponse = 0
    strReport = strReport & "Terrtex file " & strAce & Lang(567) & vbCrLf

End If



End Sub


Private Sub Command4_Click()
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If

Call CompressWFiles
End Sub

Private Sub Command40_Click()
Dim strTemp As String, NewFile As Integer, Filpath1$
Dim strFolder As String

MousePointer = 11

strFolder = Label2(0).Caption

strTemp = "ATTRIB -R " & ChrW$(34) & strFolder & "\*.*" & ChrW$(34) & " /s"

Filpath1$ = App.Path & "\TempFiles"

NewFile = FreeFile

Open Filpath1$ & "\do_read.bat" For Output As #NewFile

   Print #NewFile, strTemp
   
   Close NewFile
   strDrive = Left$(Filpath1$, 1)
      ChDrive strDrive
ChDir Filpath1$
mydir = CurDir
  DoEvents
'result = Shell(Environ$("comspec") & " /c do_read.bat", vbNormalFocus)
Call ShellAndWait("do_read.bat", True, vbNormalFocus)
 MousePointer = 0
End Sub

Private Sub Command42_Click()
cursouind = 0
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True

 
   Next i
End Sub

Private Sub Command43_Click(Index As Integer)
Select Case Index
Case 0
Text1(0).Text = "*.s"
Case 1
Text1(0).Text = "*.t"
Case 2
Text1(0).Text = "*.w"
Case 3
Text1(0).Text = "*.*"
End Select

End Sub

Private Sub Command44_Click()
Dim i As Integer, NewRouteName As String
Dim booExists As Boolean, OldRouteName As String

SparePath = App.Path & "\TempFiles"
MousePointer = 11
RoutePath = File1(0).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If
Text1(0) = "*.trk"
RouteName = File1(0).List(i)
OldRouteName = RouteName
Call FindRouteName(RoutePath & "\" & RouteName, NewRouteName, booExists)
If booExists = False Then
MsgBox Lang(339) & vbCr & Lang(340), 16, Lang(341)
Exit Sub
Else
RouteName = NewRouteName
End If
WorldPath = RoutePath & "\World"

RouteListed = True
Rem ********** Delete any spare .w files ****
SparePath = App.Path & "\TempFiles"
Close

cursouind = 1
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents





If RouteName = vbNullString Then
Call MsgBox(Lang(463), vbExclamation, App.Title)

Exit Sub
End If
Call UncompressAllW(WorldPath)
 Drive1(1).Drive = Left$(RoutePath, 2)
Dir1(1).Path = RoutePath & "\shapes"
Text1(1).Text = "*.s"
flagChange = 1
frmRepShape.Show

     DoEvents
     
MousePointer = 0






End Sub

Private Sub Command45_Click()
Dim x As Integer, strBat As String, NewFile As Integer

SparePath = App.Path & "\TempFiles"

RoutePath = File1(0).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If
If RouteName = vbNullString Then
Call MsgBox(Lang(463), vbExclamation, App.Title)

Exit Sub
End If
If Not DirExists(App.Path & "\StuffitPack") Then
Call MsgBox("You do not appear to have unzipped the file StuffitPack.zip" _
            & vbCrLf & "in the Route_Riter folder." _
            , vbExclamation, App.Title)


Exit Sub
End If


FileCopy App.Path & "\StuffitPack\EZStuff4.bat", RoutePath & "\EZStuff4.bat"
FileCopy App.Path & "\stuffit.ref", RoutePath & "\" & RouteName & ".ref"

strBat = "xcopy " & ChrW$(34) & App.Path & "\StuffitPack\stuffed\*.*" & ChrW$(34) & " " & ChrW$(34) & RoutePath & "\stuffed\*.*" & ChrW$(34) & " /s /y"
Open App.Path & "\TempFiles\do_stuff.bat" For Output As #12

Print #12, strBat
Close #12

ChDrive (Left$(App.Path, 1))
ChDir App.Path & "\TempFiles"

Call ShellAndWait("do_stuff.bat", True, vbNormalFocus)
DoEvents
ChDrive (Left$(RoutePath, 1))
ChDir RoutePath

Call ShellAndWait("EZStuff4.bat", True, vbNormalFocus)
DoEvents

NewFile = FreeFile
Open RoutePath & "\CleanUp.bat" For Append As #NewFile
Print #NewFile, "If not exist c:\windows\command\deltree.exe goto tryxp"
Print #NewFile, "Deltree /y " & ChrW$(34) & "Stuffed" & ChrW$(34)
Print #NewFile, "GoTo CarryOn"
Print #NewFile, ":tryxp"
Print #NewFile, "RD /s /q " & ChrW$(34) & "Stuffed" & ChrW$(34)
Print #NewFile, ":CarryOn"
Print #NewFile, "echo *"
Print #NewFile, "echo *"
Print #NewFile, "echo All done - Please close any open DOS windows"
Print #NewFile, "echo *"
Close #NewFile
Call ShellAndWait("CleanUp.bat", True, vbNormalFocus)
DoEvents

Kill RoutePath & "\ezstuff4.bat"
Kill RoutePath & "\cleanup.bat"
End Sub



Private Sub Command46_Click()
Dim x As Integer, strNewRoute As String, strRP As String, strNewBat As String, Y As Integer
Dim Filpath1$, NewFile As Integer, flagway As Integer, strIntName As String
Dim booExists As Boolean, strNewIntName As String, strDrv As String, strNewPath As String

MousePointer = 11
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If
NewFolder:
strNewRoute = InputBox("Enter new route folder name", "Duplicate Route", RouteName)
If strNewRoute = vbNullString Then
MousePointer = 0
Exit Sub
End If
Y = InStr(strNewRoute, " ")
If Y > 0 Then
Call MsgBox("This option will not work if there are spaces in the folder name.", vbExclamation, App.Title)
GoTo NewFolder
End If
If strNewRoute = RouteName Then
Call MsgBox("Your new route folder name must be different" _
            & vbCrLf & "to the original folder name. Please try again." _
            , vbInformation, App.Title)
GoTo NewFolder
End If
Call FindRouteInternalName(RoutePath & "\" & RouteName & ".trk", strIntName, booExists)
strNewIntName = InputBox("Enter new Route Name", "Route Name for Display", strIntName)
If strNewIntName = vbNullString Then
MousePointer = 0
Exit Sub
End If
If Not DirExists(strRP & strNewRoute) Then
MkDir strRP & strNewRoute
End If
strDrv = Left$(RoutePath, 2)
strNewBat = strDrv & vbCrLf
strNewBat = strNewBat & "chdir " & ChrW$(34) & RoutePath & ChrW$(34) & vbCrLf
strNewBat = strNewBat & "call xcopy " & "*.* " & ChrW$(34) & "..\" & strNewRoute & "\" & ChrW$(34) & " /s /y" & vbCrLf
NewFile = FreeFile

  Open Filpath1$ & "\newbat.bat" For Output As #NewFile
  Print #NewFile, strNewBat
  Close #NewFile
  
    Call ShellAndWait(ChrW$(34) & Filpath1$ & "\newbat.bat" & ChrW$(34), True, vbNormalFocus)
    DoEvents
    ChDrive strDrv
    
    ChDir strRP & strNewRoute
    
    If FileExists(strRP & strNewRoute & "\" & RouteName & ".ace") Then
    Name RouteName & ".ace" As strNewRoute & ".ace"
    End If
    If FileExists(strRP & strNewRoute & "\" & RouteName & ".mkr") Then
    Name RouteName & ".mkr" As strNewRoute & ".mkr"
    End If
    If FileExists(strRP & strNewRoute & "\" & RouteName & ".rdb") Then
    Name RouteName & ".rdb" As strNewRoute & ".rdb"
    End If
    If FileExists(strRP & strNewRoute & "\" & RouteName & ".rit") Then
    Name RouteName & ".rit" As strNewRoute & ".rit"
    End If
    If FileExists(strRP & strNewRoute & "\" & RouteName & ".ref") Then
    Name RouteName & ".ref" As strNewRoute & ".ref"
    End If
    If FileExists(strRP & strNewRoute & "\" & RouteName & ".tdb") Then
    Name RouteName & ".tdb" As strNewRoute & ".tdb"
    End If
    If FileExists(strRP & strNewRoute & "\" & RouteName & ".tit") Then
    Name RouteName & ".tit" As strNewRoute & ".tit"
    End If
    If FileExists(strRP & strNewRoute & "\" & RouteName & ".trk") Then
    Name RouteName & ".trk" As strNewRoute & ".trk"
    End If
    flagway = 0
    
    Call ConvertTrk3(strRP & strNewRoute & "\" & strNewRoute & ".trk", RouteName, strNewRoute, strNewIntName)
'    flagway = 1
'    Call ConvertTrk2(strRP & strNewRoute & "\" & strNewRoute & ".trk", flagway, RouteName, strNewRoute, strNewIntName)
    
    Rem ******************* Fix Activities - Change Routename to strnewroute
    
    strNewPath = strRP & strNewRoute
   If DirExists(strNewPath & "\activities") Then
   strNewPath = strRP & strNewRoute
cursouind = 0
Drive1(cursouind).Drive = Left$(strNewPath, 2)
Dir1(cursouind).Path = strNewPath & "\Activities"
Text1(cursouind) = "*.act"
DoEvents

    File1(cursouind).Pattern = Text1(cursouind).Text
  For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
  
   flagway = 0
   Call ConvertAct(strNewPath & "\Activities\" & File1(cursouind).List(i), flagway, strNewRoute)
   flagway = 1
   Call ConvertAct(strNewPath & "\Activities\" & File1(cursouind).List(i), flagway, strNewRoute)
   End If
   Next i
   End If
   
 MousePointer = 0
 Call MsgBox("Route duplication has completed.", vbInformation, App.Title)
 
End Sub
Private Sub Command47_Click()
Select Case MsgBox(Lang(483) & vbCrLf, vbOKCancel + vbInformation + vbDefaultButton1, App.Title)

    Case vbOK
    Call ListAce
    
File1(cursouind).Refresh
tit$ = vbNullString

Close

    Case vbCancel
Exit Sub
End Select
File1(cursouind).Refresh

End Sub


Private Sub Command48_Click()

If tit$ = vbNullString Then
frmPicView.Show 1

     DoEvents
     
Else
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
    Filpath$ = File1(cursouind).Path
   If Right$(Filpath$, 1) = "\" Then
   Filpath$ = Left$(Filpath$, Len(Filpath$) - 1)
   End If

   strPicView = Filpath$ & "\" & File1(cursouind).List(i)
  frmPicView.Show 1
  
     DoEvents
     
   If booAbort = True Then
   booAbort = False
   Exit For
   End If
   End If
   Next i
   End If
File1(cursouind).Refresh
File1(curtarind).Refresh

End Sub


Private Sub Command49_Click(Index As Integer)
Select Case Index
Case 0
Dim strAceView As String, strTGASave As String
Dim Filpath1$, result As Integer
cursouind = 0
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
    Filpath$ = File1(cursouind).Path
   If Right$(Filpath$, 1) = "\" Then
   Filpath$ = Left$(Filpath$, Len(Filpath$) - 1)
   End If
Filpath1$ = File1(1).Path

   strAceView = Filpath$ & "\" & File1(cursouind).List(i)
   strTGASave = Filpath1$ & "\" & Left$(File1(0).List(i), Len(File1(0).List(i)) - 3) & "tga"
  result = AceToTgaSquare(strAceView, strTGASave)
  
   
   End If
   Next i
 Case 1

cursouind = 0
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
    Filpath$ = File1(cursouind).Path
   If Right$(Filpath$, 1) = "\" Then
   Filpath$ = Left$(Filpath$, Len(Filpath$) - 1)
   End If
Filpath1$ = File1(1).Path

   strAceView = Filpath$ & "\" & File1(cursouind).List(i)
   strTGASave = Filpath1$ & "\" & Left$(File1(0).List(i), Len(File1(0).List(i)) - 3) & "bmp"
  result = AceToBmp(strAceView, strTGASave)
  
   
   End If
   Next i
End Select

File1(cursouind).Refresh
File1(curtarind).Refresh
End Sub


Private Sub Command5_Click()
Dim i%, Filpath1$, ii As Integer
Dim strNew As String, booIsEnv As Boolean
Dim x As Integer, Y As Integer
Dim j As Long, itExists As Boolean
Dim NewFile As Integer, NewFile2 As Integer
Dim Newfile3 As Integer, booGotIt As Boolean
Dim q As Integer


'Set tfh = New TokenFileHandler
Select Case MsgBox("This operation will Compact your Route and move files which are not required to the folder RRBackups" _
                   & vbCrLf & "This may take quite a long time to run, so please confirm you wish to proceed." _
                   & vbCrLf & "Dynatrax files will NOT be moved." _
                   , vbOKCancel Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbOK
Rem - Do nothing
    Case vbCancel
Exit Sub
End Select
numShp = 0
numAce = 0
intCars = 0
For j = 0 To 1000
Soundfile(j) = vbNullString
WavFile(j) = vbNullString
Next
ReDim strShp(0 To Shp_Chunk)
ReDim strGlobShp(0 To Shp_Chunk)
ReDim Ace1(0 To Shp_Chunk)
ReDim Cars(0 To Car_CHUNK)
For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then
If frmUtils.Controls(q).Caption <> Lang(30) And frmUtils.Controls(q).Caption <> Lang(14) Then
frmUtils.Controls(q).Enabled = False
ElseIf frmUtils.Controls(q).Caption = Lang(30) Then

frmUtils.Controls(q).Caption = Lang(637)
End If
End If
Next q
SB1.Panels(4).Text = time
SB1.Panels(6).Text = vbNullString

strKillFiles = vbNullString

On Error GoTo Errtrap

If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If
Call ClearSetup
Call KillRubbish
strReport = vbNullString
'MousePointer = 11
intTempFile = 0
Filpath1$ = App.Path & "\TempFiles"
cursouind = 0


DoEvents
cursouind = 0
If FileExists(RoutePath & "\carspawn.dat") Then
NewFile = FreeFile
NewFile2 = FreeFile
Newfile3 = FreeFile

     Open RoutePath & "\carspawn.dat" For Input As #NewFile
      Do While Not EOF(NewFile)
      Line Input #NewFile, strNew
      
     strNew = Trim$(strNew)
    If Left$(strNew, 14) = "CarSpawnerItem" Then
    
   x = InStr(strNew, ChrW$(34))
   If x > 0 Then
   Y = InStr(x + 1, strNew, ChrW$(34))
   strNew = Mid$(strNew, x + 1, Y - (x + 1))
   strNew = Trim$(strNew)
   
   For Y = 0 To intCars
   If strNew = Cars(Y) Then
   itExists = True
   Exit For
   End If
   Next Y
   If itExists = False Then
   Cars(intCars) = strNew
   intCars = intCars + 1
   If intCars > UBound(Cars) Then
     ReDim Preserve Cars(0 To intCars + Car_CHUNK)
    End If
    
 End If
 itExists = False
   End If
   End If
   strNew = vbNullString
  itExists = False
   Loop
   Close #NewFile
 End If

 Rem*************************************
 If FileExists(RoutePath & "\telepole.dat") Then
 booHaz = False
  Call CompactCheckForS(RoutePath & "\telepole.dat", booHaz)
 End If
If FileExists(RoutePath & "\speedpost.dat") Then
 Call CompactCheckForAce(RoutePath & "\speedpost.dat")
 booHaz = False
 Call CompactCheckForS(RoutePath & "\speedpost.dat", booHaz)
 
 End If
 
 If FileExists(RoutePath & "\sigcfg.dat") Then
 Call CompactCheckForAce(RoutePath & "\sigcfg.dat")
 
 
 End If

 If FileExists(RoutePath & "\forests.dat") Then
' Call CheckForAce(RoutePath & "\forests.dat")
 End If
 
Rem ************ Check Envfiles ***************************
Rem ************* Show dialog for Move or Delete files ******

'frmUnwanted.Show 1
booKillMove = True
If booKillMove = True Then
strKillPath = RoutePath & "\RRBackups"
If Not DirExists(strKillPath) Then
MkDir strKillPath
End If
If Not DirExists(strKillPath & "\Shapes") Then
MkDir strKillPath & "\Shapes"
End If
If Not DirExists(strKillPath & "\Envfiles") Then
MkDir strKillPath & "\Envfiles"
End If
If Not DirExists(strKillPath & "\Envfiles\Textures") Then
MkDir strKillPath & "\Envfiles\Textures"
End If
If Not DirExists(strKillPath & "\Sound") Then
MkDir strKillPath & "\Sound"
End If
If Not DirExists(strKillPath & "\Textures") Then
MkDir strKillPath & "\Textures"
End If
If Not DirExists(strKillPath & "\Textures\Night") Then
MkDir strKillPath & "\Textures\Night"
End If
If Not DirExists(strKillPath & "\Textures\Snow") Then
MkDir strKillPath & "\Textures\Snow"
End If
If Not DirExists(strKillPath & "\Textures\Autumn") Then
MkDir strKillPath & "\Textures\Autumn"
End If
If Not DirExists(strKillPath & "\Textures\AutumnSnow") Then
MkDir strKillPath & "\Textures\AutumnSnow"
End If
If Not DirExists(strKillPath & "\Textures\Spring") Then
MkDir strKillPath & "\Textures\Spring"
End If
If Not DirExists(strKillPath & "\Textures\SpringSnow") Then
MkDir strKillPath & "\Textures\SpringSnow"
End If
If Not DirExists(strKillPath & "\Textures\Winter") Then
MkDir strKillPath & "\Textures\Winter"
End If
If Not DirExists(strKillPath & "\Textures\WinterSnow") Then
MkDir strKillPath & "\Textures\WinterSnow"
End If
If Not DirExists(strKillPath & "\Terrtex") Then
MkDir strKillPath & "\Terrtex"
End If
If Not DirExists(strKillPath & "\Terrtex\Snow") Then
MkDir strKillPath & "\Terrtex\Snow"
End If
DoEvents
End If
'*********************************************************
cursouind = 1
Drive1(cursouind).Drive = Left$(EnvPath, 2)
Dir1(cursouind).Path = EnvPath
Text1(cursouind) = "*.env"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
For ii = 1 To 13
    If File1(cursouind).List(i) = strEnv(ii) Then
    booIsEnv = True
    End If
    Next ii
    If booIsEnv = True Then
    booIsEnv = False
    Else
        If booKillMove = False Then
        Kill EnvPath & "\" & File1(cursouind).List(i)
        ElseIf booKillMove = True Then
        FileCopy EnvPath & "\" & File1(cursouind).List(i), strKillPath & "\Envfiles\" & File1(cursouind).List(i)
        DoEvents
        Kill EnvPath & "\" & File1(cursouind).List(i)
        End If
    
    strKillFiles = strKillFiles & EnvPath & "\" & File1(cursouind).List(i) & vbCrLf
    
    End If
Next i
Dir1(cursouind).Path = EnvPath
Text1(cursouind) = "*.*"
DoEvents
Text1(cursouind) = "*.env"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i


For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
   Call CheckEnvForAce(EnvPath & "\" & File1(cursouind).List(i))
   
 End If
Next i

Drive1(cursouind).Drive = Left$(EnvPath, 2)
Dir1(cursouind).Path = EnvPath & "\Textures"
Text1(cursouind) = "*.ace"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
      For j = 1 To envAceNumber
   If File1(cursouind).List(i) = EnvAceFile(j) Then
   booGotIt = True
   GoTo GetAnother
   End If
   
   Next
   If booKillMove = False Then
   Kill EnvPath & "\Textures" & "\" & File1(cursouind).List(i)
   ElseIf booKillMove = True Then
   FileCopy EnvPath & "\Textures" & "\" & File1(cursouind).List(i), strKillPath & "\Envfiles\Textures\" & File1(cursouind).List(i)
   DoEvents
   Kill EnvPath & "\Textures" & "\" & File1(cursouind).List(i)
   End If
   strKillFiles = strKillFiles & EnvPath & "\Textures" & "\" & File1(cursouind).List(i) & vbCrLf
GetAnother:
booGotIt = False
   Next
   
   

Call CompactRoute


MousePointer = 0


DoEvents
For i = 0 To numShp - 1
strReport = strReport & strShp(i) & vbCrLf
Next i
strReport = strReport & vbCrLf & vbCrLf & "Total unique shapes = " & Str(numShp) & vbCrLf & vbCrLf

Call CompactACESMS
DoEvents

MousePointer = 0
 Select Case MsgBox(Lang(484) & vbCrLf & Lang(485), vbYesNo + vbExclamation + vbDefaultButton1, App.Title)

    Case vbYes
   
 frmReport.Rich1.Text = strReport
 frmReport.Show 1
 
     DoEvents
     

    Case vbNo
booReport = True
End Select

Call KillUnused
MousePointer = 0


SB1.Panels(6).Text = time
If booKillMove = False Then
strKillFiles = Lang(486) & vbCrLf & vbCrLf & strKillFiles
ElseIf booKillMove = True Then
strKillFiles = "The following files have been moved into " & strKillPath & vbCrLf & vbCrLf & strKillFiles
End If

frmReport.Rich1.Text = strKillFiles
 frmReport.Show 1
 
     DoEvents
     
For q = 0 To frmUtils.Controls.Count - 1
If TypeOf frmUtils.Controls(q) Is CommandButton Or TypeOf frmUtils.Controls(q) Is SSTab Then frmUtils.Controls(q).Enabled = True
Next q
Command1(15).Caption = Lang(30)
Dir1(0).Path = RoutePath
Text1(0).Text = "*.*"

Exit Sub
Errtrap:

If Err = 75 Then
MsgBox Lang(448) & vbCr & Lang(449), 48, Lang(450)
'********************
Resume Next
End If
Call MsgBox("An error " & Err & " occurred in subroutine 'Compact Route' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End Sub

Private Sub Command50_Click()
Dim intDelay As Integer

intDelay = Val(Text2)
tit$ = Dir1(0).Path

If tit$ <> vbNullString Then

Call WinSlideShow(tit$, intDelay, 0, 0, 0, 0)
End If
End Sub

Private Sub Command51_Click()
Dim strBatText As String

If tit$ = vbNullString Then
frmPicView.Show 1

     DoEvents
     
Else
      For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
 
    Filpath$ = File1(cursouind).Path
   If Right$(Filpath$, 1) = "\" Then
   Filpath$ = Left$(Filpath$, Len(Filpath$) - 1)
   End If

   strPicView = Filpath$ & "\" & File1(cursouind).List(i)

        
    strBatText = ChrW$(34) & App.Path & "\sviewRR.exe" & ChrW$(34) & " " & ChrW$(34) & strPicView & ChrW$(34) & ";"
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)
    
    End If
    Next i
   End If
  
   
End Sub

Private Sub Command52_Click()
Dim strAceView As String, strTGASave As String
Dim Filpath1$, result As Integer
MousePointer = 11
cursouind = 0
Label11.Visible = True
If Not DirExists(App.Path & "\TempFiles") Then
MkDir App.Path & "\TempFiles"
End If
FileCopy App.Path & "\AceIt.exe", App.Path & "\tempfiles\AceIt.exe"
Kill App.Path & "\TempFiles\*.*"
DoEvents
FileCopy App.Path & "\AceIt.exe", App.Path & "\tempfiles\AceIt.exe"
For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
    Filpath$ = File1(cursouind).Path
   If Right$(Filpath$, 1) = "\" Then
   Filpath$ = Left$(Filpath$, Len(Filpath$) - 1)
   End If
Filpath1$ = App.Path & "\TempFiles"
        If File1(cursouind).List(i) = "loco.ace" Then GoTo CarryON

   strAceView = Filpath$ & "\" & File1(cursouind).List(i)
   Call readFile(strAceView, bdata())
   If bdata(16) = 17 Then GoTo CarryON
   If bdata(8) <> 0 Or bdata(12) <> 0 Then GoTo CarryON
   If bdata(9) <> bdata(13) Then GoTo CarryON
   strTGASave = Filpath1$ & "\" & Left$(File1(0).List(i), Len(File1(0).List(i)) - 3) & "tga"
   
   Label11.Caption = "Converting: " & File1(cursouind).List(i) & " to TGA"
   DoEvents
 result = AceToTgaSquare(strAceView, strTGASave)
  
  ' Result = AceToTga(strAceView, strTGASave)
   End If
CarryON:
   Next i
  
   Call ConvTGAtoACE(Filpath$)
 MousePointer = 0
 File1(1).Refresh
 Label11.Visible = False
End Sub

Private Sub Command53_Click()
Text1(0).Text = "*.ace"
End Sub

Private Sub Command54_Click()
Dim strFirst As String

strFirst = Dir1(0).Path

comAbortFil.Visible = True
frmSearch.Show

     DoEvents
     
DoEvents
If booAbort = True Then
booAbort = False
Unload frmSearch
End If
comAbortFil.Visible = False
DoEvents
Dir1(0).Path = strFirst

End Sub

Private Sub Command55_Click()
Dim strPath As String, flagway As Integer, strAI As String, strDummy As String
Dim NewFile As Integer, strType As String, Y As Integer, yy As Integer, y1 As Integer
Dim y2 As Integer

yy = 1
If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
SparePath = App.Path & "\TempFiles"
Select Case MsgBox(Lang(500) & vbCrLf & Lang(501), vbOKCancel Or vbExclamation Or vbDefaultButton1, App.Title)

    Case vbOK
GoTo MakeWag
    Case vbCancel
Exit Sub
End Select

MakeWag:
For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
  
   strPath = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   If Right$(strPath, 4) <> ".eng" Then
   Call MsgBox(Lang(393) & File1(cursouind).List(i) & vbCrLf & Lang(456), vbExclamation, App.Title)
   
   GoTo CarryON
   End If
   NewFile = FreeFile
   Open strPath For Input As #NewFile
  Do While Not EOF(NewFile)
TryAgain:
  Line Input #NewFile, A$

  Y = InStr(A$, "Type ")
  If Y > 0 Then
 
  y1 = InStr(Y, A$, "(")
  y2 = InStr(Y, A$, ")")
  strType = Trim$(Mid$(A$, y1 + 1, (y2 - y1) - 1))
  
  If strType <> "Diesel" And strType <> "Electric" And strType <> "Steam" Then
  GoTo TryAgain
  Else
  Close #NewFile
  Exit Do
  End If
  End If
  Loop
   If strType = "Diesel" Or strType = "Electric" Then
   strDummy = "$" & File1(cursouind).List(i)
   strDummy = Left$(strDummy, Len(strDummy) - 3) & "wag"
   strAI = SparePath & "\" & strDummy
   FileCopy strPath, strAI
   flagway = 0
   Call ConvertDummy(strAI, flagway)
   DoEvents
   flagway = 1
   Call ConvertDummy(strAI, flagway)
   DoEvents
   FileCopy strAI, File1(cursouind).Path & "\" & strDummy
   DoEvents
   Kill strAI
   Call MsgBox(Lang(502) & vbCrLf & Lang(503) & strDummy, vbExclamation, App.Title)
   
   ElseIf strType = "Steam" Then
   
   strDummy = "$" & File1(cursouind).List(i)
   strAI = SparePath & "\" & strDummy
   FileCopy strPath, strAI
   flagway = 0
   Call ConvertDummy2(strAI, flagway)
   DoEvents
   
   flagway = 1
   Call ConvertDummy2(strAI, flagway)
   DoEvents
   FileCopy strAI, File1(cursouind).Path & "\" & strDummy
   DoEvents
   Kill strAI
   Call MsgBox(Lang(504) & vbCrLf & Lang(503) & strDummy, vbExclamation, App.Title)
   
   End If
   End If
CarryON:
   Next i
   
   Text1(cursouind) = "*.wag"
   DoEvents
   Text1(cursouind) = "*.*"
   
End Sub


Private Sub Command56_Click(Index As Integer)

Dim strBatText As String, i As Integer

Select Case Index
Case 0
    Filpath$ = File1(cursouind).Path
        
   strBatText = ChrW$(34) & App.Path & "\tview.exe" & ChrW$(34) & " " & ChrW$(34) & Filpath$ & ChrW$(34)
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)
    Case 1
   
    For i = 0 To File1(cursouind).ListCount - 1
  If File1(cursouind).Selected(i) Then
 
    Filpath$ = File1(cursouind).Path
   If Right$(Filpath$, 1) = "\" Then
   Filpath$ = Left$(Filpath$, Len(Filpath$) - 1)
   End If

   strPicView = Filpath$ & "\" & File1(cursouind).List(i)

        
   strBatText = ChrW$(34) & App.Path & "\tgatool2a.exe" & ChrW$(34) & " " & ChrW$(34) & strPicView & ChrW$(34)
   ChDrive Left$(App.Path, 1)
   ChDir App.Path
  
    Call ShellAndWait(strBatText, True, vbNormalFocus)
    
    End If
    Next i
  
    End Select
End Sub


Private Sub Command58_Click()
Dim strRefFile As String, NewFile As Integer, strTemp As String

strRefFile = Label2(0).Caption
If strRefFile = vbNullString Or Right$(strRefFile, 4) <> ".ref" Then
Call MsgBox(Lang(628), vbExclamation, App.Title)

Exit Sub
End If
FileCopy strRefFile, strRefFile & ".bak"
NewFile = FreeFile
   Open strRefFile For Binary As #NewFile
    strTemp = String(2, " ")
    Get #NewFile, , strTemp
 Close #NewFile
 
 If Asc(Mid$(strTemp, 1, 1)) <> 255 Then
 If Asc(Mid$(strTemp, 2, 1)) <> 254 Then

Call ConvertIt(strRefFile, 1)
DoEvents
End If
End If
MousePointer = 11

frmReadRef.Show

     DoEvents
 MousePointer = 0

End Sub

Private Sub Command59_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim result As String


Select Case Button
Case 1

Command59.Caption = "Con"
result = MSTSPath
Text1(Index) = "*.*"
Drive1(Index).Drive = Left$(result, 2)
Dir1(Index).Path = result & "\Trains\Consists"
Case 2

Command59.Caption = "TSet"
result = MSTSPath
Text1(Index) = "*.*"
Drive1(Index).Drive = Left$(result, 2)
Dir1(Index).Path = result & "\Trains\trainset"
End Select
End Sub


Private Sub Command6_Click(Index As Integer)
Dim strPath As String, flagway As Integer

MousePointer = 11
If tit$ = vbNullString Then
MSG = Lang(408)
Response = MsgBox(MSG, 0, Lang(407))
Exit Sub
End If
For i = 0 To File1(cursouind).ListCount - 1
   If File1(cursouind).Selected(i) Then
   strPath = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   flagway = Index
   Call ConvertIt(strPath, flagway)
   End If
   Next i
MousePointer = 0
End Sub

Private Sub Command60_Click()
Dim strSFile As String, Filpath1$, strShapeName As String, x As Integer
Dim strSpc As String, strShapePath As String, MyString As String


'Set tfh = New TokenFileHandler
strReport = vbNullString
If Right$(Label2(0).Caption, 2) <> ".s" Then
Call MsgBox("You must select an .S file in the Left Hand file window", vbExclamation, App.Title)

Exit Sub
End If
Filpath1$ = App.Path & "\TempFiles"
strSFile = Label2(0).Caption
x = InStrRev(strSFile, "\")
strShapeName = Mid$(strSFile, x + 1)
strShapePath = Left$(strSFile, x - 1)
Open strSFile For Binary As #5
    strSpc = String(2, " ")
    Get #5, , strSpc
 Close #5
 
 If Not (Asc(Mid$(strSpc, 1, 1)) = 255 And Asc(Mid$(strSpc, 2, 1)) = 254) Then
   ' result = tfh.decompress(fullpath$, fullpath$)
   TokMode = 0
     booWriteFile = True
   Call DoDeComp2(strShapeName, strShapePath, Filpath1$)
    Else
    FileCopy strSFile, Filpath1$ & "\" & strShapeName
   End If
  
'result = tfh.decompress(strSFile, Filpath1$ & strShapeName)
DoEvents
MyString = ReadUniFile(Filpath1$ & "\" & strShapeName)
'Call ConvertIt(Filpath1$ & strShapeName, 0)
'DoEvents
 
Call CheckForAce5(MyString, strShapeName)
DoEvents
frmReport.Rich1.Text = strReport
 
     frmReport.Show 1
     
     DoEvents
     
End Sub

Private Sub Command7_Click()
Dim booExists As Boolean, NewRouteName As String, Filpath1$, strTemp As String, j As Integer, A$
Dim strLogPath As String, NewFile As Integer, x As Integer, i As Integer, SparePath As String

On Error GoTo Errtrap
strLogPath = App.Path & "\Reports\Startup.log"


cursouind = 0
Filpath1$ = App.Path & "\setupfiles"
If FileExists(App.Path & "\setupfiles\master.ref") Then
Kill App.Path & "\setupfiles\master.ref"
End If
If FileExists(App.Path & "\setupfiles\InstallMe.bat") Then
Kill App.Path & "\setupfiles\InstallMe.bat"
End If
If Not DirExists(App.Path & "\TempFiles") Then
   MkDir App.Path & "\TempFiles"
End If
Kill App.Path & "\TempFiles\*.*"
DoEvents
If Not DirExists(App.Path & "\TempFiles2") Then
   MkDir App.Path & "\TempFiles2"
End If
Kill App.Path & "\TempFiles2\*.*"
DoEvents
Open App.Path & "\SetupFiles\Installme.bat" For Append As #12
Print #12, "@Echo Off"
Close #12
MousePointer = 11

RoutePath = File1(cursouind).Path
MasterRoutePath = RoutePath
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
End If
Close
NewFile = FreeFile
Open strLogPath For Append As #NewFile
A$ = "Changed main screen caption"
Print #NewFile, A$
For j = 0 To NumRoutes - 1
If RoutePath = AllRoutes(j) Then
MainRoute = j
Exit For
End If
Next
A$ = "Mainroute = " & Str(j)
Print #NewFile, A$
frmUtils.Show

Text1(0) = "*.trk"
If File1(cursouind).ListCount = 0 Then
Text1(0) = "*.off"
    If File1(cursouind).ListCount = 1 Then
    Name RoutePath & "\" & File1(cursouind).List(0) As Left(RoutePath & "\" & File1(cursouind).List(0), Len(RoutePath & "\" & File1(cursouind).List(0)) - 3) & "trk"
    Text1(0) = "*.trk"
    DoEvents
    End If
   
If File1(cursouind).ListCount = 0 Then
Text1(0) = "*.tdb"
DoEvents
If File1(cursouind).ListCount = 0 Then
Call MsgBox(Lang(431) & vbCrLf & Lang(432), vbExclamation, Lang(407))
Text1(0) = "*.*"
MousePointer = 0
Exit Sub
Else
Call MsgBox(Lang(635) & vbCrLf & Lang(636), vbExclamation, App.Title)

Text1(0) = "*.*"
MousePointer = 0
Exit Sub
End If
End If
End If

RouteName = File1(cursouind).List(i)
OldRouteName = RouteName

Call CheckForSMS(RoutePath & "\" & RouteName)

Call FindRouteName(RoutePath & "\" & RouteName, NewRouteName, booExists)
If booExists = False Then

'MsgBox Lang(339) & vbcr & Lang(340), 16, Lang(341)
MousePointer = 0

Exit Sub
Else
RouteName = NewRouteName
End If

Call FindRouteID(RoutePath & "\" & OldRouteName)
Call IsItElectric(RoutePath & "\" & OldRouteName, booElectric)

OriginalRef = RoutePath & "\" & RouteName & ".ref"
A$ = "OriginalRef = " & OriginalRef
Print #NewFile, A$
If FileExists(OriginalRef) Then
FileCopy RoutePath & "\" & RouteName & ".ref", App.Path & "\setupfiles\master.ref"
Else
flagNoRef = True
Call MsgBox(Lang(385) & vbCrLf & Lang(386), vbExclamation, App.Title)
If FileExists(App.Path & "\stuffit.ref") Then
FileCopy App.Path & "\stuffit.ref", App.Path & "\setupfiles\master.ref"
Else
Call MsgBox("The file 'stuffit.ref' is missing from your Route_Riter folder, please reinstall it.", vbExclamation, App.Title)
Exit Sub
End If
End If
A$ = "Stuffit.ref must have been found"
Print #NewFile, A$

SaveSetting "Route_Riter6", "Files", "File", RoutePath
DoEvents
SaveSetting "Route_Riter6", "MainPath", "MainPath", MSTSPath
DoEvents
A$ = "Settings saved in registry"
Print #NewFile, A$
RouteListed = True
TexturePath = RoutePath & "\Textures"
If Not DirExists(TexturePath) Then MkDir TexturePath
TexSnowPath = RoutePath & "\Textures\Snow"
If Not DirExists(TexSnowPath) Then MkDir TexSnowPath
TexNightPath = RoutePath & "\Textures\Night"
If Not DirExists(TexNightPath) Then MkDir TexNightPath
TexAutPath = RoutePath & "\Textures\Autumn"
If Not DirExists(TexAutPath) Then MkDir TexAutPath
TexAutSnowPath = RoutePath & "\Textures\AutumnSnow"
If Not DirExists(TexAutSnowPath) Then MkDir TexAutSnowPath
TexSprPath = RoutePath & "\Textures\Spring"
If Not DirExists(TexSprPath) Then MkDir TexSprPath
TexSprSnowPath = RoutePath & "\Textures\SpringSnow"
If Not DirExists(TexSprSnowPath) Then MkDir TexSprSnowPath
TexWinPath = RoutePath & "\Textures\Winter"
If Not DirExists(TexWinPath) Then MkDir TexWinPath
TexWinSnowPath = RoutePath & "\Textures\WinterSnow"
If Not DirExists(TexWinSnowPath) Then MkDir TexWinSnowPath
TilePath = RoutePath & "\Tiles"
ShapePath = RoutePath & "\Shapes"
SoundPath = RoutePath & "\Sound"
A$ = "Texturepath = " & TexturePath
A$ = A$ & vbCrLf & "Tilepath = " & TilePath
A$ = A$ & vbCrLf & "Shapepath = " & ShapePath
A$ = A$ & vbCrLf & "Soundpath = " & SoundPath
Print #NewFile, A$

If Not DirExists(SoundPath) Then MkDir SoundPath

WorldPath = RoutePath & "\World"
EnvPath = RoutePath & "\Envfiles"
If Not DirExists(EnvPath) Then MkDir EnvPath
If Not DirExists(EnvPath & "\Textures") Then MkDir EnvPath & "\Textures"
If Not FileExists(RoutePath & "\" & RouteName & ".rdb") Then
strTemp = RoutePath & "\" & RouteName & ".rdb" & Lang(578) & Lang(579)
End If
If Not FileExists(RoutePath & "\" & RouteName & ".rit") Then
strTemp = strTemp & vbCrLf & RoutePath & "\" & RouteName & ".rit" & Lang(578) & Lang(579)
End If
If Not FileExists(RoutePath & "\" & RouteName & ".tdb") Then
strTemp = strTemp & vbCrLf & RoutePath & "\" & RouteName & ".tdb" & Lang(578)
End If
If Not FileExists(RoutePath & "\" & RouteName & ".tit") Then
strTemp = strTemp & vbCrLf & RoutePath & "\" & RouteName & ".tit" & Lang(578)
End If
If Not FileExists(RoutePath & "\tsection.dat") Then
strTemp = strTemp & vbCrLf & RoutePath & "\tsection.dat" & Lang(578) & Lang(580)
End If
If Not FileExists(TexturePath & "\acleantrack1.ace") Then
strTemp = strTemp & vbCrLf & TexturePath & "\acleantrack1.ace" & Lang(578) & Lang(581)
End If
If Not FileExists(TexturePath & "\acleantrack2.ace") Then
strTemp = strTemp & vbCrLf & TexturePath & "\acleantrack2.ace" & Lang(578) & Lang(581)
End If
If Not FileExists(TexSnowPath & "\acleantrack1.ace") Then
strTemp = strTemp & vbCrLf & TexSnowPath & "\acleantrack1.ace" & Lang(578) & Lang(582)
End If
If Not FileExists(TexSnowPath & "\acleantrack2.ace") Then
strTemp = strTemp & vbCrLf & TexSnowPath & "\acleantrack2.ace" & Lang(578) & Lang(582)
End If
If strTemp <> vbNullString Then
strTemp = Lang(583) & vbCrLf & vbCrLf & strTemp
' frmReport.Rich1.Text = strTemp
'
'     frmReport.Show 1
     strMainReport = strTemp
     
     DoEvents
 Else
 strMainReport = vbNullString

End If

Text1(0) = "*.*"
Rem ********** Delete any spare .w files ****
A$ = "Delete spare files"
Print #NewFile, A$
    cursouind = 1
    SparePath = App.Path & "\TempFiles"
    
    Call KillSpare("*.w")
DoEvents
    Call KillSpare("*.s")
DoEvents
    Call KillSpare("*.t")
DoEvents
Close
NewFile = FreeFile
Open strLogPath For Append As #NewFile
A$ = "Spare files deleted"
Print #NewFile, A$
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
Rem *******************
MousePointer = 0
cursouind = 0
A$ = "Setup Complete......"
Print #NewFile, A$
Close #NewFile
Exit Sub
Errtrap:
If Err = 53 Then
Resume Next
End If
Call MsgBox("Error " & Err.Description & " occurred while Confirming route " & RouteName, vbExclamation, App.Title)

Resume Next
End Sub
Private Sub Command8_Click()
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
'********************
Exit Sub
End If

Call CompressSFiles
End Sub





Private Sub Command90_Click()

booMini = True

frmCopy.Show
DoEvents



'booMini = False
End Sub

Private Sub Command91_Click()

Dim TrainsetPath As String, j As Long, strDir As String
Dim x As Integer, strBatFile As String
MousePointer = 11
TrainsetPath = MSTSPath & "\Trains\Trainset\"
frmMini3.Show 1
DoEvents
If strEditPath = vbNullString Then Exit Sub
cursouind = 0
Drive1(cursouind).Drive = Left$(TrainsetPath, 1)
Dir1(cursouind).Path = TrainsetPath


For j = 0 To Dir1(cursouind).ListCount - 1
x = InStrRev(Dir1(cursouind).List(j), "\")
strDir = Mid$(Dir1(cursouind).List(j), x + 1)
If Not DirExists(strEditPath & "\" & strDir) Then
MkDir strEditPath & "\" & strDir
End If
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & Dir1(cursouind).List(j) & "\*.eng" & ChrW$(34) & " " & ChrW$(34) & strEditPath & "\" & strDir & ChrW$(34) & " /S /Y" & vbCrLf
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & Dir1(cursouind).List(j) & "\*.wag" & ChrW$(34) & " " & ChrW$(34) & strEditPath & "\" & strDir & ChrW$(34) & " /S /Y" & vbCrLf
strBatFile = strBatFile & "call Xcopy " & ChrW$(34) & Dir1(cursouind).List(j) & "\*.sd" & ChrW$(34) & " " & ChrW$(34) & strEditPath & "\" & strDir & ChrW$(34) & " /S /Y" & vbCrLf
DoEvents
Next j


If strBatFile <> vbNullString Then

Open App.Path & "\TempFiles\mini.bat" For Output As #1
Print #1, strBatFile
Close #1
ChDrive Left$(App.Path, 1)
 ChDir App.Path & "\TempFiles"

DoEvents
Call ShellAndWait("mini.bat", True, vbNormalFocus)

DoEvents
End If

MousePointer = 0
End Sub

Private Sub Command92_Click()
Dim strFirst As String, strPath As String

strPath = Dir1(0).Path
If Right$(strPath, 8) <> "Trainset" Then
Call MsgBox("You must select a Trainset folder to use" _
            & vbCrLf & "                  with this option." _
            , vbExclamation, App.Title)

Exit Sub
End If

strReport = vbNullString
strFirst = Dir1(0).Path
'Call CountStock
strBBoxFix = InputBox("Lower Limit of BoundingBox", , "0.9")
DoEvents
Drive1(0).Drive = Left$(MSTSPath, 2)
Dir1(0).Path = strPath
Text1(0) = "*.sd"
frmUtils.Refresh
DoEvents


booFixBB = True

frmSearch.Show
DoEvents
Dir1(0).Path = strFirst
booFixSD = False
End Sub

Private Sub Command93_Click()
Dim Filpath1$, strRP As String, x As Integer, strBatText As String


Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
MousePointer = 11
If Not DirExists(App.Path & "\Reports") Then
MkDir App.Path & "\Reports"
End If

Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If
Select Case MsgBox("Do you wish tiles with no corresponding .w files to be unselected? (Take care with this as tiles covering the sea etc may be removed as they appear to be empty).", vbYesNo Or vbExclamation Or vbDefaultButton1, "TsUtils Filter Option")

    Case vbYes
 strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_filter.log" & ChrW$(34) & "  filter -w " & ChrW$(34) & RoutePath & ChrW$(34)
    Case vbNo
 strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_filter.log" & ChrW$(34) & "  filter " & ChrW$(34) & RoutePath & ChrW$(34)
End Select

ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
  

 Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 If Not FileExists(App.Path & "\Reports\" & RouteName & "_filter.log") Then
Call MsgBox("As no .log file has been created, you can assume that TsUtils" _
            & vbCrLf & "did not complete the option you were attempting to run." _
            , vbCritical, App.Title)

MousePointer = 0
Exit Sub
End If
 
 
  MousePointer = 0
 frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_filter.log"
 frmReport.Show 1
 DoEvents
 Call MsgBox("A new 'TD' folder has been placed within the 'newRoute' folder in your route. Back up your TD folder and copy the new TD folder contents into it.", vbInformation, App.Title)
 

End Sub



Private Sub Command94_Click()
Dim Filpath1$, strBatText As String
Dim RouteTwoPath As String
On Error GoTo Errtrap
Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
MousePointer = 11
If Not DirExists(App.Path & "\Reports") Then
MkDir App.Path & "\Reports"
End If

Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
RoutePath = File1(cursouind).Path
RouteTwoPath = File1(1).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath


 strBatText = "java -Xmx256m TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_merge.log" & ChrW$(34) & "  merge  " & ChrW$(34) & RoutePath & ChrW$(34) & " " & ChrW$(34) & RouteTwoPath & ChrW$(34)
 Call CheckMissFolders(RoutePath)
DoEvents

ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
  

 Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 If Not FileExists(App.Path & "\Reports\" & RouteName & "_merge.log") Then
Call MsgBox("As no .log file has been created, you can assume that TsUtils" _
            & vbCrLf & "did not complete the option you were attempting to run." _
            , vbCritical, App.Title)

MousePointer = 0
Exit Sub
End If
 
 
  MousePointer = 0
 frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_merge.log"
 frmReport.Show 1
 DoEvents
Call MsgBox("Your route MERGE appears to have been successful. The new files are in the NewRoute folder within your first route. You may now replace the corresponding folders  and files in your original route with the equivalent folders and files from within NewRoute (but make a backup first).", vbExclamation, App.Title)

 Exit Sub
Errtrap:
 Call MsgBox("An error " & Err & " " & Err.Description & " occurred in subroutine 'TsUtils Merge' Path1=" & RoutePath _
            & vbCrLf & "Path2= " & RouteTwoPath _
            , vbExclamation, App.Title)

End Sub



Private Sub Command95_Click()
Dim strBatText As String, fullpath$, strOrigFile As String, i As Integer, NewFile As Integer


On Error GoTo Errtrap

If Not FileExists(App.Path & "\TempFiles2\ffeditc_unicode.exe") Then
    If DirExists(MSTSPath & "\utils\ffedit") Then
    FileCopy MSTSPath & "\utils\ffedit\appids.tok", App.Path & "\TempFiles2\appids.tok"
    FileCopy MSTSPath & "\utils\ffedit\coreids.tok", App.Path & "\TempFiles2\coreids.tok"
    FileCopy MSTSPath & "\utils\ffedit\ffedit.cfg", App.Path & "\TempFiles2\ffedit.cfg"
    FileCopy MSTSPath & "\utils\ffedit\ffeditc_unicode.exe", App.Path & "\TempFiles2\ffeditc_unicode.exe"
    FileCopy MSTSPath & "\utils\ffedit\forms.hdr", App.Path & "\TempFiles2\forms.hdr"
    FileCopy MSTSPath & "\utils\ffedit\loadstr.hdr", App.Path & "\TempFiles2\loadstr.hdr"
    FileCopy MSTSPath & "\utils\ffedit\sidn.txt", App.Path & "\TempFiles2\sidn.txt"
    FileCopy MSTSPath & "\utils\ffedit\worldfile.bnf", App.Path & "\TempFiles2\worldfile.bnf"
    FileCopy MSTSPath & "\utils\ffedit\newshape.bnf", App.Path & "\TempFiles2\newshape.bnf"
    Else
    
    result = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "Path")
    
        If DirExists(result & "\utils\ffedit") Then
        FileCopy result & "\utils\ffedit\appids.tok", App.Path & "\TempFiles2\appids.tok"
        FileCopy result & "\utils\ffedit\coreids.tok", App.Path & "\TempFiles2\coreids.tok"
        FileCopy result & "\utils\ffedit\ffedit.cfg", App.Path & "\TempFiles2\ffedit.cfg"
        FileCopy result & "\utils\ffedit\ffeditc_unicode.exe", App.Path & "\TempFiles2\ffeditc_unicode.exe"
        FileCopy result & "\utils\ffedit\forms.hdr", App.Path & "\TempFiles2\forms.hdr"
        FileCopy result & "\utils\ffedit\loadstr.hdr", App.Path & "\TempFiles2\loadstr.hdr"
        FileCopy result & "\utils\ffedit\sidn.txt", App.Path & "\TempFiles2\sidn.txt"
        FileCopy result & "\utils\ffedit\worldfile.bnf", App.Path & "\TempFiles2\worldfile.bnf"
        FileCopy result & "\utils\ffedit\newshape.bnf", App.Path & "\TempFiles2\newshape.bnf"
        Else
           
        Call MsgBox("Could not find the Utils\FFEDIT folder in MSTS, this folder is required to process this file.", vbExclamation, App.Title)
        Exit Sub
        End If
    
    End If
    End If

MousePointer = 11

 cursouind = 0
ShapePath = Dir1(0).Path

For i = 0 To File1(cursouind).ListCount - 1

   If File1(cursouind).Selected(i) Then
  
    fullpath$ = File1(cursouind).Path & "\" & File1(cursouind).List(i)
   strOrigFile = File1(cursouind).List(i)
        If Right$(strOrigFile, 2) <> ".s" Then
        MousePointer = 0
        Call MsgBox("Selected file is not an .S file", vbExclamation, App.Title)
        Exit Sub
        End If
        FileCopy fullpath$, App.Path & "\TempFiles2\" & strOrigFile
   
   
   
   strBatText = "ffeditc_unicode.exe " & ChrW$(34) & strOrigFile & ChrW$(34) & " /c " & ChrW$(34) & "/o:" & strOrigFile & ChrW$(34) & vbCrLf
  
   
   NewFile = FreeFile
   Open App.Path & "\TempFiles2\doFfeditc.bat" For Output As #NewFile
   Print #NewFile, strBatText
   Close #NewFile
    
   ChDrive Left(App.Path, 1)
   ChDir App.Path & "\TempFiles2"
  DoEvents
    Call ShellAndWait("doffeditc.bat", True, vbNormalFocus)
        DoEvents
'    End If

    DoEvents
    Kill fullpath$
    DoEvents
     FileCopy App.Path & "\TempFiles2\" & strOrigFile, fullpath$
     DoEvents
     End If
    Next i
    MousePointer = 0
    
Exit Sub
Errtrap:

End Sub

Private Sub Command96_Click()
Dim Filpath1$, strRP As String, x As Integer, strBatText As String


Rem *********Check World Files *********
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If
MousePointer = 11
'Call UncompW

Rem ************************************
cursouind = 0
Filpath1$ = App.Path & "\TempFiles"
RoutePath = File1(cursouind).Path
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RoutePath
x = InStrRev(RoutePath, "\")
If x Then
RouteName = Mid$(RoutePath, x + 1)
frmUtils.Caption = "Path=" & MSTSPath & "  Route=" & RouteName
strRP = Left$(RoutePath, x)
End If
ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
   strBatText = "java -Xmx512m -Xss4m TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & RouteName & "_clrdb.log" & ChrW$(34) & "  clrdb -t -w " & ChrW$(34) & RoutePath & ChrW$(34)

  Call ShellAndWait(strBatText, True, vbNormalFocus)

 DoEvents
 If Not FileExists(App.Path & "\Reports\" & RouteName & "_clrdb.log") Then
Call MsgBox("As no .log file has been created, you can assume that TsUtils" _
            & vbCrLf & "did not complete the option you were attempting to run." _
            , vbCritical, App.Title)

MousePointer = 0
Exit Sub
End If
Call MsgBox("Your route CLRDB operation appears to have been successful. The new files are in the NewRoute folder within your first route. You may now replace the corresponding folders  and files in your original route with the equivalent folders and files from within NewRoute (but make a backup first).", vbExclamation, App.Title)

 
  MousePointer = 0
 frmReport.Rich1.LoadFile App.Path & "\Reports\" & RouteName & "_clrdb.log"
 frmReport.Show 1

End Sub

Private Sub Command97_Click()

Dim i As Integer, j As Long, GlobalPath As String, strSpare As String

On Error GoTo Errtrap


Command7.value = True
DoEvents
strSpare = App.Path & "\tempfiles"
If RouteListed = False Then
MsgBox Lang(330) & vbCr & Lang(331), 48, Lang(332)
Exit Sub
End If


MousePointer = 11
ReDim Preserve strGlobShp(0 To Shp_Chunk)
numGlobShp = 1
GlobalPath = MSTSPath & "\Global\"
frmTsect.Show 1

WorldPath = RoutePath & "\world"
SB1.Panels(2).Text = "Uncompressing .w files"
Call DoDeCompFolder("w", WorldPath, strSpare)
DoEvents
frmUtils.Dir1(0).Path = strSpare
frmUtils.Text1(0).Text = "*.w"
cursouind = 0
For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
SB1.Panels(2).Text = frmUtils.File1(cursouind).List(i)
DoEvents

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = strSpare & "\" & frmUtils.File1(cursouind).List(i)


Call ReadWorld2(fullpath$)

   End If

   Next i


 
    
    ReDim Preserve strGlobShp(0 To numGlobShp - 1)
   
    QSort3 strGlobShp(), 0, numGlobShp - 1
    DoEvents
    
    RemD2 strGlobShp(), strGlobShp2()
    DoEvents
    For j = 0 To numGlobShp - 1
    strGlobShp(j) = vbNullString
    Next j
   
   numGlobShp = UBound(strGlobShp2)
   

    booTsection = True


Call ParseTsection2(strTPath, GlobalPath & RouteName & "_tsection.dat")
   

Call MsgBox("A file " & RouteName & "_tsection.dat has been placed in your Global folder.", vbExclamation Or vbDefaultButton1, App.Title)
Exit Sub

Errtrap:
Call MsgBox("An error " & Err & "=" & Err.Description & " occurred while building a route-specific tsection.dat please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)

Resume Next
End Sub

Private Sub Command98_Click()
SetVariables
DoEvents
DetInput
End Sub

Function DetInput()
'this function determines what information you entered and what information
'needs to be calculated

If Text3(0) <> vbNullString And Text3(1) <> vbNullString Then
    'user entered lon & lat
        Call ConvertL(Val(Text3(0)), Val(Text3(1))) 'calculate long and lat values
        DoEvents
        
        Call TileName(Val(Text3(3)), Val(Text3(4))) 'calculate the world tile file name
ElseIf Text3(3) <> vbNullString And Text3(4) <> vbNullString Then
    'user entered world tile coordinates
        Call ConvertWTC(Val(Text3(3)), Val(Text3(4)))                              'Calculate the lon, lat coordinates
       
        Call TileName(Val(Text3(3)), Val(Text3(4))) 'calculate the world tile file name
ElseIf Text3(2) <> vbNullString Then
    'user inputed world tile name
End If
End Function

Private Sub Command99_Click()
Dim i As Integer
For i = 0 To 4
Text3(i) = vbNullString
Next i
End Sub

Private Sub Dir1_Change(Index As Integer)

On Error GoTo Errtrap

cursouind = Index
If cursouind = 0 Then
curtarind = 1
Else
curtarind = 0
End If
Label1(cursouind).Caption = Lang(31)
Label1(cursouind).BackColor = &H80FFFF
Label1(curtarind).Caption = Lang(32)
Label1(curtarind).BackColor = &H8000000F

Label2(cursouind).Caption = vbNullString
If Dir1(cursouind).Path <> Dir1(cursouind).List(Dir1(cursouind).ListIndex) Then
Dir1(cursouind).Path = Dir1(cursouind).List(Dir1(cursouind).ListIndex)
If Right$(Dir1(cursouind).Path, 1) <> "\" Then Dir1(cursouind).Path = Dir1(cursouind).Path & "\"
End If


File1(cursouind).Path = Dir1(cursouind).Path
Label2(cursouind).Caption = File1(cursouind).Path
If booMini = True And Index = 1 Then
frmCopy.Text1(1).Text = Dir1(1).Path
End If
If booMini = True And Index = 0 Then
frmCopy.Text1(0).Text = Dir1(0).Path
End If
Exit Sub
Errtrap:
MsgBox "Error in Dir1_Change"
Resume Next
End Sub

Private Sub Dir1_Click(Index As Integer)

If Text1(0).Text = "*.*" Then
Text1(0).Text = vbNullString
DoEvents
Text1(0).Text = "*.*"
End If

If booCopy = True And Index = 1 Then

frmCopy.Text1(0).Text = Dir1(Index).List(Dir1(Index).ListIndex)

End If
If booCopy = True And Index = 0 Then
frmCopy.Text1(1).Text = Dir1(Index).List(Dir1(Index).ListIndex)

End If
Rem ***********Update

If booUpdate = True And intUpdate = 0 Then
frmUpdate.Text1(0).Text = Dir1(Index).List(Dir1(Index).ListIndex)

End If
If booUpdate = True And intUpdate = 1 Then
frmUpdate.Text1(1).Text = Dir1(Index).List(Dir1(Index).ListIndex)

End If

Rem *********************

If booCommon = True And Index = 0 Then
frmCommon.Text1(0).Text = Dir1(Index).List(Dir1(Index).ListIndex)
frmCommon.ZOrder

End If

If booComDir = True And Index = 0 Then
frmCommon.Text1(0).Text = Dir1(Index).List(Dir1(Index).ListIndex)
frmCommon.ZOrder

End If
If booBackup = True And Index = 1 Then
frmBackup.Text1 = Dir1(Index).List(Dir1(Index).ListIndex)
frmBackup.ZOrder
End If
If booRaildriver = True And Index = 0 Then
frmDialog4.Text1.Text = Dir1(Index).List(Dir1(Index).ListIndex)

End If
cursouind = Index
If cursouind = 0 Then
curtarind = 1
Else
curtarind = 0
End If
Label1(cursouind).Caption = Lang(31)
Label1(cursouind).BackColor = &H80FFFF
Label1(curtarind).Caption = Lang(32)
Label1(curtarind).BackColor = &H8000000F
frmUtils.Refresh
End Sub





Private Sub Drive1_Change(Index As Integer)
On Error GoTo drivehandler2

cursouind = Index
If cursouind = 0 Then
curtarind = 1
Else
curtarind = 0
End If
Label1(cursouind).Caption = Lang(31)
Label1(cursouind).BackColor = &H80FFFF
Label1(curtarind).Caption = Lang(32)
Label1(curtarind).BackColor = &H8000000F

Dir1(cursouind).Path = Drive1(cursouind).Drive
Label2(cursouind).Caption = Dir1(cursouind).Path
'drv$ = Left$(Drive1(cursouind).Drive, 1)
'ax& = GetDiskSpaceFree(drv$)
'Label3(cursouind).Caption = "Free space drive " & UCase$(drv$) & ": " & Str$(ax&)
Exit Sub
drivehandler2:
Drive1(cursouind).Drive = Dir1(cursouind).Path
Exit Sub

End Sub


Private Sub Drive1_GotFocus(Index As Integer)
'drv$ = Left$(Drive1(cursouind).Drive, 1)
'ax& = GetDiskSpaceFree(drv$)
'Label3(cursouind).Caption = "Free space drive " & UCase$(drv$) & ": " & Str$(ax&)
End Sub


Private Sub Drive1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
cursouind = Index
If cursouind = 0 Then
curtarind = 1
Else
curtarind = 0
End If
Label1(cursouind).Caption = Lang(31)
Label1(cursouind).BackColor = &H80FFFF
Label1(curtarind).Caption = Lang(32)
Label1(curtarind).BackColor = &H8000000F
End Sub


Private Sub File1_Click(Index As Integer)
cursouind = Index
If cursouind = 0 Then
curtarind = 1
Else
curtarind = 0
End If
Label1(cursouind).Caption = Lang(31)
Label1(cursouind).BackColor = &H80FFFF
Label1(curtarind).Caption = Lang(32)
Label1(curtarind).BackColor = &H8000000F

If Right$(File1(cursouind).Path, 1) <> "\" Then

tit$ = File1(cursouind).Path & "\" & File1(cursouind).List(File1(cursouind).ListIndex)
Label2(cursouind).Caption = tit$
Else
tit$ = File1(cursouind).Path & File1(cursouind).List(File1(cursouind).ListIndex)
Label2(cursouind).Caption = tit$
End If
If booCommon = True And Index = 1 Then
frmCommon.Text1(1).Text = File1(cursouind).Path & "\" & File1(Index).List(File1(Index).ListIndex)
frmCommon.Text2.Text = File1(cursouind).Path & "\InstallCommon.bat"
frmCommon.ZOrder
End If
'If booTsection = True And Index = 0 Then
'dlgtsection.Text1 = File1(cursouind).path & "\" & File1(Index).list(File1(Index).ListIndex)
'End If
If IntRepW = 1 And Index = 1 Then
frmRepShape.Text1 = File1(Index).List(File1(Index).ListIndex)
ElseIf IntRepW = 2 And Index = 1 Then
frmRepShape.Text2 = File1(Index).List(File1(Index).ListIndex)
End If
End Sub


Private Sub Form_Activate()
frmUtils.Refresh


End Sub

Private Sub Form_Load()
Dim i%
Dim MySettings As Variant, result As String, x As Integer
Dim EnvString As String, A$

Rem*****************************************
On Error GoTo Errtrap
If Not DirExists(App.Path & "\Reports") Then
MkDir App.Path & "\Reports"
End If
strLogPath = App.Path & "\Reports\Startup.log"
If FileExists(strLogPath) Then
Kill strLogPath
DoEvents
End If
Open strLogPath For Append As #27
A$ = "Start"
Print #27, A$
'Call KillSpare2("*.s")
ReDim AllRoutes(0 To REF_CHUNK)
ReDim AllRoutes2(0 To REF_CHUNK)
DoEvents

cursouind = 0
A$ = "Write version number"
Print #27, A$
SB1.Panels(7).Text = "RR v." & App.Major & "." & App.Minor & "." & App.Revision
A$ = "Check for Tview.exe"
Print #27, A$
If Not FileExists(App.Path & "\tview.exe") Then
Command56(0).Visible = False
End If

A$ = "Check for TgaTool2a.exe"
Print #27, A$

If Not FileExists(App.Path & "\tgatool2a.exe") Then
Command56(1).Visible = False
End If
A$ = "Set Language"
Print #27, A$
Call SetLanguage("Lang_English.txt")

If Not DirExists(App.Path & "\Reports") Then
MkDir App.Path & "\Reports"
DoEvents
End If
A$ = "Check for Envstring"
Print #27, A$
EnvString = GetWindowsSysDir()

A$ = "Check for mwgfxvb.dll"
Print #27, A$
If FileExists(EnvString & "mwgfxvb.dll") Then
Frame9.Visible = True
Command48.Visible = True
End If
A$ = "Check for TsUtil"
Print #27, A$
'If DirExists(App.Path & "\TsUtil") Then
'SSTab1.TabVisible(7) = True
'Else
'SSTab1.TabVisible(7) = False
'End If
Rem *********
A$ = "Check for TrainStore"
Print #27, A$
'Stop
MySettings = GetAllSettings("Route_Riter6", "TrainStore")
If Not IsEmpty(MySettings) Then
strTrainStore = MySettings(0, 1)
Else
strTrainStore = vbNullString
End If
MySettings = GetAllSettings("Route_Riter6", "Conbuilder")
If Not IsEmpty(MySettings) Then
strConbuilder = MySettings(0, 1)
Else
strConbuilder = vbNullString
End If
MySettings = GetAllSettings("Route_Riter6", "TEdit")
If Not IsEmpty(MySettings) Then
strTextEditor = MySettings(0, 1)
Else
strTextEditor = vbNullString
End If

Rem ******** Save screenshotlocation
strScreenShotLocation = GetSetting("Decapod", "MSTS Shape Viewer", "screenshotLocation", "")

Rem *******
A$ = "Get Startup settings"
Print #27, A$
MySettings = GetAllSettings("Route_Riter6", "Startup")

If IsEmpty(MySettings) Then
frmUtils.WindowState = 0
frmUtils.height = 9600
frmUtils.width = 11800
Else
frmUtils.Left = MySettings(0, 1)
frmUtils.Top = MySettings(1, 1)


frmUtils.height = MySettings(2, 1)
frmUtils.width = MySettings(3, 1)
End If
If frmUtils.Top < 250 Then
frmUtils.Top = 2000
frmUtils.height = 9600
frmUtils.width = 11800
End If
TryAgain:
A$ = "TryAgain: - Get file settings"
Print #27, A$
MySettings = GetAllSettings("Route_Riter6", "Files")

If IsEmpty(MySettings) Then GoTo Label1

RoutePath = MySettings(0, 1)
If Not DirExists(RoutePath) Then
RoutePath = vbNullString
Text1(cursouind) = "*.*"
End If

If RoutePath <> vbNullString And Right$(RoutePath, 6) = "Routes" Then
RoutePath = vbNullString
Text1(cursouind) = "*.*"
End If
A$ = "Check RoutePath"
Print #27, A$
x = InStrRev(RoutePath, "\Routes")
If x = 0 Then
RoutePath = vbNullString
Text1(cursouind) = "*.*"
End If
Label1:
A$ = "Check for Common Path"
Print #27, A$
MySettings = GetAllSettings("Route_Riter6", "CommonPath")
If IsEmpty(MySettings) Then
strComPath = vbNullString
GoTo Label5
End If

strComPath = MySettings(0, 1)
If Not DirExists(strComPath) Then
Select Case MsgBox(Lang(387) & strComPath & vbCrLf & Lang(388), vbRetryCancel + vbExclamation + vbDefaultButton1, Lang(389))

    Case vbRetry
GoTo Label5
    Case vbCancel
strComPath = vbNullString
End Select

End If
A$ = "Check for Languages"
Print #27, A$
Label5:
MySettings = GetAllSettings("Route_Riter6", "Languages")
If IsEmpty(MySettings) Then
strLanguage = "Lang_English.txt"
GoTo Label2
End If
strLanguage = MySettings(0, 1)
If Not FileExists(App.Path & "\" & strLanguage) Then
strLanguage = "Lang_English.txt"
End If
Label2:
A$ = "Language = " & strLanguage
Print #27, A$
MySettings = GetAllSettings("Route_Riter6", "MainPath")

If IsEmpty(MySettings) Then
result = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "Path")
MSTSPath = result
A$ = "MSTS Path = " & MSTSPath
Print #27, A$
GoTo Label3
End If
MSTSPath = MySettings(0, 1)
A$ = "MSTS Path = " & MSTSPath
Print #27, A$
If Not DirExists(MSTSPath) Then
result = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Microsoft Games\Train Simulator\1.0", "Path")
MSTSPath = result
If result = "" Then
result = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Wow6432Node\Microsoft\Microsoft Games\Train Simulator\1.0", "Path")
MSTSPath = result
End If
A$ = "MSTS Path = " & MSTSPath
Print #27, A$
End If
If Not DirExists(MSTSPath) Then
Call MsgBox(Lang(390) & vbCrLf & Lang(391), vbCritical, Lang(392))

GoTo NoMSTS
End If

Label3:
Rem **********************

  '  End If
If Not FileExists(App.Path & "\" & strLanguage) Then
MsgBox Lang(393) & App.Path & "\" & strLanguage & Lang(394) & vbCr & Lang(395), 48, Lang(396)
End If
If Not DirExists(App.Path & "\TempFiles") Then
MkDir App.Path & "\TempFiles"
DoEvents
End If
For i = 0 To 4
Label7(i).BackColor = vbWhite
Next i
Label3.BackColor = vbWhite

Coupling(3) = "Bar"
Coupling(1) = "Automatic"
Coupling(2) = "Chain"

FCoupling(1) = "Automatic"
FCoupling(2) = "Chain"
FCoupling(3) = "Bar"

Rigid(0) = vbNullString
Rigid(1) = "Comment"
Rigid(2) = "Comment"
Rigid(3) = "Rigid"
FRigid(0) = vbNullString
FRigid(1) = "Comment"
FRigid(2) = "Comment"
FRigid(3) = "Rigid"
Brake(1) = "Air_Single_Pipe"
Brake(2) = "Air_Twin_Pipe"
Brake(3) = "Vacuum_Single_Pipe"
Brake(4) = "Vacuum_Twin_Pipe"
Brake(5) = "ECP"
Brake(6) = "EP"
Brake(7) = "Air_Piped"
Brake(8) = "Vacuum_Piped"
StockType(1) = "Engine"
StockType(2) = "Freight"
StockType(3) = "Carriage"
StockType(4) = "Tender"

StockType(5) = "Engine S*"
StockType(6) = "Engine S"
StockType(7) = "Engine D"
StockType(8) = "Engine E"
A$ = "Set Language Menu"
Print #27, A$
Call SetLangMenu
Rem **************Language
A$ = "Set Language"
Print #27, A$
Call SetLanguage(strLanguage)
A$ = "Check for Global sound path"
Print #27, A$
GlobalSoundPath = MSTSPath & "\Sound"

A$ = "Global sound path = " & GlobalSoundPath
Print #27, A$

SoundNumber = SoundNumber + 1
Soundfile(SoundNumber) = "ingame.sms"

AceNumber = AceNumber + 1
   AceFile(AceNumber) = "acleantrack1.ace"
   ESD(AceNumber) = "1"
   AceNumber = AceNumber + 1
   AceFile(AceNumber) = "acleantrack2.ace"
   ESD(AceNumber) = "1"
DoEvents
cursouind = 0
A$ = "Check for Aceit"
Print #27, A$

If Not FileExists(App.Path & "\AceIt.exe") Then
    If FileExists(App.Path & "\Aceit\Aceit.exe") Then
    FileCopy App.Path & "\Aceit\Aceit.exe", App.Path & "\Aceit.exe"
    FileCopy App.Path & "\Aceit\Aceit_help.chm", App.Path & "\Aceit_help.chm"
    FileCopy App.Path & "\Aceit\Aceit_license.txt", App.Path & "\Aceit_license.txt"
    Else
    A$ = "Aceit is not at " & App.Path & "\aceit.exe"
    Print #27, A$
    MsgBox Lang(397) & App.Path & vbCr & Lang(395), 48, Lang(396)
    End If
End If
Rem ******** Get All Routes ****************************
A$ = "Looking for existing routes in " & MSTSPath & "\Routes"
Print #27, A$

Call GetAllRoutes

Print #27, "Following routes found:-"

For i = 0 To NumRoutes - 1
Print #27, AllRoutes2(i)
Next i
Rem **************************************
If RoutePath = "" Then
A$ = "No Route Path found in Registry - Starting with no route listed"
Print #27, A$
Else
A$ = "Check Route " & RoutePath
Print #27, A$
End If
Rem ***************** Check for Java ***********************
Dim strBatText As String
If FileExists(App.Path & "\Reports\" & "TsUtil_startup.log") Then
Kill App.Path & "\Reports\" & "TsUtil_startup.log"
End If
DoEvents

strBatText = "java TSUtil -l" & ChrW$(34) & App.Path & "\Reports\" & "TsUtil_startup.log" & ChrW$(34) & "  version"


ChDrive Left$(App.Path, 1)
 'ChDir App.Path & "\TSUtil"
  

 Call ShellAndWait(strBatText, True, vbHide)

 DoEvents
 If Not FileExists(App.Path & "\Reports\" & "TsUtil_startup.log") Then
Call MsgBox("You do not appear to have a Java runtime system on your PC, this program will not run until this is fixed. See my FAQ page." _
            , vbCritical, App.Title)

MousePointer = 0
Exit Sub
End If
 A$ = "Java system/TsUtils working OK,  confirmed"
Print #27, A$
Close #27
Rem ********************************************************
'Call KillSpare2("*.s")
If RoutePath <> vbNullString Then
Drive1(cursouind).Drive = Left$(RoutePath, 2)
Dir1(cursouind).Path = RoutePath
Text1(cursouind) = "*.trk"
DoEvents
Command7.value = True
Else
result = MSTSPath
Text1(cursouind) = "*.*"
Drive1(cursouind).Drive = Left$(result, 2)
Dir1(cursouind).Path = result & "\Routes"
End If
If Not DirExists(App.Path & "\TempFiles2") Then

MkDir App.Path & "\TempFiles2"
End If
If Not FileExists(App.Path & "\UHARC.exe") Then
Command122.Enabled = False
End If
NoMSTS:

Exit Sub
Errtrap:
If Err = 380 Then
frmUtils.WindowState = 0
frmUtils.height = 9600
frmUtils.width = 11800
GoTo TryAgain
ElseIf Err = 76 Then
Resume Next
'GoTo TryAgain
Else
Call MsgBox("An Error No. " & Err & " " & Err.Description & " occurred while loading the main Route_Riter screen." _
            & vbCrLf & "Check Reports\Startup.log for possible errors." _
            , vbExclamation, App.Title)

End If



End Sub
Function GetWindowsSysDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the system directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetSystemDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator(strBuf)
        AddDirSep strBuf
        
        GetWindowsSysDir = strBuf
    Else
        GetWindowsSysDir = vbNullString
    End If
End Function
Private Sub Form_Resize()
Dim Wid%, i As Integer


On Error GoTo Errtrap

TryAgain:
Frame7.Top = 1
Frame7.Left = 200
Frame7.width = Me.width - 400

Wid% = (frmUtils.width \ 4) - 300
For i = 0 To 1
Dir1(i).height = frmUtils.height \ 3.5
File1(i).height = frmUtils.height \ 3.5
Dir1(i).width = Wid%
File1(i).width = Wid%
Drive1(i).width = Wid%
Text1(i).width = Wid%
'Label1(i).width = wID% - 700
Label1(i).Top = Frame7.Top + 150
Drive1(i).Top = Frame7.Top + Frame7.height + 30
Text1(i).Top = Drive1(i).Top
Dir1(i).Top = Text1(i).Top + Text1(i).height + 50
File1(i).Top = Text1(i).Top + Text1(i).height + 50
Label2(i).width = (Wid% * 2) + 50

Label2(i).Top = Dir1(i).Top + Dir1(i).height + 110

Line2.y1 = Label2(i).Top + Label2(i).height + 50
Line2.y2 = Line2.y1
Line1.y2 = Line2.y1
Next i
Dir1(0).Left = 150
Drive1(0).Left = 150
Label2(0).Left = 150

File1(0).Left = 200 + Dir1(0).width
Text1(0).Left = 200 + Dir1(0).width
Line1.X1 = frmUtils.width / 2
Line1.X2 = frmUtils.width / 2
Line2.X2 = frmUtils.width
File1(1).Left = Line1.X1 + 100
Text1(1).Left = Line1.X1 + 100
Label2(1).Left = Line1.X1 + 100

Dir1(1).Left = 50 + File1(1).width + File1(1).Left
Drive1(1).Left = 50 + File1(1).width + File1(1).Left

'Label1(1).Left = (frmUtils.width \ 2) + (frmUtils.width \ 4) - (Label1(0).width \ 2)

SSTab1.Top = Line2.y1 + 100
SSTab1.height = frmUtils.height * 0.48
SSTab1.width = frmUtils.width * 0.96

Command42.Top = Frame7.Top + 150
Command59.Top = Frame7.Top + 150
Command59.Left = Frame7.Left + 2
Command10(0).Top = Command42.Top
Command10(0).Left = Command59.Left + Command59.width + 10
Label1(0).Left = Command10(0).Left + Command10(0).width + 20
Command42.Left = Label1(0).Left + Label1(0).width + 10
Command53.Top = Command42.Top
Command53.Left = Command42.Left + Command42.width + 10
Command78.Top = Command53.Top
Command78.Left = Command53.Left + Command53.width
Command124.Top = Command53.Top
Command124.Left = Command78.Left + Command78.width
'Command143.Top = Command53.Top
'Command143.Left = Command124.Left + Command124.width

For i = 0 To 3
Command43(i).Top = Command42.Top
Next i
Command43(0).Left = Command124.Left + Command124.width
Command43(1).Left = Command43(0).Left + Command43(0).width
Command43(2).Left = Command43(1).Left + Command43(1).width
Command43(3).Left = Command43(2).Left + Command43(2).width

Command41.Left = Text1(1).Left + 250
Command41.Top = Command42.Top
Label1(1).Left = Command41.Left + Command41.width + 10
Command10(1).Top = Command42.Top
Command10(1).Left = Label1(1).Left + Label1(1).width + 100
'Command69.Top = Command42.Top
'Command69.Left = Command10(1).Left + Command10(1).width + 10
Command1(15).Top = SSTab1.Top + SSTab1.height - (Command1(15).height + 300)


Command1(15).Left = SSTab1.Left + SSTab1.width - (Command1(15).width + 200)

Exit Sub
Errtrap:
If Err = 380 Then
Exit Sub
End If
'Exit Sub
End Sub









Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Errtrap

Close

Call ClearTmp
Select Case MsgBox(Lang(384), vbYesNo + vbExclamation + vbDefaultButton1, "Route_Riter")

    Case vbYes
 cursouind = 1
 SparePath = App.Path & "\TempFiles2"
Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.s"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
DoEvents
SparePath = App.Path & "\TempFiles"
Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.W"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
DoEvents
 Text1(1).Text = "*.s"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 DoEvents
 
 Text1(1).Text = "do*.*"
For i = 0 To File1(cursouind).ListCount - 1
    File1(cursouind).Selected(i) = True
Next i

For i = 0 To File1(cursouind).ListCount - 1
 Kill File1(cursouind).Path & "\" & File1(cursouind).List(i)
 Next i
 
 
 Drive1(1).Drive = Left$(SparePath, 2)
Dir1(1).Path = SparePath
Text1(1).Text = "*.*"
If Not frmUtils.WindowState = 1 Then
SaveSetting "Route_Riter6", "Startup", "Left", frmUtils.Left
SaveSetting "Route_Riter6", "Startup", "Top", frmUtils.Top
SaveSetting "Route_Riter6", "Startup", "Height", frmUtils.height
SaveSetting "Route_Riter6", "Startup", "Width", frmUtils.width
End If
SaveSetting "Route_Riter6", "Files", "File", MasterRoutePath
SaveSetting "Route_Riter6", "Languages", "Language", strLanguage
SaveSetting "Route_Riter6", "MainPath", "MainPath", MSTSPath
SaveSetting "Route_Riter6", "CommonPath", "CommonPath", strComPath
SaveSetting "Route_Riter6", "User", "UserName", ""
SaveSetting "Decapod", "MSTS Shape Viewer", "ScreenShotLocation", strScreenShotLocation

    
End

DoEvents
Call KillProcess("Route_Riter.exe")
    Case vbNo
    Cancel = 1
Exit Sub
End Select
Exit Sub
Errtrap:
If Err = 76 Then
Resume Next
Else
Call MsgBox("An error " & Err & " occurred in subroutine 'frmUtils Unload' please advise" _
            & vbCrLf & "Support with details of operation being processed." _
            , vbExclamation, App.Title)
Resume Next
End If
End Sub









Private Sub mnu1_Click(Index As Integer)
strLanguage = "Lang_" & Language(Index) & ".txt"
Call SetLanguage(strLanguage)
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show

     DoEvents
     
End Sub

Private Sub mnuCommon_Click()
booComDir = True
frmCommon.Show

DoEvents
'booComDir = False
End Sub


Private Sub mnuConPath_Click()
CDL1.DialogTitle = "Select Conbuilder.exe"
CDL1.Flags = cdlOFNExplorer
If strConbuilder <> vbNullString Then
CDL1.Filename = strConbuilder
End If
CDL1.ShowOpen
strConbuilder = CDL1.Filename
SaveSetting "Route_Riter6", "Conbuilder", "Conbuilder", strConbuilder
End Sub

Private Sub mnuCont_Click()
HTMLHelpContents 1, ""
End Sub





Private Sub mnuExit_Click()
Unload Me
End Sub



Private Sub mnuFAQ_Click()
flagInternet = 3
frmInternet.Show
End Sub

Private Sub mnuHome_Click()
flagInternet = 2
frmInternet.Show
End Sub

Private Sub mnuPath_Click()
frmPath.Show 1
If booWrongMSTS = True Then
booWrongMSTS = False
Exit Sub
End If
DoEvents
'SaveSetting "Route_Riter6", "MainPath", "MainPath", MSTSPath
If Not frmUtils.WindowState = 1 Then
SaveSetting "Route_Riter6", "Startup", "Left", frmUtils.Left
SaveSetting "Route_Riter6", "Startup", "Top", frmUtils.Top
SaveSetting "Route_Riter6", "Startup", "Height", frmUtils.height
SaveSetting "Route_Riter6", "Startup", "Width", frmUtils.width
End If
SaveSetting "Route_Riter6", "Files", "File", RoutePath
SaveSetting "Route_Riter6", "Languages", "Language", strLanguage
SaveSetting "Route_Riter6", "MainPath", "MainPath", MSTSPath
SaveSetting "Route_Riter6", "CommonPath", "CommonPath", strComPath
SaveSetting "Route_Riter6", "User", "UserName", strUser
SaveSetting "Decapod", "MSTS Shape Viewer", "ScreenShotLocation", strScreenShotLocation

DoEvents
Call GetAllRoutes

End Sub




Private Sub mnuTE_Click()
CDL1.DialogTitle = "Select a Text Editor"
CDL1.Flags = cdlOFNExplorer
If strTextEditor <> vbNullString Then
CDL1.Filename = strTextEditor
End If
CDL1.ShowOpen
strTextEditor = CDL1.Filename
SaveSetting "Route_Riter6", "TEdit", "TEdit", strTextEditor
End Sub

Private Sub mnuTER_Click()
Dim retval As Variant

retval = Shell(strTextEditor, 1)
End Sub

Private Sub mnuTrainstore_Click()

CDL1.DialogTitle = "Select Trainstore.exe"
CDL1.Flags = cdlOFNExplorer
If strTrainStore <> vbNullString Then
CDL1.Filename = strTrainStore
End If
CDL1.ShowOpen
strTrainStore = CDL1.Filename

If strTrainStore <> vbNullString Then
retval = Shell(strTrainStore, 1)
End If
SaveSetting "Route_Riter6", "Trainstore", "Trainstore", strTrainStore
End Sub


Private Sub mnuUpdates_Click()
flagInternet = 2
frmInternet.Show
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 8 Then

      Call MsgBox("Do NOT use this option if you are using the VISTA or Windows 7 operating systems." _
                  & vbCrLf & "Neither operating system is compatible with HardLinks." _
                  , vbExclamation, "Warning")
                  End If
      
'Text1(0) = "*.*"
End Sub

Private Sub Text1_Change(Index As Integer)
On Error GoTo Errtrap
If Text1(Index).Text = vbNullString Then
Text1(Index).Text = "*.*"
End If

cursouind = Index
If cursouind = 0 Then
curtarind = 1
Else
curtarind = 0
End If
Label1(cursouind).Caption = Lang(31)
Label1(cursouind).BackColor = &H80FFFF
Label1(curtarind).Caption = Lang(32)
Label1(curtarind).BackColor = &H8000000F
File1(cursouind).Pattern = Text1(cursouind).Text
Exit Sub
Errtrap:
If Err = 380 Then Resume Next
If Err > 0 Then
Call MsgBox("Invalid character entered in the FILTER string", vbCritical, App.Title)

End If
End Sub





