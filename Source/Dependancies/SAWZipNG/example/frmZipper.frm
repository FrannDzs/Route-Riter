VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmZipper 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Zipper (A SAWZipNG example)"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   Icon            =   "frmZipper.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgAdd 
      Left            =   2355
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Add a file to the archive"
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   4050
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   3495
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open an archive"
      Filter          =   "Zip files (*.zip) |*.zip|Jar files (*.jar) |*.jar|All files(*.*)|*.*|"
   End
   Begin MSComDlg.CommonDialog dlgNew 
      Left            =   2910
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".zip"
      Filter          =   "Zip files (*.zip) |*.zip|All files(*.*)|*.*|"
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   210
      Left            =   0
      TabIndex        =   2
      Top             =   2985
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   370
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   2205
      Left            =   255
      TabIndex        =   1
      Top             =   780
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   3889
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      Icons           =   "imgIcons"
      SmallIcons      =   "imgIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Modified"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Ratio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Compressed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "intro"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   10
      TabIndex        =   0
      Top             =   3
      Width           =   2640
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Archive"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Archive"
         Begin VB.Menu mnuFileNewNormal 
            Caption         =   "&Normal Archive"
         End
         Begin VB.Menu mnuFileNewPKZip 
            Caption         =   "&PKZip Diskspanning"
         End
         Begin VB.Menu mnuFileNewTD 
            Caption         =   "&TD Diskspanning"
         End
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close Archive"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuMruSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMru 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuArchive 
      Caption         =   "&Archive"
      Enabled         =   0   'False
      Begin VB.Menu mnuArchiveAddFile 
         Caption         =   "Add &File"
      End
      Begin VB.Menu mnuArchiveAddDir 
         Caption         =   "Add &Directory"
      End
      Begin VB.Menu mnuArchiveSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchiveExtract 
         Caption         =   "&Extract"
      End
   End
End
Attribute VB_Name = "frmZipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private exitApp As Boolean ' Is it allowed to exit the app
Public WithEvents archive As SAWZipNG.archive
Attribute archive.VB_VarHelpID = -1

Private Const MaxMRU = 4      'Maximum number of MRUs in list (-1 for no limit)
Private Const NotFound = -1   'Indicates a duplicate entry was not found
Private Const NoMRUs = -1     'Indicates no MRUs are currently defined
Private MRUCount As Long      'Maintains a count of MRUs defined

Private Sub Form_Load()

    exitApp = False
    lblIntro.Caption = vbNewLine + "Please open or create a new zip archive" _
                     + vbNewLine + vbNewLine _
                     + "(c) 2003 - S.A.W. Franky Braem"
   
   ' Initialize the count of MRUs
   MRUCount = NoMRUs

   ' Call sub to retrieve the MRU filenames
   GetMRUFileList
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    lblIntro.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - statusBar.Height
    lstFiles.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - statusBar.Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not exitApp Then
        If Not archive Is Nothing Then
            CloseArchive
            Cancel = True
            Exit Sub
        End If
    End If
    
    SaveMRUFileList
    End
    
End Sub
Private Sub lstFiles_Click()

End Sub

Private Sub mnuArchiveAddDir_Click()

    Dim dlg As dlgDir
    Dim oldRoot As String
    
    dlgDir.Show vbModal, Me
    If dlgDir.ok Then
    
        oldRoot = archive.RootPath
        If dlgDir.chkFullpath.Value <> vbChecked Then
            archive.RootPath = dlgDir.txtDir.Text
        End If
        archive.AddFolderWithWildcard dlgDir.txtDir.Text, _
                                      dlgDir.txtWildCard.Text, _
                                      dlgDir.chkSubDir.Value, _
                                      dlgDir.chkFullpath.Value
        If dlgDir.chkFullpath.Value <> vbChecked Then
            archive.RootPath = oldRoot
        End If
        
        RefreshList
    
    End If
    
End Sub

Private Sub mnuArchiveAddFile_Click()

    dlgAdd.ShowOpen
    If Len(dlgAdd.filename) > 0 Then
        If archive.AddFile(dlgAdd.filename) Then
            RefreshList
        Else
            MsgBox "Unable to add the file"
        End If
    End If

End Sub

Private Sub mnuFileClose_Click()

    CloseArchive

End Sub

Private Sub mnuFileExit_Click()

    If Not archive Is Nothing Then
        CloseArchive
    End If
    exitApp = True
    Unload Me

End Sub

Private Sub CloseArchive()

    'Don't accept files from now
    DragAcceptFiles Me.hwnd, False
    
    archive.Close
    Set archive = Nothing

    lstFiles.ListItems.Clear
    lstFiles.Visible = False

    mnuArchive.Enabled = False
    mnuFileClose.Enabled = False

End Sub

Private Sub mnuFileNewNormal_Click()

    dlgNew.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt
    dlgNew.DialogTitle = "Create a new archive"
    dlgNew.ShowSave
    If Len(dlgNew.filename) > 0 Then
        If archive Is Nothing Then
            Set archive = New SAWZipNG.archive
        Else
            archive.Close
        End If
        archive.Create dlgNew.filename
        
        OpenArchive
    End If

End Sub

Private Sub mnuFileOpen_Click()

    Dim mode As SAWZipNG.OpenMode
    Dim f As SAWZipNG.FileInfo

    dlgOpen.ShowOpen
    If Len(dlgOpen.filename) > 0 Then
        If archive Is Nothing Then
            Set archive = New SAWZipNG.archive
        Else
            archive.Close
        End If
        
        If GetAttr(dlgOpen.filename) = vbReadOnly Then
            mode = OM_READONLY
        Else
            mode = OM_OPEN
        End If
        archive.Open dlgOpen.filename, mode
        
        If archive.Count > 0 Then
            
            Set f = archive.GetFileInfo(0)
            If f.Encrypted Then
                MsgBox "Archive is password protected!"
            End If
        
        End If
        
        OpenArchive
    End If

End Sub

Public Sub RefreshList()
    
    Dim it As ListItem
    Dim f As SAWZipNG.FileInfo
    Dim fileparts As Variant
    Dim filename As String
    Dim iconIdx As Long
    
    Me.MousePointer = vbHourglass
    
    If archive Is Nothing Then
        Return
    End If
    
    lstFiles.ListItems.Clear
    For i = 0 To archive.Count - 1
      Set f = archive.GetFileInfo(i)
      fileparts = Split(f.filename, "\")
      filename = fileparts(UBound(fileparts))
      If Len(filename) <> 0 Then ' we don't show directories
        Set it = lstFiles.ListItems.Add(, f.filename, filename)
        iconIdx = GetIconIndex(filename)
        If iconIdx <> -1 Then
            it.SmallIcon = iconIdx
        End If
        it.ListSubItems.Add , , GetFileType(filename)
        it.ListSubItems.Add , , f.ModificationDate
        it.ListSubItems.Add , , f.UncompressedSize
        it.ListSubItems.Add , , CStr(Round(100 - f.CompressionRatio)) + " %"
        it.ListSubItems.Add , , f.CompressionSize
        it.ListSubItems.Add , , f.filename
      End If
    Next i
    
    Me.MousePointer = vbNormal
    
End Sub

Private Sub OpenArchive()
        
    mnuArchive.Enabled = True
    mnuFileClose.Enabled = True
    frmZipper.Caption = "Zipper - " + archive.ArchivePath
    lstFiles.Visible = True
    
    DragAcceptFiles Me.hwnd, True        ' Accept files from now
    RefreshList
   
    AddMRUItem archive.ArchivePath
    
End Sub


Private Function GetIconIndex(filename As String) As Long

    Dim pic As IPicture
    Dim img As ListImage
    
    Dim dot As Long
    Dim ext As String
    
    dot = InStrRev(filename, ".")
    If dot Then
      ext = Mid(filename, dot + 1)
    Else
      GetIconIndex = -1
      Exit Function
    End If

    On Error Resume Next
    Set img = imgIcons.ListImages.Item(ext)
    If img Is Nothing Then
        Set pic = PictureFromHandle(GetSmallIcon(filename), vbPicTypeIcon)
        Set img = imgIcons.ListImages.Add(, ext, pic)
    End If
    
    GetIconIndex = img.Index

End Function

Private Sub AddMRUItem(NewItem As String)
   Dim result As Long

   ' Call sub to check for duplicates
   result = CheckForDuplicateMRU(NewItem)

   ' Handle case if duplicate found
   If result <> NotFound Then
      ' Call sub to reorder MRU list
      ReorderMRUList NewItem, result
   Else
      ' Call sub to add new item to MRU menu
      AddMenuElement NewItem
   End If
End Sub

Private Function CheckForDuplicateMRU(ByVal NewItem As String) As Long
   Dim i As Long

   ' Uppercase newitem for string comparisons
   NewItem = UCase$(NewItem)

   ' Check all existing MRUs for duplicate
   For i = 0 To MRUCount
      If UCase$(Me.mnuMru(i).Caption) = NewItem Then
         ' Duplicate found, return the location of the duplicate
         CheckForDuplicateMRU = i

         ' Stop searching
         Exit Function
      End If
   Next i

   ' No duplicate found, so return -1
   CheckForDuplicateMRU = -1
End Function

Private Sub mnuQuit_Click()
   ' Close the program
   Unload Me
End Sub

Private Sub AddMenuElement(NewItem As String)
   Dim i As Long

   ' Check that we will not exceed maximum MRUs
   If (MRUCount < (MaxMRU - 1)) Or (MaxMRU = -1) Then
      'Increment the menu count
      MRUCount = MRUCount + 1

      ' Check if this is the first item
      If MRUCount <> 0 Then
         ' Add a new element to the menu
         Load mnuMru(MRUCount)
      End If

      ' Make new element visible
      mnuMru(MRUCount).Visible = True
   End If

   ' Shift items to maintain most recent to least recent
   For i = (MRUCount) To 1 Step -1
      ' Set the captions
      mnuMru(i).Caption = mnuMru(i - 1).Caption
   Next i

   ' Set caption for new item
   mnuMru(0).Caption = NewItem
   mnuMruSep.Visible = True
End Sub

Private Sub ReorderMRUList(DuplicateMRU As String, DuplicateLocation As Long)
   Dim i As Long

   ' Move entries previously "more recent" than the
   ' duplicate down one in the MRU list
   For i = DuplicateLocation To 1 Step -1
      mnuMru(i).Caption = mnuMru(i - 1).Caption
   Next i

   ' Set caption of newitem
   mnuMru(0).Caption = DuplicateMRU
End Sub

Private Sub GetMRUFileList()
   Dim i As Long           'Loop control variable
   Dim result As String    'Name of MRU from registry

   ' Loop through all entries
   Do
      ' Retrieve entry from registry
      result = GetSetting(App.Title, "MRUFiles", Trim$(CStr(i)), "")

      ' Check if a value was returned
      If result <> "" Then
         ' Call sub to additem to MRU list
         AddMRUItem result
      End If

      ' Increment counter
      i = i + 1
   Loop Until (result = "")
   If i = 1 Then
    mnuMruSep.Visible = False
   End If
End Sub

Private Sub SaveMRUFileList()
   Dim i As Long           ' Loop control variable

   ' Loop through all MRU
   For i = 0 To MRUCount
      ' Write MRU to registry with key as it's position in list
      SaveSetting App.Title, "MRUFiles", Trim$(CStr(i)), mnuMru(i).Caption
   Next i
End Sub

Private Sub mnuMRU_Click(Index As Integer)

    Dim filename As String
    
    ' Call sub to reorder the MRU list
    ReorderMRUList mnuMru(Index).Caption, CLng(Index)
    
    If Not archive Is Nothing Then
        CloseArchive
    Else
        Set archive = New SAWZipNG.archive
    End If
    
    If Len(Dir(mnuMru(Index).Caption)) > 0 Then
        filename = mnuMru(Index).Caption
        If GetAttr(filename) = vbReadOnly Then
            mode = OM_READONLY
        Else
            mode = OM_OPEN
        End If
        archive.Open filename, mode
        
        If archive.Count > 0 Then
            
            Set f = archive.GetFileInfo(0)
            If f.Encrypted Then
                MsgBox "Archive is password protected!"
            End If
        
        End If
        
        OpenArchive
        
    Else
        mnuMru(Index).Visible = False
    End If
End Sub
