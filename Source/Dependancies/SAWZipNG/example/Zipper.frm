VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmZipper 
   Caption         =   "Zipper"
   ClientHeight    =   4725
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6885
   Icon            =   "Zipper.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4470
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6597
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Number of files in the archive"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "The size of the archive in bytes"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgAdd 
      Left            =   6600
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Add a file to the archive"
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   6600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open an archive file"
   End
   Begin MSComDlg.CommonDialog dlgNew 
      Left            =   6600
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".zip"
      DialogTitle     =   "Create a new archive"
      Filter          =   "Zip files (*.zip) |*.zip|All files(*.*)|*.*|"
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   3855
      Left            =   15
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Uncompressed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Compressed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Comment"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   180
      Left            =   75
      TabIndex        =   2
      Top             =   30
      Width           =   840
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuArchive 
      Caption         =   "&Archive"
      Enabled         =   0   'False
      Begin VB.Menu mnuArchiveAdd 
         Caption         =   "Add &File"
      End
      Begin VB.Menu mnuArchiveAddDirectory 
         Caption         =   "Add &Directory"
      End
      Begin VB.Menu mnuArchiveRemoveFile 
         Caption         =   "&Remove File"
      End
      Begin VB.Menu mnuArchiveRelativePath 
         Caption         =   "Relative Path"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuArchiveSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchiveComment 
         Caption         =   "&Comment"
      End
      Begin VB.Menu mnuArchiveCommentFile 
         Caption         =   "C&omment File"
      End
      Begin VB.Menu mnuArchiveSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchiveExtract 
         Caption         =   "&Extract"
      End
      Begin VB.Menu mnuArchiveExtractAll 
         Caption         =   "E&xtract All"
      End
   End
End
Attribute VB_Name = "frmZipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public arc As SAWZIPLib.Archive
Private exitApp As Boolean

Private Sub Form_Load()

    exitApp = False
    lblIntro.Caption = vbNewLine + "Please open or create a new zipfile"

    statusBar.Panels.Item(1).Text = "Ready."
    statusBar.Panels.Item(2).Text = "No archive open"
    statusBar.Panels.Item(3).Text = "No archive open"
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    lblIntro.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - statusBar.Height
    lstFiles.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - statusBar.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not exitApp Then
        If Not arc Is Nothing Then
            CloseArchive
            Cancel = True
            Exit Sub
        End If
    End If
    
    End
    
End Sub

Private Sub mnuArchive_Click()

    If lstFiles.SelectedItem Is Nothing Then
        mnuArchiveRemoveFile.Enabled = False
        mnuArchiveCommentFile.Enabled = False
        mnuArchiveExtract.Enabled = False
    Else
        mnuArchiveRemoveFile.Enabled = True
        mnuArchiveCommentFile.Enabled = True
        mnuArchiveExtract.Enabled = True
    End If
    
End Sub

Private Sub mnuArchiveAdd_Click()
    
    Dim f As SAWZIPLib.File
    Dim Item As ListItem
    
    dlgAdd.ShowOpen
    If Len(dlgAdd.filename) > 0 Then
        Set f = New SAWZIPLib.File
        f.FullPath = False
        f.Name = dlgAdd.filename
        arc.Files.Add f
        statusBar.Panels.Item(1).Text = dlgAdd.filename + " added to " + arc.Name + "."
        RefreshList
    End If
    
End Sub

Private Sub mnuArchiveAddDirectory_Click()

  Dim dir As String
  Dim count As Long
  Dim wildcard As String
  
  Dim dlg As New dlgDir
  dlg.Show vbModal, Me
  
  If dlg.ok Then
    count = arc.Files.count
    dir = dlg.txtDir.Text
    If Right(dir, 1) <> "\" Then
        dir = dir + "\"
    End If
    wildcard = dlg.txtWildCard.Text
    If Len(Trim(wildcard)) = 0 Then
        wildcard = "*.*"
    End If
    If dlg.chkSubDir.Value Then
      arc.StoreRelativePaths = True
      arc.Files.AddDir dir, wildcard, 9, True, True
    Else
      arc.Files.AddFileByName dir + wildcard, 9, True, True
    End If
    RefreshList
    statusBar.Panels.Item(1).Text = CStr(arc.Files.count - count) + " files added."
 Else
    statusBar.Panels.Item(1).Text = "No files added."
 End If

End Sub

Private Sub mnuArchiveComment_Click()
    
    Dim dlg As New dlgComment
    
    If Not arc Is Nothing Then
        
        dlg.lblMessage.Caption = "Enter some comment for the archive ..."
        dlg.lblFile.Visible = False
        dlg.txtComment.Text = arc.Comment
        dlg.Show vbModal, Me
        If dlg.ok Then
            arc.Comment = dlg.txtComment.Text
        End If
        
    End If

End Sub

Private Sub mnuArchiveCommentFile_Click()
    
    Dim dlg As New dlgComment
    Dim f As SAWZIPLib.File
    
    If Not lstFiles.SelectedItem Is Nothing Then
        Set f = arc.Files.Item(lstFiles.SelectedItem.Index - 1)
        If f Is Nothing Then
            MsgBox "Can't select file"
            Return
        End If
        dlg.lblMessage.Caption = "Enter some comment for file ..."
        dlg.lblFile.Caption = " " + f.Name + " "
        dlg.txtComment.Text = f.Comment
        dlg.Show vbModal, Me
        If dlg.ok Then
            f.Comment = dlg.txtComment.Text
        End If
    End If

End Sub

Private Sub mnuArchiveExtract_Click()
    
    Dim f As SAWZIPLib.File
    Dim dir As String
    Dim extractedName As String
    
    If Not lstFiles.SelectedItem Is Nothing Then
        Set f = arc.Files.Item(lstFiles.SelectedItem.Index - 1)
        If f Is Nothing Then
            MsgBox "Can't select file"
            Return
        End If
        
        dir = SelectDir(Me.hwnd, "Select destination directory")
        If Len(dir) > 0 Then
            extractedName = f.Extract(dir)
            statusBar.Panels.Item(1).Text = "Extracted to " + extractedName
            MsgBox "Extracted to " + extractedName
        End If
    End If
    
End Sub

Private Sub mnuArchiveExtractAll_Click()
    
    Dim dirName As String

    dirName = SelectDir(Me.hwnd, "Select a directory to add to the archive")
    If Len(Trim(dirName)) > 0 Then
        arc.Files.ExtractToDirectory dirName
        statusBar.Panels.Item(1).Text = "All files extracted to " + dirName
    End If
    
End Sub

Private Sub mnuArchiveRelativePath_Click()

    mnuArchiveRelativePath.Checked = Not mnuArchiveRelativePath.Checked
    arc.StoreRelativePaths = mnuArchiveRelativePath.Checked

End Sub

Private Sub mnuArchiveRemoveFile_Click()

    If lstFiles.SelectedItem Is Nothing Then
        Return
    End If
    
    arc.Files.Remove lstFiles.SelectedItem.Index - 1
    RefreshList
    
End Sub

Private Sub mnuFileClose_Click()

    CloseArchive
    
End Sub

Private Sub mnuFileExit_Click()
    
    If Not arc Is Nothing Then
        arc.Close
    End If
    exitApp = True
    Unload Me
    
End Sub

Private Sub mnuFileNew_Click()
    
    dlgNew.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt
    dlgNew.ShowSave
    If Len(dlgNew.filename) > 0 Then
        If arc Is Nothing Then
            Set arc = New Archive
            arc.StoreRelativePaths = mnuArchiveRelativePath.Checked
        Else
            arc.Close
        End If
        arc.Create dlgNew.filename
        
        ' Accept files from now
        DragAcceptFiles Me.hwnd, True
        
        mnuArchive.Enabled = True
        mnuFileClose.Enabled = True
        frmZipper.Caption = "Zipper - " + dlgNew.filename
        lstFiles.Visible = True
        RefreshList
    End If
    
End Sub
Private Sub mnuFileOpen_Click()
    
    dlgOpen.ShowOpen
    If Len(dlgOpen.filename) > 0 Then
        If arc Is Nothing Then
            Set arc = New Archive
        Else
            arc.Close
        End If
        arc.StoreRelativePaths = mnuArchiveRelativePath.Checked
        arc.Open dlgOpen.filename
        
        ' Accept files from now
        DragAcceptFiles Me.hwnd, True
        
        mnuArchive.Enabled = True
        mnuFileClose.Enabled = True
        frmZipper.Caption = "Zipper - " + dlgOpen.filename
        lstFiles.Visible = True
        RefreshList
    End If

End Sub

Public Sub RefreshList()

    Dim it As ListItem

    If arc Is Nothing Then
        Return
    End If
    
    lstFiles.ListItems.Clear
    For Each f In arc.Files
        If Len(f.RelativePath) > 0 Then
            Set it = lstFiles.ListItems.Add(, f.RelativePath, f.RelativePath)
        Else
            Set it = lstFiles.ListItems.Add(, f.Name, f.Name)
        End If
        it.ListSubItems.Add , , f.UncompressedSize
        it.ListSubItems.Add , , f.CompressedSize
        it.ListSubItems.Add , , f.ModificationDate
        it.ListSubItems.Add , , f.Comment
    Next f
    
    Select Case arc.Files.count
    Case 0
        statusBar.Panels.Item(2).Text = "Empty archive"
    Case 1
        statusBar.Panels.Item(2).Text = "1 file"
    Case Else
        statusBar.Panels.Item(2).Text = CStr(arc.Files.count) + " files"
    End Select
    statusBar.Panels.Item(3).Text = CStr(FileLen(arc.Name)) + " bytes"
    
End Sub

Private Sub CloseArchive()
    
    ' Don't accept files from now
    DragAcceptFiles Me.hwnd, False
    
    arc.Close
    Set arc = Nothing
    
    mnuArchive.Enabled = False
    mnuFileClose.Enabled = False
    frmZipper.Caption = "Zipper"
    lstFiles.ListItems.Clear
    lstFiles.Visible = False

    statusBar.Panels.Item(1).Text = "Archive closed."
    statusBar.Panels.Item(2).Text = "No archive open"
    statusBar.Panels.Item(3).Text = "No archive open"

End Sub
