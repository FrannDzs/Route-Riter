Attribute VB_Name = "global"
' Copyright ©1996-2002 VB<EM>net</EM>, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you many not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source on
'               any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const MAX_PATH = 260

Public Declare Function SHGetPathFromIDList Lib "shell32" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Public Declare Function SHBrowseForFolder Lib "shell32" _
   Alias "SHBrowseForFolderA" _
  (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Sub CoTaskMemFree Lib "ole32" _
   (ByVal pv As Long)

Public Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long                      '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80          '  out: type name
End Type
Public Const SHGFI_TYPENAME = &H400                     '  get type name
Public Const SHGFI_ICON = &H100                         '  get icon
Public Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute
Public Const SHGFI_SMALLICON = &H1                      '  get small icon

Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Type PictureDesc
  cbSize    As Long
  lPicType  As PictureTypeConstants
  hImage    As Long
  lWidth    As Long
  lHeight   As Long
End Type

Private Type Guid
  Data1         As Long
  Data2         As Integer
  Data3         As Integer
  Data4(0 To 7) As Byte
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictureDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you many not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source on
'               any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type POINTAPI
   x As Long
   y As Long
End Type
 
Public Type MSG
   hwnd As Long
   message As Long
   wParam As Long
   lParam As Long
   time As Long
   pt As POINTAPI
End Type

Public Declare Sub DragAcceptFiles Lib "shell32" _
  (ByVal hwnd As Long, _
   ByVal fAccept As Long)

Public Declare Sub DragFinish Lib "shell32" _
  (ByVal hDrop As Long)

Public Declare Function DragQueryFile Lib "shell32" _
   Alias "DragQueryFileA" _
  (ByVal hDrop As Long, _
   ByVal UINT As Long, _
   ByVal lpStr As String, _
   ByVal ch As Long) As Long

Public Declare Function PeekMessage Lib "user32" _
   Alias "PeekMessageA" _
  (lpMsg As MSG, _
   ByVal hwnd As Long, _
   ByVal wMsgFilterMin As Long, _
   ByVal wMsgFilterMax As Long, _
   ByVal wRemoveMsg As Long) As Long

Public Const PM_NOREMOVE = &H0
Public Const PM_NOYIELD = &H2
Public Const PM_REMOVE = &H1
Public Const WM_DROPFILES = &H233

Public Sub Main()
  
  'In order for this to function properly you should place of of your program
  'execution code in the Sub Main(), Make sure you change the project startup
  'to sub Main
   frmZipper.Show
  
  'This must be the last line! Nothing gets called after this
   WatchForFiles

End Sub
Public Sub WatchForFiles()
   
  'This watches for all WM_DROPFILES messages

   Dim FileDropMessage As MSG    'Msg Type
   Dim fileDropped As Boolean    'True if Files where dropped
   Dim hDrop As Long             'Pointer to the dropped file structure
   Dim filename As String * 128  'the dropped filename
   Dim numOfDroppedFiles As Long 'the number of dropped files
   Dim curFile As Long           'the current file number
   Dim ret As Long
   
  'loop to keep checking for files
  'NOTE: Do any code you want to execute before this set
   
   Do
      
      'check for Dropped file messages
       fileDropped = PeekMessage(FileDropMessage, 0, _
                     WM_DROPFILES, WM_DROPFILES, PM_REMOVE Or PM_NOYIELD)

       If fileDropped Then
        
         'get the pointer to the dropped file structure
          hDrop = FileDropMessage.wParam
         
         'get the total number of files
          numOfDroppedFiles = DragQueryFile(hDrop, True, filename, 127)

          For curFile = 1 To numOfDroppedFiles
             
             'get the file name
              ret = DragQueryFile(hDrop, curFile - 1, filename, 127)
             
              frmZipper.archive.AddFile filename

          Next curFile
         
          frmZipper.RefreshList
         'we are now done with the structure, tell windows to discard it
          DragFinish (hDrop)

   End If
  
  'be nice
   DoEvents

   Loop

End Sub

Public Function SelectDir(hwnd As Long, Title As String) As String
  
  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Integer

 'Fill the BROWSEINFO structure with the
 'needed data. To accomodate comments, the
 'With/End With sytax has not been used, though
 'it should be your 'final' version.

 'hwnd of the window that receives messages
 'from the call. Can be your application
 'or the handle from GetDesktopWindow().
  bi.hOwner = hwnd

 'Pointer to the item identifier list specifying
 'the location of the "root" folder to browse from.
 'If NULL, the desktop folder is used.
  bi.pidlRoot = 0&

 'message to be displayed in the Browse dialog
  bi.lpszTitle = Title

 'the type of folder to return.
  bi.ulFlags = BIF_RETURNONLYFSDIRS

 'show the browse for folders dialog
  pidl = SHBrowseForFolder(bi)

 'the dialog has closed, so parse & display the
 'users returned folder selection contained in pidl
  path = Space$(MAX_PATH)

  If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     SelectDir = Left(path, pos - 1)
  Else
     SelectDir = ""
  End If

  Call CoTaskMemFree(pidl)

End Function

Public Function GetFileType(sFile As String) As String

    Dim sfi As SHFILEINFO

    If SHGetFileInfo(sFile, 0&, _
                     sfi, Len(sfi), _
                     SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
        GetFileType = TrimNull(sfi.szTypeName)
    End If

End Function

Function TrimNull(s As String) As String

  Dim pos As Integer
  
  pos = InStr(s, Chr(0))
  If pos Then
    TrimNull = Left(s, pos - 1)
  Else
    TrimNull = s
  End If

End Function

Function GetSmallIcon(sFile As String) As Long

    Dim sfi As SHFILEINFO
    Dim ret As Long

    If SHGetFileInfo(sFile, 0&, sfi, Len(sfi), SHGFI_ICON Or _
        SHGFI_TYPENAME Or SHGFI_SMALLICON Or SHGFI_USEFILEATTRIBUTES) Then
        
        GetSmallIcon = sfi.hIcon
        
    End If
    
End Function


Public Function PictureFromHandle(ByVal hImage As Long, ByVal ePicType As PictureTypeConstants) As IPicture

'Create a Picture object from an Icon handle

Dim NewPic  As Picture
Dim PicConv As PictureDesc
Dim IGuid   As Guid
    
    If hImage <> 0 Then
        'Setup the descriptor
        PicConv.cbSize = Len(PicConv)
        PicConv.lPicType = ePicType
        PicConv.hImage = hImage
        
        'IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        With IGuid
            .Data1 = &H7BF80980
            .Data2 = &HBF32
            .Data3 = &H101A
            .Data4(0) = &H8B
            .Data4(1) = &HBB
            .Data4(2) = &H0
            .Data4(3) = &HAA
            .Data4(4) = &H0
            .Data4(5) = &H30
            .Data4(6) = &HC
            .Data4(7) = &HAB
        End With
        
        'Create the Picture object...
        OleCreatePictureIndirect PicConv, IGuid, True, NewPic
        
        '...and return it
        Set PictureFromHandle = NewPic
        
        'Cleanup memory
        Set NewPic = Nothing
    
    Else
        'Invalid icon handle
        Set PictureFromHandle = Nothing
        
    End If
    
End Function
