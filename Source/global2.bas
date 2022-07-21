Attribute VB_Name = "global2"
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
   Y As Long
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


Public Function SelectDir(hwnd As Long, Title As String) As String
  
  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim Pos As Integer

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
     Pos = InStr(path, Chr$(0))
     SelectDir = Left$(path, Pos - 1)
  Else
     SelectDir = vbNullString
  End If

  Call CoTaskMemFree(pidl)

End Function

