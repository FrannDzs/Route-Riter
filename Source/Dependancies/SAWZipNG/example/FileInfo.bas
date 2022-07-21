Attribute VB_Name = "basFileInfo"
Option Explicit

'Constants for File Info (Assoc file (Reg), Icons, etc)
Private Const MAX_PATH                      As Long = &H104&        'Max Path Len
Private Const SHGFI_SYSICONINDEX            As Long = &H4000&       'Get System Icon Index
Private Const SHGFI_LARGEICON               As Long = &H0&          'Get Large Icon
Private Const SHGFI_SMALLICON               As Long = &H1&          'Get Small Icon
Private Const SHGFI_DISPLAYNAME             As Long = &H200&        'Get File Display Name
Private Const SHGFI_TYPENAME                As Long = &H400&        'Get File Type Name
Private Const SHGFI_ICON                    As Long = &H100&        'Get icon
Private Const INVALID_HANDLE_VALUE          As Long = &HFFFFFFFF    'File not found
Private Const FILE_ATTRIBUTE_READONLY       As Long = &H1&          'Read Only File
Private Const FILE_ATTRIBUTE_HIDDEN         As Long = &H2&          'Hidden File
Private Const FILE_ATTRIBUTE_SYSTEM         As Long = &H4&          'System File
Private Const FILE_ATTRIBUTE_DIRECTORY      As Long = &H10&         'Folder
Private Const FILE_ATTRIBUTE_ARCHIVE        As Long = &H20&         'Archive File
Private Const HKEY_CLASSES_ROOT             As Long = &H80000000
Private Const HKEY_CURRENT_CONFIG           As Long = &H80000005
Private Const HKEY_CURRENT_USER             As Long = &H80000001
Private Const HKEY_DYN_DATA                 As Long = &H80000006
Private Const HKEY_LOCAL_MACHINE            As Long = &H80000002
Private Const HKEY_PERFORMANCE_DATA         As Long = &H80000004
Private Const HKEY_USERS                    As Long = &H80000003
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const ERROR_SUCCESS                 As Long = &H0&
Private Const SYNCHRONIZE                   As Long = &H100000
Private Const READ_CONTROL                  As Long = &H20000
Private Const STANDARD_RIGHTS_READ          As Long = (READ_CONTROL)
Private Const KEY_QUERY_VALUE               As Long = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS        As Long = &H8&
Private Const KEY_NOTIFY                    As Long = &H10&
Private Const KEY_READ                      As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
                                                KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

'Constants for File Version Info
Private Const VS_FF_DEBUG                   As Long = &H1&
Private Const VS_FF_INFOINFERRED            As Long = &H10&
Private Const VS_FF_PATCHED                 As Long = &H4&
Private Const VS_FF_PRERELEASE              As Long = &H2&
Private Const VS_FF_PRIVATEBUILD            As Long = &H8&
Private Const VS_FF_SPECIALBUILD            As Long = &H20&
Private Const VS_FFI_FILEFLAGSMASK          As Long = &H3F&
Private Const VS_FFI_SIGNATURE              As Long = &HFEEF04BD
Private Const VS_FFI_STRUCVERSION           As Long = &H10000
Private Const VFT_APP                       As Long = &H1&
Private Const VFT_DLL                       As Long = &H2&
Private Const VFT_DRV                       As Long = &H3&
Private Const VFT_FONT                      As Long = &H4&
Private Const VFT_STATIC_LIB                As Long = &H7&
Private Const VFT_UNKNOWN                   As Long = &H0&
Private Const VFT_VXD                       As Long = &H5&
Private Const VFT2_DRV_COMM                 As Long = &HA&
Private Const VFT2_DRV_DISPLAY              As Long = &H4&
Private Const VFT2_DRV_INSTALLABLE          As Long = &H8&
Private Const VFT2_DRV_KEYBOARD             As Long = &H2&
Private Const VFT2_DRV_LANGUAGE             As Long = &H3&
Private Const VFT2_DRV_MOUSE                As Long = &H5&
Private Const VFT2_DRV_NETWORK              As Long = &H6&
Private Const VFT2_DRV_PRINTER              As Long = &H1&
Private Const VFT2_DRV_SOUND                As Long = &H9&
Private Const VFT2_DRV_SYSTEM               As Long = &H7&
Private Const VFT2_FONT_RASTER              As Long = &H1&
Private Const VFT2_FONT_TRUETYPE            As Long = &H3&
Private Const VFT2_FONT_VECTOR              As Long = &H2&
Private Const VFT2_UNKNOWN                  As Long = &H0&
Private Const VOS__BASE                     As Long = &H0&
Private Const VOS__PM16                     As Long = &H2&
Private Const VOS__PM32                     As Long = &H3&
Private Const VOS__WINDOWS16                As Long = &H1&
Private Const VOS__WINDOWS32                As Long = &H4&
Private Const VOS_DOS                       As Long = &H10000
Private Const VOS_DOS_WINDOWS16             As Long = &H10001
Private Const VOS_DOS_WINDOWS32             As Long = &H10004
Private Const VOS_NT                        As Long = &H40000
Private Const VOS_NT_WINDOWS32              As Long = &H40004
Private Const VOS_OS216                     As Long = &H20000
Private Const VOS_OS216_PM16                As Long = &H20002
Private Const VOS_OS232                     As Long = &H30000
Private Const VOS_OS232_PM32                As Long = &H30003
Private Const VOS_UNKNOWN                   As Long = &H0&

'Types for File Version Info
Private Type VS_FIXEDFILEINFO
   dwSignature          As Long
   dwStrucVersionl      As Integer
   dwStrucVersionh      As Integer
   dwFileVersionMSl     As Integer
   dwFileVersionMSh     As Integer
   dwFileVersionLSl     As Integer
   dwFileVersionLSh     As Integer
   dwProductVersionMSl  As Integer
   dwProductVersionMSh  As Integer
   dwProductVersionLSl  As Integer
   dwProductVersionLSh  As Integer
   dwFileFlagsMask      As Long
   dwFileFlags          As Long
   dwFileOS             As Long
   dwFileType           As Long
   dwFileSubtype        As Long
   dwFileDateMS         As Long
   dwFileDateLS         As Long
End Type

Public Type FileVersionInfo
    sFileVer    As String
    sProdVer    As String
    sFileFlags  As String
    sFileOS     As String
    sFileType   As String
    sSubType    As String
End Type

'Types for File Info (Assoc file (Reg), Icons, etc)
Public Type FileInfo
    picLgIcon   As StdPicture
    picSmIcon   As StdPicture
    sDispName   As String
    sTypeName   As String
    lSizeKB     As Long
    tModDate    As Date
    sAttribs    As String
    bError      As Boolean
    sErrorMsg   As String
End Type

Private Type SHFILEINFO
    hIcon           As Long                 'Icon handle
    iIcon           As Long                 'Icon index
    dwAttributes    As Long                 'SFGAO_flags
    szDisplayName   As String * MAX_PATH    'Display name (or path)
    szTypeName      As String * 80          'Type name
End Type

Private Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Private Type SYSTEMTIME
    wYear       As Integer
    wMonth      As Integer
    wDayOfWeek  As Integer
    wDay        As Integer
    wHour       As Integer
    wMinute     As Integer
    wSecond     As Integer
    wMillisecs  As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

'APIs for File Info (Assoc file (Reg), Icons, etc)
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'APIs for File Version Info
Private Declare Function GetFileVersionInfo1 Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

Public Function GetFileInfo(ByVal sFilename As String) As FileInfo

Dim lRet    As Long
Dim lFlags  As Long
Dim hFind   As Long
Dim lPos    As Long
Dim lAttr   As Long
Dim dSize   As Double
Dim fiFile  As FileInfo
Dim sfiFile As SHFILEINFO
Dim fdData  As WIN32_FIND_DATA
Dim ftTime  As FILETIME     'Local file date/time
Dim stTime  As SYSTEMTIME   'System file date/time

    'Set the FileInfo Pictures to Nothing
    With fiFile
        Set .picLgIcon = Nothing
        Set .picSmIcon = Nothing
    End With
    
    'Setup the flags
    lFlags = SHGFI_ICON Or SHGFI_DISPLAYNAME Or _
        SHGFI_TYPENAME Or SHGFI_SYSICONINDEX

    'Find the file
    hFind = FindFirstFile(sFilename, fdData)
    If hFind <> INVALID_HANDLE_VALUE Then
        
        'Get the file info and small icon
        lRet = SHGetFileInfo(sFilename, &H0&, sfiFile, _
          Len(sfiFile), lFlags Or SHGFI_SMALLICON)
        If lRet <> 0 Then
            With fiFile
                'Create a picture from the icon handle
                If sfiFile.hIcon <> 0 Then
                    Set .picSmIcon = PictureFromHandle(sfiFile.hIcon, vbPicTypeIcon)
                    'Free the memory
                    Call DestroyIcon(sfiFile.hIcon)
                    sfiFile.hIcon = 0
                End If
                
                'Get the file info again for the large icon
                lRet = SHGetFileInfo(sFilename, &H0&, sfiFile, _
                  Len(sfiFile), lFlags)
                
                'Create a picture from the icon handle
                If sfiFile.hIcon <> 0 Then
                    Set .picLgIcon = PictureFromHandle(sfiFile.hIcon, vbPicTypeIcon)
                    'Free the memory
                    Call DestroyIcon(sfiFile.hIcon)
                    sfiFile.hIcon = 0
                End If
                
                'Get the display name
                lPos = InStr(1, sfiFile.szDisplayName, Chr$(0))
                If lPos > 1 Then
                    .sDispName = Left$(sfiFile.szDisplayName, lPos - 1)
                End If
                
                'File Size
                dSize = (CDbl(fdData.nFileSizeHigh) * CDbl(2# ^ 32#)) _
                  + CDbl(fdData.nFileSizeLow) / 1024#
                .lSizeKB = CLng(dSize)
                If dSize > .lSizeKB Then
                    .lSizeKB = .lSizeKB + 1
                End If
                
                'File Type
                .sTypeName = Left$(sfiFile.szTypeName, _
                    InStr(1, sfiFile.szTypeName, vbNullChar) - 1)
        
                'Last Modified (Translate to local system time)
                Call FileTimeToLocalFileTime(fdData.ftLastWriteTime, ftTime)
                Call FileTimeToSystemTime(ftTime, stTime)
                .tModDate = DateSerial(stTime.wYear, stTime.wMonth, _
                    stTime.wDay) + TimeSerial(stTime.wHour, _
                    stTime.wMinute, stTime.wSecond)
                
                'Attributes
                lAttr = fdData.dwFileAttributes
                .sAttribs = IIf((lAttr And FILE_ATTRIBUTE_READONLY) > 0, "R", "") _
                    & IIf((lAttr And FILE_ATTRIBUTE_HIDDEN) > 0, "H", "") _
                    & IIf((lAttr And FILE_ATTRIBUTE_SYSTEM) > 0, "S", "") _
                    & IIf((lAttr And FILE_ATTRIBUTE_ARCHIVE) > 0, "A", "")
            
            End With
        
        Else
            fiFile.bError = True
            fiFile.sErrorMsg = "No system icons associated with this file type."
        End If
        
        Call FindClose(hFind)
    
    Else
        fiFile.bError = True
        fiFile.sErrorMsg = "File not found."
        
    End If
    
    GetFileInfo = fiFile
    
End Function

Public Function GetLongFilename(ByVal sShortFilename As String) As String

'Returns the Long Filename associated with sShortFilename.
'Note: sShortFilename must be a valid filename.

'Note: The GetLongPathName() API will not work on Win95 or WinNT.
'This function will work for all Windows versions 95/NT 3.1 and later.

Dim lRet    As Long
Dim lPos    As Long
Dim lChars  As Long
Dim hFind   As Long
Dim sFile   As String
Dim WFData  As WIN32_FIND_DATA

    'FindFirstFile returns the last element of the path
    '(file/folder name) as a long filename.

    'Work backwards through the path, getting the long
    'file/folder name for the last element in the path and
    'then removing that element from the path in each loop.
    
    'For example: "C:\Progra~1\Common~1\Services\yahoo.bmp"
    'The first loop will return "yahoo.bmp"
    '   (sFile = "\yahoo.bmp")
    'The second loop will return "Services"
    '   (sFile = "\Services\yahoo.bmp")
    'The third loop will return "Common Files"
    '   (sFile = "\Common Files\Services\yahoo.bmp")
    'The fourth loop will return "Program Files"
    '   (sFile = "\Program Files\Common Files\Services\yahoo.bmp")
    'The last loop will fail to find "\", so it
    'will drop to the else and prepend "C:")
    '   (sFile = "C:\Program Files\Common Files\Services\yahoo.bmp")
    'So, sFile is now the long filename
    
    lPos = InStrRev(sShortFilename, "\")
    Do While Len(sShortFilename) > 0
        If lPos > 0 Then
            'Get the long file/folder name for the last element
            hFind = FindFirstFile(sShortFilename, WFData)
            
            'If the file/folder is found...
            If hFind <> INVALID_HANDLE_VALUE Then
                lChars = InStr(1, WFData.cFileName, Chr$(0)) - 1
                If lChars > 0 Then
                    'Prepend sFile with "\" and the file/folder name
                    sFile = "\" & Left$(WFData.cFileName, lChars) & sFile
                End If
            End If
            lRet = FindClose(hFind)
            
            'Remove the last element of the path
            sShortFilename = Left$(sShortFilename, lPos - 1)
            
            'Move to the previous element
            lPos = InStrRev(sShortFilename, "\")
        Else
            'Prepend what's left of the original filename (ie., "C:")
            sFile = sShortFilename & sFile
            
            'Drop out of the loop
            Exit Do
        End If
    Loop
    
    'Return the long filename.
    GetLongFilename = sFile
    
End Function

Public Function GetShortFilename(ByVal sLongFilename As String) As String

Dim lLen    As Long
Dim sBuffer As String

    sBuffer = String(1024, Chr$(0))
    lLen = GetShortPathName(sLongFilename, sBuffer, Len(sBuffer))
    If lLen > 0 Then
        GetShortFilename = Mid$(sBuffer, 1, lLen)
    End If
    
End Function


Public Function GetAssocExeFilename(ByVal sFileExt As String) As String

Dim lPos1   As Long
Dim lPos2   As Long
Dim sRegKey As String
Dim sValue  As String
Dim sTemp1  As String
Dim sTemp2  As String

    If Left$(sFileExt, 1) <> "." Then
        sFileExt = "." & sFileExt
    End If
    
    sRegKey = "HKEY_CLASSES_ROOT\" & sFileExt & "\(Default)"
    sValue = GetRegString(sRegKey, True)
    
    If Len(sValue) > 0 Then
        sRegKey = "HKEY_CLASSES_ROOT\" & sValue & "\SHell\Open\Command\(Default)"
        sValue = Trim$(GetRegString(sRegKey))
        If Len(sValue) > 0 Then
            'Strip quotes (also removes command line parameters, if any)
            If InStr(1, sValue, Chr$(34)) = 1 Then
                lPos1 = InStr(2, sValue, Chr$(34))
                If lPos1 > 0 Then
                    sValue = Mid$(sValue, 2, lPos1 - 2)
                Else
                    'Error, if no end quotes
                    sValue = ""
                End If
            Else
                'Remove command line parameters, if any
                lPos1 = InStrRev(sValue, ".")
                If lPos1 > 0 Then
                    sValue = Left$(sValue, lPos1 + 3)
                Else
                    'Error, if no file extension
                    sValue = ""
                End If
            End If
            If Len(sValue) > 0 Then
                'Replace path variables (%ProgramFiles%)
                lPos1 = InStr(1, sValue, "%")
                Do While lPos1 > 0
                    lPos2 = InStr(lPos1 + 1, sValue, "%")
                    If lPos2 > 0 Then
                        'Get the path variable with delimiters(%)
                        sTemp1 = Mid$(sValue, lPos1, lPos2 - lPos1 + 1)
                        'Get the path variable without delimiters(%)
                        sTemp2 = Replace(sTemp1, "%", "")
                        'Get the real path
                        sTemp2 = Environ(sTemp2)
                        'Replace the path variable with the real path
                        sValue = Replace(sValue, sTemp1, sTemp2)
                        lPos1 = InStr(1, sValue, "%")
                    Else
                        'Error, if no end delimiter
                        sValue = ""
                        Exit Do
                    End If
                Loop
            End If
        End If
    End If
    
    GetAssocExeFilename = sValue

End Function

Public Function GetRegString(ByVal sKeyPath As String, Optional ByVal bSilent As Boolean = True) As String

'Returns a String Value from the registry.

'sKeyPath = Full path to the Registry Key.
'(ex. HKEY_CLASSES_ROOT\giffile\shell\Open\command\command")

'For Default Key use "(Default)"
'(ex. HKEY_CLASSES_ROOT\giffile\shell\Open\command\(Default)")

'bSilent: If True, don't show any error messages

'Returns: The Key's string value or an empty string, if
'         an error occurred or the value is not set.

Dim hKey        As Long
Dim lPos        As Long
Dim lRet        As Long
Dim lHandle     As Long
Dim lType       As Long
Dim lSize       As Long
Dim sKey        As String
Dim sSubKey     As String
Dim sReturn     As String
Const lReserved As Long = &H0&

    'Find the first path separator
    lPos = InStr(sKeyPath, "\")
    
    If lPos > 0 Then
        
        'Setup the Buffer
        lSize = 1024
        sReturn = String$(lSize, " ")
        
        'Extract the Hive from the Path
        hKey = GetHiveKey(Left$(UCase$(sKeyPath), lPos - 1))
        sSubKey = Mid$(sKeyPath, lPos + 1)
        
        'Find the last path separator
        lPos = InStrRev(sSubKey, "\")
        If lPos > 0 Then
            
            'Extract the Key and SubKey
            If lPos < Len(sSubKey) Then
                sKey = Trim$(Mid$(sSubKey, lPos + 1))
            Else
                'Extract default value
                sKey = ""
            End If
            sSubKey = Trim$(Left$(sSubKey, lPos - 1))
            
            'To get the (Default) value Key must be an empty string ("")
            If LCase$(sKey) = "(default)" Or LCase$(sKey) = "default" Then
                sKey = ""
            End If
            
            'Open the Key
            lRet = RegOpenKeyEx(hKey, sSubKey, lReserved, KEY_READ, lHandle)
            
            'If successful opening the Key...
            If lRet = ERROR_SUCCESS Then
            
                'Get the Value
                lRet = RegQueryValueEx(lHandle, sKey, lReserved, _
                  lType, ByVal sReturn, lSize)
                
                'If successful getting the Value...
                If lRet = ERROR_SUCCESS Then
                    'Return the string Value
                    GetRegString = Left$(sReturn, lSize - 1)
                Else
                    'Error getting the Value
                    GetRegString = ""
                    If Not bSilent Then
                        Call ShowRegError(lRet)
                    End If
                End If
                
                lRet = RegCloseKey(lHandle)
            
            Else
                'Error opening the Key
                GetRegString = ""
                If Not bSilent Then
                    Call ShowRegError(lRet)
                End If
            End If
        
        Else
            'Invalid path; No Key/Subkey defined
            GetRegString = ""
            If Not bSilent Then
                MsgBox "Invalid Registry Key passed to GetRegString", vbExclamation, "Key Not Found"
            End If
        End If
        
    End If
    
End Function

Private Sub ShowRegError(ByVal lErrNbr As Long)

'Get the description for the registry error and show it.

Dim lRet    As Long
Dim lBuffSz As Long
Dim sBuffer As String
Dim sMsg    As String

    'Setup the buffer
    lBuffSz = 1024
    sBuffer = String$(lBuffSz, " ")
    
    'Get the error description for lErrNbr
    lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
      FORMAT_MESSAGE_IGNORE_INSERTS, &H0&, CLng(lErrNbr), _
      &H0&, sBuffer, lBuffSz, vbNull)
    
    'Setup the message
    If lRet > 0 Then
        sMsg = Left$(sBuffer, lRet)
    Else
        MsgBox "Unknown error (#" & CStr(lErrNbr) & _
          ") occurred while attempting to read from the registry"
    End If
    
    'Show the error.
    MsgBox sMsg, vbExclamation, "Error Reading Registry"
    
End Sub

Private Function GetHiveKey(ByVal sKeyName As String) As Long

'Translate the string Hive Key name to it's long value.

    Select Case sKeyName
        Case "HKEY_CLASSES_ROOT"
            GetHiveKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_CONFIG"
            GetHiveKey = HKEY_CURRENT_CONFIG
        Case "HKEY_CURRENT_USER"
            GetHiveKey = HKEY_CURRENT_USER
        Case "HKEY_DYN_DATA"
            GetHiveKey = HKEY_DYN_DATA
        Case "HKEY_LOCAL_MACHINE"
            GetHiveKey = HKEY_LOCAL_MACHINE
        Case "HKEY_PERFORMANCE_DATA"
            GetHiveKey = HKEY_PERFORMANCE_DATA
        Case "HKEY_USERS"
            GetHiveKey = HKEY_USERS
    End Select

End Function

Public Function GetFileVersionInfo(ByVal sFilename As String) As FileVersionInfo

Dim lRet        As Long
Dim lTemp       As Long
Dim lDummy      As Long
Dim yaBuff()    As Byte
Dim lBuffLen    As Long
Dim lPointer    As Long
Dim sfiFile     As SHFILEINFO
Dim fviTemp     As FileVersionInfo
Dim ffiVersion  As VS_FIXEDFILEINFO

    'Get the size of the structure needed for this file's version info
    lBuffLen = GetFileVersionInfoSize(sFilename, lDummy)
    If lBuffLen < 1 Then
        Exit Function
    End If
     
    'Read the data into a Byte array
    ReDim yaBuff(lBuffLen)
    lRet = GetFileVersionInfo1(sFilename, 0&, lBuffLen, yaBuff(0))
    
    'Retrieve a pointer to the root block and move
    'the bytes into the VS_FIXEDFILEINFO structure.
    lRet = VerQueryValue(yaBuff(0), "\", lPointer, lDummy)
    MoveMemory ffiVersion, lPointer, Len(ffiVersion)
    
    With fviTemp
        'Get the File Version number
        .sFileVer = Format$(ffiVersion.dwFileVersionMSh) _
          & "." & Format$(ffiVersion.dwFileVersionMSl) _
          & "." & Format$(ffiVersion.dwFileVersionLSh) _
          & "." & Format$(ffiVersion.dwFileVersionLSl)
        
        'Get the Product Version number
        .sProdVer = Format$(ffiVersion.dwProductVersionMSh) _
          & "." & Format$(ffiVersion.dwProductVersionMSl) _
          & "." & Format$(ffiVersion.dwProductVersionLSh) _
          & "." & Format$(ffiVersion.dwProductVersionLSl)
        
        'Get the Flag attributes
        lTemp = ffiVersion.dwFileFlags
        .sFileFlags = ""
        If (lTemp And VS_FF_DEBUG) <> 0 Then
            .sFileFlags = .sFileFlags & "Debug "
        End If
        If (lTemp And VS_FF_PRERELEASE) <> 0 Then
            .sFileFlags = .sFileFlags & "PreRelease "
        End If
        If (lTemp And VS_FF_PATCHED) <> 0 Then
            .sFileFlags = .sFileFlags & "Patched "
        End If
        If (lTemp And VS_FF_PRIVATEBUILD) <> 0 Then
            .sFileFlags = .sFileFlags & "Private "
        End If
        If (lTemp And VS_FF_INFOINFERRED) <> 0 Then
            .sFileFlags = .sFileFlags & "Info "
        End If
        If (lTemp And VS_FF_SPECIALBUILD) <> 0 Then
            .sFileFlags = .sFileFlags & "Special "
        End If
        If (lTemp And VFT2_UNKNOWN) <> 0 Then
            .sFileFlags = .sFileFlags + "Unknown "
        End If
        .sFileFlags = Trim(.sFileFlags)
        If Len(.sFileFlags) = 0 Then
            .sFileFlags = "(None)"
        End If
        
        'Get the OS that the file was designed for
        Select Case ffiVersion.dwFileOS
            Case VOS__PM16
                .sFileOS = "PM-16"
            Case VOS__PM32
                .sFileOS = "PM-32"
            Case VOS__WINDOWS16
                .sFileOS = "Win16"
            Case VOS__WINDOWS32
                .sFileOS = "Win32"
            Case VOS_DOS
                .sFileOS = "DOS"
            Case VOS_DOS_WINDOWS16
                .sFileOS = "DOS-Win16"
            Case VOS_DOS_WINDOWS32
                .sFileOS = "DOS-Win32"
            Case VOS_NT
                .sFileOS = "NT"
            Case VOS_NT_WINDOWS32
                .sFileOS = "NT-Win32"
            Case VOS_OS216
                .sFileOS = "OS/2-16"
            Case VOS_OS232
                .sFileOS = "OS/2-32"
            Case VOS_OS216_PM16
                .sFileOS = "OS/2-16 PM-16"
            Case VOS_OS232_PM32
                .sFileOS = "OS/2-32 PM-32"
            Case Else
                .sFileOS = "(Unknown)"
        End Select
        
        'Get the File Type
        lRet = SHGetFileInfo(sFilename, &H0&, sfiFile, Len(sfiFile), SHGFI_TYPENAME)
        If lRet <> 0 Then
            .sFileType = Left$(sfiFile.szTypeName, InStr(1, sfiFile.szTypeName, vbNullChar) - 1)
        Else
            .sFileType = "(Unknown)"
        End If
        
        'Get the File SubType
        Select Case ffiVersion.dwFileType
            Case VFT_DRV
                Select Case ffiVersion.dwFileSubtype
                    Case VFT2_DRV_PRINTER
                        .sSubType = "Printer drv"
                    Case VFT2_DRV_KEYBOARD
                        .sSubType = "Keyboard drv"
                    Case VFT2_DRV_LANGUAGE
                        .sSubType = "Language drv"
                    Case VFT2_DRV_DISPLAY
                        .sSubType = "Display drv"
                    Case VFT2_DRV_MOUSE
                        .sSubType = "Mouse drv"
                    Case VFT2_DRV_NETWORK
                        .sSubType = "Network drv"
                    Case VFT2_DRV_SYSTEM
                        .sSubType = "System drv"
                    Case VFT2_DRV_INSTALLABLE
                        .sSubType = "Installable"
                    Case VFT2_DRV_SOUND
                        .sSubType = "Sound drv"
                    Case VFT2_DRV_COMM
                        .sSubType = "Comm drv"
                    Case VFT2_UNKNOWN
                        .sSubType = "(Unknown)"
                End Select
            Case VFT_FONT
               Select Case ffiVersion.dwFileSubtype
                    Case VFT2_FONT_RASTER
                        .sSubType = "Raster Font"
                    Case VFT2_FONT_VECTOR
                        .sSubType = "Vector Font"
                    Case VFT2_FONT_TRUETYPE
                        .sSubType = "TrueType Font"
               End Select
        End Select
    End With
    
    GetFileVersionInfo = fviTemp
    
End Function

Public Function FileExists(ByVal sFilename As String) As Boolean

Dim hFind   As Long
Dim fdData  As WIN32_FIND_DATA
    
    'Attempt to Find the file
    hFind = FindFirstFile(sFilename, fdData)
    If hFind <> INVALID_HANDLE_VALUE Then
        Call FindClose(hFind)
        FileExists = True
    Else
        FileExists = False
    End If
    
End Function

Public Function GetFileTitle(ByVal sFilename As String) As String

Dim lPos    As Long

    lPos = InStrRev(sFilename, "\")
    If lPos > 0 Then
        If lPos < Len(sFilename) Then
            GetFileTitle = Mid$(sFilename, lPos + 1)
        Else
            GetFileTitle = ""
        End If
    Else
        GetFileTitle = sFilename
    End If
    
End Function

Public Function GetFilePath(ByVal sFilename As String, Optional ByVal bAddBackslash As Boolean = False) As String

Dim lPos    As Long

    lPos = InStrRev(sFilename, "\")
    If lPos > 0 Then
        If Not bAddBackslash Then
            lPos = lPos - 1
        End If
        GetFilePath = Mid$(sFilename, 1, lPos)
    Else
        GetFilePath = ""
    End If
    
End Function

