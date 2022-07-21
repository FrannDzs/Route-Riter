VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompressZIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Property Variables:
Public m_CompressedSize As Long
Public m_OriginalSize As Long

'Declares
'Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
'Private Declare Function GetFocus Lib "user32" () As Long
'Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Const ZLIB_NOERROR = 0

'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappingAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
'Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
'Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
'Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
'Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
'Private Const GENERIC_WRITE = &H40000000
'Private Const GENERIC_READ = &H80000000
'Private Const FILE_SHARE_WRITE = &H2
'Private Const FILE_SHARE_READ = &H1
'Private Const CREATE_ALWAYS = 2
'Private Const OPEN_EXISTING = 3
'Private Const PAGE_NOACCESS = &H1&
'Private Const PAGE_READONLY = &H2&
'Private Const PAGE_READWRITE = &H4&
'Private Const PAGE_WRITECOPY = &H8&
'Private Const PAGE_EXECUTE = &H10&
'Private Const PAGE_EXECUTE_READ = &H20&
'Private Const PAGE_EXECUTE_READWRITE = &H40&
'Private Const PAGE_EXECUTE_WRITECOPY = &H80&
'Private Const PAGE_GUARD = &H100&
'Private Const PAGE_NOCACHE = &H200&
'Private Const FILE_BEGIN = 0
'Private Const SECTION_QUERY = &H1
'Private Const SECTION_MAP_WRITE = &H2
'Private Const SECTION_MAP_READ = &H4
'Private Const SECTION_MAP_EXECUTE = &H8
'Private Const SECTION_EXTEND_SIZE = &H10

'Private Const FILE_MAP_WRITE = SECTION_MAP_WRITE
'Private Const FILE_MAP_COPY = SECTION_QUERY
'Private Const FILE_MAP_READ = SECTION_MAP_READ

'Private Type SYSTEM_INFO
'   dwOemID As Long
'   dwPageSize As Long
'   lpMinimumApplicationAddress As Long
'   lpMaximumApplicationAddress As Long
'   dwActiveProcessorMask As Long
'   dwNumberOrfProcessors As Long
'   dwProcessorType As Long
'   dwAllocationGranularity As Long
'   dwReserved As Long
'End Type
'Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

' To Report API errors:
'Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
'Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
'Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
'Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
'Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Enum CZErrors
    [Insufficient Buffer] = -5
End Enum

Public Function WinAPIError(ByVal lLastDLLError As Long) As String
Dim sBuff As String
Dim lCount As Long
    
   ' Return the error message associated with LastDLLError:
   sBuff = String$(256, 0)
   lCount = FormatMessage( _
      FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
      0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
   If lCount Then
      WinAPIError = Left$(sBuff, lCount)
   End If

End Function



Public Function CompressData(TheData() As Byte) As Long
Dim lResult As Long
'Allocate memory for byte array
Dim lBufferSize As Long
Dim bTempBuffer() As Byte
Dim iL As Long, iU As Long

   m_OriginalSize = 0
   m_CompressedSize = 0
   
   GetBounds TheData, iL, iU
   If iU > iL Then
      ' Store the original size of this data:
      m_OriginalSize = iU - iL + 1
   
      ' Prepare the area to compress into.
      ' Ensure we have sufficient space for the worst possible case (no
      ' compression plus additional space for the compression info):
      lBufferSize = m_OriginalSize
      lBufferSize = lBufferSize + (lBufferSize * 0.01) + 12
      ReDim bTempBuffer(0 To lBufferSize - 1) As Byte

      'Compress byte array (data):
      lResult = compress(bTempBuffer(0), lBufferSize, TheData(iL), m_OriginalSize)
      
      ' Result is an error code
      If lResult = ZLIB_NOERROR Then
      
         ' lBufferSize will have been set by zlib:
         m_CompressedSize = lBufferSize
         
         ' If we got data back:
         If m_CompressedSize > 0 Then
            'Truncate to actual compressed size
            ReDim Preserve TheData(0 To lBufferSize - 1) As Byte
            ' Return data in buffer:
            CopyMemory TheData(0), bTempBuffer(0), lBufferSize
            
         Else
            Erase TheData
         End If
            
         'Set properties if no error occurred
         m_CompressedSize = lBufferSize
   
      End If
      
      'Cleanup
      Erase bTempBuffer

   End If

   'Return error code (if any)
   CompressData = lResult

End Function

Public Function CompressString(TheString As String) As Long
Dim lResult As Long
Dim lCmpSize As Long
Dim sTBuff As String

   m_CompressedSize = 0
   m_OriginalSize = Len(TheString)

   'Allocate string space for the buffers
   
   lCmpSize = m_OriginalSize
   lCmpSize = lCmpSize + (lCmpSize * 0.01) + 12
   sTBuff = String$(lCmpSize, 0)

   'Compress string (temporary string buffer) data
   lResult = compress(ByVal sTBuff, lCmpSize, ByVal TheString, Len(TheString))

   If lResult = ZLIB_NOERROR Then
      
      'Crop the string and set it to the actual string.
      TheString = Left$(sTBuff, lCmpSize)

      'Set compressed size of string.
      m_CompressedSize = lCmpSize

      'Cleanup
      sTBuff = vbNullString
   
   Else
      ' Error
      m_OriginalSize = 0
      
   End If

   'Return error code (if any)
   CompressString = lResult
   
End Function

Public Function DecompressData( _
      TheData() As Byte, _
      OrigSize As Long _
   ) As Long
Dim lResult As Long
Dim lBufferSize As Long
Dim bTempBuffer() As Byte
Dim iL As Long, iU As Long
 On Error GoTo Errtrap
   m_OriginalSize = 0
   m_CompressedSize = 0

   GetBounds TheData, iL, iU
   If iU > iL Then

      m_OriginalSize = 0
      m_CompressedSize = iU - iL + 1

      lBufferSize = OrigSize + 1
      ReDim bTempBuffer(0 To lBufferSize - 1) As Byte

      'Decompress data
      lResult = uncompress(bTempBuffer(0), lBufferSize, TheData(0), UBound(TheData) + 1)

      'Reset properties
      If lResult = ZLIB_NOERROR Then
         m_OriginalSize = lBufferSize
         'Truncate buffer to compressed size
         ' removed preserve
         ReDim TheData(0 To lBufferSize - 1) As Byte
         CopyMemory TheData(0), bTempBuffer(0), lBufferSize
      Else
         ' error
         m_CompressedSize = 0
         m_OriginalSize = 0
      End If
      
   End If
   
   'Return error code (if any)
   DecompressData = lResult
   Exit Function
Errtrap:
   
DecompressData = lResult
End Function
Public Function DecompressString( _
      TheString As String, _
      OrigSize As Long _
   ) As Long
Dim lResult As Long

'Allocate string space
Dim lCmpSize As Long
Dim sTBuff As String
    
   m_CompressedSize = Len(TheString)
   m_OriginalSize = 0
   
   sTBuff = String$(OrigSize + 1, 0)
   lCmpSize = Len(sTBuff)

   'Decompress
   lResult = uncompress(ByVal sTBuff, lCmpSize, ByVal TheString, m_CompressedSize)
    
   If lResult = ZLIB_NOERROR Then

      'Make string the size of the uncompressed string
      TheString = Left$(sTBuff, lCmpSize)
      m_OriginalSize = lCmpSize
   
   Else
      ' Error:
      m_CompressedSize = 0
      m_OriginalSize = 0
   End If
   
   'Return error code (if any)
   DecompressString = lResult
   

End Function

Private Sub GetBounds(TheData() As Byte, iL As Long, iU As Long)
On Error Resume Next ' irritating issue with Ubound & LBound
   iL = LBound(TheData)
   If Err.Number = 0 Then
      iU = UBound(TheData)
   Else
      iL = 0: iU = 0
   End If
End Sub

'Private Function UnsignedAdd _
'   (Start As Long, Incr As Long) As Long
'   ' only works with positive increments
'   If Start And &H80000000 Then 'Start < 0
'      UnsignedAdd = Start + Incr
'   ElseIf (Start Or &H80000000) < -Incr Then
'      UnsignedAdd = Start + Incr
'   Else
'      UnsignedAdd = (Start + &H80000000) + _
'            (Incr + &H80000000)
'   End If
'End Function

