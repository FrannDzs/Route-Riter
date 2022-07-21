Attribute VB_Name = "myfile"
Option Explicit

Declare Function GetShortPathName Lib "kernel32" _
         Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
         ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
         
        

Public Enum tdcErrorHandlerResponse
    tdcResume = 0
    tdcResumeNext = 1
    tdcExitSub = 2
    tdcUnrecoverable = 3
    tdcUnrecognized = 4
End Enum




         
Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long



Type FILE_HEADER                    ' Structure for document file header
    lVersion As Long
End Type




Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STILL_ACTIVE = &H103

  Private Declare Function OpenProcess Lib "kernel32" (ByVal _
   dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal _
   dwProcessID As Long) As Long
   Private Declare Sub Sleep Lib "kernel32" (ByVal _
   dwMilliseconds As Long)
   Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
   As Long) As Long


    
   



''*********************************************************************************
' Purpose:  Extract various information about a path
'*********************************************************************************

Public Sub ExtractFileParts(strSourceFileName As String, strExtractedPath As String _
, strExtractedFile As String, strExtractedFileType As String)
    On Error GoTo errorHandler
    
    Dim lngDotPosition As Long
    Dim lngBackSlashPosition As Long
    Dim lngCount As Long
    
    'set impossible values for lngDotPosition and lngBackSlashPosition
    lngDotPosition = -1
    lngBackSlashPosition = -1
    'Determine the "." position and the "\" position so we can extract the file title - extention
    For lngCount = Len(strSourceFileName) To 1 Step -1
        If Mid$(strSourceFileName, lngCount, 1) = "." And lngDotPosition = -1 Then
            lngDotPosition = lngCount
        ElseIf Mid$(strSourceFileName, lngCount, 1) = "\" And lngBackSlashPosition = -1 Then
            lngBackSlashPosition = lngCount
        End If
    Next

    'if no dot was found or dot is before the last "\" then set to the end
    If lngDotPosition = -1 Or lngDotPosition < lngBackSlashPosition Then
        lngDotPosition = Len(strSourceFileName) + 1
    End If
    
    'determine the extracted path
    If lngBackSlashPosition > 0 Then
        strExtractedPath = Left$(strSourceFileName, lngBackSlashPosition)
        
        'determine the extracted file
        strExtractedFile = Mid$(strSourceFileName, lngBackSlashPosition + 1, lngDotPosition - lngBackSlashPosition - 1)
    Else
        strExtractedPath = ""
        
        strExtractedFile = Mid$(strSourceFileName, 1, lngDotPosition - 1)
    End If
    
    
    'determine the extracted Type
    If lngDotPosition <= Len(strSourceFileName) Then
        strExtractedFileType = UCase(Right$(strSourceFileName, Len(strSourceFileName) - lngDotPosition))
    End If

    Exit Sub 'avoid executing the error handler
    
errorHandler:
'    If lngHandleError(Err.Number, mstrcModule, Err.Description, False) = tdcResume Then
'        Resume
'    Else
'        Err.Raise Err.Number, mstrcModule, Err.Description
'    End If
End Sub



'*********************************************************************************
' Purpose:  Gets a short version of a file name
'*********************************************************************************

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

Public Function ShellAndWait(PathName As String, blnDoEvents As Boolean, WS As VbAppWinStyle, frmFocusForm As Form) As Double
 
 ' Drop-in replacement for VB's Shell command- except it
 ' doesn't return until shelled app is done. J.LeVasseur
 ' <lvasseur@tiac.net> Pieced together from a couple of
 ' other people's ideas, actually. Works for me ...
 '-------------------------------------------------------
    Dim lhProcess    As Long
    Dim lExitcode    As Long
    Dim dProcessID    As Double
    '----------------------------
    On Error GoTo errShellAndWait
    dProcessID = Shell(PathName, WS)
    lhProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, dProcessID)
        
'    If Not frmFocusForm Is Nothing Then
'        'bring back into focus
'        frmFocusForm.SetFocus
'    End If
    
    Do
      Call Sleep(50) ': DoEvents
      If blnDoEvents Then
        DoEvents
      End If
      Call GetExitCodeProcess(lhProcess, lExitcode)
    Loop While (lExitcode = STILL_ACTIVE)
    CloseHandle (lhProcess)
    ShellAndWait = dProcessID
    Exit Function
    
errShellAndWait:
    If lhProcess <> 0 Then
      CloseHandle (lhProcess)
    End If
    ShellAndWait = dProcessID
 End Function


Sub SetImagePath(ByVal path As String, ctl As Control)
    Dim RStr As String
    Dim plen As Integer
    Dim MSG As String

    plen = Len(path)

    While plen > 0
        RStr = Right$(path, 1)
        If RStr = ":" Or RStr = "\" Then
            plen = 0
        Else
            plen = plen - 1
            path = Left$(path, plen)
        End If
    Wend

    On Error GoTo PathError
    ctl.ViewImagePath = path
    Exit Sub

PathError:
    MSG = "Invalid image path: " & path
    MsgBox MSG, 48, "Viewer"
    ctl.ViewImagePath = ""
    Exit Sub

End Sub

