Attribute VB_Name = "mdINIBas"
Option Explicit
Option Compare Text

' *ONLY* API declaration, returns the default windows path.
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' Characters to use in reading/writing sections.
Const T1 = "["
Const T2 = "]"

' Character to use in reading/writing keys.
Const KSEP = "="

' Protects format of multi-line strings.
Const NL_TAG = "/n"
Const NL_HEAD = "/"
Const NL_PROTECTED_HEAD = "/~"

' See ProtectStr function for more information.
' Set this constant to true to enable line protection.
Const PROTECT_STR = False

' These characters are invalid in an *.INI key or section.
Const INVALID_CHARS = T1 & T2 & KSEP

' An array to store the lines loaded from the INI file.
Dim FileContents() As String


' All featured proceedures follow this line.
' ------------------------------------------------------------------------

' Loads a section from an INI file
Public Function GetPrivateProfileSection(Section As String, Default As String, INIFile As String) As String
Dim LoadStr As String, BeginLine As Long, EndLine As Long, i As Long
    On Error GoTo LoadERR
    ' Is it a valid section?
    If Not ValidStr(Section) Then GoTo LoadERR
    ' Load the INI file into the FileContents database.
    OpenINIFile INIFile
    ' If the section doesn't exist, then screw it.
    If Not FindLine(BeginLine, T1 & Section & T2) Then GoTo LoadERR
    EndLine = BeginLine
    ' If there is no other section, select all data after the section
    ' until the end of the file.
    If Not FindLine(EndLine, T1) Then EndLine = GetLineCnt
    ' Gather all the data in the section.
    For i = BeginLine + 1 To EndLine - 1
        LoadStr = LoadStr & FileContents(i)
        If Not i = EndLine - 1 Then LoadStr = LoadStr & vbCrLf
    Next i
    ' Interpret decryptable one-line string value.
    ProtectStr LoadStr, False
    ' Return the data.
    GetPrivateProfileSection = LoadStr
    Exit Function
LoadERR:
    GetPrivateProfileSection = Default
End Function

' Returns a string inside an INI file.
Public Function GetPrivateProfileString(Section As String, Key As String, Default As String, INIFile As String) As String
Dim LoadStr As String, CurLine As Long, CharPos As Long
    On Error GoTo LoadERR
    ' Is it a valid key?
    If Not ValidStr(Key) And ValidStr(Section) Then GoTo LoadERR
    ' Load the contents of the INI file into the FileContents database.
    OpenINIFile INIFile
    ' If the specified section doesn't exist, screw it.
    If Not FindLine(CurLine, T1 & Section & T2) Then GoTo LoadERR
    ' If the specified key doesn't exist, screw it.
    If Not FindLine(CurLine, Key & KSEP, True) Then GoTo LoadERR
    ' Look for the equal sign in the key.
    CharPos = InStr(1, FileContents(CurLine), KSEP)
    ' Get all data after the equal sign.
    LoadStr = Mid$(FileContents(CurLine), CharPos + 1)
    ' Interpret decryptable one-line string value.
    ProtectStr LoadStr, False
    ' Return the loaded string.
    GetPrivateProfileString = LoadStr
    Exit Function
LoadERR:
    GetPrivateProfileString = Default
End Function

Public Sub GetPrivateProfileSections(ByRef StrArr() As String, INIFile As String)
Dim i As Integer, CurPos As Long, CharPos As Integer
    ' Load the contents of the INI file into the FileContents database.
    OpenINIFile INIFile
    Do
        ' If there are no remaining sections, then exit the proceedure.
        If Not FindLine(CurPos, T1) Then Exit Do
        i = i + 1
        ReDim Preserve StrArr(1 To i)
        CharPos = InStr(1, FileContents(CurPos), T2)
        StrArr(i) = Mid$(FileContents(CurPos), 2, CharPos - 2)
    Loop
End Sub

Public Sub GetPrivateProfileKeys(ByRef StrArr() As String, Section As String, INIFile As String)
Dim i As Integer, CurPos As Long, CharPos As Integer
    ' Is it a valid section?
    If Not ValidStr(Section) Then GoTo FindERR
    ' Load the contents of the INI file into the FileContents database.
    OpenINIFile INIFile
    ' Find the specified section.
    If Not FindLine(CurPos, T1 & Section & T2) Then Exit Sub
    ' Loop through the following lines.
    Do
        CurPos = CurPos + 1
        If CurPos > GetLineCnt Then Exit Sub
        ' If it is a section format, then our search is over.
        If Left$(FileContents(CurPos), 1) = T1 Then Exit Do
        ' If it is a key format, then store it to the database.
        If FileContents(CurPos) Like "*" & KSEP & "*" Then
            i = i + 1
            CharPos = InStr(1, FileContents(CurPos), KSEP)
            ReDim Preserve StrArr(1 To i)
            StrArr(i) = Left$(FileContents(CurPos), CharPos - 1)
        End If
    Loop
FindERR:
End Sub


Public Sub RemovePrivateProfileSection(Section As String, INIFile As String)
Dim BeginPos As Long, EndPos As Long
    On Error GoTo DeleteERR
    ' Is it a valid section?
    If Not ValidStr(Section) Then GoTo DeleteERR
    ' Load the INI file into the FileContents array.
    OpenINIFile INIFile
    ' If the section doesn't exist, don't bother...
    If Not FindLine(BeginPos, T1 & Section & T2) Then Exit Sub
    EndPos = BeginPos
    ' Find the beginning of the next section.
    FindLine EndPos, T1, True
    ' Remove everything in the section.
    RemoveLines BeginPos, EndPos - 1
    ' Save the file to disk.
    SaveINIFile INIFile
    
DeleteERR:
End Sub

Public Sub RemovePrivateProfileString(Section As String, Key As String, INIFile As String)
Dim CurPos As Long
    On Error GoTo DeleteERR
    ' Is it a valid section? ...valid key?
    If Not ValidStr(Section) And ValidStr(Key) Then GoTo DeleteERR
    ' Load the INI file into the FileContents array.
    OpenINIFile INIFile
    ' If the specified section doesn't exist, screw it.
    If Not FindLine(CurPos, T1 & Section & T2) Then Exit Sub
    ' If the specified key doesn't exist, screw it.
    If Not FindLine(CurPos, Key & KSEP, True) Then Exit Sub
    ' Otherwise, delete the key.
    RemoveLines CurPos, CurPos
    ' Then save the INI file.
    SaveINIFile INIFile
DeleteERR:
End Sub

' Writes a string to an INI file.
Public Sub WritePrivateProfileSection(Section As String, Value As String, INIFile As String)
Dim BeginLine As Long, EndLine As Long
    On Error GoTo WriteERR
    ' Is it a valid Section?
    If Not ValidStr(Section) Then GoTo WriteERR
    ' Load the INIFile to the FileContents array.
    OpenINIFile INIFile
    ' Make Value decryptable one-line string.
    ProtectStr Value
    ' If the specified section doesn't exist, then...
    If Not FindLine(BeginLine, T1 & Section & T2) Then
        ' If we're not at the beginning of the file, then add
        ' a new line.
        If FileContents(BeginLine) > "" Then AddLine ""
        ' Add the section.
        AddLine T1 & Section & T2
        ' Add the key. Even if the key is multi-lined, it still
        ' works in the FileContents array.
        AddLine Value
    Else
        EndLine = BeginLine
        ' Search for the beginning of the next section.
        FindLine EndLine, T1, True
        ' Remove all data between the current and next sections.
        RemoveLines BeginLine + 1, EndLine - 1
        ' Construct the new section.
        AddLine Value, BeginLine + 1
        AddLine "", BeginLine + 2
    End If
    ' Save the new file to disk.
    SaveINIFile INIFile
WriteERR:
End Sub

' Writes a string to the INI file.
Public Sub WritePrivateProfileString(Section As String, Key As String, Value As String, INIFile As String)
Dim LastLine As Long, EndLine As Long, NewKey As String
    On Error GoTo WriteERR
    ' Is it a valid section? ...valid key?
    If Not ValidStr(Section) And ValidStr(Key) Then GoTo WriteERR
    ' Load the INI file into the FileContents array.
    OpenINIFile INIFile
    ' Make Value decryptable one-line string.
    ProtectStr Value
    ' The key to search for/create
    NewKey = Key & KSEP
    ' If the specified section doesn't exist, then...
    If Not FindLine(LastLine, T1 & Section & T2) Then
        ' If we're not at the beginning of the file, add a new line.
        If FileContents(LastLine) > "" Then AddLine ""
        ' Create the specified section at the end of the file.
        AddLine T1 & Section & T2
        ' Create the new key inside the new section.
        AddLine NewKey & Value
    Else
        EndLine = LastLine
        ' If the specified key doesn't exist, then...
        If Not FindLine(EndLine, NewKey, True) Then
            ' Create the specified key after the section header.
            AddLine NewKey & Value, LastLine + 1
        Else
            ' Replace the value of the old key with the new value.
            FileContents(LastLine + 1) = NewKey & Value
        End If
    End If
    ' Save the file to disk
    SaveINIFile INIFile
WriteERR:
End Sub

' ------------------------------------------------------------------------
' All featured procedures proceed this line.

' Finds and replaces a given string inside another string.
Function ReplaceChars(Chars As String, Optional ReplaceChr As String, Optional ReplaceWith As String) As String
Dim ChrCnt As Long
    If ReplaceChr = vbNullString Then ReplaceChr = " "
    ChrCnt = 1
    Do
        ChrCnt = InStr(ChrCnt, Chars, ReplaceChr)
        If ChrCnt = 0 Then Exit Do
        Chars = Left$(Chars, ChrCnt - 1) & ReplaceWith & Right$(Chars, Len(Chars) + 1 - Len(ReplaceChr) - ChrCnt)
        ChrCnt = ChrCnt + Len(ReplaceWith)
    Loop
    ReplaceChars = Chars
End Function

' This function is optional. If you set the PROTECT_STR constant to true,
' multi-line values stored in the INI file are combined to one line
' to avoid writing unreadable strings. For instance, if you called:

' WritePrivateProfileSection "Hello", "This is a test" & vbCrLf & "[Does it work?]", "FSoft.ini"

' You would return "This is a test" if you ran the GetPrivateProfileSection
' function to return the string. This is true in the traditional API
' commands also. If PROTECT_STR is true, this problem won't happen,
' although most INI files then are unreadable by the functions in this
' module. This is a bug-free function, and works in all situations.
Sub ProtectStr(ByRef StrVal As String, Optional Protect As Boolean = True, Optional Section As Boolean)
    If Not PROTECT_STR Then Exit Sub
    If Protect Then
        ReplaceChars StrVal, NL_HEAD, NL_PROTECTED_HEAD
        ReplaceChars StrVal, vbCrLf, NL_TAG
    Else
        ReplaceChars StrVal, NL_TAG, vbCrLf
        ReplaceChars StrVal, NL_PROTECTED_HEAD, NL_HEAD
    End If
End Sub

' Return the windows directory with a "\" at the end.
Function GetWindowsDir() As String
Dim RetStr As String, RetLen As Long
    RetStr = Space$(1024)
    RetLen = GetWindowsDirectory(RetStr, Len(RetStr))
    RetStr = Left$(RetStr, RetLen)
    ' Add a "\" at the end of the windows directory if not present.
    If Not Right$(RetStr, 1) = "\" Then RetStr = RetStr & "\"
    GetWindowsDir = RetStr
End Function

' Gets the number of lines total in the INI file.
Function GetLineCnt()
    On Error Resume Next
    GetLineCnt = UBound(FileContents)
End Function

' Searches through the INI file line database for a specified string.
Function FindLine(ByRef CurLine As Long, LineStr As String, Optional FindKey As Boolean) As Boolean
Dim LastCodeLine As Long
    ' Loop through the database starting with CurLine.
    For CurLine = CurLine + 1 To GetLineCnt
        ' If it isn't a blank line, keep it!
        If FileContents(CurLine) > "" Then LastCodeLine = CurLine
        ' If we're looking for a key and hit a section, we know the
        ' key doesn't exist.
        If FindKey And Left$(FileContents(CurLine), Len(T1)) = T1 Then Exit For
        ' If there is a match, then we found what we're looking for.
        If Left$(FileContents(CurLine), Len(LineStr)) = LineStr Then
            FindLine = True
            Exit For
        End If
    Next CurLine
    ' Make sure the last line has code on it.
    If CurLine > LastCodeLine Then ReDim Preserve FileContents(1 To GetLineCnt + 1)
    CurLine = LastCodeLine
End Function

' Checks keys and sections for invalid characters.
Function ValidStr(StrVal As String) As Boolean
Dim i As Integer
    ' Assume true.
    ValidStr = True
    ' Loop through each invalid character.
    For i = 1 To Len(INVALID_CHARS)
        ' If StrVal contains that character, it is a bad string.
        If InStr(1, StrVal, Mid$(INVALID_CHARS, i, 1)) > 0 Then
            ValidStr = False
            Exit For
        End If
    Next i
End Function

' Inserts a new line into the INI file line database.
Sub AddLine(NewStr As String, Optional Pos As Long)
Dim i As Long
    ' Add a line at the end of the database.
    ReDim Preserve FileContents(1 To GetLineCnt + 1)
    ' If Pos = 0, we will insert the new string just added to the
    ' database.
    If Pos = 0 Then Pos = GetLineCnt
    ' Loop from the end of the database to where the new string
    ' is supposed to be coppied.
    For i = GetLineCnt To Pos + 1 Step -1
        ' Shift the previous line up one.
        FileContents(i) = FileContents(i - 1)
    Next i
    ' Add the new string to the open line.
    FileContents(Pos) = NewStr
End Sub

' Removes a stack of lines from the INI file line database.
Sub RemoveLines(BeginStack As Long, Optional EndStack As Long)
Dim i As Integer, j As Integer
    ' Default is to delete only one line.
    If EndStack = 0 Then EndStack = BeginStack
    ' Loop from the line after the last-to-remove lines to the
    ' end of the database.
    For i = EndStack + 1 To GetLineCnt
        ' Replace the lines to be removed with the selected line.
        FileContents(BeginStack + j) = FileContents(i)
        j = j + 1
    Next i
    ' Remove the now-unused lines.
    ReDim Preserve FileContents(1 To GetLineCnt - (EndStack - BeginStack + 1))
End Sub



' Store the FileContents array to file.
Sub SaveINIFile(FileName As String)
Dim i As Integer
    ' Set "C:\Windows\" as the default directory.
    ChDir GetWindowsDir
    Open FileName For Output As #1
    ' Loop through each line
    For i = 1 To GetLineCnt
        ' Add the current line to the file.
        Print #1, Trim$(FileContents(i))
    Next i
    Close
End Sub

' Load contents from file into FileContents array.
Sub OpenINIFile(FileName As String)
Dim i As Integer, NewLine As String
    ' Delete any data already stored in FileContents array.
    Erase FileContents
    ' Set "C:\Window\" as the default directory.
    ChDir GetWindowsDir
    On Error GoTo FileERR
    Open FileName For Input As #1
    ' Loop to the end of the file.
    Do While Not EOF(1)
        i = i + 1
        ' Retrieve the current line of data into the array.
        Line Input #1, NewLine
        ' Avoid blank lines.
        If NewLine = vbNullString Then
            i = i - 1
        Else
            If Left$(NewLine, 1) = T1 Then
                i = i + 1
            End If
            ReDim Preserve FileContents(1 To i)
            FileContents(i) = NewLine
        End If
    Loop
    Close
FileERR:
End Sub
