Attribute VB_Name = "Module3"

Public Function CompStrings(strOne As String, strTwo As String, booSame As Boolean)
If strOne = strTwo Then
booSame = True
ElseIf strOne <> strTwo Then
booSame = False
End If
End Function

Public Function FixCon(MyString As String, strCon As String, booFailed As Boolean, strThisPath As String)
Dim x As Long, strStart As String, strEnd As String, strnew3 As String
Dim Engname As String, Engpath As String, booEntry As Boolean, zz As Long, Z As Long
Dim strEngname As String, strEngpath As String, Wagname As String, Wagonpath As String
Dim strWagName As String, strWagPath As String, oldWagName As String, NewWagName As String
Dim OldEngName As String, NewEngName As String, strNewEng As String, j As Long, strNewWag As String

On Error GoTo Errtrap
If booFixAct = True Then GoTo ActStart

x = InStr(1, MyString, strCon, vbTextCompare)
If x = 0 Then
strReport = strReport & strCon & " does not have a valid TrainCfg entry (should be same as filename less .con)" & vbCrLf
booFailed = True
Exit Function
End If

strStart = Left$(MyString, x - 1)
strEnd = Mid$(MyString, x + Len(strCon))

MyString = strStart & strCon & strEnd
Rem ************* Check EngineData
ActStart:
x = 1
Do

x = InStr(x, MyString, "EngineData")
If x = 0 Then Exit Do
Z = InStr(x, MyString, vbCr)
zz = InStr(x, MyString, vbLf)
If zz < Z Then Z = zz
strnew3 = Mid$(MyString, x, Z - x)

Call CheckEngineData(strnew3, Engname, Engpath, booEntry)
Engname = Engname & ".eng"
strEngname = vbNullString
strEngpath = vbNullString

Call CheckNamePath(Engname, Engpath, strEngname, strEngpath)
strStart = Left$(MyString, x + 10)
strEnd = Mid$(MyString, Z)
If strEngpath = vbNullString Then

strReport = strReport & "Folder " & Engpath & " was missing from your Trainset (Referred by " & strThisPath & ")" & vbCrLf & vbCrLf
GoTo GetNext2
End If

If strEngname = vbNullString Then

strReport = strReport & Engpath & "\" & Engname & " are missing from your trainset (Referred by " & strThisPath & ")" & vbCrLf & vbCrLf
GoTo GetNext
End If
OldEngName = Left$(Engname, Len(Engname) - 4)
NewEngName = Left$(strEngname, Len(strEngname) - 4)
j = InStr(NewEngName, " ")
If j > 0 Then
NewEngName = ChrW$(34) & NewEngName & ChrW$(34)
End If
j = InStr(NewEngName, "(")
If j > 0 Then
NewEngName = ChrW$(34) & NewEngName & ChrW$(34)
End If
j = InStr(strEngpath, " ")
If j > 0 Then
strEngpath = ChrW$(34) & strEngpath & ChrW$(34)
End If
j = InStr(strEngpath, "(")
If j > 0 Then
strEngpath = ChrW$(34) & strEngpath & ChrW$(34)
End If
strEngpath = Replace(strEngpath, ChrW$(34) & ChrW$(34), ChrW$(34))
strNewEng = "( " & NewEngName & " " & strEngpath & " )"
MyString = strStart & strNewEng & strEnd

GetNext:
x = Z
Loop

DoEvents
x = 1
Do
x = InStr(x, MyString, "WagonData")

If x = 0 Then Exit Do
Z = InStr(x, MyString, vbCr)
zz = InStr(x, MyString, vbLf)
If zz < Z Then Z = zz
strnew3 = Mid$(MyString, x, Z - x)


Wagname = vbNullString
Wagonpath = vbNullString
Call CheckWagonData(strnew3, Wagname, Wagonpath, booEntry)
Wagname = Wagname & ".wag"

strWagName = vbNullString
strWagPath = vbNullString
Call CheckWagPath(Wagname, Wagonpath, strWagName, strWagPath)

strStart = Left$(MyString, x + 9)
strEnd = Mid$(MyString, Z)
If strWagPath = vbNullString Then

strReport = strReport & "Folder " & Wagonpath & " was missing from your Trainset (Referred by " & strThisPath & ")" & vbCrLf & vbCrLf
GoTo GetNext2
End If

If strWagName = vbNullString Then

strReport = strReport & Wagonpath & "\" & Wagname & " was missing from your Trainset (Referred by " & strThisPath & ")" & vbCrLf & vbCrLf
GoTo GetNext2
End If

oldWagName = Left$(Wagname, Len(Wagname) - 4)
NewWagName = Left$(strWagName, Len(strWagName) - 4)

j = InStr(NewWagName, " ")
If j > 0 Then
NewWagName = ChrW$(34) & NewWagName & ChrW$(34)
End If
j = InStr(NewWagName, "(")
If j > 0 Then
NewWagName = ChrW$(34) & NewWagName & ChrW$(34)
End If
j = InStr(strWagPath, " ")
If j > 0 Then
strWagPath = ChrW$(34) & strWagPath & ChrW$(34)
End If
j = InStr(strWagPath, "(")
If j > 0 Then
strWagPath = ChrW$(34) & strWagPath & ChrW$(34)
End If
strWagPath = Replace(strWagPath, ChrW$(34) & ChrW$(34), ChrW$(34))
strNewWag = "( " & NewWagName & " " & strWagPath & " )"
MyString = strStart & strNewWag & strEnd


GetNext2:
DoEvents
x = Z
Loop

Exit Function
Errtrap:

Call MsgBox("An error occurred processing " & strCon, vbExclamation, App.Title)
booFailed = True

End Function

Public Sub FixSrv(strOld As String, strService As String, booSame As Boolean)

If strOld <> strService Then
booSame = False
Else
booSame = True
End If
End Sub


