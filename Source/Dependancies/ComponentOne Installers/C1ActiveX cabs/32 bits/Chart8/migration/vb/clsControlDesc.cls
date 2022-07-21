VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private libName As String
Private libGuid As String
Private verMajor As Integer
Private verMinor As Integer
Private ocxName As String
Private description As String
Private groupID As Integer

Private idString As String
Private mcString As String
Private pcString As String

Private isInProject As Boolean

Private Sub Class_Initialize()
  groupID = 0
  libName = ""
  libGuid = ""
  verMajor = 0
  verMinor = 0
  ocxName = ""
  description = ""
  idString = ""
  mcString = ""
  pcString = ""
  InProject = False
End Sub

Public Sub SetInfo(gid As Integer, tlibName As String, _
    tlibGuid As String, vMajor As Integer, vMinor As Integer, _
    ctrlName As String, describe As String)
  groupID = gid
  libName = tlibName
  libGuid = tlibGuid
  verMajor = vMajor
  verMinor = vMinor
  ocxName = ctrlName
  description = describe
  idString = ""
  mcString = ""
  pcString = ""
  InProject = False
End Sub

Public Property Get Group() As Integer
  Group = groupID
End Property

Public Property Get TypeLibName() As String
  TypeLibName = libName
End Property

Public Property Get TypeLibGuid() As String
  TypeLibGuid = libGuid
End Property

Public Property Get VersionMajor() As Integer
  VersionMajor = verMajor
End Property

Public Property Get VersionMinor() As Integer
  VersionMinor = verMinor
End Property

Public Property Get ControlFilename() As String
  ControlFilename = ocxName
End Property

Public Property Get ControlDescription() As String
  ControlDescription = description
End Property

Public Property Get ControlString() As String
  ' Combines other properties of this control to form a string
  ' for building references
  If idString = "" Then
    idString = "{" + TypeLibGuid + "}#" + CStr(VersionMajor) + "." + CStr(VersionMinor) + "#0"
  End If
  ControlString = idString
End Property

Public Property Get ModuleControlString() As String
  ' Generates references to the control object as they appear in a coded module,
  ' including Forms, Modules and Classes
  If mcString = "" Then
    mcString = "Object = """ + ControlString + """; """ + ControlFilename + """"
  End If
  ModuleControlString = mcString
End Property

Public Property Get ProjectControlString() As String
  ' Generates references to the control object as they appear in a project file.
  If pcString = "" Then
    pcString = "Object=" + ControlString + "; " + ControlFilename
  End If
  ProjectControlString = pcString
End Property

Public Property Get InProject() As Boolean
  InProject = isInProject
End Property

Public Property Let InProject(newValue As Boolean)
  isInProject = newValue
End Property

Private Function convertLine(bufLine, oStr As String, tStr As String) As Boolean
  convertLine = False
  If oStr = tStr Then Exit Function
  If Len(oStr) = 0 Then Exit Function

  Dim i As Integer
  i = InStr(1, bufLine, oStr, vbTextCompare)
  If i = 0 Then Exit Function
  
  While i > 0
    bufLine = Left(bufLine, i - 1) + tStr + Mid(bufLine, i + Len(oStr))
    i = i + Len(tStr)
    i = InStr(i, bufLine, oStr, vbTextCompare)
  Wend
  
  convertLine = True
End Function

Public Function ConvertProjectControlString(bufLine As String, other As clsControlDesc) As Boolean
   ConvertProjectControlString = convertLine(bufLine, other.ProjectControlString(), ProjectControlString())
End Function

Public Function ConvertFromOther(bufLine As String, other As clsControlDesc) As Boolean
  ConvertFromOther = False
  If InStr(1, bufLine, other.ModuleControlString(), vbTextCompare) = 1 Then
    ConvertFromOther = True
    bufLine = ModuleControlString()
  Else
    If convertLine(bufLine, other.TypeLibName, TypeLibName) Then ConvertFromOther = True
    If convertLine(bufLine, other.TypeLibGuid, TypeLibGuid) Then ConvertFromOther = True
    If convertLine(bufLine, other.ControlFilename, ControlFilename) Then ConvertFromOther = True
  End If
End Function

Public Function MarkIfInProject(bufLine As String) As Boolean
  Dim pcs As String, nPos As Long
  pcs = ProjectControlString()

  nPos = InStr(1, bufLine, pcs, vbTextCompare)
  
  If nPos > 0 Then
    InProject = True
  Else
    InProject = False
  End If
  MarkIfInProject = InProject
End Function
