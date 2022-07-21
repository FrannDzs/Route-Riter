VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const controlTotal = 12     ' total number of controls considered
Const groupTotal = 6        ' total number of control groups considered

Const controlsPerGroup = controlTotal \ groupTotal

Private firstToGroup As Integer
Private Controls(0 To controlTotal - 1) As New clsControlDesc
Private Groups(0 To groupTotal - 1) As String
Private Title As String
Private targetGroup As Integer

''''''''''''''''''''''''
' LoadControlData() sets the Title, and information about each control
'
' When other controls are involved, only this routine and constants
' above should need to change
'
Public Sub LoadControlData()
  Title = "ComponentOne Chart Controls"

  Dim grp As Integer
  grp = 0
  Groups(grp) = "Olectra Chart 5.0 Controls"
  Call Controls(0).SetInfo(grp, "OlectraChart2D", "92D71E90-25A8-11CF-A640-9986B64D9618", 5, 0, "olch2x32.ocx", "Olectra Chart 2D, Version 5.0")
  Call Controls(1).SetInfo(grp, "OlectraChart3D", "3D931500-4601-11CF-80B2-0020AF19EE14", 5, 0, "olch3x32.ocx", "Olectra Chart 3D, Version 5.0")

  grp = 1
  Groups(grp) = "Olectra Chart 6.0 Controls"
  Call Controls(2).SetInfo(grp, "OlectraChart2D", "92D71E90-25A8-11CF-A640-9986B64D9618", 6, 0, "olch2x32.ocx", "Olectra Chart 2D, Version 6.0")
  Call Controls(3).SetInfo(grp, "OlectraChart3D", "3D931500-4601-11CF-80B2-0020AF19EE14", 6, 0, "olch3x32.ocx", "Olectra Chart 3D, Version 6.0")

  grp = 2
  Groups(grp) = "ComponentOne Chart 7.0 ANSI Controls"
  Call Controls(4).SetInfo(grp, "C1Chart2D7", "C643EB3F-235C-4181-9B55-36A268833718", 7, 0, "Olch2x7.ocx", "ComponentOne Chart 2D 7.0 (ANSI)")
  Call Controls(5).SetInfo(grp, "C1Chart3D7", "A4F5504C-4D7B-4827-87C7-7CA6D5794D06", 7, 0, "Olch3x7.ocx", "ComponentOne Chart 3D 7.0 (ANSI)")

  grp = 3
  Groups(grp) = "ComponentOne Chart 7.0 UNICODE Controls"
  Call Controls(6).SetInfo(grp, "C1Chart2D7U", "9487F13A-8164-4CB5-97BD-CFA9A776D71F", 7, 0, "Olch2xu7.ocx", "ComponentOne Chart 2D 7.0 (UNICODE)")
  Call Controls(7).SetInfo(grp, "C1Chart3D7U", "7DA9DE68-6056-4010-8A8D-B76808352C30", 7, 0, "Olch3xu7.ocx", "ComponentOne Chart 3D 7.0 (UNICODE)")

  grp = 4
  firstToGroup = grp
  Groups(grp) = "ComponentOne Chart 8.0 ANSI Controls"
  Call Controls(8).SetInfo(grp, "C1Chart2D8", "0BE3824E-5AFE-4B11-A6BC-4B3AD564982A", 8, 0, "Olch2x8.ocx", "ComponentOne Chart 2D 8.0 (ANSI)")
  Call Controls(9).SetInfo(grp, "C1Chart3D8", "AC82DA6D-CD3F-43BF-AF2E-56591B5585D8", 8, 0, "Olch3x8.ocx", "ComponentOne Chart 3D 8.0 (ANSI)")

  grp = 5
  Groups(grp) = "ComponentOne Chart 8.0 UNICODE Controls"
  Call Controls(10).SetInfo(grp, "C1Chart2D8U", "75634CE7-D088-4C44-8F7C-3C117CE5857B", 8, 0, "Olch2xu8.ocx", "ComponentOne Chart 2D 8.0 (UNICODE)")
  Call Controls(11).SetInfo(grp, "C1Chart3D8U", "5C9704A4-FE02-45EC-A1E3-7773F3CB0D5A", 8, 0, "Olch3xu8.ocx", "ComponentOne Chart 3D 8.0 (UNICODE)")

End Sub

Public Property Get ControlsTitle() As String
  ControlsTitle = Title
End Property

''''''''''''''''''''''''''''''''''''
'  Fills the migration FROM listbox
'
Public Sub FillFromListbox(listb As ListBox)
  Dim i As Integer, j As Integer, k As Integer
  listb.Clear
  k = 0
  For i = 0 To UBound(Groups)
    For j = 0 To UBound(Controls)
      If Controls(j).Group = i And Controls(j).InProject Then
        Call listb.AddItem(Groups(i), k)
        listb.ItemData(k) = i
        k = k + 1
        Exit For
      End If
    Next j
  Next i
  If listb.ListCount > 0 Then
    listb.Selected(0) = True
  End If
End Sub

''''''''''''''''''''''''''''''''''''
'  Fills the migration TO listbox
'
Public Sub FillToListbox(listb As ListBox)
  Dim i As Integer, k As Integer
  listb.Clear
  For i = firstToGroup To UBound(Groups)
    k = i - firstToGroup
    Call listb.AddItem(Groups(i), k)
    listb.ItemData(k) = i
  Next i
  listb.ListIndex = 0
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'  Clears the InProject flag for all controls that
'  are not in a group selected in the FROM listbox
'
Public Sub ClearUnselectedControls(listb As ListBox)
  Dim i As Integer, j As Integer, k As Integer
  For i = 0 To listb.ListCount - 1
    If Not listb.Selected(i) Then
      k = listb.ItemData(i)
      For j = 0 To UBound(Controls)
        If Controls(j).Group = k Then
          Controls(j).InProject = False
        End If
      Next j
    End If
  Next i
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'  Sets the InProject flag for all controls that
'  appear in an VBP Object reference in the specified
'  buffer.
'
Public Function CheckForControlUsage(buffer As String) As Boolean
  Dim i As Integer
  CheckForControlUsage = False
  For i = 0 To UBound(Controls)
    If Controls(i).MarkIfInProject(buffer) Then
      CheckForControlUsage = True
    End If
  Next i
End Function

Public Sub SetTargetControls(listb As ListBox)
  targetGroup = listb.ItemData(listb.ListIndex)
End Sub

Public Function ConvertControls(buffer As String) As Boolean
  Dim cFrom As clsControlDesc
  Dim firstTarget As Integer
  
  firstTarget = targetGroup * controlsPerGroup
  ConvertControls = False
  
  Dim grp As Integer, firstFrom As Integer, fromIndex As Integer
  
  For grp = 0 To UBound(Groups)
    firstFrom = grp * controlsPerGroup
    If firstFrom <> firstTarget Then
      For fromIndex = 0 To controlsPerGroup - 1
        Set cFrom = Controls(firstFrom + fromIndex)
        If cFrom.InProject Then
          If Controls(firstTarget + fromIndex).ConvertFromOther(buffer, cFrom) Then
            ConvertControls = True
          End If
        End If
      Next fromIndex
    End If
  Next grp
End Function

Public Function ConvertVBPControls(buffer As String) As Boolean
  Dim cFrom As clsControlDesc
  Dim firstTarget As Integer
  
  firstTarget = targetGroup * controlsPerGroup
  ConvertVBPControls = False
  
  Dim grp As Integer, firstFrom As Integer, fromIndex As Integer
  
  For grp = 0 To UBound(Groups)
    firstFrom = grp * controlsPerGroup
    If firstFrom <> firstTarget Then
      For fromIndex = 0 To controlsPerGroup - 1
        Set cFrom = Controls(firstFrom + fromIndex)
        If cFrom.InProject Then
          If Controls(firstTarget + fromIndex).ConvertProjectControlString(buffer, cFrom) Then
            ConvertVBPControls = True
          End If
        End If
      Next fromIndex
    End If
  Next grp
End Function


Private Sub Class_Initialize()
  Call LoadControlData
End Sub

