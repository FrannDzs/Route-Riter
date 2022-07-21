VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAtObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#####################################################################################
'#
'#  clsAObj - Helper for clsRsize to manage resize events on forms
'#
'#####################################################################################

Dim AttachedObj As Object
Dim bVaries As Boolean
Dim edge(4) As Boolean
Dim offset(4) As Integer
Dim vedge(4) As Integer ' -1 = False, 0 = 50%, 1 = True

Public Function SetGUI(GUIobject As Object, bleft As Boolean, bright As Boolean, btop As Boolean, bbottom As Boolean)

Set AttachedObj = GUIobject
edge(1) = bleft
edge(2) = bright
edge(3) = btop
edge(4) = bbottom

offset(1) = GUIobject.Left
offset(2) = GUIobject.width
offset(3) = GUIobject.Top
offset(4) = GUIobject.height

bVaries = False
End Function

Public Function VSetGUI(GUIobject As Object, vbleft As Integer, vbright As Integer, vbtop As Integer, vbbottom As Integer)

Set AttachedObj = GUIobject
vedge(1) = vbleft
vedge(2) = vbright
vedge(3) = vbtop
vedge(4) = vbbottom

offset(1) = GUIobject.Left
offset(2) = GUIobject.width
offset(3) = GUIobject.Top
offset(4) = GUIobject.height

bVaries = True
End Function


Public Property Get Left() As Integer
Left = edge(1)
End Property

Public Property Get Right() As Integer
Right = edge(2)
End Property

Public Property Get Top() As Integer
Top = edge(3)
End Property

Public Property Get Bottom() As Integer
Bottom = edge(4)
End Property

Public Property Get VLeft() As Integer
VLeft = vedge(1)
End Property

Public Property Get VRight() As Integer
VRight = vedge(2)
End Property

Public Property Get VTop() As Integer
VTop = vedge(3)
End Property

Public Property Get VBottom() As Integer
VBottom = vedge(4)
End Property


Public Property Get Attached() As Object
Set Attached = AttachedObj
End Property

Public Property Get OrgTop() As Integer
OrgTop = offset(3)
End Property

Public Property Get OrgHeight() As Integer
OrgHeight = offset(4)
End Property

Public Property Get OrgLeft() As Integer
OrgLeft = offset(1)
End Property

Public Property Get OrgWidth() As Integer
OrgWidth = offset(2)
End Property

Public Property Get Varies() As Boolean
Varies = bVaries
End Property

Public Property Let Varies(bNewValue As Boolean)
bVaries = bNewValue
End Property

