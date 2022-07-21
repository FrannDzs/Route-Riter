VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTokens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarTokName As String 'local copy
Private mvarTokNumber As Long 'local copy
Private mvarTokEmbed As Boolean 'local copy
Private mvarTokCount As Long 'local copy
Private mvarTokPrecis As Long 'local copy
Private mvarTokType As Long 'local copy
'local variable(s) to hold property value(s)

Public Property Let TokType(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TokType = 5
    mvarTokType = vData
End Property


Public Property Get TokType() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TokType
    TokType = mvarTokType
End Property



Public Property Let TokPrecis(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TokPrecis = 5
    mvarTokPrecis = vData
End Property


Public Property Get TokPrecis() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TokPrecis
    TokPrecis = mvarTokPrecis
End Property



Public Property Let TokCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TokCount = 5
    mvarTokCount = vData
End Property


Public Property Get TokCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TokCount
    TokCount = mvarTokCount
End Property



Public Property Let TokEmbed(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TokEmbed = 5
    mvarTokEmbed = vData
End Property


Public Property Get TokEmbed() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TokEmbed
    TokEmbed = mvarTokEmbed
End Property



Public Property Let TokNumber(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TokNumber = 5
    mvarTokNumber = vData
End Property


Public Property Get TokNumber() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TokNumber
    TokNumber = mvarTokNumber
End Property



Public Property Let TokName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TokName = 5
    mvarTokName = vData
End Property


Public Property Get TokName() As String
Attribute TokName.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TokName
    TokName = mvarTokName
End Property



