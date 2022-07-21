VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MovableObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private objModel As sfModel

Private sX As Single
Private sY As Single
Private sZ As Single
Private sR As Single

Private sLastError As String


Public Property Get LastError() As String
    LastError = sLastError
End Property

Public Function LoadShape(Shapefile As String) As Boolean
    Dim sFileStem As String
    Dim sDescriptorFile As String
    Dim sRollingstockFile As String
   Dim ThisShapePath As String, x As Integer
   
    Me.Model.Filename = Shapefile
    sLastError = Trim$(Me.Model.info)
    If sLastError = vbNullString Then
        LoadShape = True
        x = InStrRev(Shapefile, "\")
        ThisShapePath = Left$(Shapefile, x)
        sFileStem = Left$(Shapefile, Len(Shapefile) - 2)
        
        ' Acquire associated descriptor and wagon/engine files
        sDescriptorFile = Dir(sFileStem & ".sd")
        If sDescriptorFile <> vbNullString Then
            sDescriptorFile = sFileStem & ".sd"
        End If
        ' Grab bounding box from sd file if exists
        
        sRollingstockFile = Dir(ThisShapePath & conWagon(ThisCon))
    
       ' LoadShape = False
    End If
    If sRollingstockFile = vbNullString Then
    
    Call MsgBox(sFileStem & ".eng/.wag" & Lang(353), vbExclamation, App.Title)
    
    LoadShape = False
    Exit Function
    End If
 Call GetSize(sRollingstockFile, EngSize)

End Function

'Public Property Get Model() As sfModel
'    Set Model = objModel
'End Property

Public Property Get x() As Single
    x = sX
End Property

Public Property Get Y() As Single
    Y = sY
End Property

Public Property Get Z() As Single
    Z = sZ
End Property

Public Property Get R() As Single
    R = sR
End Property

Public Property Let x(NewX As Single)

    MoveTo NewX, Y, Z

End Property

Public Property Let Y(NewY As Single)

    MoveTo x, NewY, Z

End Property

Public Property Let Z(NewZ As Single)

    MoveTo x, Y, NewZ

End Property

Public Property Let R(NewR As Single)

    RotateTo NewR

End Property

Sub MoveBy(Optional DX As Single = 0#, Optional DY As Single = 0#, Optional DZ As Single = 0#)

    If DX = 0# And DY = 0# And DZ = 0# Then Exit Sub
    
    MoveTo sX + DX, sY + DY, sZ + DZ
    
End Sub
    
Sub MoveTo(NewX As Single, NewY As Single, NewZ As Single)

    ' Set internal record of position
    sX = NewX
    sY = NewY
    sZ = NewZ

    ApplyNewOrientation
    
End Sub

Sub MoveAndRotateTo(NewX As Single, NewY As Single, NewZ As Single, NewR As Single)

    ' Set internal record of position
    sX = NewX
    sY = NewY
    sZ = NewZ
    sR = NewR

    ApplyNewOrientation
    
End Sub

Sub RotateBy(Optional DR As Single = 0#)
    If DR = 0# Then Exit Sub

    RotateTo sR + DR
End Sub

Sub RotateTo(NewR As Single)
    
    ' May have to draw the object back to zero X/Y/Z first before rotate....

   ' Dim mw As D3DMATRIX, mt As D3DMATRIX
    
    ' Set internal record of position
    sR = NewR
    
    ApplyNewOrientation
    
End Sub

Private Sub ApplyNewOrientation()
    Dim mw As D3DMATRIX, mr As D3DMATRIX, mt As D3DMATRIX
    
    ' set up identity matrix
    D3DXMatrixIdentity mw
    
    ' create rotation matrix about the vertical axis
    D3DXMatrixRotationY mr, sR
    
    ' multiply the two...  (rotation first)
    D3DXMatrixMultiply mw, mr, mw
    
    ' create translation matrix
    D3DXMatrixTranslation mt, sX, sY, sZ

    ' multiply the two...  (translation first)
    D3DXMatrixMultiply mw, mt, mw
    
    Me.Model.setMatrix mw

End Sub

Private Sub Class_Initialize()
    Set objModel = New sfModel
End Sub

Private Sub Class_Terminate()
    Set objModel = Nothing
End Sub
