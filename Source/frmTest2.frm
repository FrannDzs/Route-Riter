VERSION 5.00
Begin VB.Form frmTest2 
   Caption         =   "Test2"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Textures2 
      Caption         =   "Shape Textures"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Writing 
      Caption         =   "Writing"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Textures 
      Caption         =   "Textures"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Out 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   6735
   End
   Begin VB.CommandButton Test 
      Caption         =   "Test"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   1815
   End
End
Attribute VB_Name = "frmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim TilePath As String
Dim HazShp() As Variant
Dim HazShp2() As Variant
Dim strShp() As Variant
Dim strShapes() As Variant
Dim numShp As Long
Dim ForTex() As Variant
Dim ForTex2() As Variant
Dim Ace1() As Variant
Dim Ace2() As Variant
Dim numAce As Long
Dim numFor As Integer
Dim numHaz As Integer
Const Shp_Chunk = 5000
Const For_Chunk = 500
Dim booCoal As Boolean, booWat As Boolean, booDies As Boolean, booCross As Boolean, booSig As Boolean



Public Function RemD2(ByRef rArray() As Variant, xArray() As Variant) As Variant

    'Declare variables
    Dim ii As Long, jj As Long
    
    'Initialize variables
    count3 = 1
    high = UBound(rArray)
    'Declare temp array
    
    ReDim xArray(1 To high)
    
    'Start duplicates removal code

xArray(1) = rArray(1)
jj = 2
    For ii = 2 To high
        If rArray(ii) <> rArray(ii - 1) Then
        xArray(jj) = rArray(ii)
        jj = jj + 1
End If
Next ii
If rArray(high) <> rArray(high - 1) Then
xArray(jj) = rArray(high)
jj = jj + 1
End If
        

  
ReDim Preserve xArray(1 To jj - 1)

End Function

 Private Sub ReadWorld(fullpath$)
   Dim tw As value, fname As value, j As Integer, jj As Integer
   
    Dim tfh As New TokenFileHandler
    Dim tf As TokenFile
    Set tf = tfh.readFile(fullpath$)
   
    Dim i As Integer
         Dim tv As value
        Set tv = tf.getbyindex(0)   'was token
         ' Stop
        If tv.Count > 0 Then
            For j = 0 To tv.Count - 1
                'Dim tw As value
                Set tw = tv.getbyindex(j)
            Select Case tw.Name
               Case "Static"
                Set fname = tw.getbyindex(1)
                'List1.AddItem fname.tostring()
                If Right(fname.tostring(), 2) <> ".s" Then Stop
                strShp(numShp) = fname.tostring()
                numShp = numShp + 1
                    If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(1 To numShp + Shp_Chunk)
                    End If
             ' End If
                DoEvents
                Case "TrackObj"
                For jj = 1 To tw.Count - 1
                Set fname = tw.getbyindex(jj)
                If Right(fname.tostring(), 2) = ".s" Then
                strShp(numShp) = fname.tostring()
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(1 To numShp + Shp_Chunk)
                    End If
                DoEvents
                Exit For
                End If
                Next jj
                Case "Hazard"
                Set fname = tw.getbyindex(2)
                HazShp(numHaz) = fname.tostring()
                numHaz = numHaz + 1
                If numHaz > UBound(HazShp) Then
                    ReDim Preserve HazShp(1 To numHaz + For_Chunk)
                    End If
               
                DoEvents
                Case "Pickup"
                
                For jj = 1 To tw.Count - 1
                Set fname = tw.getbyindex(jj)
                If jj = 2 Then
               ' Stop
                    Select Case fname
                    Case "5"
                    booCoal = True
                    Case "6"
                    booWat = True
                    Case "7"
                    booDies = True
                    End Select
                    End If
                If Right(fname.tostring(), 2) = ".s" Then
                
                strShp(numShp) = fname.tostring()
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(1 To numShp + Shp_Chunk)
                    End If
                DoEvents
                Exit For
                End If
                Next jj
                Case "Signal"
                booSig = True
                For jj = 1 To tw.Count - 1
                Set fname = tw.getbyindex(jj)
                If Right(fname.tostring(), 2) = ".s" Then
                strShp(numShp) = fname.tostring()
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(1 To numShp + Shp_Chunk)
                    End If
                DoEvents
                Exit For
                End If
                Next jj
                Case "Gantry"
               For jj = 1 To tw.Count - 1
                Set fname = tw.getbyindex(jj)
                If Right(fname.tostring(), 2) = ".s" Then
                strShp(numShp) = fname.tostring()
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(1 To numShp + Shp_Chunk)
                    End If
                DoEvents
                Exit For
                End If
                Next jj
                DoEvents
                Case "LevelCr"
                booCross = True
               For jj = 1 To tw.Count - 1
                Set fname = tw.getbyindex(jj)
                If Right(fname.tostring(), 2) = ".s" Then
                strShp(numShp) = fname.tostring()
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(1 To numShp + Shp_Chunk)
                    End If
                DoEvents
                Exit For
                End If
                Next jj
                Case "Forest"
                For jj = 1 To tw.Count - 1
                Set fname = tw.getbyindex(jj)
                If Right(fname.tostring(), 4) = ".ace" Then
                ForTex(numFor) = fname.tostring()
                numFor = numFor + 1
                If numFor > UBound(ForTex) Then
                    ReDim Preserve ForTex(1 To numFor + For_Chunk)
                    End If
                Exit For
                End If
                Next jj
                Case "SpeedPost"
                For jj = 1 To tw.Count - 1
                Set fname = tw.getbyindex(jj)
                If Right(fname.tostring(), 2) = ".s" Then
               ' List1.AddItem fname.tostring()
               strShp(numShp) = fname.tostring()
               If Right(fname.tostring(), 2) <> ".s" Then Stop
                numShp = numShp + 1
                If numShp > UBound(strShp) Then
                    ReDim Preserve strShp(0 To numShp + Shp_Chunk)
                    End If
                DoEvents
                ElseIf Right(fname.tostring(), 4) = ".ace" Then
                Ace1(numAce) = fname.tostring
                numAce = numAce + 1
                'Debug.Print fname.tostring()
                
                End If
                Next jj
            End Select
       ' End If
        Next j
        End If
       ' Stop
       
End Sub
Private Sub ReadShape(fullpath$)
   
    Dim tfh As New TokenFileHandler
    Dim tf As TokenFile
    Set tf = tfh.readFile(fullpath$)
   
    Dim i As Integer
         Dim tv As value
        Set tv = tf.getbyindex(0)   'was token
               
        If tv.Count > 0 Then
            For j = 0 To tv.Count - 1
                Dim tw As value
                Set tw = tv.getbyindex(j)
                If tw.Name = "images" Then
                
                Dim tex As value
                Dim numtex As value
                Set numtex = tw.getbyindex(0)
                
           For q = 1 To numtex
           
            Set tex = tw.getbyindex(q)
            Ace1(numAce) = tex.tostring()
                numAce = numAce + 1
                If numAce > UBound(Ace1) Then
                    ReDim Preserve Ace1(1 To numAce + Shp_Chunk)
                    End If
                DoEvents
               
             Next q
        End If
        Next j
        End If
       
       
End Sub


Private Sub Form_Load()
Dim i As Integer, j As Long
Dim booLow As Boolean
Dim GlobalPath As String
Dim strOrigFile As String
ReDim strShp(1 To Shp_Chunk)
ReDim Ace1(1 To Shp_Chunk)
ReDim ForTex(1 To For_Chunk)
ReDim HazShp(1 To For_Chunk)
Debug.Print Now
numShp = 1
numHaz = 1
numFor = 1
numAce = 1
TilePath = RoutePath & "\world"
'TilePath = RoutePath & "\shapes"
GlobalPath = MSTSPath & "\Global"
frmUtils.Dir1(0).path = TilePath
'frmUtils.Text1(0).Text = "*.s"
frmUtils.Text1(0).Text = "*.w"
 cursouind = 0

booOKAll = False

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
    frmUtils.File1(cursouind).Selected(i) = True
Next i

For i = 0 To frmUtils.File1(cursouind).ListCount - 1
frmUtils.Label12.Caption = "Reading: " & frmUtils.File1(cursouind).list(i)

   If frmUtils.File1(cursouind).Selected(i) Then
    fullpath$ = TilePath & "\" & frmUtils.File1(cursouind).list(i)
  'Call GetACES(fullpath$)
Call ReadWorld(fullpath$)

   End If
GetNext:
   Next i
  
  
   'Debug.Print Now
    
    ReDim Preserve strShp(1 To numShp - 1)
   
    QSort3 strShp(), 1, numShp - 1
    DoEvents
    RemD2 strShp(), strShapes()
    DoEvents
    For j = 1 To numShp - 1
    strShp(j) = ""
    Next j
    ReDim Preserve ForTex(1 To numFor - 1)
    QSort3 ForTex(), 1, numFor - 1
    DoEvents
    RemD2 ForTex(), ForTex2()
    DoEvents
   ' Set ForTex() = Nothing
    If numHaz > 1 Then
    ReDim Preserve HazShp(1 To numHaz - 1)
    QSort3 HazShp(), 1, numHaz - 1
    DoEvents
    RemD2 HazShp(), HazShp2()
    DoEvents
   ' Set HazShp() = Nothing
    
    End If
    numShp = UBound(strShapes)
    numFor = UBound(ForTex2)
    numHaz = UBound(HazShp2)
    For j = 1 To numShp
    'Stop
    If FileExists(MSTSPath & "\routes\" & RouteName & "\shapes\" & strShapes(j)) Then
    Call ReadShape(MSTSPath & "\routes\" & RouteName & "\shapes\" & strShapes(j))
    ElseIf FileExists(GlobalPath & "\shapes\" & strShapes(j)) Then
    Call ReadShape(GlobalPath & "\shapes\" & strShapes(j))
    Else
    'Look for shape
    Debug.Print "Missing - " & strShapes(j)
    Stop
    End If
    Next j
    'Stop
    ReDim Preserve Ace1(1 To numAce - 1)
    QSort3 Ace1(), 1, numAce - 1
    DoEvents
    RemD2 Ace1(), Ace2()
    Debug.Print Now
   Stop
frmTest2.Show
Stop
End Sub
Public Function QSort3(strList() As Variant, lLbound As Long, lUbound As Long)
    
    Dim strTemp As String
    Dim strBuffer As String
    Dim lngCurLow As Long
    Dim lngCurHigh As Long
    Dim lngCurMidpoint As Long
    
    lngCurLow = lLbound ' Start current low and high at actual low/high
    lngCurHigh = lUbound
    
    If lUbound <= lLbound Then Exit Function ' Error!
    lngCurMidpoint = (lLbound + lUbound) \ 2 ' Find the approx midpoint of the array
    
    strTemp = strList(lngCurMidpoint) ' Pick as a starting point (we are making
    ' an assumption that the data *might* be
    '
    ' in semi-sorted order already!
    


    Do While (lngCurLow <= lngCurHigh)


        Do While strList(lngCurLow) < strTemp
            lngCurLow = lngCurLow + 1
            If lngCurLow = lUbound Then Exit Do
        Loop
        


        Do While strTemp < strList(lngCurHigh)
            lngCurHigh = lngCurHigh - 1
            If lngCurHigh = lLbound Then Exit Do
        Loop


        If (lngCurLow <= lngCurHigh) Then ' if low is <= high then swap
            strBuffer = strList(lngCurLow)
            strList(lngCurLow) = strList(lngCurHigh)
            strList(lngCurHigh) = strBuffer
            '
            lngCurLow = lngCurLow + 1 ' CurLow++
            lngCurHigh = lngCurHigh - 1 ' CurLow--
        End If
        
    Loop
    


    If lLbound < lngCurHigh Then ' Recurse if necessary
        QSort3 strList(), lLbound, lngCurHigh
    End If
    


    If lngCurLow < lUbound Then ' Recurse if necessary
        QSort3 strList(), lngCurLow, lUbound
    End If
    
End Function





Private Sub Test_Click()
    'On Error Resume Next
    Dim tfh As New TokenFileHandler
    Dim tf As TokenFile
   ' Set tf = tfh.readFile("test.t")
   Set tf = tfh.readFile(fullpath$)
    'Out.Text = tf.Header + ChrW(13) + ChrW(10)
    
 '   Dim Token As value
'    'Set token = tf.GetByName("Terrain")
'    Set Token = tf.GetByName("Images")
'    Out.Text = Out.Text + Token.Name + ChrW(13) + ChrW(10)
'
    Dim i As Integer
   ' For i = 0 To Token.Count - 1
        Dim tv As value
        Set tv = tf.getbyindex(0)   'was token
       ' Out.Text = Out.Text + "   " + tv.Name + ChrW(13) + ChrW(10)
        
        If tv.Count > 0 Then
            For j = 0 To tv.Count - 1
                Dim tw As value
                Set tw = tv.getbyindex(j)
                If tw.Name = "images" Then
                
                'Debug.Print tw.Name, tw.tostring()
                'Out.Text = Out.Text + "      " + tw.Name + ": " + tw.tostring() + ChrW(13) + ChrW(10)
                Dim tex As value
                Dim numtex As value
                Set numtex = tw.getbyindex(0)
                'Debug.Print "Numtex=" & numtex
           For q = 1 To numtex
           ' Dim tex As Token
            Set tex = tw.getbyindex(q)
            
            Debug.Print tex.tostring()
            
        Next q
'            Set tex = tw.getbyname("image")
            End If
            Next j
        End If
   ' Next i
    
'    Set Token = Token.GetByName("Shaders")
'    Set Token = Token.GetByName("numShaders")
'    Out.Text = Out.Text + ChrW(13) + ChrW(10) + "shaders=" + str$(Token.ToUInt()) + ChrW(13) + ChrW(10)
End Sub

Private Sub Textures_Click()
    Dim tfh As New TokenFileHandler
    Dim tf As TokenFile
    Set tf = tfh.readFile(fullpath$)
   ' Out.Text = tf.Header + ChrW(13) + ChrW(10)
    
    Dim terrain As Token
    Set terrain = tf.getbyname("Terrain")
    Dim shaders As Token
    Set shaders = terrain.getbyname("Shaders")
   ' Out.Text = Out.Text + "NumShaders=" + shaders.GetByIndex(0).ToString() + ChrW(13) + ChrW(10)
    
    Dim i As Integer
    For i = 1 To shaders.Count - 1
        Dim shader As Token
        Set shader = shaders.getbyindex(i)
        Dim texSlots As Token
        Set texSlots = shader.getbyname("TexSlots")
        Dim j As Integer
        For j = 1 To texSlots.Count - 1
            Dim texSlot As Token
            Set texSlot = texSlots.getbyindex(j)
            Dim texture As value
            Set texture = texSlot.getbyindex(0)
            'Out.Text = Out.Text + "Shader[" + str$(i) + "], TexSlot[" + str$(j) + "]: " + texture.ToString() + ChrW(13) + ChrW(10)
            Out.Text = Out.Text + texture.tostring() + ChrW(13) + ChrW(10)
            
        Next j
    Next i
End Sub

Private Sub Textures2_Click()
  Dim tfh As New TokenFileHandler
    Dim tf As TokenFile
    Set tf = tfh.readFile(fullpath$)
   ' Out.Text = tf.Header + ChrW(13) + ChrW(10)
    Stop
    Dim Shape As Token
    Set Shape = tf.getbyname("Shape")
    Dim Images As Token
    Set Images = Shape.getbyname("Images")
   ' Out.Text = Out.Text + "NumShaders=" + Images.GetByIndex(0).ToString() + ChrW(13) + ChrW(10)
    
    Dim i As Integer
    For i = 1 To Images.Count - 1
        Dim image As Token
        Set image = Images.getbyname("Image")
'        Dim texSlots As token
'        Set texSlots = Image.GetByName("TexSlots")
        Dim j As Integer
        For j = 1 To image.Count - 1
                Dim texSlot As Token
                Set texSlot = texSlots.getbyindex(j)
            Dim texture As value
            Set texture = image.getbyindex(0)
            'Out.Text = Out.Text + "Image[" + str$(i) + "], TexSlot[" + str$(j) + "]: " + texture.ToString() + ChrW(13) + ChrW(10)
            Out.Text = Out.Text + texture.tostring() + ChrW(13) + ChrW(10)
            
        Next j
    Next i
End Sub

Private Sub Writing_Click()
    Dim tfh As New TokenFileHandler
    Dim tf As TokenFile
    Set tf = tfh.readFile("bush.s")
    
    Dim Shape As Token
    Set Shape = tf.getbyname("shape")
    Dim Images As Token
    Set Images = Shape.getbyname("images")
    Dim image As Token
    Set image = Images.getbyindex(1)
    Dim ace As value
    Set ace = image.getbyindex(0)
    Out.Text = Out.Text + "Before:" + image.tostring() + ChrW(13) + ChrW(10)
    image.Val = "test.ace"
    Out.Text = Out.Text + "After:" + image.tostring() + ChrW(13) + ChrW(10)
    tf.WriteFile ("bush2.s")
End Sub
