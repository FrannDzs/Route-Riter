Attribute VB_Name = "mwace"
Type Pic
    width As Long
    height As Long
    depth As Long
    numcols As Long
    fin As Long
    comment(80) As Byte
    cmap(768) As Byte
    jlib As Long
    tlib As Long
    plib As Long
    buff As Long
    fout As Long
    ptr As Long
    progress As Long
    ilib As Long
    stype As Long
    spare(120) As Byte
End Type
    
    
Public Declare Function AceToBmp Lib "mwacevb.dll" Alias "AceToBmpVB" (ByVal s As String, ByVal d As String) As Long
Public Declare Function AceToBmps Lib "mwacevb.dll" Alias "AceToBmpsVB" (ByVal s As String, ByVal d As String, ByVal a As String) As Long
Public Declare Function AceToTga Lib "mwacevb.dll" Alias "AceToTgaVB" (ByVal s As String, ByVal d As String) As Long
Public Declare Function AceToTgaSquare Lib "mwacevb.dll" Alias "AceToTgaSquareVB" (ByVal s As String, ByVal d As String) As Long
Public Declare Function BmpsToTga Lib "mwacevb.dll" Alias "BmpsToTgaVB" (ByVal s As String, ByVal d As String, ByVal a As String) As Long
Public Declare Function BmpsToTgaSquare Lib "mwacevb.dll" Alias "BmpsToTgaSquareVB" (ByVal s As String, ByVal d As String, ByVal a As String) As Long
Public Declare Function CheckAce Lib "mwacevb.dll" Alias "CheckAceVB" (ByVal s As String, p As Pic) As Long

