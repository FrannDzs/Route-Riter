Attribute VB_Name = "MWGFX"


'basic conversion and processing functions
Public Declare Function anytobmps Lib "mwgfxvb.dll" Alias "anytobmpsVB" (ByVal s As String, ByVal d As String, p As Pic, ByVal a As Long, ByVal b As Long) As Long
Public Declare Function anytogrey Lib "mwgfxvb.dll" Alias "anytogreyVB" (ByVal s As String, ByVal d As String, p As Pic, ByVal a As Long, ByVal b As Long) As Long
Public Declare Function anyto256 Lib "mwgfxvb.dll" Alias "anyto256VB" (ByVal s As String, ByVal d As String, p As Pic, ByVal a As Long, ByVal b As Long) As Long
Public Declare Function Pic_Convert Lib "mwgfxvb.dll" Alias "Pic_ConvertVB" (ByVal s As String, ByVal d As String, p As Pic, ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long) As Long
Public Declare Function bmptoanys Lib "mwgfxvb.dll" Alias "bmptoanysVB" (ByVal s As String, ByVal d As String, p As Pic, ByVal a As Long, ByVal b As Long) As Long
Public Declare Function checkbmp Lib "mwgfxvb.dll" Alias "checkbmpVB" (ByVal s As String, p As Pic) As Long
Public Declare Function bmprocess Lib "mwgfxvb.dll" Alias "bmprocessVB" (ByVal s As String, ByVal d As String, p As Pic, ByVal a As Long) As Long
'mwgfx24.dll functions (accessed through mwgfxvb.dll)
Public Declare Function WinImagePrint Lib "mwgfxvb.dll" Alias "WinImagePrintVB" (ByVal s As String) As Long
Public Declare Function WinImageBrowse Lib "mwgfxvb.dll" Alias "WinImageBrowseVB" (ByVal s As String) As Long
Public Declare Function WinImageCopy Lib "mwgfxvb.dll" Alias "WinImageCopyVB" (ByVal s As String) As Long
Public Declare Function WinImageSize Lib "mwgfxvb.dll" Alias "WinImageSizeVB" (ByVal s As String) As Long
Public Declare Function WinImageCrop Lib "mwgfxvb.dll" Alias "WinImageCropVB" (ByVal s As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Long
Public Declare Function WinImageAdjust Lib "mwgfxvb.dll" Alias "WinImageAdjustVB" (ByVal s As String) As Long
Public Declare Function WinImageShow Lib "mwgfxvb.dll" Alias "WinImageShowVB" (ByVal s As String, ByVal c As Long) As Long
Public Declare Function WinSlideShow Lib "mwgfxvb.dll" Alias "WinSlideShowVB" (ByVal s As String, ByVal secs As Long, ByVal locol As Long, ByVal sloop As Long, ByVal validext As Long, ByVal sbeep As Long) As Long










