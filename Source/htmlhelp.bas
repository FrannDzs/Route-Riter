Attribute VB_Name = "htmlhelp"
Option Explicit

'Public gmfv As D3DVECTOR
Public gflip As Boolean

Const HH_DISPLAY_TOPIC = &H0

'Const HH_HELP_CONTEXT = &HF         ' Display mapped numeric value in
                                    ' dwData.

'''''''''''''''''''''''''''''''''''
Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
         (ByVal hwndCaller As Long, ByVal pszFile As String, _
         ByVal uCommand As Long, ByVal dwData As Long) As Long

' HTML Help file launched in response to a button click:
Public Sub HH_DISPLAY_Click(hwnd As Long)
'hWnd is a Long defined elsewhere to be the window handle
'that will be the parent to the help window.
   Dim hwndHelp As Long
   'The return value is the window handle of the created help window.
   hwndHelp = HtmlHelp(hwnd, App.Path & "\" & App.HelpFile, HH_DISPLAY_TOPIC, 0)
End Sub


