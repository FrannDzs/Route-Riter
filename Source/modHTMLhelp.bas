Attribute VB_Name = "modHTMLhelp"

' Visual Basic code for implementing HTML Help 1.1
Declare Function HtmlHelp Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As Long) As Long

Declare Function htmlHelpTopic Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As String) As Long

Private Const HH_DISPLAY_TOC = &H1            ' WinHelp equivalent

Public Function SetHTMLHelpStrings(ByVal intSelHelpFile As Integer) As String

  ' Set the string variable to
  ' include the application path
  Select Case intSelHelpFile
  Case 1
    SetHTMLHelpStrings = App.path & _
       "\route_riter.chm"
  'Case 2
  '  SetHTMLHelpStrings = App.Path & _
  '     "\HelpTutorial.chm"
  End Select

End Function

'
'End Sub
'
'
Public Sub HTMLHelpContents(ByVal intHelpFile As Integer, _
    strWindow As String)
'
'  ' Force the Help window to display
'  ' the Contents file (*.hhc) in the left pane
  HtmlHelp hwnd, SetHTMLHelpStrings(intHelpFile) _
         & ">" & strWindow, HH_DISPLAY_TOC, 0
'
End Sub
'
'
'
'
