VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmProcess 
   Caption         =   "TsUtils - Process Options"
   ClientHeight    =   10800
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   ScaleHeight     =   10800
   ScaleWidth      =   13860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Complete Processing"
      Height          =   375
      Left            =   8760
      TabIndex        =   41
      Top             =   10080
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog OpenFileDialog1 
      Left            =   240
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Button5 
      Caption         =   "Display Report"
      Height          =   375
      Left            =   10560
      TabIndex        =   14
      Top             =   10080
      Width           =   1335
   End
   Begin VB.CommandButton Button3 
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   12
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Button2 
      Caption         =   "Process this Option"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   10080
      Width           =   1575
   End
   Begin VB.CommandButton Button1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   12000
      TabIndex        =   10
      Top             =   10080
      Width           =   1095
   End
   Begin VB.TextBox Textbox5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   9600
      Width           =   13215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options Specific to this Command"
      Height          =   5055
      Left            =   240
      TabIndex        =   8
      Top             =   4440
      Width           =   12855
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmProcess.frx":0000
         Left            =   10440
         List            =   "frmProcess.frx":000D
         TabIndex        =   46
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmProcess.frx":0029
         Left            =   10440
         List            =   "frmProcess.frx":0036
         TabIndex        =   45
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   5
         Left            =   10440
         TabIndex        =   44
         Top             =   3120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   4
         Left            =   10440
         TabIndex        =   43
         Top             =   2760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Button6 
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10560
         TabIndex        =   42
         Top             =   3960
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   3
         Left            =   10440
         TabIndex        =   40
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   2
         Left            =   10440
         TabIndex        =   39
         Top             =   2040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   1
         Left            =   10440
         TabIndex        =   38
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   0
         Left            =   10440
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox Checkbox11 
         Caption         =   "Check11"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   3120
         Width           =   10095
      End
      Begin VB.CheckBox Checkbox10 
         Caption         =   "Check10"
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   2760
         Width           =   10095
      End
      Begin VB.CheckBox Checkbox9 
         Caption         =   "Check9"
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   2400
         Width           =   10095
      End
      Begin VB.CheckBox Checkbox8 
         Caption         =   "Check8"
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   2040
         Width           =   10095
      End
      Begin VB.CheckBox Checkbox7 
         Caption         =   "Check7"
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   1680
         Width           =   9255
      End
      Begin VB.CheckBox Checkbox6 
         Caption         =   "Check6"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   1320
         Width           =   9135
      End
      Begin VB.TextBox Textbox8 
         Height          =   405
         Left            =   3480
         TabIndex        =   22
         Top             =   4440
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.TextBox Textbox7 
         Height          =   405
         Left            =   3480
         TabIndex        =   21
         Top             =   3960
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.TextBox Textbox6 
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   3480
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.CommandButton Button4 
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10560
         TabIndex        =   13
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   375
         Left            =   480
         TabIndex        =   34
         Top             =   3480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Left            =   480
         TabIndex        =   33
         Top             =   4440
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Top             =   3960
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   855
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   12135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Global Options"
      Height          =   2535
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   12855
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   7200
         TabIndex        =   50
         Top             =   1200
         Width           =   5055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "As above, but do NOT extend the accesspath"
         Height          =   255
         Left            =   6960
         TabIndex        =   49
         Top             =   1560
         Width           =   5655
      End
      Begin VB.CheckBox Check1 
         Caption         =   $"frmProcess.frx":0043
         Height          =   375
         Left            =   6960
         TabIndex        =   48
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox Textbox9 
         Height          =   375
         Left            =   10680
         TabIndex        =   36
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox Checkbox13 
         Caption         =   "Increase Heap-memory to 256mb"
         Height          =   195
         Left            =   4800
         TabIndex        =   30
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox Checkbox12 
         Caption         =   "Increase Heap-memory to 512mb"
         Height          =   255
         Left            =   7920
         TabIndex        =   29
         Top             =   360
         Width           =   3255
      End
      Begin VB.CheckBox Checkbox5 
         Caption         =   "Suppress specific messages from the log file and screen"
         Height          =   195
         Left            =   960
         TabIndex        =   19
         Top             =   2040
         Width           =   4935
      End
      Begin VB.CheckBox Checkbox4 
         Caption         =   "Suppress all 'Information' messages from the log file and screen"
         Height          =   195
         Left            =   960
         TabIndex        =   18
         Top             =   1620
         Width           =   5775
      End
      Begin VB.CheckBox Checkbox3 
         Caption         =   "Suppress all 'Information' and 'Warning' messages from the log file and screen"
         Height          =   195
         Left            =   960
         TabIndex        =   17
         Top             =   1200
         Width           =   5895
      End
      Begin VB.CheckBox Checkbox2 
         Caption         =   "Include message keys in the log file"
         Height          =   195
         Left            =   960
         TabIndex        =   16
         Top             =   780
         Width           =   3615
      End
      Begin VB.CheckBox Checkbox1 
         Caption         =   "Write a log file for this option"
         Height          =   195
         Left            =   960
         TabIndex        =   15
         Top             =   360
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Message numbers to suppress (separate with space)"
         Height          =   195
         Left            =   5640
         TabIndex        =   35
         Top             =   2040
         Width           =   5055
      End
   End
   Begin VB.TextBox Textbox4 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   8655
   End
   Begin VB.TextBox Textbox3 
      Height          =   375
      Left            =   11040
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Textbox2 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   8655
   End
   Begin VB.TextBox Textbox1 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   8655
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "If this appears OK, click 'Complete Processing' to finalise. The above command may be edited manually"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   47
      Top             =   10080
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Log File Path"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "On Route"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Using TsUtils Command"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strCheck As String

Public CvrtFlag As Integer

Public strBatText As String
Private Sub ParseH(strText As String, strFlag As String, strTemp As String)
Dim x As Integer, strM(0 To 10) As String, intM(0 To 10) As Integer, i As Integer

 On Error GoTo Errtrap
            x = InStr(strText, " ")
            If x = 0 Then
                strM(0) = Trim(strText)
                numM = 1
                'strworldoptions = strworldoptions & " -m" & strM(0)
            Else
                x = 0
                i = 0
                Do
                    x = InStr(x + 1, strText, " ")
                    If x > 0 Then
                        intM(i) = x
                        i = i + 1
                    Else: Exit Do
                    End If
                Loop
                
                strM(0) = Left(strText, (intM(0) - 1))
                For x = 1 To i - 1
                    strM(x) = Mid(strText, intM(x - 1) + 1, intM(x) - (intM(x - 1) + 1))
                Next
                strM(i) = Mid(strText, intM(i - 1) + 1)
            End If
            strTemp = ""
            For x = 0 To i
                strTemp = strTemp & strFlag & strM(x)
            Next
'            strworldoptions = strworldoptions & strTemp
'            strTemp = ""
       ' End If
       Exit Sub
Errtrap:
       Call MsgBox("Error " & Err & " " & Err.Description & " occurred while running ParseH.", vbExclamation, App.Title)

End Sub

Private Sub ParseV(strV As String)
Dim x As Integer, Y As Integer, i As Integer
Dim strvv() As String, strTemp As String

x = InStr(strV, " ")
If x = 0 Then
Y = InStr(strV, ":")
    If Y = 0 Then
    Call MsgBox("Invalid format in the 'Manually assign shapes box' must be Old number:New number etc" _
                & vbCrLf & "e.g. 100:137 101:138 etc" _
                , vbExclamation, App.Title)
    
    Exit Sub
    ElseIf Y = 1 Then
    strCheck = " -v" & strV
    Exit Sub
    End If
End If
x = 1
Y = 0
Do While x > 0
x = InStr(x, strV, " ")
If x > 0 Then
Y = Y + 1
x = x + 1
End If
Loop
ReDim strvv(0 To Y) As String
x = 1

Do While x > 0
Y = x
x = InStr(x, strV, " ")
If x > 0 Then
strvv(i) = Mid(strV, Y, x - Y)
i = i + 1
x = x + 1
End If
Loop

strvv(i) = Mid(strV, Y)
For x = 0 To i
strTemp = " -v" & strvv(x)
strCheck = strCheck & strTemp
Next x

End Sub

Private Sub Button1_Click()
Unload Me
End Sub

Private Sub Button2_Click()
Dim strX As String, strMask As String
        Dim DX As Integer, DY As Integer, strTemp As String, x As Integer
        Dim strM(0 To 10) As String, i As Integer, numM As Integer
        Dim intM(0 To 20) As Integer, strMerge As String
        Dim strType As String, strTemp2 As String

On Error GoTo Errtrap
        DX = 0
        DY = 0
        strworldoptions = ""
        strX = ""
        If Checkbox1.value = 1 Then
      
            If strOptionCode = "chkup" Then
                strworldoptions = "-l" & ChrW$(34) & strLogPath & "\" & strOptionCode & ".log"
                strReport = strLogPath & "\" & strOptionCode & ".log"
            Else
                strworldoptions = "-l" & ChrW$(34) & strLogPath & "\" & strRouteName & "_" & strOptionCode & ".log" & ChrW$(34)
                strReport = strLogPath & "\" & strRouteName & "_" & strOptionCode & ".log"
            End If
        End If
        If Checkbox2.value = 1 Then
            strworldoptions = strworldoptions & " -k"
        End If
        If Checkbox3.value = 1 Then
            strworldoptions = strworldoptions & " -e"
        End If
        If Checkbox4.value = 1 Then
            strworldoptions = strworldoptions & " -w"
        End If
        If Checkbox5.value = 1 And Textbox9.Text <> "" Then
        
            x = InStr(Textbox9.Text, " ")
            If x = 0 Then
                strM(0) = Trim(Textbox9.Text)
                numM = 1
                'strworldoptions = strworldoptions & " -m" & strM(0)
            Else
                x = 0
                i = 0
                Do
                    x = InStr(x + 1, Textbox9.Text, " ")
                    If x > 0 Then
                        intM(i) = x
                        i = i + 1
                    Else: Exit Do
                    End If
                Loop
                
                strM(0) = Left(Textbox9.Text, (intM(0) - 1))
                For x = 1 To i - 1
                    strM(x) = Mid(Textbox9.Text, intM(x - 1) + 1, intM(x) - (intM(x - 1) + 1))
                Next
                strM(i) = Mid(Textbox9.Text, intM(i - 1) + 1)
            End If
            strTemp = ""
            For x = 0 To i
                strTemp = strTemp & " -m" & strM(x)
            Next
            strworldoptions = strworldoptions & strTemp
            strTemp = ""
        End If
        If Check1.value = 1 Then
            If Text2 = "" Then
            strTemp = " -h"
            Else
            strTemp = " -h" & Trim(Text2)
            End If
        End If
        If Check2.value = 1 Then
        Check1.value = 0
        Text2 = ""
        strTemp = " -H"
        End If
        strworldoptions = strworldoptions & strTemp
            strTemp = ""
        If Checkbox13.value = 1 Then
            strX = " -Xmx256m"
        End If
        If Checkbox12.value = 1 Then
            strX = " -Xmx512m"
        End If
        
        strworldoptions = "java" & strX & " TSUtil " & strworldoptions
        Select Case TSOption
            Case 1
               

                    Textbox5.Text = strworldoptions & " wunc " & ChrW$(34) & strSelectedPath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True: Command1.Enabled = True

                    
                            
               
            Case 2
                

                    Textbox5.Text = strworldoptions & " wcmp " & ChrW$(34) & strSelectedPath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True: Command1.Enabled = True


                
            Case 3
                
                    Textbox5.Text = strworldoptions & " wcmpo " & ChrW$(34) & strSelectedPath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
        

            Case 4
               
                    Textbox5.Text = "java" & strX & " TSUtil cmkr " & ChrW$(34) & Textbox6.Text & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
                
            Case 5
           
                For x = 0 To 2
                If List1.Selected(x) = True Then
                strType = "-" & List1.List(x)
                strType = LCase(strType)
                End If
                Next x
                strType = " " & strType
                If Trim(strType) = "" Then strType = ""
                For x = 0 To 2
                If List2.Selected(x) = True Then
                Select Case List2.List(x)
                Case "COMP"
                strTemp = "-c"
                Case "UnC Txt"
                strTemp = "-e"
                Case "UnC Bin"
                strTemp = "-r"
                End Select
                End If
                Next x
                strTemp = " " & strTemp
                If Trim(strTemp) = "" Then strTemp = ""
                If Textbox8 <> "" Then
                strMask = Trim(Textbox8)
                End If
                If Checkbox8.value = 1 Then
                strMask = "-m" & strMask
                ElseIf Checkbox9.value = 1 Then
                strMask = "-n" & strMask
                End If
                strMask = " " & strMask
                If Trim(strMask) = "" Then strMask = ""
                If Checkbox10.value = 1 Then
                strTemp2 = " -o"
                End If
                Textbox5.Text = strworldoptions & " fmgr " & strType & strTemp & strMask & strTemp2 & " " & ChrW$(34) & Textbox6 & ChrW$(34) & " " & ChrW$(34) & Textbox7 & ChrW$(34)
                Label9.Visible = True: Command1.Enabled = True



            Case 6
            strTemp = Trim(Text1(2)) & " " & Trim(Text1(3)) & " " & Trim(Text1(0)) & " " & Trim(Text1(1))
            strTemp = Trim(strTemp)
            
            Textbox5.Text = strworldoptions & " chgdb " & ChrW$(34) & Textbox6.Text & ChrW$(34) & " " & strTemp
            Label9.Visible = True: Command1.Enabled = True


            Case 7
            x = InStrRev(Textbox2, "\")
            strTemp = Left(Textbox2, x - 1)
            strTemp = ChrW$(34) & strTemp & ChrW$(34) & " " & ChrW$(34) & Textbox6 & ChrW$(34) & " " & ChrW$(34) & Textbox7 & ChrW$(34)
            If Checkbox6.value = 1 Then
                    strCheck = "-a "
                End If
                If Checkbox7.value = 1 Then
                    strCheck = strCheck & "-o "
                End If
             Textbox5.Text = strworldoptions & " dcpy " & strCheck & " " & strTemp
             Label9.Visible = True: Command1.Enabled = True

            Case 8
            strTemp = ChrW$(34) & Textbox6 & ChrW$(34) & " " & Text1(0) & " " & Text1(1) & " " & Text1(2) & " " & Text1(3) & " " & Text1(4) & " " & Text1(5)
            Textbox5.Text = strworldoptions & " shift " & strCheck & " " & strTemp
            Label9.Visible = True: Command1.Enabled = True

            Case 9
            
                If Checkbox6.value = 1 Then
                    strCheck = "-w "
                    ElseIf Checkbox7.value = 1 Then
                    strCheck = "-c -w"
                    Checkbox6.value = 0
                End If
                

                    Textbox5.Text = strworldoptions & " filter " & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
                    


                
                strCheck = ""

            Case 10
                

                    Textbox5.Text = strworldoptions & " cmpw " & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
                    
                
                strCheck = ""
            Case 11
                If Checkbox6.value = 1 Then
                    strCheck = "-s "
                End If
                If Checkbox7.value = 1 Then
                    strCheck = strCheck & "-u "
                End If
                If Checkbox8.value = 1 Then
                    strCheck = strCheck & "-v"
                End If
               

                    Textbox5.Text = strworldoptions & " tsconv " & strCheck & " " & ChrW$(34) & strMSTSpath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
                    


                strCheck = ""
            Case 12
                If Checkbox6.value = 1 Then
                    strCheck = " -s"
                ElseIf Checkbox7.value = 1 Then
                    strCheck = " -S"
                End If
                If Checkbox8.value = 1 Then
                    strCheck = strCheck & " -k"
                End If
                If Checkbox9.value = 1 Then
                Call ParseH(Text1(3), " -z", strTemp)
                End If
               If strTemp <> "" Then
               strCheck = strCheck & strTemp
                End If
                    Textbox5.Text = strworldoptions & " ichk" & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
                    
                    


                strCheck = ""
            Case 13
                If Checkbox6.value = 1 Then
                    strCheck = "-o "
                    Checkbox7.value = 0
                    Checkbox8.value = 0
                    Checkbox9.value = 0
                    Checkbox10.value = 0
                    Checkbox11.value = 0
                End If
                If Checkbox7.value = 1 Then
                    strCheck = strCheck & "-w "
                End If
                If Checkbox8.value = 1 Then
                    strCheck = strCheck & "-a"
                End If
                If Checkbox9.value = 1 Then
                    strCheck = strCheck & "-m "
                End If
                If Checkbox10.value = 1 Then
                    strCheck = strCheck & "-c"
                    Checkbox11.value = 0
                End If
                If Checkbox11.value = 1 Then
                    strCheck = strCheck & "-u "
                    Checkbox10.value = 0
                End If

               

                    Textbox5.Text = strworldoptions & " rendb " & strCheck & " " & ChrW$(34) & strMSTSpath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True


                strCheck = ""
            Case 14
                If Checkbox6.value = 1 Then
                    strCheck = "-c "
                End If
                If Checkbox8.value = 1 Then
                    strCheck = strCheck & "-p "
                End If
                DX = Val(Textbox7.Text)
                DY = Val(Textbox8.Text)
               

                    Textbox5.Text = strworldoptions & " move " & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34) & Str(DX) & " " & Str(DY)
                    Label9.Visible = True: Command1.Enabled = True
                    

                
                strCheck = ""

            Case 15
                If Textbox7.Text <> "" And Textbox8.Text <> "" Then
                    strCheck = " -v" & Trim$(Textbox7.Text) & ":" & Trim$(Textbox8.Text) & " "
                End If
                x = InStrRev(Textbox6.Text, "\")
                strSelectedPath = Left$(Textbox6.Text, (x - 1))

              

                    Textbox5.Text = strworldoptions & " chkup " & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
                    

                
                strCheck = ""
            Case 16
                If Checkbox6.value = 1 Then
                    strCheck = "-c "
                End If
                DX = Val(Textbox7.Text)
                

                    Textbox5.Text = strworldoptions & " adjh " & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34) & Str(DX)
                    Label9.Visible = True: Command1.Enabled = True
                    

                strCheck = ""
            Case 17
                If Checkbox6.value = 1 Then
                    strCheck = " -s"
                ElseIf Checkbox7.value = 1 Then
                    strCheck = " -t"
                ElseIf Checkbox8.value = 1 Then
                    strCheck = " -T"
                ElseIf Checkbox9.value = 1 Then
                    strCheck = " -w"
                End If
                If Checkbox10.value = 1 And strCheck = " -s" Or strCheck = " -w" Then
                    strCheck = strCheck & " -c"
                End If
                If Checkbox11.value = 1 Then
                    strCheck = strCheck & " -x"
                End If
                

                    Textbox5.Text = strworldoptions & " cmp " & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
                    

                
                strCheck = ""
            Case 18
                If Checkbox6.value = 1 Then
                    strCheck = " -s"
                ElseIf Checkbox7.value = 1 Then
                    strCheck = " -t"
                ElseIf Checkbox8.value = 1 Then
                    strCheck = " -T"
                ElseIf Checkbox9.value = 1 Then
                    strCheck = " -w"
                End If
                If Checkbox11.value = 1 Then
                    strCheck = strCheck & " -x"
                End If
               

                    Textbox5.Text = strworldoptions & " unc " & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
                    

                
                strCheck = ""
            Case 19
                If Checkbox6.value = 1 Then
                    strCheck = " -p"
                End If
                If Val(Textbox7.Text) > 0 Then
                    strCheck = strCheck & " -h" & Textbox7.Text
                End If
                strMerge = Textbox6.Text
                

                    Textbox5.Text = strworldoptions & " merge " & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34) & " " & ChrW$(34) & strMerge & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
                    

               
                strCheck = ""
            Case 20
                If Checkbox6.value = 1 Then
                    strCheck = " -t"
                End If
                If Checkbox8.value = 1 Then
                    strCheck = strCheck & " -p"
                End If
                If Checkbox9.value = 1 Then
                    strCheck = strCheck & " -w"
                End If
                

                    Textbox5.Text = strworldoptions & " clrdb " & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
                    

                strCheck = ""
            Case 21
                

                    Textbox5.Text = strworldoptions & " zusi " & ChrW$(34) & strSelectedPath & ChrW$(34)
                    Label9.Visible = True: Command1.Enabled = True
                    
                
                strCheck = ""
            Case 22
                If Checkbox6.value = 1 Then
                    strCheck = " -c"

                ElseIf Checkbox7.value = 1 Then
                    strCheck = " -u"
                End If
                If Checkbox8.value = 1 Then
                    strCheck = strCheck & " -w"
                End If
                If Textbox7.Text <> "" And Textbox8.Text <> "" Then
                    strTemp = " " & Textbox7.Text & " " & Textbox8.Text

                End If

               

                    Textbox5.Text = strworldoptions & " shftdyn " & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34) & strTemp
                    Label9.Visible = True: Command1.Enabled = True
                    

                
            Case 23

                    strBatText = strworldoptions & " version"
                    Textbox5.Text = strworldoptions & " version"
                    Label9.Visible = True: Command1.Enabled = True

               
            Case 24
                If Checkbox6.value = 1 Then strCheck = " -r"
                If Checkbox7.value = 1 Then strCheck = strCheck & " -s"
                If Checkbox8.value = 1 Then
                strTemp = " -c"
                ElseIf Checkbox9.value = 1 Then
                strTemp = " -u"
                End If
                strCheck = strCheck & strTemp
                If Checkbox10.value = 1 Then strCheck = strCheck & " -k"
                If Text1(5) <> "" Then
                Call ParseV(Text1(5))
                End If
                If Textbox6 <> "" Then
                strTemp2 = " -b" & ChrW$(34) & Textbox6 & ChrW$(34)
                End If
                
                Textbox5.Text = strworldoptions & " cvrt " & strCheck & strTemp2 & " " & ChrW$(34) & strSelectedPath & ChrW$(34)
                Label9.Visible = True: Command1.Enabled = True
                

            Case 25
                Textbox5.Text = strworldoptions & " cvrt -v99:99 " & ChrW$(34) & strSelectedPath & ChrW$(34)
                Label9.Visible = True: Command1.Enabled = True
                

            Case 26
                If Checkbox6.value = 1 Then
                
                Call ParseH(Text1(0), " -h", strTemp)
                strCheck = strTemp
                End If
                
                If Checkbox7.value = 1 Then strCheck = " -H"
                If Checkbox8.value = 1 Then
                Call ParseH(Text1(2), " -m", strTemp)
                strCheck = strTemp
                End If
                                
                If Checkbox9.value = 1 Then strCheck = " -m"
                If Checkbox10.value = 1 Then
                Call ParseH(Text1(4), " -z", strTemp)
                strCheck = strTemp
                
                End If
                Textbox5.Text = strworldoptions & " cvrt" & strCheck & " " & ChrW$(34) & strSelectedPath & ChrW$(34)
                Label9.Visible = True: Command1.Enabled = True
                

                
        End Select
        DoEvents
Exit Sub
Errtrap:
Call MsgBox("Error " & Err & " " & Err.Description & " occurred while processing this option.", vbExclamation, App.Title)


End Sub

Private Sub Button3_Click()
TSUFlag = 2
TsUtil_CD.Show 1
End Sub

Private Sub Button4_Click()
        If strOptionCode = "dcpy" Then
        TSUFlag = 3
        TsUtil_CD.Show 1
        Exit Sub
        End If
        If strOptionCode = "fmgr" Then
        TSUFlag = 3
        TsUtil_CD.Show 1
        Exit Sub
        End If


        If strOptionCode = "cmkr" Then
            OpenFileDialog1.Filter = "Coo Files|*.coo"
            OpenFileDialog1.DialogTitle = "Select a Coo File"
        ElseIf strOptionCode = "cvrt" Then
            OpenFileDialog1.Filter = "dat Files|*.dat"
            OpenFileDialog1.DialogTitle = "Select a tsection.dat File"
        ElseIf strOptionCode = "chgdb" Then
            OpenFileDialog1.Filter = "Database Files(*.rdb;*tdb)|*.rdb;*tdb"
            OpenFileDialog1.DialogTitle = "Select a .tdb or .rdb File"
            OpenFileDialog1.InitDir = strSelectedPath
        ElseIf strOptionCode = "shift" Then
            OpenFileDialog1.Filter = "tsection.dat Files(*.dat)|*.dat"
            OpenFileDialog1.DialogTitle = "Select a tsection.dat File"
            OpenFileDialog1.InitDir = MSTSPath & "\Global"
        End If
            OpenFileDialog1.Action = 1

        If OpenFileDialog1.Filename <> "" Then
            Textbox6.Text = OpenFileDialog1.Filename
          

        End If
End Sub

Private Sub Button5_Click()
 MousePointer = 0
If FileExists(strReport) Then
 frmReport.Rich1.LoadFile strReport
 frmReport.Show 1
 DoEvents
 Else
 Call MsgBox("No Log file has been produced for this option. The process may have failed.", vbExclamation, App.Title)
 
 End If
  End Sub

Private Sub Button6_Click()
If strOptionCode <> "fmgr" Then
TSUFlag = 4
TsUtil_CD.Show 1
        ElseIf strOptionCode = "fmgr" Then
        TSUFlag = 4
        TsUtil_CD.Show 1
        
        End If

End Sub

Private Sub Checkbox11_Click()
If Checkbox11.value = 1 And CvrtFlag = 3 Then
            Label5.Visible = True
            Textbox7.Visible = True
            Label5.Caption = "StaticDetailLevel  :  Old track type  :  New track type  ---  e.g. 1:A:Xt (up to 3 entries)"
End If
End Sub

Private Sub Checkbox12_Click()
If Checkbox12.value = 1 Then
            Checkbox13.value = 0
        End If
End Sub

Private Sub Checkbox13_Click()
If Checkbox13.value = 1 Then
            Checkbox12.value = 0
        End If
End Sub

Private Sub Checkbox3_Click()
If Checkbox3.value = 1 Then
            Checkbox4.value = 0
        End If
End Sub

Private Sub Checkbox4_Click()
If Checkbox4.value = 1 Then
            Checkbox3.value = 0
        End If
End Sub

Private Sub Checkbox6_Click()
If strOptionCode = "ichk" Then
If Checkbox6.value = 1 Then
            Checkbox7.value = 0
        End If
        End If
End Sub

Private Sub Checkbox7_Click()
If strOptionCode = "ichk" Then
If Checkbox7.value = 1 Then
            Checkbox6.value = 0
        End If
        End If
End Sub


Private Sub Command1_Click()

                        ChDrive Left$(App.Path, 1)
                            ''ChDir App.Path & "\TSUtil"
                            strBatText = Textbox5.Text
                            Call ShellAndWait(strBatText, True, vbNormalFocus)
DoEvents
Button5.value = True
End Sub

Private Sub Form_Load()
Dim i As Integer

Textbox1.Text = strTSoption
        If TSOption = 11 Then
            Textbox2.Text = strMSTSpath
        Else
            Textbox2.Text = strSelectedPath
        End If
        Textbox3.Text = TSOption
        Textbox4.Text = App.Path & "\Reports"
        strLogPath = Textbox4.Text
        Checkbox6.Visible = False
        Checkbox7.Visible = False
        Checkbox8.Visible = False
        Checkbox9.Visible = False
        Checkbox10.Visible = False
        Checkbox11.Visible = False
        Textbox6.Visible = False
        Button4.Visible = False
        Button6.Visible = False
        For i = 0 To 5
        Text1(i).Visible = False
        Next
        Command1.Enabled = False
        Label6.Visible = False
        Label9.Visible = False
        Textbox8.Visible = False
        List1.Visible = False
        List2.Visible = False
        Text1(0).Alignment = 1
                Text1(1).Alignment = 1
        Select Case TSOption
            Case 1
                Label4.Caption = "The function 'wunc' is used to convert ALL world-files to 'uncompressed text'using ffeditc_unicode.exe" _
                & vbCrLf & "No other options are applicable. The uncompressed files will be found in the World folder (suggest you should use the similar option in Route_Riter)"
                
                strOptionCode = "wunc"
            Case 2
                Label4.Caption = "The function 'wcmp' is used to convert ALL world-files to 'compressed binary'using ffeditc_unicode.exe" _
                & vbCrLf & "No other options are applicable. The uncompressed files will be found in the World folder (suggest you should use the similar option in Route_Riter)"
                
                strOptionCode = "wcmp"
            Case 3
                Label4.Caption = "The function 'wcmpo' is used to convert world-files to 'compressed binary' if they are currently uncompressed Text files (otherwise" _
                & vbCrLf & "they are unchanged) using ffeditc_unicode.exe. No other options are applicable. " _
                & vbCrLf & "The uncompressed files will be found in the World folder (suggest you should use the similar option in Route_Riter)"
                strOptionCode = "wcmpo"
            Case 4
                Label4.Caption = "The function 'cmkr' converts '.coo' files to a MSTS .mkr file." _
                & vbCrLf & "e.g. c:\Test.coo is converted to c:\Test.mkr"
                Textbox6.Visible = True
                Button4.Visible = True
                strOptionCode = "cmkr"
            Case 5
                Label4.Caption = "The function 'fmgr' is used to compress/expand .s/.t/.w files or groups of files, including all files in a folder. " _
                & vbCrLf & "This function only works on one file type at a time. See Tsutils_en.txt for full details as to how you can use either" _
                & vbCrLf & "'Regular Expressions' or Windows wild-cards to select those batches of files you wish to process."
                Label7.Visible = True
                Label5.Visible = True
                Textbox6.Visible = True
                Textbox7.Visible = True
                Button4.Visible = True
                Button6.Visible = True
                Label6.Visible = True
                Textbox8.Visible = True
                Label7.Caption = "Source folder (Source/Target may be the same)"
                Label5.Caption = "Target folder"
                Label6.Caption = "Mask for file-selection (Not needed when converting whole folder)"
                Checkbox6.Visible = True
                Checkbox7.Visible = True
                Checkbox8.Visible = True
                Checkbox9.Visible = True
                Checkbox10.Visible = True
                List1.Visible = True
                List2.Visible = True
                Checkbox6.Caption = "Type of files to process? (s, t or w)"
                Checkbox7.Caption = "Target file format - Compressed binary, Uncompressed text or Uncomp binary (CB,UT or UB)"
                Checkbox8.Caption = "Use regular-expression mask for file selection"
                Checkbox9.Caption = "Use Windows style pattern matching (* or ?) for file selection"
                Checkbox10.Caption = "Automatically create any missing sub-directories for target"
                strOptionCode = "fmgr"
            
            Case 6
                Label4.Caption = "The function 'chgdb' makes it possible to adapt the track/road database for changes" _
                & vbCrLf & "in the Global\tsection.dat file, this function may be used if a rebuild of .tdb has failed"
               Checkbox6.Caption = "Maximum Shapes entry in original tsection.dat (Default=263)"
                Checkbox7.Caption = "Maximum Sections entry in original tsection.dat (Default=376)"
                Checkbox8.Caption = "New Shapes entry which can be used (e.g. 40000 for x-tracks)"
                Checkbox9.Caption = "New Sections entry which can be used (e.g. 40000 for x-tracks)"
                Checkbox6.Visible = True
                Checkbox7.Visible = True
                Checkbox8.Visible = True
                Checkbox9.Visible = True
                For i = 0 To 3
                Text1(i).Visible = True
                Next i
                Button4.Visible = True
                Text1(0) = "263"
                Text1(1) = "376"
                Text1(2) = "40000"
                Text1(3) = "40000"
                Label7.Caption = "Select .tdb or .rdb file to modify"
                Label7.Visible = True
                Textbox6.Visible = True
                strOptionCode = "chgdb"
            Case 7
                Label4.Caption = "The function 'dcpy' compares NewRoute with OldRoute and copies any Files/Folders which" _
                & vbCrLf & "have been changed into a third folder. This can be used to provide a patch to update Old to New"
                Checkbox6.Visible = True
                Checkbox7.Visible = True
                Checkbox6.Caption = "Include files with suffix '.bk' in the comparison"
                Checkbox7.Caption = "Only copy files from Dir1 to Dir3 if also in Dir2 and has a newer date."
                Label7.Visible = True
                Label5.Visible = True
                Textbox6.Visible = True
                Textbox7.Visible = True
                Button4.Visible = True
                Button6.Visible = True
                Label7.Caption = "Old version path"
                Label5.Caption = "Patch copy path"
                strOptionCode = "dcpy"
                
            Case 8
                Label4.Caption = "The function 'shift' shifts ranges of ShapeID and/or Section-IDs within a tsection.dat" _
                & vbCrLf & "file. e.g. values 0 400 2000 0 400 2000 moves all shapes and sections between 0 and 400" _
                & vbCrLf & "to positions starting at 2000"
                Checkbox6.Visible = False
                Checkbox7.Visible = True
                Checkbox8.Visible = True
                Checkbox9.Visible = True
                Checkbox10.Visible = True
                Checkbox11.Visible = True
                For i = 0 To 5
                Text1(i).Visible = True
                Next
                Checkbox6.Caption = "Lower value of Shape-ID range"
                Checkbox7.Caption = "Upper value of Shape-ID range"
                Checkbox8.Caption = "Shift value for Shape-IDs"
                Checkbox9.Caption = "Lower value of  Section-ID range"
                Checkbox10.Caption = "Upper value of Section-ID range"
                Checkbox11.Caption = "Shift value for Section-IDs"
                 Label7.Visible = True
                Label7.Caption = "Full path of tsection.dat"
                Textbox6.Visible = True
                Button4.Visible = True
                strOptionCode = "shift"
            Case 9
                Label4.Caption = "The function 'filter' checks all marked tiles in the Route Geometry for completeness." _
                & vbCrLf & "If no .t file exists, the tile is unmarked. If the -w flag is selected, the tile is also" _
                & vbCrLf & "unmarked if no corresponding .w file exists. Both the Tile and Lo_Tile folders are checked"
                Checkbox6.Caption = "Unmark tile if no .w tile is present"
                Checkbox6.Visible = True
                Checkbox7.Caption = "Unmark tile if no .w tile is present AND delete .t file"
                Checkbox7.Visible = True
                strOptionCode = "filter"
            Case 10
                Label4.Caption = "The function 'cmpw' compresses .w files using Martyn Griffin's 'Comp' program." _
                & vbCrLf & "This program must be installed in the 'Utils\Comp' folder for the current MSTS instance."
                strOptionCode = "cmpw"
               
                If Not FileExists(MSTSPath & "\utils\comp\comp.exe") Then
                Call MsgBox("Utils\Comp\Comp.exe not found, this option can not be used.", vbExclamation, App.Title)
                
                Exit Sub
                End If
            Case 11
                Label4.Caption = "The function 'tsconv' modifies ALL routes in the current MSTS instance for use with a" _
                & vbCrLf & "new tsection.dat file (see 'cvrt' which only modifies a single route). The original tsection.dat" _
                & vbCrLf & "must be placed in the Routes folder. Alternatively if a route has used a different tsection.dat" _
                & vbCrLf & "This may be placed within that route's folder and named $tsection.dat"
                Checkbox6.Caption = "Check that all shape files referred to in the new tsection.dat are present"
                Checkbox7.Caption = "Write new world files in uncompressed Unicode text."
                Checkbox8.Caption = "Manually assign different definitions between the Old and New tsection.dat files"
                Checkbox6.Visible = True
                Checkbox7.Visible = True
                Checkbox8.Visible = True
                strOptionCode = "tsconv"
            Case 12
                Label4.Caption = "The function 'ichk' checks the syntax of the global and local tsection.dat files and" _
                & vbCrLf & "whether references between road/track databases and world files are correct. It also checks" _
                & vbCrLf & "if duplicate UiD identifications are used and if these may be changed. However NO changes are made."
                Checkbox6.Caption = "Check if Shape Files listed in Global tsection.dat are present"
                Checkbox7.Caption = "As above, but also checks all shapes/textures referred by World files are present"
                Checkbox8.Caption = "Enable 'Correction mode' for .W files - See TsUtil_en.txt"
                Checkbox9.Caption = "Check  Static Objects which are hidden 'Hidewire' objects (use format 1:Xt:A etc)"
                Checkbox6.Visible = True
                Checkbox7.Visible = True
                Checkbox8.Visible = True
                Checkbox9.Visible = True
                Text1(3).Visible = True
                strOptionCode = "ichk"
            Case 13
                Label4.Caption = "The function 'rendb' can repair references within track and road databases following" _
                & vbCrLf & "manual deletion of defective Track-Nodes. Gaps are removed etc. See TSUtil_en.txt for full details"
                Checkbox6.Visible = True
                Checkbox7.Visible = True
                Checkbox8.Visible = True
                Checkbox9.Visible = True
                Checkbox10.Visible = True
                Checkbox11.Visible = True
                Checkbox6.Caption = "Only reorganize the TDB (no manual changes must have been made)"
                Checkbox7.Caption = "Always create new databases"
                Checkbox8.Caption = "Remove all TRitem-definitions which are no longer referenced"
                Checkbox9.Caption = "If UID entries contain errors, replace them with '?'"
                Checkbox10.Caption = "Write all .w files in Compressed format"
                Checkbox11.Caption = "Write all .w files in Unicode format"

                strOptionCode = "rendb"
            Case 14
                Label4.Caption = "The function 'move' shifts a complete route by the defined number of squares in the X" _
                & vbCrLf & "and/or Z directions."
                Checkbox6.Visible = True
                Checkbox6.value = 1
                Checkbox8.Visible = True
                Checkbox6.Caption = "Write .t files in 'reduced format'"
                Checkbox8.Caption = "Write new *_e.raw and *n_raw files"
                Label5.Caption = "Number of tiles to shift E\W (W is negative)"
                Label6.Caption = "Number of tiles to shift N\S (S is negative)"
                Label5.Visible = True
                Label6.Visible = True
                Textbox7.Visible = True
                Textbox8.Visible = True
                strOptionCode = "move"
            Case 15
                Label4.Caption = "The function 'chkup' checks two versions of tsection.dat (both must be in same folder)" _
                & vbCrLf & "to see if the newer one is compatible with the older version." _
                & vbCrLf & "The shape/section definitions may have been altered between the two files."
                Label7.Visible = True
                Textbox6.Visible = True
                Label7.Caption = "Old tsection.dat file"
                Label5.Caption = "Manually change Shape Number:"
                Label6.Caption = "To:"
                Label5.Visible = True
                Label6.Visible = True
                Textbox7.Visible = True
                Textbox8.Visible = True
                strOptionCode = "chkup"
            Case 16
                Label4.Caption = "The function 'adjh' adjusts the altitude of the route. This adjustment figure which can" _
                & vbCrLf & "be negative is added to all height values in the route definition."
                Checkbox6.Visible = True
                Checkbox6.value = 1
                Checkbox6.Caption = "Write .t files in 'reduced format'"
                Label5.Caption = "Change altitude by (in metres)"
                Label5.Visible = True
                Textbox7.Visible = True
                strOptionCode = "adjh"

            Case 17
                Label4.Caption = "The function 'cmp' will compress all .s/.t/.w files which are in uncompressed text" _
                & vbCrLf & "format to uncompressed binary. With the -c flag, the .s/.w files are then compressed to " _
                & vbCrLf & "compressed binary (.t files cannot be compressed further)"
                Checkbox6.Visible = True
                Checkbox7.Visible = True
                Checkbox8.Visible = True
                Checkbox9.Visible = True
                Checkbox10.Visible = True
                Checkbox11.Visible = True
                Checkbox6.Caption = "Compact to uncompressed binary .s files in the Shapes folder to 'newroute'"
                Checkbox7.Caption = "Compact to uncompressed binary .t files in the Tiles folder to 'newroute'"
                Checkbox8.Caption = "Compact to uncompressed binary .t files in the Lo_Tiles folder to 'newroute'"
                Checkbox9.Caption = "Compact to uncompressed binary .w files in the World folder to 'newroute'"
                Checkbox10.Caption = "Further compress to compressed binary (.s/.w files only)"
                Checkbox11.Caption = "Copy ALL files in folder to 'newroute' whether processed or not."

                strOptionCode = "cmp"
            Case 18
                Label4.Caption = "The function 'unc' uncompresses all .s/.t/.w files in the route"
                Checkbox6.Visible = True
                Checkbox7.Visible = True
                Checkbox8.Visible = True
                Checkbox9.Visible = True
                Checkbox10.Visible = True
                Checkbox6.Caption = "Uncompress .s files in the Shapes folder to 'newroute'"
                Checkbox7.Caption = "Uncompress .t files in the Tiles folder to 'newroute'"
                Checkbox8.Caption = "Uncompress .t files in the Lo_Tiles folder to 'newroute'"
                Checkbox9.Caption = "Uncompress .w files in the World folder to 'newroute'"
                Checkbox10.Caption = "Copy ALL files in folder to 'newroute' whether processed or not."
                strOptionCode = "unc"
            Case 19
                Label4.Caption = "The function 'merge' creates a common, merged, route from two route definitions." _
                & vbCrLf & "See the TsUtils_en.txt file for restrictions/limitations of this option"

                Label7.Visible = True
                Textbox6.Visible = True
                Label7.Caption = "Route to be merged:-"
                Checkbox6.Visible = True
                Label5.Visible = True
                Label5.Caption = "Modify 'StaticDetailLevel' to (1, 2 or 3)"
                Checkbox6.Caption = "Write new *_e.raw and *n_raw files"
                Textbox7.Visible = True
                strOptionCode = "merge"
            Case 20
                Label4.Caption = "The function 'clrdb' analyzes track- and road-database and deletes all definition-chains" _
                & vbCrLf & "(TrEndNode-TrVectorNode/TrJunctionNode - TrEndNode), which are no longer used. "
                Checkbox6.Visible = True
                Checkbox8.Visible = True
                Checkbox9.Visible = True
                Checkbox6.Caption = "Copy all .t files to the 'newRoute' folder"
                Checkbox8.Caption = "Write new *_e.raw and *n_raw files"
                Checkbox9.Caption = "Automatically generate any missing .w files"
                Checkbox6.value = 1
                Checkbox8.value = 1
                Checkbox9.value = 1
                strOptionCode = "clrdb"
            Case 21
                Label4.Caption = "The function 'zusi' Prepares a route for import into the Zusi train simulator"
                strOptionCode = "zusi"
            Case 22
                Label4.Caption = "The function 'shftdyn' is used to shift the range of dynamic track definitions in the" _
                & vbCrLf & "local tsection.dat to their correct position as defined in the Global\tsection.dat"
                Checkbox6.Visible = True
                Checkbox7.Visible = True
                Checkbox8.Visible = True
                Checkbox6.Caption = "Write .w files in compressed format"
                Checkbox7.Caption = "Write .w files as uncompressed unicode files"
                Checkbox8.Caption = "Automatically create missing DynTrack objects (see TsUtil_en.txt)"
                Checkbox7.value = 1
                Label5.Visible = True
                Label6.Visible = True
                Textbox7.Visible = True
                Textbox8.Visible = True
                Label5.Caption = "ID-Bias for dynamic paths"
                Label6.Caption = "ID-Bias for dynamic sections"
                strOptionCode = "shftdyn"
            Case 23
                Label4.Caption = "The function 'version' gives the current version of all TsUtil classes currently installed."
                strOptionCode = "version"
            Case 24
                Label4.Caption = "The function 'cvrt' is mainly used to convert a route-definition for use with a new tsection.dat" _
                & vbCrLf & "file (similar to the Horace utility). The original tsection.dat file must also be available." _
                & vbCrLf & "See  TsUtil_en.txt for full details of possible commands including use to produce a 'Hidewire' functionality."
                Checkbox6.Caption = "Activate renumbering of UiD numbers in World Files"    ' -r
                Checkbox7.Caption = "Check the new 'tsection.dat' for missing shape files"  ' -s
                Checkbox8.Caption = "Write World files in Compressed format"                ' -c
                Checkbox9.Caption = "Write all World files in uncompressed format"          ' -u
                Checkbox10.Caption = "Check .W files in 'Correction Mode'"                  ' -k
                Checkbox11.Caption = "Manually assign shape nos (enter as OLD:NEW) multiples allowed."  ' -v
                Checkbox6.Visible = True
                Checkbox7.Visible = True
                Checkbox8.Visible = True
                Checkbox9.Visible = True
                Checkbox10.Visible = True
                Checkbox11.Visible = True
                Textbox6.Visible = True
                Button4.Visible = True
                Label7.Caption = "Original tsection.dat(if any)"
                Label7.Visible = True
                Textbox6.Visible = True
                Button4.Visible = True
                If Checkbox11.value = 1 Then
                    Label5.Visible = True
                    Textbox7.Visible = True
                End If
                Text1(5).Visible = True
              
                Rem ************* Carry on from here - not finished ********************
                
                strOptionCode = "cvrt"
                CvrtFlag = 1
            Case 25
                Label4.Caption = "The function 'cvrt' is mainly used to convert a route-definition for use with a new tsection.dat" _
                & vbCrLf & "However, 'cvrt' is multifunctional, so use this option to fix invalid SoundSource positions in a route"
                strOptionCode = "cvrt"
                CvrtFlag = 2
            Case 26
            Label4.Caption = "The function 'cvrt' may also be used to give the functionality of the 'Hidewire' utility" _
                & vbCrLf & "See  TsUtil_en.txt for full details of using TsUtils to build a route with some of the overhead wires hidden." _
                & vbCrLf & "To use this, enter, for example 2:A:Xt in the top box - this results in all TrkObjs with a StaticdetailLevel of 2 (may be 1, 2 or 3) being" _
                & vbCrLf & "converted to their static representation. If the filename starts with 'A' it is replaced with 'Xt', e.g. (A1t10mstrt.s -> Xt1t10mstrt.s)"
                Checkbox6.Caption = "Convert TrkObjs to Static objects with new filenames based on e.g. 1:A:Xt :- "    ' -h
                Checkbox7.Caption = "Restore Static objects back to their original TrkObj status based on:- " ' -H
                Checkbox8.Caption = "Give TrkObjs new filenames, but do not convert to Static - based on:- "  ' -m
                Checkbox9.Caption = "Roll-back above option to revert filenames to their original state."          ' -m
                Checkbox10.Caption = "Change Static obj back to a new TrkObj"                  ' -z
                
                Checkbox6.Visible = True
                Checkbox7.Visible = True
                Checkbox8.Visible = True
                Checkbox9.Visible = True
                Checkbox10.Visible = True
                
                Textbox6.Visible = True
                Button4.Visible = True
'                Label7.Caption = "Original tsection.dat(if any)"
'                Label7.Visible = True
                Textbox6.Visible = True
                Button4.Visible = True
                If Checkbox11.value = 1 Then
                    Label5.Visible = True
                    Textbox7.Visible = True
                End If
                Text1(0).Visible = True
               ' Text1(1).Visible = True
                Text1(2).Visible = True
                Text1(4).Visible = True
            strOptionCode = "cvrt"
                CvrtFlag = 3
        End Select

End Sub


Private Sub Text1_Change(Index As Integer)
If strOptionCode = "fmgr" And Index = 0 Then
Text1(0) = UCase(Text1(0))
If Text1(0) <> "S" And Text1(0) <> "T" And Text1(0) <> "W" Then
Text1(0) = ""
End If
End If
If strOptionCode = "fmgr" And Index = 1 Then
'Text1(1) = UCase(Text1(1))
If Len(Text1(1)) = 2 Then

If UCase(Text1(1)) <> "CB" And UCase(Text1(1)) <> "UT" And UCase(Text1(1)) <> "UB" Then
Text1(1) = ""
Else
Text1(1) = UCase(Text1(1))
End If
End If
End If
If strOptionCode = "cvrt" Then
Select Case Index
Case 0
If Val(Left(Text1(0), 1)) < 1 Or Val(Left(Text1(0), 1)) > 3 Then
Call MsgBox("StaticDetailLevel of TrackObjects must be 1, 2 or 3 for this option to work.", vbExclamation, App.Title)
Exit Sub
End If
Case 2
If Val(Left(Text1(2), 1)) < 4 Or Val(Left(Text1(2), 1)) > 9 And Val(Left(Text1(2), 1)) <> 0 Then
Call MsgBox("StaticDetailLevel of TrackObjects must be 0 or between 4 or 9 for this option to work.", vbExclamation, App.Title)
Exit Sub
End If
Case 4
If Val(Left(Text1(4), 1)) < 1 Or Val(Left(Text1(4), 1)) > 3 Then
Call MsgBox("StaticDetailLevel of TrackObjects must be 1, 2 or 3 for this option to work.", vbExclamation, App.Title)
Exit Sub
End If
End Select
End If
End Sub





