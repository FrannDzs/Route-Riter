VERSION 5.00
Begin VB.Form frmActive 
   Caption         =   "Active/Inactive Routes"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Select All/None"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   6840
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   6135
      Left            =   360
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select/unselect routes you wish to Activate or Deactivate then click Update."
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   6600
      Width           =   1815
   End
End
Attribute VB_Name = "frmActive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me

End Sub


Private Sub Command2_Click()
Dim i As Integer

For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
If Right$(List1.list(i), 11) = "Train-Store" Then
Call MsgBox(List1.list(i) _
            & vbCrLf & "is stored in Train-Store, you must use Train-Store to activate it." _
            , vbExclamation, App.Title)

GoTo CarryOn
End If

frmUtils.Drive1(0) = Left$(AllRoutes2(i), 1)
frmUtils.Dir1(0).path = AllRoutes2(i)
frmUtils.Text1(0) = "*.trk"
DoEvents
    If frmUtils.File1(0).ListCount = 0 Then
    frmUtils.Text1(0) = "*.off"
    DoEvents
            If frmUtils.File1(0).ListCount = 0 Then
            Call MsgBox("There is no '.trk' or '.off' file in this folder - Route may have been stored" _
                        & vbCrLf & "with Train-Store." _
                        , vbExclamation, App.Title)
            GoTo CarryOn
            Else
            Name frmUtils.Dir1(0).path & "\" & frmUtils.File1(0).list(0) As Left$(frmUtils.Dir1(0).path & "\" & frmUtils.File1(0).list(0), Len(frmUtils.Dir1(0).path & "\" & frmUtils.File1(0).list(0)) - 3) & "trk"
            End If
    
    End If
ElseIf List1.Selected(i) = False Then
frmUtils.Drive1(0) = Left$(AllRoutes2(i), 1)
frmUtils.Dir1(0).path = AllRoutes2(i)
frmUtils.Text1(0) = "*.off"
DoEvents
    If frmUtils.File1(0).ListCount = 0 Then
    frmUtils.Text1(0) = "*.trk"
    DoEvents
            If frmUtils.File1(0).ListCount = 0 Then
            Call MsgBox("There is no '.trk' or '.off' file in " & List1.list(i) _
                        & vbCrLf & "Route may have been stored with Train-Store." _
                        , vbExclamation, App.Title)
            GoTo CarryOn
            Else
            Name frmUtils.Dir1(0).path & "\" & frmUtils.File1(0).list(0) As Left$(frmUtils.Dir1(0).path & "\" & frmUtils.File1(0).list(0), Len(frmUtils.Dir1(0).path & "\" & frmUtils.File1(0).list(0)) - 3) & "off"
            End If
    
    End If
    
End If
CarryOn:
Next i




End Sub

Private Sub Command3_Click()
Dim i As Integer
If List1.Selected(0) = True Then
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next i
Else
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next i
End If
End Sub


