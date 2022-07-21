VERSION 5.00
Object = "{605925BE-4799-4093-A2E7-39323147E70E}#1.0#0"; "C1Query8.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.UserControl C1QueryUI 
   Alignable       =   -1  'True
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   ScaleHeight     =   6165
   ScaleWidth      =   6045
   ToolboxBitmap   =   "C1QueryUI.ctx":0000
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo1"
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "C1QueryUI.ctx":00FA
      Left            =   3120
      List            =   "C1QueryUI.ctx":0104
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22675457
      CurrentDate     =   36864
   End
   Begin VB.PictureBox ctlSplitter 
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   5415
      TabIndex        =   2
      Top             =   2640
      Width           =   5415
   End
   Begin C1Query80Ctl.C1QueryFrame ctlFields 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _cx             =   9551
      _cy             =   4683
      DesignTemplates =   ""
      ManualRender    =   0   'False
      Enabled         =   -1  'True
      DebugContextMenu=   0   'False
      Border          =   -1  'True
      TabInQuery      =   0   'False
      FullFieldNames  =   0   'False
      SchemaControl   =   ""
      ContentsType    =   2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DesignTimeTemplates=   -1  'True
      TypedEditing    =   0   'False
      FormatDate      =   2
      CheckBoxes      =   0   'False
      CheckValues     =   -1  'True
   End
   Begin C1Query80Ctl.C1QueryFrame ctlCond 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   5415
      _cx             =   9551
      _cy             =   5530
      DesignTemplates =   ""
      ManualRender    =   0   'False
      Enabled         =   -1  'True
      DebugContextMenu=   0   'False
      Border          =   -1  'True
      TabInQuery      =   0   'False
      FullFieldNames  =   0   'False
      SchemaControl   =   ""
      ContentsType    =   1
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DesignTimeTemplates=   -1  'True
      TypedEditing    =   0   'False
      FormatDate      =   2
      CheckBoxes      =   0   'False
      CheckValues     =   -1  'True
   End
End
Attribute VB_Name = "C1QueryUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type LookupInfo
    FullFieldName As String
    RowSource As Object
    RowMember As String
    ListField As String
    BoundColumn As String
    ShowListField As Boolean
End Type

Private SchemaControlName As String
Private Initialized As Boolean
Private ConditionsVisible As Boolean
Private FieldsVisible As Boolean
Private SplitActive As Boolean
Private SplitRatioNumber As Single
Private CustomEditor As Object
Private EditorText As String
Private EditorType As Integer
Private EditorSelStart As Integer
Private EditorSelLength As Integer
Private InternalChange As Boolean
Private LookupCount As Integer
Private LookupInfos() As LookupInfo
Private LookupCombosCount As Integer
Private LookupCombos() As Object
Private LookupEditing As Boolean
Private LookupShowListField As Boolean

Event Error(ByVal Source As Object, ByVal ErrorNumber As Long, Description As String, fCancelDisplay As Boolean)
Event ShowCustomEditor(ByVal Field As C1Query80Ctl.IC1QField, ByVal Value As Variant, ByVal ValueText As String, ControlObj As Object, ControlWND As Long)
Event AfterEditing(ByVal Field As C1Query80Ctl.IC1QField, Value As Variant, ValueText As String)

Public Property Get ControlFields() As Object
Attribute ControlFields.VB_Description = "Returns the constituent C1QueryFrame control displaying result fields of the query."
   Set ControlFields = ctlFields
End Property

Public Property Get ControlConditions() As Object
Attribute ControlConditions.VB_Description = "Returns the constituent C1QueryFrame control displaying query conditions."
   Set ControlConditions = ctlCond
End Property

Public Property Get SchemaControl() As String
Attribute SchemaControl.VB_Description = "Returns/sets the name of a C1Query control to which this UI control is attached."
   SchemaControl = SchemaControlName
End Property

Public Property Let SchemaControl(ByVal NewVal As String)
   If UserControl.Ambient.UserMode Then
      Err.Raise Number:=382, Description:="Let/Set not supported at run time."
   End If
   SchemaControlName = NewVal
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in  an object."
    BackColor = ctlCond.BackColor
End Property

Public Property Let BackColor(ByVal NewVal As OLE_COLOR)
    ctlCond.BackColor = NewVal
    ctlFields.BackColor = NewVal
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in  an object."
    ForeColor = ctlCond.ForeColor
End Property

Public Property Let ForeColor(ByVal NewVal As OLE_COLOR)
    ctlCond.ForeColor = NewVal
    ctlFields.ForeColor = NewVal
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
    Set Font = ctlCond.Font
End Property

Public Property Set Font(ByVal NewVal As Font)
    Set ctlCond.Font = NewVal
    Set ctlFields.Font = NewVal
End Property

Public Property Get VisibleFields() As Boolean
Attribute VisibleFields.VB_Description = "Returns/sets a value that determines whether the fields control is visible."
   VisibleFields = FieldsVisible
End Property

Public Property Let VisibleFields(ByVal NewVal As Boolean)
   Dim Change As Boolean
   Change = FieldsVisible <> NewVal
   FieldsVisible = NewVal
   ctlFields.Visible = NewVal
   If Change Then
      ctlSplitter.Visible = Not ctlFields.Visible Or ctlCond.Visible
      UserControl_Resize
   End If
End Property

Public Property Get VisibleConditions() As Boolean
Attribute VisibleConditions.VB_Description = "Returns/sets a value that determines whether the conditions control is visible."
   VisibleConditions = ConditionsVisible
End Property

Public Property Let VisibleConditions(ByVal NewVal As Boolean)
   Dim Change As Boolean
   Change = ConditionsVisible <> NewVal
   ctlCond.Visible = NewVal
   ConditionsVisible = NewVal
   If Change Then
      ctlSplitter.Visible = Not ctlFields.Visible Or ctlCond.Visible
      UserControl_Resize
   End If
End Property

Public Property Get FullFieldNames() As Boolean
Attribute FullFieldNames.VB_Description = "Returns/sets a value that determines whether to show field names with full dot-separated folder path."
   FullFieldNames = ctlCond.FullFieldNames
End Property

Public Property Let FullFieldNames(ByVal NewVal As Boolean)
   ctlCond.FullFieldNames = NewVal
   ctlFields.FullFieldNames = NewVal
End Property

Public Property Get SplitRatio() As Single
Attribute SplitRatio.VB_Description = "Returns/sets the ratio of the fields control height relative to the overall control height."
   If UserControl.Height < 0 Then SplitRatio = 100
   SplitRatio = SplitRatioNumber
End Property

Public Property Let SplitRatio(ByVal NewVal As Single)
   ctlSplitter.Top = ctlFields.Top + ((UserControl.Height - ctlSplitter.Height) * NewVal)
   SplitRatioNumber = NewVal
   If Initialized Then UserControl_Resize
End Property

Public Sub SetLookup(ByVal FullFieldName As String, ByVal RowSource As Object, ByVal RowMember As String, _
        ByVal ListField As String, ByVal BoundColumn As String, ByVal ShowListField As Boolean)
    ReDim Preserve LookupInfos(LookupCount)
    If BoundColumn = "" Then BoundColumn = ListField
    If ListField = "" Then ListField = BoundColumn
    LookupInfos(LookupCount).FullFieldName = FullFieldName
    Set LookupInfos(LookupCount).RowSource = RowSource
    LookupInfos(LookupCount).RowMember = RowMember
    LookupInfos(LookupCount).ListField = ListField
    LookupInfos(LookupCount).BoundColumn = BoundColumn
    LookupInfos(LookupCount).ShowListField = ShowListField
    LookupCount = LookupCount + 1
End Sub

Private Sub UserControl_Initialize()
    ConditionsVisible = True
    FieldsVisible = True
    SplitRatioNumber = 0.5
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   SchemaControlName = PropBag.ReadProperty("SchemaControl", "")
   ctlCond.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
   ctlFields.BackColor = ctlCond.BackColor
   ctlCond.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
   ctlFields.ForeColor = ctlCond.ForeColor
   Set Font = PropBag.ReadProperty("Font", Ambient.Font)
   VisibleFields = PropBag.ReadProperty("VisibleFields")
   VisibleConditions = PropBag.ReadProperty("VisibleConditions")
   FullFieldNames = PropBag.ReadProperty("FullFieldNames", True)
   SplitRatio = PropBag.ReadProperty("SplitRatio")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SchemaControl", SchemaControlName, "")
    Call PropBag.WriteProperty("BackColor", ctlCond.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", ctlCond.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("VisibleFields", VisibleFields)
    Call PropBag.WriteProperty("VisibleConditions", VisibleConditions)
    Call PropBag.WriteProperty("FullFieldNames", FullFieldNames, True)
    Call PropBag.WriteProperty("SplitRatio", SplitRatio)
End Sub

Private Sub UserControl_Show()
    If Not Initialized Then
        Initialized = True
        Dim Ctrl As Object
        For Each Ctrl In UserControl.Extender.Parent.Controls
           If Ctrl.Name = SchemaControlName Then
              ctlCond.SchemaControlObject = Ctrl
              ctlFields.SchemaControlObject = Ctrl
              Exit For
           End If
        Next Ctrl
        If Not UserControl.Ambient.UserMode Then
            ctlCond.DesignMode = 1
            ctlFields.DesignMode = 1
        End If
    End If
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
   On Error Resume Next
   If ctlSplitter.Top < 0 Then
      ctlSplitter.Top = 0
   End If
   If ctlSplitter.Top > UserControl.Height - ctlSplitter.Height - 10 Then
      ctlSplitter.Top = UserControl.Height - ctlSplitter.Height - 10
   End If
   ctlFields.Width = UserControl.Width
   ctlCond.Width = UserControl.Width
   ctlSplitter.Width = UserControl.Width
   ctlFields.Height = ctlSplitter.Top - ctlFields.Top
   ctlCond.Top = ctlSplitter.Top + ctlSplitter.Height
   ctlCond.Height = UserControl.Height - ctlCond.Top
End Sub

Private Sub ctlSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.MousePointer = vbSizeNS
    SplitActive = True
End Sub

Private Sub ctlSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ctlSplitter_MouseMove(Button, Shift, X, Y)
    SplitActive = False
    UserControl.MousePointer = vbNormal
End Sub

Private Sub ctlSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.MousePointer = vbSizeNS
    If SplitActive Then
        ctlSplitter.Top = ctlSplitter.Top + Y
        UserControl_Resize
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.MousePointer = vbNormal
End Sub

Private Sub ctlCond_Error(ByVal ErrorNumber As Long, Description As String, CancelDisplay As Boolean)
    RaiseEvent Error(ctlCond, ErrorNumber, Description, CancelDisplay)
End Sub

Private Sub ctlFields_Error(ByVal ErrorNumber As Long, Description As String, CancelDisplay As Boolean)
    RaiseEvent Error(ctlFields, ErrorNumber, Description, CancelDisplay)
End Sub


Private Sub ctlCond_BeforeEdit(ByVal Field As C1Query80Ctl.IC1QField, ByVal Value As Variant, ValueText As String, ModalEditing As Boolean, ShowCustomEditor As Boolean, TypedEditing As Boolean, ControlWidth As Double, ControlHeight As Double)
   Dim I As Integer
   If LookupCount <> 0 Then
        For I = LBound(LookupInfos) To UBound(LookupInfos)
             If LookupInfos(I).FullFieldName = Field.FullFolderName Then
                 If LookupCombosCount <= I Then
                     ReDim Preserve LookupCombos(I)
                     LookupCombosCount = I + 1
                 End If
                 If LookupCombos(I) Is Nothing Then
                     Load DataCombo1(I + 1)
                     Set LookupCombos(I) = DataCombo1(I + 1)
                     LookupCombos(I).RowMember = LookupInfos(I).RowMember
                     LookupCombos(I).ListField = LookupInfos(I).ListField
                     LookupCombos(I).BoundColumn = LookupInfos(I).BoundColumn
                     Set LookupCombos(I).RowSource = LookupInfos(I).RowSource
                 End If
                 ShowCustomEditor = True
                 ControlWidth = LookupCombos(I).Width
                 ControlHeight = LookupCombos(I).Height
                 Set CustomEditor = LookupCombos(I)
                 LookupShowListField = LookupInfos(I).ShowListField
                 LookupEditing = True
                 Exit Sub
             End If
        Next I
   End If
   EditorType = ctlCond.VarTypeFromDataType(Field.Type)
   Select Case EditorType
      Case 16, 17 'VT_I1, VT_UI1
         EditorType = vbByte
      Case 2, 18 'VT_I2, VT_UI2
         EditorType = vbInteger
      Case 3, 19 'VT_I4, VT_UI4
         EditorType = vbLong
      Case 4, 5 'VT_R4, VT_R8
         EditorType = vbDouble
   End Select
   If EditorType = vbDate Then
      ShowCustomEditor = True
      ControlWidth = DTPicker1.Width
      ControlHeight = DTPicker1.Height
      Set CustomEditor = DTPicker1
      Exit Sub
   End If
   If EditorType = vbBoolean Then
      ShowCustomEditor = True
      ControlWidth = Combo1.Width
      ControlHeight = Combo1.Height
      Set CustomEditor = Combo1
      Exit Sub
   End If
   If EditorType = vbByte Or EditorType = vbInteger Or EditorType = vbLong Or EditorType = vbDouble Or EditorType = vbCurrency Then
      ShowCustomEditor = True
      ControlWidth = Text1.Width
      ControlHeight = Text1.Height
      Set CustomEditor = Text1
      Exit Sub
   End If
End Sub

Private Sub ctlCond_ShowCustomEditor(ByVal Field As C1Query80Ctl.IC1QField, ByVal Value As Variant, ByVal ValueText As String, ControlObj As Object, ControlWND As Long)
   If CustomEditor Is DTPicker1 Then
      Set DTPicker1.Font = Font
      DTPicker1.Visible = True
      DTPicker1.SetFocus
      Set ControlObj = DTPicker1
      ControlWND = DTPicker1.hWnd
      DTPicker1.Value = Value
      If Field.Type = c1qTypeTime Then
         DTPicker1.Format = dtpTime
      Else
         DTPicker1.Format = dtpShortDate
      End If
   ElseIf CustomEditor Is Combo1 Then
      Set Combo1.Font = Font
      Combo1.Visible = True
      Combo1.SetFocus
      Set ControlObj = Combo1
      ControlWND = Combo1.hWnd
      If Value Then
         Combo1.ListIndex = 0
      Else
         Combo1.ListIndex = 1
      End If
   ElseIf CustomEditor Is Text1 Then
      Set Text1.Font = Font
      Text1.Visible = True
      Text1.SetFocus
      Set ControlObj = Text1
      ControlWND = Text1.hWnd
      Text1.Text = ValueText
      EditorText = ValueText
      EditorSelStart = Text1.SelStart
      EditorSelLength = Text1.SelLength
   ElseIf LookupEditing Then
      Set CustomEditor.Font = Font
      CustomEditor.Visible = True
      CustomEditor.SetFocus
      Set ControlObj = CustomEditor
      ControlWND = CustomEditor.hWnd
      CustomEditor.BoundText = Value
   End If
   RaiseEvent ShowCustomEditor(Field, Value, ValueText, ControlObj, ControlWND)
End Sub

Private Sub ctlCond_EndEdit(ByVal Field As C1Query80Ctl.IC1QField, Value As Variant, ValueText As String)
    If CustomEditor Is DTPicker1 Then
       Value = DTPicker1.Value
       ValueText = CStr(DTPicker1.Value)
    ElseIf CustomEditor Is Combo1 Then
       ValueText = Combo1.Text
       Value = ValueText = "Yes"
    ElseIf CustomEditor Is Text1 Then
       ValueText = Text1.Text
       If ValueText = "" Then
          ValueText = "0"
       End If
       Value = ValueText
       Dim FmtStr As String
       If GetFormattedNumber(FmtStr, ValueText) Then ValueText = FmtStr
    ElseIf LookupEditing Then
       If LookupShowListField Then
          If CustomEditor.Text <> "" Then
             ValueText = CustomEditor.Text
             Value = CustomEditor.BoundText
          End If
       Else
          If CustomEditor.BoundText <> "" Then
             ValueText = CustomEditor.BoundText
             Value = CustomEditor.BoundText
          End If
       End If
    End If
    RaiseEvent AfterEditing(Field, Value, ValueText)
End Sub

Private Sub ctlCond_HideCustomEditor(ByVal Field As C1Query80Ctl.IC1QField, ByVal ControlObj As Object, ByVal ControlWND As Long)
   CustomEditor.Visible = False
   Set CustomEditor = Nothing
   LookupEditing = False
End Sub

Private Sub Text1_Change()
    If InternalChange Then Exit Sub
    Dim OK As Boolean
    OK = True
    If Text1.Text <> "" Then
       Dim SelStart As Integer, SelLength As Integer
       Dim Str As String
       Dim FmtStr As String
       Str = Text1.Text
       OK = GetFormattedNumber(FmtStr, Str)
       If Not OK Then
          Str = Str + "0"
       End If
       OK = GetFormattedNumber(FmtStr, Str)
    End If
    If OK Then
       EditorText = Text1.Text
       EditorSelStart = Text1.SelStart
       EditorSelLength = Text1.SelLength
    Else
       On Error GoTo Err
       InternalChange = True
       Text1.Text = EditorText
       Text1.SelStart = EditorSelStart
       Text1.SelLength = EditorSelLength
Err:
       InternalChange = False
    End If
End Sub

Private Sub Text1_Click()
    EditorSelStart = Text1.SelStart
    EditorSelLength = Text1.SelLength
End Sub

Private Function GetFormattedNumber(DstStr As String, ByVal SrcStr As String) As Boolean
    GetFormattedNumber = True
    DstStr = ""
    If SrcStr = "" Then Exit Function
    DstStr = "0"
    Dim V As Variant
    On Error GoTo Err
    Select Case EditorType
       Case vbByte
          V = CByte(SrcStr)
       Case vbInteger
          V = CInt(SrcStr)
       Case vbLong
          V = CLng(SrcStr)
       Case vbDouble
          V = CDbl(SrcStr)
       Case vbCurrency
          V = CCur(SrcStr)
    End Select
    DstStr = CStr(V)
    Exit Function
Err:
    GetFormattedNumber = False
End Function


