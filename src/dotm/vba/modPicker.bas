Attribute VB_Name = "modPicker"
Option Explicit

' ============================================================================
' 通用列表选择器：frmPicker（Word 风格、可调整大小、可滚动、可勾选）
' 用于文献引用、图/表/公式引用。构建时按 forms.json 创建窗体。
' ============================================================================

#If VBA7 Then
Private Declare PtrSafe Function FindWindowW Lib "user32" ( _
    ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLongW Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLongW Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long
#Else
Private Declare Function FindWindowW Lib "user32" ( _
    ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetWindowLongW Lib "user32" ( _
    ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" ( _
    ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" ( _
    ByVal hWnd As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000

' 窗体尺寸限制（磅）
Public Const PICKER_MIN_WIDTH As Single = 420
Public Const PICKER_MIN_HEIGHT As Single = 300
Public Const PICKER_MAX_WIDTH As Single = 880
Public Const PICKER_MAX_HEIGHT As Single = 640

Private mPickerLastW As Single
Private mPickerLastH As Single
Private mPickerClamping As Boolean

Public Const DIALOG_MARGIN As Single = 12
Public Const DIALOG_BUTTON_WIDTH As Single = 78
Public Const DIALOG_BUTTON_HEIGHT As Single = 23

Public Type PickerItem
    Text As String
    Key As String
End Type

Public gPickerItems() As PickerItem
Public gPickerSelected() As Long
Public gPickerSelectedCount As Long
Public gPickerAccepted As Boolean
Public gPickerTitle As String
Public gPickerPrompt As String

Public Sub ResetPicker()
    Erase gPickerItems
    Erase gPickerSelected
    gPickerSelectedCount = 0
    gPickerAccepted = False
    gPickerTitle = "选择条目"
    gPickerPrompt = "请选择："
End Sub

Public Sub AddPickerItem(ByVal text As String, ByVal key As String)
    Dim n As Long

    n = 0
    On Error Resume Next
    n = UBound(gPickerItems) + 1
    On Error GoTo 0
    ReDim Preserve gPickerItems(n)
    gPickerItems(n).Text = text
    gPickerItems(n).Key = key
End Sub

Public Function PickerItemCount() As Long
    On Error Resume Next
    PickerItemCount = UBound(gPickerItems) + 1
    On Error GoTo 0
End Function

' 返回 True 表示用户点了“插入”；选中项在 gPickerSelected / gPickerSelectedCount
Public Function ShowPicker(ByVal title As String, ByVal prompt As String) As Boolean
    Dim frm As Object

    gPickerTitle = title
    gPickerPrompt = prompt
    gPickerAccepted = False
    gPickerSelectedCount = 0
    Erase gPickerSelected

    On Error Resume Next
    Set frm = VBA.UserForms.Add("frmPicker")
    If Err.Number <> 0 Or frm Is Nothing Then
        Err.Clear
        ShowPicker = False
        Exit Function
    End If
    On Error GoTo 0
    frm.Show
    ShowPicker = gPickerAccepted
End Function

Public Sub LoadPickerForm(frm As Object)
    Dim lst As Object
    Dim i As Long

    frm.Caption = gPickerTitle
    frm.Controls("lblPrompt").Caption = gPickerPrompt
    Set lst = frm.Controls("lstItems")
    lst.Clear
    For i = 0 To PickerItemCount() - 1
        lst.AddItem gPickerItems(i).Text
    Next i
    If lst.ListCount > 0 Then
        lst.ListIndex = 0
        lst.Selected(0) = False
    End If
    UpdatePickerSelection frm
End Sub

Public Sub UpdatePickerSelection(frm As Object)
    Dim lst As Object
    Dim i As Long
    Dim selected As Long

    Set lst = frm.Controls("lstItems")
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then selected = selected + 1
    Next i
    frm.Controls("lblHint").Caption = "已选择 " & CStr(selected) & " 项"
    With frm.Controls("btnOK")
        .Enabled = (selected > 0)
        .Default = .Enabled
    End With
End Sub

Public Sub AcceptPickerForm(frm As Object)
    Dim lst As Object
    Dim i As Long
    Dim n As Long

    Set lst = frm.Controls("lstItems")
    n = 0
    If lst.ListCount = 0 Then
        Erase gPickerSelected
        gPickerSelectedCount = 0
        gPickerAccepted = False
        Exit Sub
    End If
    ReDim gPickerSelected(lst.ListCount - 1)
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            gPickerSelected(n) = i
            n = n + 1
        End If
    Next i
    gPickerSelectedCount = n
    gPickerAccepted = (n > 0)
End Sub

Public Sub ApplyWordDialogStyle(frm As Object)
    Dim c As Object

    frm.BackColor = vbButtonFace
    For Each c In frm.Controls
        c.Font.Name = "Microsoft YaHei UI"
        c.Font.Size = 9
        If TypeName(c) = "Label" Then
            c.ForeColor = vbButtonText
            c.BackStyle = fmBackStyleTransparent
        End If
    Next c
End Sub

Public Sub PickerApplyChrome(frm As Object)
    ApplyWordDialogStyle frm
    With frm.Controls("lstItems")
        .BackColor = vbWindowBackground
        .ForeColor = vbWindowText
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = vbButtonShadow
    End With
End Sub

' ---------------------------------------------------------------------------
' 窗口：加可拉伸边框（WS_THICKFRAME），并把宽高限制在最小/最大之间
' ---------------------------------------------------------------------------
Public Sub PickerMakeResizable(frm As Object)
#If VBA7 Then
    Dim hWnd As LongPtr
#Else
    Dim hWnd As Long
#End If
    Dim style As Long

    On Error Resume Next
    hWnd = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(frm.Caption))
    If hWnd = 0 Then Exit Sub
    style = GetWindowLongW(hWnd, GWL_STYLE)
    If (style And WS_THICKFRAME) = 0 Then
        SetWindowLongW hWnd, GWL_STYLE, (style Or WS_THICKFRAME)
        DrawMenuBar hWnd
    End If
End Sub

Public Sub PickerClampSize(frm As Object)
    If mPickerClamping Then Exit Sub
    mPickerClamping = True
    On Error Resume Next
    If frm.Width > PICKER_MAX_WIDTH Then frm.Width = PICKER_MAX_WIDTH
    If frm.Width < PICKER_MIN_WIDTH Then frm.Width = PICKER_MIN_WIDTH
    If frm.Height > PICKER_MAX_HEIGHT Then frm.Height = PICKER_MAX_HEIGHT
    If frm.Height < PICKER_MIN_HEIGHT Then frm.Height = PICKER_MIN_HEIGHT
    mPickerClamping = False
End Sub

' ---------------------------------------------------------------------------
' 布局：内容区随窗口宽度伸缩；列表高度随窗口高度伸缩；按钮固定在右下角
' ---------------------------------------------------------------------------
Public Sub PickerLayout(frm As Object)
    Dim w As Single
    Dim h As Single
    Dim margin As Single
    Dim btnW As Single
    Dim gap As Single
    Dim btnTop As Single

    On Error Resume Next
    margin = DIALOG_MARGIN
    btnW = DIALOG_BUTTON_WIDTH
    gap = 6
    w = frm.InsideWidth
    h = frm.InsideHeight
    If w < 240 Or h < 180 Then Exit Sub

    btnTop = h - margin - DIALOG_BUTTON_HEIGHT
    frm.Controls("lblPrompt").Move margin, margin, w - 2 * margin, 16
    frm.Controls("lstItems").Move margin, 31, w - 2 * margin, btnTop - 41
    frm.Controls("btnCancel").Move w - margin - btnW, btnTop, btnW, DIALOG_BUTTON_HEIGHT
    frm.Controls("btnOK").Move w - margin - 2 * btnW - gap, btnTop, btnW, DIALOG_BUTTON_HEIGHT
    frm.Controls("lblHint").Move margin, btnTop + 5, _
        w - 3 * margin - 2 * btnW - gap, 16

    mPickerLastW = w
    mPickerLastH = h
End Sub

Public Sub PickerSyncLayout(frm As Object)
    On Error Resume Next
    PickerClampSize frm
    If Abs(frm.InsideWidth - mPickerLastW) > 0.5 Or _
       Abs(frm.InsideHeight - mPickerLastH) > 0.5 Then
        PickerLayout frm
    End If
End Sub

Private Function PickerBoolText(ByVal value As Boolean) As String
    If value Then
        PickerBoolText = "True"
    Else
        PickerBoolText = "False"
    End If
End Function

' 测试入口：选择器窗体能否加载并填充条目
Public Function TestPickerForm() As String
    Dim frm As Object
    Dim lst As Object
    Dim initiallyEnabled As Boolean

    On Error GoTo EH
    ResetPicker
    gPickerTitle = "测试选择器"
    gPickerPrompt = "勾选："
    AddPickerItem "条目一", "1"
    AddPickerItem "条目二", "2"
    Set frm = VBA.UserForms.Add("frmPicker")
    Load frm
    LoadPickerForm frm
    Set lst = frm.Controls("lstItems")
    initiallyEnabled = frm.Controls("btnOK").Enabled
    lst.Selected(1) = True
    UpdatePickerSelection frm
    AcceptPickerForm frm
    TestPickerForm = "ok|" & CStr(lst.ListCount) & "|" & CStr(lst.ListStyle) & "|" & _
        PickerBoolText(initiallyEnabled) & "|" & PickerBoolText(frm.Controls("btnOK").Enabled) & "|" & _
        frm.Controls("lblHint").Caption & "|" & CStr(gPickerSelected(0))
    On Error Resume Next
    Unload frm
    Exit Function
EH:
    TestPickerForm = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：窗口尺寸限制与随窗口变化的布局
Public Function TestPickerLayout() As String
    Dim frm As Object
    Dim wMax As Single
    Dim wMin As Single
    Dim hMax As Single
    Dim listW As Single

    On Error GoTo EH
    ResetPicker
    AddPickerItem "条目一", "1"
    AddPickerItem "条目二", "2"
    Set frm = VBA.UserForms.Add("frmPicker")
    Load frm
    LoadPickerForm frm
    PickerApplyChrome frm

    frm.Width = PICKER_MAX_WIDTH + 300
    PickerSyncLayout frm
    wMax = frm.Width

    frm.Width = PICKER_MIN_WIDTH - 120
    PickerSyncLayout frm
    wMin = frm.Width

    frm.Height = PICKER_MAX_HEIGHT + 300
    PickerSyncLayout frm
    hMax = frm.Height

    frm.Width = 700
    frm.Height = 500
    PickerSyncLayout frm
    listW = frm.Controls("lstItems").Width

    TestPickerLayout = "ok|" & PickerBoolText(wMax <= PICKER_MAX_WIDTH) & "|" & _
        PickerBoolText(wMin >= PICKER_MIN_WIDTH) & "|" & _
        PickerBoolText(hMax <= PICKER_MAX_HEIGHT) & "|" & _
        PickerBoolText(listW > 600)

    On Error Resume Next
    Unload frm
    Exit Function
EH:
    TestPickerLayout = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function
