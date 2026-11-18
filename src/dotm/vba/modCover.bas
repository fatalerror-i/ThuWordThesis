Attribute VB_Name = "modCover"
Option Explicit

' ============================================================================
' 封面：原生内容控件与级联刷新
' 封面页用内容控件承载数据，不再使用自定义窗体：
'   - 学位类别：下拉列表（按配置列明层次的学位选项，来自 modDegreeOptions）
'   - 学科（专业学位改为专业领域）：可输入组合框，候选随学位类别刷新
'   - 成文日期：文本框，"2026-10" 自动转为中英文日期
' 学位选项确认后由 ThisDocument 的选区事件同步，退出控件与保存时补充刷新；
' 成文日期在退出控件时同步，功能区“刷新封面”按钮可手动重跑。
' 副/联合导师行由功能区按钮增删，作者信息块始终保持 6 段（6×26 磅）。
' ============================================================================

Private Const STYLE_COVER_INFO As String = "封面作者信息"

Public Const TAG_TITLE_CN As String = "论文题目"
Private Const TAG_DEGREE As String = "学位类别"
Private Const TAG_DEPARTMENT As String = "培养单位"
Private Const TAG_DISCIPLINE As String = "学科（专业学位改为专业领域）"
Private Const TAG_AUTHOR_CN As String = "研究生（专业学位改为申请人）"
Private Const TAG_SUPERVISOR_CN As String = "指导教师"
Private Const TAG_ASSOC_SUP As String = "副指导教师（可选）"
Private Const TAG_CO_SUP As String = "联合指导教师（可选）"
Private Const TAG_DATE_CN As String = "成文日期"
Private Const TAG_TITLE_EN As String = "英文论文题目"
Private Const TAG_SUBMIT_EN As String = "英文提交声明"
Private Const TAG_DEGREE_LINE_EN As String = "英文degree行"
Private Const TAG_DEGREE_NAME_EN As String = "英文学位名称"
Private Const TAG_DISCIPLINE_EN As String = "英文学科"
Private Const TAG_FIELD_EN As String = "英文专业领域"
Private Const TAG_AUTHOR_EN As String = "英文作者"
Private Const TAG_SUPERVISOR_EN As String = "指导教师（英文）"
Private Const TAG_DATE_EN As String = "英文成文日期"
Private Const TAG_SECURITY As String = "密级"

Private Const LABEL_DISCIPLINE As String = "学科"
Private Const LABEL_DISCIPLINE_PRO As String = "专业领域"
Private Const LABEL_AUTHOR As String = "研究生"
Private Const LABEL_AUTHOR_PRO As String = "申请人"
Private Const LABEL_ASSOC_SUP As String = "副指导教师"
Private Const LABEL_CO_SUP As String = "联合指导教师"

' 密级开关添加时的默认值，用户可直接修改。
Private Const SECURITY_DEFAULT As String = "秘密★5年"

' 文档事件与功能区回调共用同一把锁，重建控件时不允许再次进入级联。
Private mUpdatingCover As Boolean

Public Function CoverUpdateInProgress() As Boolean
    CoverUpdateInProgress = mUpdatingCover
End Function

' 选项确认、文档切换及保存时补齐派生字段；相同学位不重复改写文档。
Public Sub RefreshDegreeIfChanged()
    Dim cc As ContentControl
    If mUpdatingCover Then Exit Sub
    Set cc = FindCC(TAG_DEGREE)
    If cc Is Nothing Then Exit Sub
    If RecallCoverText("THU_DEGREE_APPLIED", "") <> Trim$(cc.Range.Text) Then
        ApplyDegreeCascade cc
    End If
End Sub

' ============================================================================
' 级联入口
' ============================================================================

' 学位类别变化：刷新英文提交声明、英文 degree 行、英文学位名称、
' 英文导师标题、中文标签（学科/专业领域、研究生/申请人）与学科候选。
Public Sub ApplyDegreeCascade(ByVal cc As ContentControl)
    Dim level As String
    Dim degreeType As String
    Dim category As String
    Dim degreeName As String
    Dim isDoctor As Boolean
    Dim isProfessional As Boolean
    Dim submitLine As String
    Dim degreeLine As String
    Dim hasDiscipline As Boolean
    Dim oldScreenUpdating As Boolean
    Dim originalSelection As Range
    Dim originalTag As String

    On Error GoTo Failed
    oldScreenUpdating = Application.ScreenUpdating
    If mUpdatingCover Then Exit Sub
    If cc Is Nothing Then Exit Sub
    If Not ParseDegreeText(cc.Range.Text, level, degreeType, category) Then Exit Sub
    degreeName = modDegreeOptions.CoverDegreeEnglish(category, level)
    If Len(degreeName) = 0 Then Exit Sub

    isDoctor = (level = "博士")
    If Len(degreeType) > 0 Then
        isProfessional = (degreeType = "专业学位")
    Else
        isProfessional = modDegreeOptions.CoverIsProfessionalCategory(category)
    End If
    If isDoctor Then
        submitLine = "Dissertation submitted to"
    Else
        submitLine = "Thesis submitted to"
    End If
    If isProfessional Then
        degreeLine = "for the professional degree of"
    Else
        degreeLine = "for the degree of"
    End If

    Set originalSelection = Selection.Range.Duplicate
    originalTag = SelectedCoverTag()
    mUpdatingCover = True
    Application.ScreenUpdating = False
    WriteCC TAG_SUBMIT_EN, submitLine
    WriteCC TAG_DEGREE_LINE_EN, degreeLine
    If WriteCC(TAG_DEGREE_NAME_EN, degreeName) Then
        FormatCC TAG_DEGREE_NAME_EN, "Arial", 16, True
    End If
    UpdateSupervisorTitle isDoctor
    SyncEnglishCoSupervisor
    hasDiscipline = (Not isProfessional) Or modDegreeOptions.CoverHasProfessionalField(category)
    ' 硕博切换通常不改变行结构，保留原控件及其中的作者数据。
    If hasDiscipline <> (Not (FindCC(TAG_DISCIPLINE) Is Nothing)) Then
        RebuildCoverAuthorInfo hasDiscipline
    End If
    SetCoverLabels isProfessional
    SyncEnglishDiscipline isProfessional, modDegreeOptions.CoverHasProfessionalField(category)
    RefreshDisciplineList category, degreeType
    ApplyLabelFitWidths
    RememberCoverText "THU_DEGREE_APPLIED", Trim$(cc.Range.Text)
    RestoreCoverSelection originalSelection, originalTag
    mUpdatingCover = False
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

Failed:
    mUpdatingCover = False
    Application.ScreenUpdating = oldScreenUpdating
    ReportCascadeError Err.Number, Err.Description
End Sub

' 成文日期变化：把 "2026-10" 写成中文日期，并同步英文日期；
' 已经写成中文日期的（含“年”）不再改动英文日期，避免覆盖手填值。
Public Sub ApplyDateCascade(ByVal cc As ContentControl)
    Dim s As String
    Dim cn As String
    Dim en As String

    On Error GoTo Failed
    If cc Is Nothing Then Exit Sub
    s = Trim$(cc.Range.Text)
    If Len(s) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    cn = ChineseDate(s)
    If Len(cn) > 0 Then
        If cn <> s Then WriteCC TAG_DATE_CN, cn
    End If
    If InStr(s, "年") = 0 Then
        en = EnglishDate(s)
        If Len(en) > 0 Then
            If en <> s Then
                If WriteCC(TAG_DATE_EN, en) Then FormatCC TAG_DATE_EN, "Arial", 16, True
            End If
        End If
    End If
    Application.ScreenUpdating = True
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    ReportCascadeError Err.Number, Err.Description
End Sub

' 功能区“刷新封面”：按当前内容控件值重跑学位与日期级联。
Public Sub RefreshCover(ByVal control As IRibbonControl)
    Dim cc As ContentControl

    On Error GoTo Failed
    If Not EnsureReady() Then Exit Sub

    Application.ScreenUpdating = False
    Set cc = FindCC(TAG_DEGREE)
    If Not cc Is Nothing Then ApplyDegreeCascade cc
    Set cc = FindCC(TAG_DATE_CN)
    If Not cc Is Nothing Then ApplyDateCascade cc
    SyncEnglishCoSupervisor
    Application.ScreenUpdating = True
    Application.StatusBar = "封面已刷新。"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    ReportCascadeError Err.Number, Err.Description
End Sub

' 级联过程出错时只写状态栏，不用 MsgBox：避免在内容控件退出事件里阻塞编辑。
Public Sub ReportCascadeError(ByVal errNo As Long, ByVal errDesc As String)
    On Error Resume Next
    Application.StatusBar = "封面自动更新失败：" & errDesc
End Sub

' ============================================================================
' 学位解析
' ============================================================================

' 解析学位下拉的显示文本：
'   "工学硕士学位"       -> level=硕士, dtype=学术学位, category=工学
'   "金融硕士专业学位"   -> level=硕士, dtype=专业学位, category=金融
' 只有 modDegreeOptions 能给出英文名时才认为解析成功。
Private Function ParseDegreeText(ByVal text As String, ByRef level As String, _
                                 ByRef degreeType As String, ByRef category As String) As Boolean
    Dim s As String

    s = Trim$(text)
    If Len(s) = 0 Then Exit Function
    If Right$(s, 4) = "专业学位" Then
        degreeType = "专业学位"
        s = Left$(s, Len(s) - 4)
    ElseIf Right$(s, 2) = "学位" Then
        degreeType = ""
        s = Left$(s, Len(s) - 2)
    Else
        Exit Function
    End If
    If Len(s) < 3 Then Exit Function
    level = Right$(s, 2)
    If level <> "硕士" And level <> "博士" Then Exit Function
    category = Left$(s, Len(s) - 2)
    If Len(modDegreeOptions.CoverDegreeEnglish(category, level)) = 0 Then Exit Function
    ParseDegreeText = True
End Function

' 硕士/博士切换英文导师标题：Thesis Supervisor: <-> Dissertation Supervisor:
Private Sub UpdateSupervisorTitle(ByVal isDoctor As Boolean)
    Dim cc As ContentControl
    Dim s As String
    Dim findText As String
    Dim replText As String
    Dim pos As Long
    Dim r As Range

    Set cc = FindCC(TAG_SUPERVISOR_EN)
    If cc Is Nothing Then Exit Sub
    s = cc.Range.Text
    If isDoctor Then
        findText = "Thesis Supervisor:"
        replText = "Dissertation Supervisor:"
    Else
        findText = "Dissertation Supervisor:"
        replText = "Thesis Supervisor:"
    End If
    pos = InStr(1, s, findText, vbTextCompare)
    If pos = 0 Then Exit Sub
    Set r = cc.Range.Duplicate
    r.Start = r.Start + pos - 1
    r.End = r.Start + Len(findText)
    r.Text = replText
End Sub

' ============================================================================
' 学科候选
' ============================================================================

' 按学位类别重建学科组合框候选；保留用户已填写的文本。
Private Sub RefreshDisciplineList(ByVal category As String, ByVal degreeType As String)
    Dim cc As ContentControl
    Dim items As Variant
    Dim current As String
    Dim isProfessional As Boolean
    Dim item As Variant

    On Error GoTo Failed
    Set cc = FindCC(TAG_DISCIPLINE)
    If cc Is Nothing Then Exit Sub
    current = ReadCC(TAG_DISCIPLINE)
    If Len(degreeType) > 0 Then
        isProfessional = (degreeType = "专业学位")
    Else
        isProfessional = modDegreeOptions.CoverIsProfessionalCategory(category)
    End If
    If isProfessional Then
        items = modDegreeOptions.CoverDisciplineOptions(category, "专业学位")
    Else
        items = modDegreeOptions.CoverDisciplineOptions(category, "学术学位")
    End If

    cc.DropdownListEntries.Clear
    If IsArray(items) Then
        For Each item In items
            If Len(CStr(item)) > 0 Then
                cc.DropdownListEntries.Add Text:=CStr(item), Value:=CStr(item)
            End If
        Next item
    End If
    If Len(current) > 0 Then cc.Range.Text = current
    Exit Sub

Failed:
    ReportCascadeError Err.Number, Err.Description
End Sub

' ============================================================================
' 中文标签：学科/专业领域、研究生/申请人
' ============================================================================

Private Sub SetCoverLabels(ByVal professional As Boolean)
    If professional Then
        ReplaceParagraphLabel LABEL_DISCIPLINE, LABEL_DISCIPLINE_PRO
        ReplaceParagraphLabel LABEL_AUTHOR, LABEL_AUTHOR_PRO
    Else
        ReplaceParagraphLabel LABEL_DISCIPLINE_PRO, LABEL_DISCIPLINE
        ReplaceParagraphLabel LABEL_AUTHOR_PRO, LABEL_AUTHOR
    End If
End Sub

Private Sub ReplaceParagraphLabel(ByVal findLabel As String, ByVal replLabel As String)
    Dim cell As Cell
    Dim p As Paragraph
    Dim r As Range
    Dim i As Long

    Set cell = AuthorInfoCell()
    If cell Is Nothing Then Exit Sub
    For i = cell.Range.Paragraphs.Count To 1 Step -1
        Set p = cell.Range.Paragraphs(i)
        If ParagraphLabel(p) = findLabel Then
            Set r = p.Range.Duplicate
            r.Start = r.Start + 1
            r.End = r.Start + Len(findLabel)
            r.Text = replLabel
        End If
    Next i
End Sub

' 取作者信息段落的标签文字（去掉前导制表符和冒号后的内容）。
Private Function ParagraphLabel(ByVal p As Paragraph) As String
    Dim s As String
    Dim pos As Long

    s = Replace(Replace(p.Range.Text, vbCr, ""), Chr$(13), "")
    If Left$(s, 1) = vbTab Then s = Mid$(s, 2)
    pos = InStr(s, "：")
    If pos > 1 Then ParagraphLabel = Left$(s, pos - 1)
End Function

' 统一作者信息块所有标签的 FitText 宽度：取最长标签字数 × 16 磅
' （4 字=1280 twips，5 字=1600，6 字=1920），保证冒号对齐。
Private Sub ApplyLabelFitWidths()
    Dim cell As Cell
    Dim p As Paragraph
    Dim label As String
    Dim r As Range
    Dim maxChars As Long
    Dim i As Long

    Set cell = AuthorInfoCell()
    If cell Is Nothing Then Exit Sub
    maxChars = 4
    For i = 1 To cell.Range.Paragraphs.Count
        Set p = cell.Range.Paragraphs(i)
        label = ParagraphLabel(p)
        If Len(label) > maxChars Then maxChars = Len(label)
    Next i
    For i = 1 To cell.Range.Paragraphs.Count
        Set p = cell.Range.Paragraphs(i)
        label = ParagraphLabel(p)
        If Len(label) > 0 Then
            Set r = p.Range.Duplicate
            r.Start = r.Start + 1
            r.End = r.Start + Len(label)
            r.FitTextWidth = maxChars * 16
        End If
    Next i
End Sub

' ============================================================================
' 副/联合导师行：功能区按钮增删，不弹窗
' ============================================================================

Public Sub ToggleAssociateSupervisor(ByVal control As IRibbonControl)
    ToggleSupervisorRow TAG_ASSOC_SUP, LABEL_ASSOC_SUP
End Sub

Public Sub ToggleCoSupervisor(ByVal control As IRibbonControl)
    ToggleSupervisorRow TAG_CO_SUP, LABEL_CO_SUP
End Sub

Private Sub ToggleSupervisorRow(ByVal tagName As String, ByVal label As String)
    Dim exists As Boolean
    Dim existing As ContentControl
    Dim oldScreenUpdating As Boolean
    Dim originalSelection As Range
    Dim originalTag As String

    On Error GoTo Failed
    oldScreenUpdating = Application.ScreenUpdating
    If mUpdatingCover Then Exit Sub
    If Not EnsureReady() Then Exit Sub
    RefreshDegreeIfChanged
    Set originalSelection = Selection.Range.Duplicate
    originalTag = SelectedCoverTag()
    mUpdatingCover = True
    Set existing = FindCC(tagName)
    exists = Not (existing Is Nothing)

    Application.ScreenUpdating = False
    If exists Then
        RemoveSupervisorRow tagName
    Else
        InsertSupervisorRow tagName, label
    End If
    RebuildOptionalSupervisorRows
    SyncEnglishCoSupervisor
    ApplyLabelFitWidths
    RestoreCoverSelection originalSelection, originalTag
    mUpdatingCover = False
    Application.ScreenUpdating = oldScreenUpdating

    If exists Then
        Application.StatusBar = "已删除" & label & "行。"
    Else
        Application.StatusBar = "已添加" & label & "行，请直接填写。"
    End If
    Exit Sub

Failed:
    mUpdatingCover = False
    Application.ScreenUpdating = oldScreenUpdating
    ' 选填信息按钮只写状态栏，不弹出错误窗口。
    Application.StatusBar = "封面导师行操作失败：" & Err.Description
End Sub

' 中英文副导师、联合导师分别对应 Associate Supervisor、Co-supervisor。
' 保留用户已输入的英文姓名；删除一行时不覆盖其他导师行。
Private Sub SyncEnglishCoSupervisor()
    Dim cc As ContentControl
    Dim parts As Variant
    Dim part As Variant
    Dim base As String
    Dim assoc As String
    Dim co As String
    Dim newText As String
    Set cc = FindCC(TAG_SUPERVISOR_EN)
    If cc Is Nothing Then Exit Sub
    parts = Split(cc.Range.Text, Chr$(11))
    base = CStr(parts(0))
    If InStr(1, base, "Supervisor:", vbTextCompare) = 0 Then Exit Sub
    For Each part In parts
        If InStr(1, part, "Associate Supervisor:", vbTextCompare) > 0 Then assoc = CStr(part)
        If InStr(1, part, "Co-supervisor:", vbTextCompare) > 0 Then co = CStr(part)
    Next part
    newText = base
    If Not FindCC(TAG_ASSOC_SUP) Is Nothing Then
        If Len(assoc) = 0 Then assoc = vbTab & "Associate Supervisor:" & vbTab & "Professor ×××"
        newText = newText & Chr$(11) & assoc
    End If
    If Not FindCC(TAG_CO_SUP) Is Nothing Then
        If Len(co) = 0 Then co = vbTab & "Co-supervisor:" & vbTab & "Professor ×××"
        newText = newText & Chr$(11) & co
    End If
    If newText <> cc.Range.Text Then
        cc.Range.Text = newText
        FormatEnglishSupervisor
    End If
End Sub

Private Sub RememberCoverText(ByVal key As String, ByVal value As String)
    On Error Resume Next
    ActiveDocument.Variables(key).Value = value
    If Err.Number <> 0 Then
        Err.Clear
        ActiveDocument.Variables.Add key, value
    End If
    On Error GoTo 0
End Sub

Private Function RecallCoverText(ByVal key As String, ByVal fallback As String) As String
    RecallCoverText = fallback
    On Error Resume Next
    RecallCoverText = ActiveDocument.Variables(key).Value
    On Error GoTo 0
End Function

Private Sub SyncEnglishDiscipline(ByVal professional As Boolean, ByVal hasField As Boolean)
    Dim value As String
    Dim cc As ContentControl
    Dim prefix As Paragraph
    Dim r As Range
    Set cc = FindCC(TAG_DISCIPLINE_EN)
    If Not cc Is Nothing Then
        Set prefix = cc.Range.Paragraphs(1).Previous
        If Not prefix Is Nothing Then
            Set r = ParagraphContentRange(prefix)
            ' in 是独立段落中的普通文字，专业学位隐藏文字但保留段落高度。
            If Trim$(r.Text) = "in" Or Len(Trim$(r.Text)) = 0 Then
                If professional Then
                    r.Text = ""
                Else
                    r.Text = "in"
                End If
            End If
        End If
    End If
    value = ReadCC(TAG_DISCIPLINE_EN)
    If professional Then
        If Len(Trim$(value)) > 0 Then RememberCoverText "THU_ACADEMIC_EN", value
        WriteCC TAG_DISCIPLINE_EN, ""
    ElseIf Len(Trim$(value)) = 0 Then
        WriteCC TAG_DISCIPLINE_EN, RecallCoverText("THU_ACADEMIC_EN", "Field of Study")
    End If
    value = ReadCC(TAG_FIELD_EN)
    If professional And hasField Then
        If Len(Trim$(value)) = 0 Then WriteCC TAG_FIELD_EN, RecallCoverText("THU_FIELD_EN", "(Field of Study)")
    Else
        If Len(Trim$(value)) > 0 Then RememberCoverText "THU_FIELD_EN", value
        WriteCC TAG_FIELD_EN, ""
    End If
End Sub

' ============================================================================
' 密级开关：按内容控件是否存在切换。无控件 → 插入控件并写默认密级；
' 有控件且已有内容 → 删除控件并清空该行文字（密级行保留）；有控件但为空 → 视为添加。
' ============================================================================

Public Sub ToggleSecurity(ByVal control As IRibbonControl)
    Dim cc As ContentControl
    Dim p As Paragraph

    On Error GoTo Failed
    If Not EnsureReady() Then Exit Sub
    Set cc = FindCC(TAG_SECURITY)
    Application.ScreenUpdating = False

    If cc Is Nothing Then
        InsertSecurityControl
        SyncPublicDeclaration
        Application.ScreenUpdating = True
        Application.StatusBar = "已添加密级信息，请直接修改。"
        Exit Sub
    End If

    If Len(Trim$(cc.Range.Text)) = 0 Then
        cc.Range.Text = SECURITY_DEFAULT
        SyncPublicDeclaration
        Application.ScreenUpdating = True
        Application.StatusBar = "已添加密级信息，请直接修改。"
        Exit Sub
    End If

    ' 删除控件并清空密级段落，保留表格行或页眉及其格式。
    Set p = cc.Range.Paragraphs(1)
    cc.Delete False
    ClearSecurityParagraph p
    SyncPublicDeclaration
    Application.ScreenUpdating = True
    Application.StatusBar = "已删除密级信息。"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    ' 选填信息按钮只写状态栏，不弹出错误窗口。
    Application.StatusBar = "封面密级操作失败：" & Err.Description
End Sub

' 封面密级切换后同步研究生公开/涉密声明。
Public Sub SyncPublicDeclaration()
    Dim p As Paragraph
    Dim r As Range
    Dim text As String
    Dim phrase As String
    Dim pos As Long
    Dim secure As Boolean
    If THU_UNDERGRADUATE Then Exit Sub
    secure = (Len(Trim$(ReadCC(TAG_SECURITY))) > 0)
    phrase = "，不包含涉及国家秘密的内容"
    For Each p In ActiveDocument.Paragraphs
        text = p.Range.Text
        If Left$(text, Len("本人郑重声明：所呈交的学位论文")) = "本人郑重声明：所呈交的学位论文" Then
            pos = InStr(text, phrase)
            If secure And pos > 0 Then
                Set r = p.Range.Duplicate
                r.Start = r.Start + pos - 1
                r.End = r.Start + Len(phrase)
                r.Delete
            ElseIf Not secure And pos = 0 Then
                pos = InStr(text, "独立进行研究工作所取得的成果")
                If pos > 0 Then
                    Set r = p.Range.Duplicate
                    r.Start = r.Start + pos - 1 + Len("独立进行研究工作所取得的成果")
                    r.End = r.Start
                    r.InsertAfter phrase
                End If
            End If
            Exit Sub
        End If
    Next p
End Sub

' 在本科封面首行或研究生封面页眉插入密级控件。
Private Sub InsertSecurityControl()
    Dim p As Paragraph
    Dim r As Range
    Dim cc As ContentControl

    Set p = SecurityParagraph()
    If p Is Nothing Then
        Err.Raise vbObjectError + 513, "modCover.InsertSecurityControl", "未找到封面密级段落。"
    End If
    ClearSecurityParagraph p
    Set r = ParagraphContentRange(p)
    Set cc = p.Range.ContentControls.Add(wdContentControlText, r)
    cc.Tag = TAG_SECURITY
    cc.Range.Text = SECURITY_DEFAULT
End Sub

' 本科只查封面表格内的密级样式，避免旧封面误写到 logo；研究生用首节页眉。
Private Function SecurityParagraph() As Paragraph
    Dim hdr As HeaderFooter
    Dim p As Paragraph

    If THU_UNDERGRADUATE Then
        For Each p In ActiveDocument.Tables(1).Range.Paragraphs
            If p.Style.NameLocal = "封面密级" Then
                Set SecurityParagraph = p
                Exit Function
            End If
        Next p
        Exit Function
    End If
    Set hdr = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary)
    hdr.LinkToPrevious = False
    Set p = hdr.Range.Paragraphs(1)
    p.Style = ActiveDocument.Styles("封面密级")
    Set SecurityParagraph = p
End Function

' 只清段落文字，保留段落样式与格式（不套用作者信息样式）。
Private Sub ClearSecurityParagraph(ByVal p As Paragraph)
    Dim r As Range

    Set r = ParagraphContentRange(p)
    If r.Characters.Count > 0 Then r.Text = ""
End Sub

Private Function IsSecurityParagraphBlank() As Boolean
    Dim p As Paragraph
    Dim s As String

    Set p = SecurityParagraph()
    If p Is Nothing Then Exit Function
    s = Replace(Replace(Replace(p.Range.Text, vbCr, ""), Chr$(7), ""), vbTab, "")
    IsSecurityParagraphBlank = (Len(Trim$(s)) = 0)
End Function

Private Sub InsertSupervisorRow(ByVal tagName As String, ByVal label As String)
    Dim p As Paragraph

    Set p = BlankOptionalParagraph()
    If p Is Nothing Then Exit Sub
    FillSupervisorRow p, label, tagName, ""
End Sub

' 作者信息块最后两段中第一个空段。
Private Function BlankOptionalParagraph() As Paragraph
    Dim cell As Cell
    Dim p As Paragraph
    Dim i As Long
    Dim count As Long

    Set cell = AuthorInfoCell()
    If cell Is Nothing Then Exit Function
    count = cell.Range.Paragraphs.Count
    For i = count To count - 1 Step -1
        If i < 1 Then Exit For
        Set p = cell.Range.Paragraphs(i)
        If Len(Trim$(Replace(Replace(Replace(p.Range.Text, vbCr, ""), Chr$(7), ""), vbTab, ""))) = 0 Then
            Set BlankOptionalParagraph = p
            Exit Function
        End If
    Next i
End Function

Private Sub RemoveSupervisorRow(ByVal tagName As String)
    Dim cc As ContentControl
    Dim cell As Cell
    Dim idx As Long
    Dim p As Paragraph

    Set cc = FindCC(tagName)
    If cc Is Nothing Then Exit Sub
    Set cell = AuthorInfoCell()
    If cell Is Nothing Then Exit Sub
    idx = ParagraphIndexOf(cell, cc.Range.Paragraphs(1))
    cc.Delete False
    If idx < 1 Then Exit Sub
    Set p = cell.Range.Paragraphs(idx)
    ClearCoverParagraph p
End Sub

Private Function ParagraphIndexOf(ByVal cell As Cell, ByVal target As Paragraph) As Long
    Dim p As Paragraph
    Dim i As Long
    Dim targetStart As Long

    targetStart = target.Range.Start
    For i = 1 To cell.Range.Paragraphs.Count
        Set p = cell.Range.Paragraphs(i)
        If p.Range.Start = targetStart Then
            ParagraphIndexOf = i
            Exit Function
        End If
    Next i
End Function

' 在指定段写入“\t标签：\t + 文本内容控件”，值可以为空。
' 段落正文范围：去掉末端段落标记；单元格末段的段落标记与单元格标记
' 在 Word 对象模型中合为一个字符，因此只需移除一次。
Private Function ParagraphContentRange(ByVal p As Paragraph) As Range
    Dim r As Range
    Dim lastChar As String

    Set r = p.Range.Duplicate
    If r.Characters.Count = 0 Then
        Set ParagraphContentRange = r
        Exit Function
    End If
    lastChar = r.Characters(r.Characters.Count).Text
    If InStr(lastChar, Chr$(13)) > 0 Or InStr(lastChar, Chr$(7)) > 0 Then
        r.MoveEnd Unit:=wdCharacter, Count:=-1
    End If
    Set ParagraphContentRange = r
End Function

Private Sub FillSupervisorRow(ByVal p As Paragraph, ByVal label As String, _
                              ByVal tagName As String, ByVal value As String)
    Dim r As Range
    Dim rngCC As Range
    Dim cc As ContentControl
    Dim placeholder As String

    Dim existing As ContentControl

    Set existing = FindCC(tagName)
    If Not existing Is Nothing Then existing.Delete False
    ClearCoverParagraph p
    Set r = ParagraphContentRange(p)
    r.Text = vbTab & label & "：" & vbTab

    placeholder = value
    If Len(placeholder) = 0 Then placeholder = " "
    Set r = ParagraphContentRange(p)
    r.Collapse Direction:=wdCollapseEnd
    r.Text = placeholder

    Set rngCC = ParagraphContentRange(p)
    rngCC.Start = rngCC.End - Len(placeholder)
    ' Word 在选区仍位于另一内容控件时可能忽略 Range 参数。
    ' 在重入锁保护下先定位目标文字，再从该段的集合创建控件。
    rngCC.Select
    Set cc = p.Range.ContentControls.Add(wdContentControlText, rngCC)
    cc.Tag = tagName
    If Len(value) > 0 Then cc.Range.Text = value
    FormatAuthorParagraph p
End Sub

Private Function SelectedCoverTag() As String
    Dim cc As ContentControl
    For Each cc In ActiveDocument.ContentControls
        If Selection.InRange(cc.Range) Then
            SelectedCoverTag = cc.Tag
            Exit Function
        End If
    Next cc
End Function

Private Sub RestoreCoverSelection(ByVal original As Range, ByVal tagName As String)
    Dim cc As ContentControl
    If Len(tagName) > 0 Then
        Set cc = FindCC(tagName)
        If cc Is Nothing Then Set cc = FindCC(TAG_AUTHOR_CN)
        If Not cc Is Nothing Then cc.Range.Select
    ElseIf Not original Is Nothing Then
        original.Select
    End If
End Sub

' 清空段内内容控件与文字，保留段落标记和封面作者信息样式。
Private Sub ClearCoverParagraph(ByVal p As Paragraph)
    Dim r As Range

    Set r = ParagraphContentRange(p)
    If r.Characters.Count > 0 Then r.Text = ""
    FormatAuthorParagraph p
End Sub

' 按实际封面类型紧凑排列，始终保留六行总高度。
Private Sub RebuildOptionalSupervisorRows()
    RebuildCoverAuthorInfo Not (FindCC(TAG_DISCIPLINE) Is Nothing)
End Sub

Private Sub RebuildCoverAuthorInfo(ByVal hasDiscipline As Boolean)
    Dim cell As Cell
    Dim tags As Variant
    Dim values(0 To 5) As String
    Dim hasAssoc As Boolean
    Dim hasCo As Boolean
    Dim i As Long
    Dim row As Long
    Dim cc As ContentControl
    Dim level As String
    Dim degreeType As String
    Dim category As String
    Dim changed As Boolean
    Dim recording As Boolean
    Dim errNo As Long
    Dim errDesc As String
    On Error GoTo Failed
    Set cell = AuthorInfoCell()
    If cell Is Nothing Then Exit Sub
    If cell.Range.Paragraphs.Count < 6 Then Exit Sub
    tags = Array(TAG_DEPARTMENT, TAG_DISCIPLINE, TAG_AUTHOR_CN, TAG_SUPERVISOR_CN, TAG_ASSOC_SUP, TAG_CO_SUP)
    hasAssoc = Not (FindCC(TAG_ASSOC_SUP) Is Nothing)
    hasCo = Not (FindCC(TAG_CO_SUP) Is Nothing)
    For i = 0 To 5
        values(i) = ReadCC(CStr(tags(i)))
        Set cc = FindCC(CStr(tags(i)))
        If Not cc Is Nothing Then
            If cc.LockContents Or cc.LockContentControl Then
                Err.Raise vbObjectError + 514, "modCover.RebuildCoverAuthorInfo", "请先解锁封面作者信息控件。"
            End If
        End If
    Next i
    If Len(Trim$(values(1))) > 0 Then RememberCoverText "THU_DISCIPLINE_CN", values(1)
    If hasDiscipline And Len(Trim$(values(1))) = 0 Then values(1) = RecallCoverText("THU_DISCIPLINE_CN", " ")
    ' 把整次重建作为一次可撤销操作；创建失败时恢复原作者区。
    Application.UndoRecord.StartCustomRecord "更新封面作者信息"
    recording = True
    For i = 0 To 5
        Set cc = FindCC(CStr(tags(i)))
        If Not cc Is Nothing Then
            cc.Delete False
            changed = True
        End If
    Next i
    For i = 1 To 6
        ClearCoverParagraph cell.Range.Paragraphs(i)
    Next i
    FillSupervisorRow cell.Range.Paragraphs(1), "培养单位", TAG_DEPARTMENT, values(0)
    row = 2
    If hasDiscipline Then
        FillSupervisorRow cell.Range.Paragraphs(row), LABEL_DISCIPLINE, TAG_DISCIPLINE, values(1)
        Set cc = FindCC(TAG_DISCIPLINE)
        cc.Type = wdContentControlComboBox
        row = row + 1
    End If
    FillSupervisorRow cell.Range.Paragraphs(row), LABEL_AUTHOR, TAG_AUTHOR_CN, values(2)
    row = row + 1
    FillSupervisorRow cell.Range.Paragraphs(row), "指导教师", TAG_SUPERVISOR_CN, values(3)
    row = row + 1
    If hasAssoc Then
        FillSupervisorRow cell.Range.Paragraphs(row), LABEL_ASSOC_SUP, TAG_ASSOC_SUP, values(4)
        row = row + 1
    End If
    If hasCo Then FillSupervisorRow cell.Range.Paragraphs(row), LABEL_CO_SUP, TAG_CO_SUP, values(5)
    SetCoverLabels (InStr(ReadCC(TAG_DEGREE), "专业学位") > 0)
    If ParseDegreeText(ReadCC(TAG_DEGREE), level, degreeType, category) Then
        RefreshDisciplineList category, degreeType
    End If
    Application.UndoRecord.EndCustomRecord
    Exit Sub
Failed:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If recording Then
        Application.UndoRecord.EndCustomRecord
        If changed Then ActiveDocument.Undo 1
    End If
    On Error GoTo 0
    Err.Raise errNo, "modCover.RebuildCoverAuthorInfo", errDesc
End Sub

' ============================================================================
' 日期
' ============================================================================

Private Function ChineseDate(ByVal s As String) As String
    Dim parts() As String
    Dim y As Integer
    Dim m As Integer
    Dim d As Integer

    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If InStr(s, "年") > 0 Then
        ChineseDate = s
        Exit Function
    End If
    parts = Split(Replace(Replace(s, "/", "-"), ".", "-"), "-")
    If UBound(parts) < 1 Then
        ChineseDate = s
        Exit Function
    End If
    On Error Resume Next
    y = CInt(parts(0))
    m = CInt(parts(1))
    If UBound(parts) >= 2 Then d = CInt(parts(2))
    On Error GoTo 0
    If y = 0 Or m < 1 Or m > 12 Then
        ChineseDate = s
        Exit Function
    End If
    ChineseDate = ChineseYear(y) & ChineseMonth(m)
    If d > 0 Then ChineseDate = ChineseDate & ChineseDay(d)
End Function

Private Function ChineseYear(ByVal y As Integer) As String
    Dim s As String
    Dim i As Long
    Dim out As String
    Dim digits As String

    digits = "〇一二三四五六七八九"
    s = CStr(y)
    For i = 1 To Len(s)
        out = out & Mid$(digits, CInt(Mid$(s, i, 1)) + 1, 1)
    Next i
    ChineseYear = out & "年"
End Function

Private Function ChineseMonth(ByVal m As Integer) As String
    ChineseMonth = ChineseNumber(m) & "月"
End Function

Private Function ChineseDay(ByVal d As Integer) As String
    ChineseDay = ChineseNumber(d) & "日"
End Function

Private Function ChineseNumber(ByVal n As Integer) As String
    Dim digits As String
    Dim t As Integer
    Dim o As Integer

    digits = "一二三四五六七八九"
    If n <= 0 Or n > 31 Then Exit Function
    If n < 10 Then
        ChineseNumber = Mid$(digits, n, 1)
    ElseIf n = 10 Then
        ChineseNumber = "十"
    ElseIf n < 20 Then
        ChineseNumber = "十" & Mid$(digits, n - 10, 1)
    Else
        t = n \ 10
        o = n Mod 10
        ChineseNumber = Mid$(digits, t, 1) & "十"
        If o > 0 Then ChineseNumber = ChineseNumber & Mid$(digits, o, 1)
    End If
End Function

Private Function EnglishDate(ByVal s As String) As String
    Dim parts() As String
    Dim y As Integer
    Dim m As Integer
    Dim months As Variant

    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    parts = Split(Replace(Replace(s, "/", "-"), ".", "-"), "-")
    If UBound(parts) < 1 Then
        EnglishDate = s
        Exit Function
    End If
    On Error Resume Next
    y = CInt(parts(0))
    m = CInt(parts(1))
    On Error GoTo 0
    If y = 0 Or m < 1 Or m > 12 Then
        EnglishDate = s
        Exit Function
    End If
    months = Array("January", "February", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")
    EnglishDate = CStr(months(m - 1)) & ", " & CStr(y)
End Function

' ============================================================================
' 内容控件存取
' ============================================================================

Private Function FindCC(ByVal tagName As String) As ContentControl
    Dim cc As ContentControl

    If tagName = TAG_SECURITY And Not THU_UNDERGRADUATE Then
        For Each cc In ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.ContentControls
            If cc.Tag = tagName Then
                Set FindCC = cc
                Exit Function
            End If
        Next cc
        Exit Function
    End If
    For Each cc In ActiveDocument.ContentControls
        If cc.Tag = tagName Then
            Set FindCC = cc
            Exit Function
        End If
    Next cc
End Function

Public Function ReadCC(ByVal tagName As String) As String
    Dim cc As ContentControl

    Set cc = FindCC(tagName)
    If cc Is Nothing Then Exit Function
    If cc.ShowingPlaceholderText Then Exit Function
    ReadCC = Trim$(cc.Range.Text)
End Function

Private Function WriteCC(ByVal tagName As String, ByVal value As String) As Boolean
    Dim cc As ContentControl

    Set cc = FindCC(tagName)
    If cc Is Nothing Then Exit Function
    If cc.LockContents Then
        MsgBox "内容控件“" & tagName & "”已锁定，无法写入。", vbExclamation, "封面信息"
        Exit Function
    End If
    On Error Resume Next
    cc.Range.Text = value
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    WriteCC = True
End Function

Private Function AuthorInfoCell() As Cell
    Dim cc As ContentControl

    On Error Resume Next
    Set cc = FindCC(TAG_DEPARTMENT)
    If cc Is Nothing Then Set cc = FindCC(TAG_SUPERVISOR_CN)
    If Not cc Is Nothing Then Set AuthorInfoCell = cc.Range.Cells(1)
    On Error GoTo 0
End Function

' ============================================================================
' 格式
' ============================================================================

Private Sub FormatAuthorParagraph(ByVal p As Paragraph)
    On Error Resume Next
    p.Style = ActiveDocument.Styles(STYLE_COVER_INFO)
    On Error GoTo 0
    p.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    p.Range.ParagraphFormat.SpaceBefore = 0
    p.Range.ParagraphFormat.SpaceAfter = 0
    p.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    p.Range.ParagraphFormat.LineSpacing = 26
    p.TabStops.ClearAll
    p.TabStops.Add Position:=159, Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    p.TabStops.Add Position:=164, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
End Sub

Private Sub FormatEnglishSupervisor()
    Dim cc As ContentControl
    Dim p As Paragraph
    Dim centerPos As Single

    Set cc = FindCC(TAG_SUPERVISOR_EN)
    If cc Is Nothing Then Exit Sub
    Set p = cc.Range.Paragraphs(1)
    p.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    p.Range.ParagraphFormat.LeftIndent = 0
    p.Range.ParagraphFormat.RightIndent = 0
    p.TabStops.ClearAll
    centerPos = 195.5
    On Error Resume Next
    If Not cc.Range.Cells(1) Is Nothing Then centerPos = cc.Range.Cells(1).Width / 2
    On Error GoTo 0
    p.TabStops.Add Position:=centerPos, Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    p.TabStops.Add Position:=centerPos + 5, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
End Sub

Private Sub FormatCC(ByVal tagName As String, ByVal fontName As String, _
                     ByVal fontSize As Single, ByVal boldValue As Boolean)
    Dim cc As ContentControl

    Set cc = FindCC(tagName)
    If cc Is Nothing Then Exit Sub
    On Error Resume Next
    cc.Range.Font.Name = fontName
    cc.Range.Font.NameAscii = fontName
    cc.Range.Font.NameFarEast = fontName
    cc.Range.Font.Size = fontSize
    cc.Range.Font.Bold = boldValue
    On Error GoTo 0
End Sub

' ============================================================================
' 测试入口（由 tests 通过 Word.Run 调用，返回 "ok|..." 字符串）
' ============================================================================

Public Function TestCoverDegreeCascade() As String
    Dim cc As ContentControl
    Dim degreeText As String
    Dim isProfessional As Boolean
    Dim isThesisTitle As Boolean

    On Error GoTo Failed
    Set cc = FindCC(TAG_DEGREE)
    If cc Is Nothing Then
        TestCoverDegreeCascade = "error|缺少学位类别控件"
        Exit Function
    End If
    If Not SelectDegreeByText("工学硕士学位") Then
        TestCoverDegreeCascade = "error|下拉列表中没有工学硕士学位"
        Exit Function
    End If
    ApplyDegreeCascade cc
    degreeText = ReadCC(TAG_DEGREE)
    isProfessional = (InStr(1, ReadCC(TAG_DEGREE_LINE_EN), "professional", vbTextCompare) > 0)
    isThesisTitle = (InStr(1, ReadCC(TAG_SUPERVISOR_EN), "Thesis Supervisor:", vbTextCompare) > 0)
    TestCoverDegreeCascade = "ok|" & degreeText & "|" & ReadCC(TAG_SUBMIT_EN) & "|" & _
        ReadCC(TAG_DEGREE_LINE_EN) & "|" & ReadCC(TAG_DEGREE_NAME_EN) & "|" & _
        BoolText(isProfessional) & "|" & BoolText(isThesisTitle) & "|" & CStr(DisciplineItemCount())
    Exit Function

Failed:
    TestCoverDegreeCascade = "error|" & Err.Description
End Function

Public Function TestCoverSelectDegree(ByVal displayText As String) As Boolean
    If Not SelectDegreeByText(displayText) Then Exit Function
    ApplyDegreeCascade FindCC(TAG_DEGREE)
    TestCoverSelectDegree = True
End Function

Public Function TestCoverDegreeCascadeProfessional() As String
    Dim cc As ContentControl
    Dim isProfessional As Boolean
    Dim isThesisTitle As Boolean

    On Error GoTo Failed
    Set cc = FindCC(TAG_DEGREE)
    If cc Is Nothing Then
        TestCoverDegreeCascadeProfessional = "error|缺少学位类别控件"
        Exit Function
    End If
    If Not SelectDegreeByText("金融硕士专业学位") Then
        TestCoverDegreeCascadeProfessional = "error|下拉列表中没有金融硕士专业学位"
        Exit Function
    End If
    ApplyDegreeCascade cc
    isProfessional = (InStr(1, ReadCC(TAG_DEGREE_LINE_EN), "professional", vbTextCompare) > 0)
    isThesisTitle = (InStr(1, ReadCC(TAG_SUPERVISOR_EN), "Thesis Supervisor:", vbTextCompare) > 0)
    TestCoverDegreeCascadeProfessional = "ok|" & ReadCC(TAG_DEGREE) & "|" & _
        ReadCC(TAG_SUBMIT_EN) & "|" & ReadCC(TAG_DEGREE_LINE_EN) & "|" & _
        ReadCC(TAG_DEGREE_NAME_EN) & "|" & BoolText(isProfessional) & "|" & _
        BoolText(isThesisTitle) & "|" & CStr(DisciplineItemCount())
    Exit Function

Failed:
    TestCoverDegreeCascadeProfessional = "error|" & Err.Description
End Function

Public Function TestCoverDegreeCascadeDoctor() As String
    Dim cc As ContentControl
    Dim isDoctor As Boolean

    On Error GoTo Failed
    Set cc = FindCC(TAG_DEGREE)
    If cc Is Nothing Then
        TestCoverDegreeCascadeDoctor = "error|缺少学位类别控件"
        Exit Function
    End If
    If Not SelectDegreeByText("工学博士学位") Then
        TestCoverDegreeCascadeDoctor = "error|下拉列表中没有工学博士学位"
        Exit Function
    End If
    ApplyDegreeCascade cc
    isDoctor = (InStr(1, ReadCC(TAG_SUPERVISOR_EN), "Dissertation Supervisor:", vbTextCompare) > 0)
    TestCoverDegreeCascadeDoctor = "ok|" & ReadCC(TAG_DEGREE) & "|" & _
        ReadCC(TAG_SUBMIT_EN) & "|" & ReadCC(TAG_DEGREE_LINE_EN) & "|" & _
        ReadCC(TAG_DEGREE_NAME_EN) & "|" & BoolText(isDoctor)
    Exit Function

Failed:
    TestCoverDegreeCascadeDoctor = "error|" & Err.Description
End Function

Public Function TestCoverDateCascade() As String
    Dim cc As ContentControl
    Dim boldOk As Boolean

    On Error GoTo Failed
    Set cc = FindCC(TAG_DATE_CN)
    If cc Is Nothing Then
        TestCoverDateCascade = "error|缺少成文日期控件"
        Exit Function
    End If
    cc.Range.Text = "2026-10"
    ApplyDateCascade cc
    Set cc = FindCC(TAG_DATE_EN)
    If Not cc Is Nothing Then boldOk = (cc.Range.Font.Bold = True)
    TestCoverDateCascade = "ok|" & ReadCC(TAG_DATE_CN) & "|" & ReadCC(TAG_DATE_EN) & "|" & BoolText(boldOk)
    Exit Function

Failed:
    TestCoverDateCascade = "error|" & Err.Description
End Function

Public Function TestCoverRefresh() As String
    Dim cc As ContentControl

    On Error GoTo Failed
    If Not SelectDegreeByText("金融硕士专业学位") Then
        TestCoverRefresh = "error|下拉列表中没有金融硕士专业学位"
        Exit Function
    End If
    Set cc = FindCC(TAG_DATE_CN)
    If Not cc Is Nothing Then cc.Range.Text = "2027-03"
    RefreshCover Nothing
    TestCoverRefresh = "ok|" & ReadCC(TAG_DEGREE) & "|" & ReadCC(TAG_DEGREE_LINE_EN) & "|" & _
        ReadCC(TAG_DEGREE_NAME_EN) & "|" & ReadCC(TAG_DATE_CN) & "|" & ReadCC(TAG_DATE_EN)
    Exit Function

Failed:
    TestCoverRefresh = "error|" & Err.Description
End Function

Public Function TestCoverToggleRows() As String
    Dim parts As String

    On Error GoTo Failed
    parts = "ok|" & CStr(AuthorParagraphCount()) & "|" & CStr(ActiveAuthorRowCount()) & "|" & CStr(LabelFitWidth(1))
    ToggleAssociateSupervisor Nothing
    parts = parts & "|" & CStr(ActiveAuthorRowCount()) & "|" & CStr(LabelFitWidth(1))
    ToggleCoSupervisor Nothing
    parts = parts & "|" & CStr(ActiveAuthorRowCount()) & "|" & CStr(LabelFitWidth(1))
    ToggleAssociateSupervisor Nothing
    parts = parts & "|" & CStr(ActiveAuthorRowCount()) & "|" & CStr(LabelFitWidth(1))
    ToggleCoSupervisor Nothing
    parts = parts & "|" & CStr(ActiveAuthorRowCount()) & "|" & CStr(LabelFitWidth(1))
    TestCoverToggleRows = parts
    Exit Function

Failed:
    TestCoverToggleRows = "error|" & Err.Description
End Function

Public Function TestEnglishCoverFormat() As String
    Dim cc As ContentControl
    Dim boldName As Long
    Dim boldDate As Long

    On Error GoTo Failed
    Set cc = FindCC(TAG_DEGREE_NAME_EN)
    If Not cc Is Nothing Then boldName = CLng(cc.Range.Font.Bold)
    Set cc = FindCC(TAG_DATE_EN)
    If Not cc Is Nothing Then boldDate = CLng(cc.Range.Font.Bold)
    TestEnglishCoverFormat = "ok|" & CStr(boldName) & "|" & CStr(boldDate) & "|" & _
        BoolText(SupervisorTabOk()) & "|" & BoolText(SecurityBlank())
    Exit Function

Failed:
    TestEnglishCoverFormat = "error|" & Err.Description
End Function

Public Function TestCoverSecurityBlank() As String
    On Error GoTo Failed
    TestCoverSecurityBlank = "ok|" & BoolText(SecurityBlank()) & "|"
    Exit Function

Failed:
    TestCoverSecurityBlank = "error|" & Err.Description
End Function

' 密级开关：无控件 → 添加；有内容 → 删除控件；再点 → 重新添加。
Public Function TestCoverSecurityClick() As String
    ToggleSecurity Nothing
    TestCoverSecurityClick = ReadCC(TAG_SECURITY)
End Function

Public Function TestCoverSecurityToggle() As String
    Dim cc As ContentControl
    Dim p As Paragraph
    Dim firstValue As String
    Dim secondHasCC As Boolean
    Dim secondBlank As Boolean
    Dim thirdValue As String
    Dim count1 As Long
    Dim count2 As Long
    Dim count3 As Long

    On Error GoTo Failed
    ' 归一化：删除现有密级控件并清空该行文字，从“无控件”状态开始。
    Set cc = FindCC(TAG_SECURITY)
    If Not cc Is Nothing Then
        Set p = cc.Range.Paragraphs(1)
        cc.Delete False
        ClearSecurityParagraph p
    End If

    ToggleSecurity Nothing
    firstValue = ReadCC(TAG_SECURITY)
    count1 = ActiveDocument.ContentControls.Count

    ToggleSecurity Nothing
    secondHasCC = Not (FindCC(TAG_SECURITY) Is Nothing)
    secondBlank = IsSecurityParagraphBlank()
    count2 = ActiveDocument.ContentControls.Count

    ToggleSecurity Nothing
    thirdValue = ReadCC(TAG_SECURITY)
    count3 = ActiveDocument.ContentControls.Count

    TestCoverSecurityToggle = "ok|" & firstValue & "|" & BoolText(secondHasCC) & "|" & _
        BoolText(secondBlank) & "|" & thirdValue & "|" & _
        CStr(count1) & "|" & CStr(count2) & "|" & CStr(count3)
    Exit Function

Failed:
    TestCoverSecurityToggle = "error|" & Err.Description
End Function

' 英文 Co-supervisor 行随中文联合导师行同步增删，且保留导师标题行。
Public Function TestCoverEnglishCoSupervisorSync() As String
    Dim cc As ContentControl
    Dim beforeText As String
    Dim onText As String
    Dim offText As String

    On Error GoTo Failed
    Set cc = FindCC(TAG_SUPERVISOR_EN)
    If cc Is Nothing Then
        TestCoverEnglishCoSupervisorSync = "error|缺少英文导师控件"
        Exit Function
    End If
    beforeText = cc.Range.Text
    ToggleCoSupervisor Nothing
    onText = cc.Range.Text
    ToggleCoSupervisor Nothing
    offText = cc.Range.Text
    TestCoverEnglishCoSupervisorSync = "ok|" & _
        BoolText(InStr(1, beforeText, "Co-supervisor:", vbTextCompare) > 0) & "|" & _
        BoolText(InStr(1, onText, "Co-supervisor:", vbTextCompare) > 0) & "|" & _
        BoolText(InStr(1, onText, "Thesis Supervisor:", vbTextCompare) > 0) & "|" & _
        BoolText(InStr(1, offText, "Co-supervisor:", vbTextCompare) = 0) & "|" & _
        BoolText(InStr(1, offText, "Thesis Supervisor:", vbTextCompare) > 0) & "|" & _
        BoolText(InStr(1, onText, "Professor", vbTextCompare) > 0)
    Exit Function

Failed:
    TestCoverEnglishCoSupervisorSync = "error|" & Err.Description
End Function

' 导师行开关：控件存在时再次点击即删除（即使已填写内容）；不存在时插入空行。
Public Function TestCoverSupervisorToggle() As String
    Dim cc As ContentControl
    Dim parts As String

    On Error GoTo Failed
    ' 归一化：确保两个可选导师行都不存在。
    If Not FindCC(TAG_ASSOC_SUP) Is Nothing Then RemoveSupervisorRow TAG_ASSOC_SUP
    If Not FindCC(TAG_CO_SUP) Is Nothing Then RemoveSupervisorRow TAG_CO_SUP

    ' 副导师行：添加 → 填写 → 再点删除（有内容也删除）。
    ToggleAssociateSupervisor Nothing
    Set cc = FindCC(TAG_ASSOC_SUP)
    If cc Is Nothing Then
        TestCoverSupervisorToggle = "error|添加副指导教师行失败"
        Exit Function
    End If
    cc.Range.Text = "张三　教授"
    ToggleAssociateSupervisor Nothing
    parts = "ok|" & BoolText(FindCC(TAG_ASSOC_SUP) Is Nothing) & "|" & _
        CStr(AuthorParagraphCount()) & "|" & CStr(ActiveAuthorRowCount()) & "|" & _
        BoolText(FindCC(TAG_CO_SUP) Is Nothing)

    ' 再次点击：重新插入空行，原内容不恢复。
    ToggleAssociateSupervisor Nothing
    parts = parts & "|" & BoolText(Not (FindCC(TAG_ASSOC_SUP) Is Nothing)) & "|" & ReadCC(TAG_ASSOC_SUP)

    ' 联合导师行：添加 → 填写 → 再点删除（有内容也删除）。
    ToggleCoSupervisor Nothing
    Set cc = FindCC(TAG_CO_SUP)
    If cc Is Nothing Then
        TestCoverSupervisorToggle = "error|添加联合指导教师行失败"
        Exit Function
    End If
    cc.Range.Text = "李四　教授"
    ToggleCoSupervisor Nothing
    parts = parts & "|" & BoolText(FindCC(TAG_CO_SUP) Is Nothing) & "|" & _
        CStr(AuthorParagraphCount()) & "|" & CStr(ActiveAuthorRowCount())
    TestCoverSupervisorToggle = parts
    Exit Function

Failed:
    TestCoverSupervisorToggle = "error|" & Err.Description
End Function

Private Function AuthorParagraphCount() As Long
    Dim cell As Cell

    Set cell = AuthorInfoCell()
    If cell Is Nothing Then Exit Function
    AuthorParagraphCount = cell.Range.Paragraphs.Count
End Function

Private Function ActiveAuthorRowCount() As Long
    Dim cell As Cell
    Dim p As Paragraph
    Dim count As Long
    Dim i As Long

    Set cell = AuthorInfoCell()
    If cell Is Nothing Then Exit Function
    For i = 1 To cell.Range.Paragraphs.Count
        Set p = cell.Range.Paragraphs(i)
        If Len(ParagraphLabel(p)) > 0 Then count = count + 1
    Next i
    ActiveAuthorRowCount = count
End Function

' 第 index 段标签的 FitText 宽度，换算成 twips（Word 对象模型用磅）。
Private Function LabelFitWidth(ByVal index As Long) As Long
    Dim cell As Cell
    Dim p As Paragraph
    Dim label As String
    Dim r As Range

    Set cell = AuthorInfoCell()
    If cell Is Nothing Then Exit Function
    If index < 1 Or index > cell.Range.Paragraphs.Count Then Exit Function
    Set p = cell.Range.Paragraphs(index)
    label = ParagraphLabel(p)
    If Len(label) = 0 Then Exit Function
    Set r = p.Range.Duplicate
    r.Start = r.Start + 1
    r.End = r.Start + Len(label)
    LabelFitWidth = CLng(r.FitTextWidth * 20)
End Function

Private Function DisciplineItemCount() As Long
    Dim cc As ContentControl

    Set cc = FindCC(TAG_DISCIPLINE)
    If cc Is Nothing Then Exit Function
    On Error Resume Next
    DisciplineItemCount = cc.DropdownListEntries.Count
    On Error GoTo 0
End Function

Private Function SupervisorTabOk() As Boolean
    Dim cc As ContentControl
    Dim p As Paragraph
    Dim hasRight As Boolean
    Dim hasLeft As Boolean
    Dim i As Long

    Set cc = FindCC(TAG_SUPERVISOR_EN)
    If cc Is Nothing Then Exit Function
    Set p = cc.Range.Paragraphs(1)
    For i = 1 To p.TabStops.Count
        If Abs(p.TabStops(i).Position - 195.5) < 0.6 Then
            If p.TabStops(i).Alignment = wdAlignTabRight Then hasRight = True
        End If
        If Abs(p.TabStops(i).Position - 200.5) < 0.6 Then
            If p.TabStops(i).Alignment = wdAlignTabLeft Then hasLeft = True
        End If
    Next i
    SupervisorTabOk = (hasRight And hasLeft)
End Function

Private Function SecurityBlank() As Boolean
    Dim cc As ContentControl

    Set cc = FindCC(TAG_SECURITY)
    If cc Is Nothing Then
        ' 没有密级控件即视为公开论文。
        SecurityBlank = True
        Exit Function
    End If
    SecurityBlank = (Len(Trim$(cc.Range.Text)) = 0)
End Function

Private Function BoolText(ByVal value As Boolean) As String
    If value Then
        BoolText = "True"
    Else
        BoolText = "False"
    End If
End Function

' 测试用：在下拉列表中选择指定显示文本（下拉内容控件不能直接改写文本）。
Private Function SelectDegreeByText(ByVal displayText As String) As Boolean
    Dim cc As ContentControl
    Dim i As Long

    Set cc = FindCC(TAG_DEGREE)
    If cc Is Nothing Then Exit Function
    For i = 1 To cc.DropdownListEntries.Count
        If Trim$(cc.DropdownListEntries(i).Text) = displayText Then
            cc.DropdownListEntries(i).Select
            SelectDegreeByText = True
            Exit Function
        End If
    Next i
End Function
