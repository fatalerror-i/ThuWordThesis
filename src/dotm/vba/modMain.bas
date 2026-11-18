Attribute VB_Name = "modMain"
Option Explicit

' ============================================================================
' 清华大学毕业论文模板（.dotm）主模块
' 说明：所有对外入口都做错误捕获；失败时中文提示，不把 Word 内部错误直接抛给作者。
' ============================================================================

Public Const STYLE_FORMULA As String = "公式"
Public Const STYLE_INLINE_FORMULA As String = "行内公式"
Public Const STYLE_FIG_CAPTION As String = "图注"
Public Const STYLE_SUBFIG_CAPTION As String = "子图注"
Public Const STYLE_TAB_CAPTION As String = "表注"
Public Const STYLE_APP_FIG_CAPTION As String = "附录图注"
Public Const STYLE_APP_TAB_CAPTION As String = "附录表注"
Public Const STYLE_UNNUM_SUBSECTION As String = "章内无编号小节"
Public Const STYLE_REF_ENTRY As String = "参考文献条目"
Public Const STYLE_REF_ENTRY_AY As String = "参考文献条目（著者出版年）"
Public Const STYLE_CITATION As String = "文献引用"
Public Const STYLE_IMAGE As String = "图片"
Public Const STYLE_SYMBOL_ENTRY As String = "符号说明条目"

Public Const BM_EQ_PREFIX As String = "THU_EQ_"
Public Const BM_FIG_PREFIX As String = "THU_FIG_"
Public Const BM_TAB_PREFIX As String = "THU_TAB_"
Public Const BM_CH_PREFIX As String = "THU_CH_"
Public Const BM_REF_ENTRY_PREFIX As String = "THU_REF_ENTRY_"
Public Const PROP_REF_NUM_PREFIX As String = "_THU_REF_NUM_"

Public gDirty As Boolean
Public gSilentClicks As Boolean

' 旧版本模板生成的附录复位域缺少 \r，Word 会把 0 当作书签名并显示
' “错误!未定义书签”。这里记录当前文档是否已尝试修复，避免每次调用
' EnsureReady 都遍历全部域。
Private gRepairDocName As String
Private gRepairDone As Boolean

' ---------------------------------------------------------------------------
' 功能区回调：关于 / 阶段开关
' ---------------------------------------------------------------------------

Public Sub ShowAbout(control As IRibbonControl)
    MsgBox APP_TITLE & vbCrLf & vbCrLf & _
           "当前版本：" & APP_VERSION, vbInformation, "关于"
End Sub

Public Sub GetPhase1Enabled(control As IRibbonControl, ByRef returnedVal As Boolean)
    returnedVal = False
End Sub

' ---------------------------------------------------------------------------
' 统一入口包装
' ---------------------------------------------------------------------------

Public Sub UpdateAllFields(control As IRibbonControl)
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    UpdateAllFieldsNow
    MsgBox "已更新目录、题注、公式编号与引用域。", vbInformation, "更新域"
    Exit Sub
EH:
    ReportError "更新域", Err.Number, Err.Description
End Sub

Public Sub BackupCopy(control As IRibbonControl)
    On Error GoTo EH
    Dim doc As Document
    Dim base As String
    Dim path As String
    Dim ts As String

    If Not EnsureReady() Then Exit Sub
    Set doc = ActiveDocument
    If Len(doc.Path) = 0 Then
        MsgBox "请先保存文档，再执行备份。", vbExclamation, "备份"
        Exit Sub
    End If
    base = doc.FullName
    If InStrRev(base, ".") > 0 Then base = Left$(base, InStrRev(base, ".") - 1)
    ts = Format$(Now, "yyyymmdd-hhnnss")
    path = base & "_备份_" & ts & ".docx"
    doc.SaveCopyAs path
    MsgBox "已备份到：" & vbCrLf & path, vbInformation, "备份"
    Exit Sub
EH:
    ReportError "备份", Err.Number, Err.Description
End Sub

' ---------------------------------------------------------------------------
' 文档状态与域
' ---------------------------------------------------------------------------

' 打开模板或启用宏时，先修复旧版附录复位域。
Public Sub AutoOpen()
    On Error Resume Next
    RepairAppendixResetFields True
    On Error GoTo 0
End Sub

Public Function EnsureReady() As Boolean
    If Documents.Count = 0 Then
        MsgBox "请先打开一篇论文文档。", vbExclamation, APP_TITLE
        Exit Function
    End If
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "文档处于保护状态，请先停止保护再使用宏。", vbExclamation, APP_TITLE
        Exit Function
    End If
    If (Not gRepairDone) Or (ActiveDocument.FullName <> gRepairDocName) Then
        RepairAppendixResetFields True
        gRepairDocName = ActiveDocument.FullName
        gRepairDone = True
    End If
    EnsureReady = True
End Function

' 修复旧版模板写坏的附录复位域：
'   SEQ 表 0 \h  ->  SEQ 表 \r 0 \h（附录表/附录图复位域同理）
' Word 把 0 解释为书签名，书签不存在时显示“错误!未定义书签”。
Public Function RepairAppendixResetFields(Optional ByVal silent As Boolean = True) As Long
    Dim f As Field
    Dim code As String
    Dim fixedCode As String
    Dim parts() As String
    Dim ident As String
    Dim i As Long
    Dim n As Long

    On Error Resume Next
    For Each f In ActiveDocument.Fields
        code = f.Code.Text
        code = Replace(code, vbCr, " ")
        code = Replace(code, vbTab, " ")
        Do While InStr(code, "  ") > 0
            code = Replace(code, "  ", " ")
        Loop
        code = Trim$(code)
        If Left$(UCase$(code), 4) = "SEQ " Then
            parts = Split(code, " ")
            If UBound(parts) >= 2 Then
                ident = parts(1)
                If (ident = "表" Or ident = "图" Or ident = "式" Or _
                    ident = "附录表" Or ident = "附录图") And InStr(code, "\h") > 0 Then
                    If parts(2) = "0" And InStr(code, "\r") = 0 Then
                        fixedCode = "SEQ " & ident & " \r 0"
                        For i = 3 To UBound(parts)
                            fixedCode = fixedCode & " " & parts(i)
                        Next i
                        f.Code.Text = fixedCode
                    End If
                    ' 旧文件里可能已有“错误!未定义书签”的缓存结果；
                    ' 代码已正确时也要更新一次，清掉缓存。
                    If parts(2) = "0" Or InStr(code, "\r 0") > 0 Then
                        f.Update
                        n = n + 1
                    End If
                End If
            End If
        End If
    Next f
    On Error GoTo 0
    RepairAppendixResetFields = n
    If n > 0 And Not silent Then
        MsgBox "已修复 " & CStr(n) & " 个附录复位域。", vbInformation, "修复附录域"
    End If
End Function

' 按附录与对象类别重置首个编号域。重排对象时移走旧首项的复位开关。
Public Sub RefreshAppendixSequences()
    Dim seen As Object
    Dim f As Field
    Dim heading As Range
    Dim code As String
    Dim ident As String
    Dim key As String
    Dim expected As String
    Dim legacyReset As Boolean

    Set seen = CreateObject("Scripting.Dictionary")
    For Each f In ActiveDocument.Fields
        code = ""
        On Error Resume Next
        code = Trim$(f.Code.Text)
        On Error GoTo 0
        If Left$(UCase$(code), 4) = "SEQ " Then
            ident = Split(code, " ")(1)
            If ident = "附录图" Or ident = "附录表" Or ident = "式" Then
                legacyReset = (InStr(code, "\h") > 0 And InStr(code, "\r 0") > 0)
                If legacyReset Or InStr(code, "\h") = 0 Then
                    Set heading = ActiveDocument.Range(0, f.Code.Start)
                    With heading.Find
                        .ClearFormatting
                        .Text = ""
                        .Style = ActiveDocument.Styles("附录标题")
                        .Forward = False
                        .Wrap = wdFindStop
                    End With
                    If heading.Find.Execute Then
                        key = CStr(heading.Paragraphs(1).Range.Start) & "|" & ident
                        If legacyReset Then
                            ' 旧模板把复位域放在附录导语中。用户在导语前插入对象时，
                            ' 只让该旧域复读当前值，避免导语之后再从 1 编号。
                            If seen.Exists(key) Then f.Code.Text = " SEQ " & ident & " \c \h "
                        ElseIf ident <> "式" Or HasAppendixNumberSource(f.Code.Paragraphs(1)) Then
                            expected = "SEQ " & ident
                            ' OMath 将直接赋值的 * 替换为乘号，SEQ 默认采用阿拉伯数字。
                            If ident <> "式" Then expected = expected & " \* ARABIC"
                            If seen.Exists(key) Then
                                expected = expected & " \s 1"
                            Else
                                expected = expected & " \r 1"
                                seen.Add key, True
                            End If
                            If code <> expected Then f.Code.Text = " " & expected & " "
                        End If
                    End If
                End If
            End If
        End If
    Next f
End Sub

Private Function HasAppendixNumberSource(p As Paragraph) As Boolean
    Dim f As Field
    For Each f In p.Range.Fields
        If InStr(1, f.Code.Text, "STYLEREF ""附录标题""", vbTextCompare) > 0 Then
            HasAppendixNumberSource = True
            Exit Function
        End If
    Next f
End Function

Public Sub UpdateAllFieldsNow()
    Dim doc As Document
    Dim s As Section
    Dim hf As HeaderFooter
    Dim f As Field
    Dim i As Long

    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    On Error Resume Next
    RepairAppendixResetFields True
    RefreshAppendixSequences
    NormalizeFootnoteReferences
    doc.Fields.Update
    For Each s In doc.Sections
        For i = 1 To 3
            For Each hf In s.Footers
                If hf.Exists Then hf.Range.Fields.Update
            Next hf
        Next i
        For i = 1 To 3
            For Each hf In s.Headers
                If hf.Exists Then hf.Range.Fields.Update
            Next hf
        Next i
    Next s
    If doc.TablesOfContents.Count > 0 Then doc.TablesOfContents(1).Update
    If doc.TablesOfFigures.Count > 0 Then doc.TablesOfFigures(1).Update
    RefreshFormulaNumbers False
    For Each f In doc.Fields
        If Left$(UCase$(Trim$(f.Code.Text)), 11) = "REF THU_EQ_" Then f.Update
    Next f
    On Error GoTo 0
    Application.ScreenUpdating = True
    ' 文献引用单独更新。普通 F9 只更新其余 Word 域。
    RefreshCitationProperties
    NormalizeCitations
End Sub

Public Sub RefreshFields(rng As Range)
    On Error Resume Next
    rng.Fields.Update
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' 段落 / 样式 / 章节判断
' ---------------------------------------------------------------------------

Public Function GetSelectionParagraph() As Paragraph
    If Selection.Paragraphs.Count < 1 Then Exit Function
    Set GetSelectionParagraph = Selection.Paragraphs(1)
End Function

Public Function StyleExists(ByVal styleName As String) As Boolean
    Dim st As Style
    On Error Resume Next
    Set st = ActiveDocument.Styles(styleName)
    StyleExists = Not (st Is Nothing)
    On Error GoTo 0
End Function

Public Function PrimaryStyleName(ByVal st As Style) As String
    ' 内置样式 NameLocal 为“原名,替换名”，逻辑判断统一使用原名。
    PrimaryStyleName = Trim$(Split(st.NameLocal, ",")(0))
End Function

Public Function IsChapterHeading(p As Paragraph) As Boolean
    Dim txt As String
    Dim st As String

    st = PrimaryStyleName(p.Style)
    txt = Trim$(CleanText(p.Range.Text))
    If st = "标题 1" Or st = "Heading 1" Or st = "章标题" Then
        IsChapterHeading = True
        Exit Function
    End If
    If p.OutlineLevel = wdOutlineLevel1 Then
        If Left$(txt, 1) = "第" And InStr(txt, "章") > 0 Then IsChapterHeading = True
    End If
End Function

Public Function IsAppendixHeading(p As Paragraph) As Boolean
    Dim txt As String
    Dim st As String

    st = PrimaryStyleName(p.Style)
    txt = Trim$(CleanText(p.Range.Text))
    If InStr(st, "附录") > 0 Then
        IsAppendixHeading = True
        Exit Function
    End If
    If Left$(txt, 2) = "附录" Then IsAppendixHeading = True
End Function

Public Function IsAppendixAt(rng As Range) As Boolean
    Dim p As Paragraph
    Dim bApp As Boolean

    bApp = False
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= rng.Start Then Exit For
        If IsAppendixHeading(p) Then
            bApp = True
        ElseIf IsChapterHeading(p) Then
            bApp = False
        End If
    Next p
    IsAppendixAt = bApp
End Function

Public Function CleanText(ByVal s As String) As String
    s = Replace(s, Chr$(13), "")
    s = Replace(s, Chr$(7), "")
    s = Replace(s, Chr$(11), "")
    s = Replace(s, Chr$(9), " ")
    s = Trim$(s)
    Do While Len(s) > 0
        If Left$(s, 1) <> ChrW(&H3000) Then Exit Do
        s = Trim$(Mid$(s, 2))
    Loop
    Do While Len(s) > 0
        If Right$(s, 1) <> ChrW(&H3000) Then Exit Do
        s = Trim$(Left$(s, Len(s) - 1))
    Loop
    CleanText = s
End Function

Public Function GetHeadingListString(p As Paragraph) As String
    Dim s As String
    On Error Resume Next
    s = p.Range.ListFormat.ListString
    On Error GoTo 0
    GetHeadingListString = Trim$(Replace(s, ChrW(&H3000), " "))
End Function

' 从“第2章”/“附录A”/“2”中取数字或字母；取不到返回空串
Public Function GetHeadingNumberText(p As Paragraph) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    s = GetHeadingListString(p)
    If Len(s) = 0 Then s = CleanText(p.Range.Text)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            out = out & ch
        ElseIf (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            out = out & UCase$(ch)
        End If
    Next i
    GetHeadingNumberText = out
End Function

' ---------------------------------------------------------------------------
' 域 / 书签
' ---------------------------------------------------------------------------

Public Function AddFieldAt(rng As Range, ByVal code As String) As Field
    Dim f As Field
    Set f = rng.Document.Fields.Add(rng, wdFieldEmpty, code, False)
    Set AddFieldAt = f
End Function

Public Sub AddTextAt(rng As Range, ByVal text As String)
    rng.Text = text
    rng.Collapse wdCollapseEnd
End Sub

Public Function GetUniqueBookmarkName(ByVal prefix As String) As String
    Dim i As Long
    Dim name As String
    i = 1
    Do
        name = prefix & Format$(i, "0000")
        If Not ActiveDocument.Bookmarks.Exists(name) Then
            GetUniqueBookmarkName = name
            Exit Function
        End If
        i = i + 1
    Loop While i < 10000
    GetUniqueBookmarkName = prefix & Format$(Timer * 1000, "0000")
End Function

' 在 rng 中查找 text，找到后把该段文本加书签；找不到返回 False
Public Function BookmarkText(ByVal bmName As String, rng As Range, ByVal text As String) As Boolean
    Dim r As Range
    Set r = rng.Duplicate
    With r.Find
        .ClearFormatting
        .Text = text
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        .MatchCase = True
    End With
    If r.Find.Execute Then
        If ActiveDocument.Bookmarks.Exists(bmName) Then ActiveDocument.Bookmarks(bmName).Delete
        ActiveDocument.Bookmarks.Add bmName, r
        BookmarkText = True
    End If
End Function

Public Sub DeleteBookmarkIfExists(ByVal bmName As String)
    On Error Resume Next
    If ActiveDocument.Bookmarks.Exists(bmName) Then ActiveDocument.Bookmarks(bmName).Delete
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Undo / 错误
' ---------------------------------------------------------------------------

Public Function StartUndo(ByVal name As String) As Object
    Dim ur As Object
    On Error Resume Next
    Set ur = Application.UndoRecord
    If Not ur Is Nothing Then ur.StartCustomRecord name
    Set StartUndo = ur
    On Error GoTo 0
End Function

Public Sub FinishUndo(ur As Object, ByVal ok As Boolean)
    On Error Resume Next
    If Not ur Is Nothing Then
        ur.EndCustomRecord
        If Not ok And gDirty Then Application.Undo
    End If
    gDirty = False
    On Error GoTo 0
End Sub

Public Sub ReportError(ByVal procName As String, ByVal errNo As Long, ByVal errDesc As String)
    On Error Resume Next
    Application.ScreenUpdating = True
    MsgBox "操作失败：" & procName & vbCrLf & vbCrLf & _
           "错误 " & CStr(errNo) & "：" & errDesc & vbCrLf & vbCrLf & _
           "文档内容未被自动修复，请撤销（Ctrl+Z）后重试。", _
           vbCritical, APP_TITLE
End Sub

' ---------------------------------------------------------------------------
' Ribbon 回调包装（Ribbon XML 中的名字与实现模块解耦）
' ---------------------------------------------------------------------------

Public Sub NormalizeFormula(control As Object)
    ' 用户要求：点“公式”后不弹消息窗口，成功/失败都静默处理。
    RunNormalizeFormula control, True
End Sub

Public Sub InsertCrossReference(control As IRibbonControl)
    Dim tag As String

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    tag = LCase$(Trim$(control.Tag))
    Select Case tag
        Case "figure"
            InsertFigureReference control
        Case "table"
            InsertTableReference control
        Case "formula"
            RunInsertFormulaReference control
        Case "chapter"
            InsertChapterReference control
        Case Else
            MsgBox "未知的引用类型：" & control.Tag, vbExclamation, "引用"
    End Select
    Exit Sub
EH:
    ReportError "交叉引用", Err.Number, Err.Description
End Sub

' ---------------------------------------------------------------------------
' 结构：正文/附录/其他章节
' ---------------------------------------------------------------------------

Public Sub InsertChapter(control As Object)
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    InsertSectionHeading "标题 1", "章标题"
    Exit Sub
EH:
    ReportError "插入正文章节", Err.Number, Err.Description
End Sub

Public Sub InsertAppendix(control As Object)
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    InsertSectionHeading "附录标题", "附录标题"
    Exit Sub
EH:
    ReportError "插入附录章节", Err.Number, Err.Description
End Sub

Public Sub InsertOtherSection(control As Object)
    Dim labels As Variant
    Dim i As Long
    Dim idx As Long
    Dim txt As String

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub

    labels = Array("摘　要", "Abstract", "目　录", "插图清单", "附表清单", _
                   "符号和缩略语说明", "参考文献", "致　谢", "声　明")
    If THU_UNDERGRADUATE Then
        ReDim Preserve labels(9)
        labels(9) = "在学期间参加课题的研究成果"
    Else
        ReDim Preserve labels(11)
        labels(9) = "个人简历、在学期间完成的相关学术成果"
        labels(10) = "指导教师（小组）评语"
        labels(11) = "答辩委员会决议书"
    End If
    ResetPicker
    For i = LBound(labels) To UBound(labels)
        AddPickerItem CStr(labels(i)), CStr(i)
    Next i
    If Not ShowPicker("插入论文其他章节", "选择要插入的章节：") Then Exit Sub
    If gPickerSelectedCount = 0 Then Exit Sub
    idx = gPickerSelected(0)
    txt = CStr(labels(idx))
    If THU_UNDERGRADUATE And idx <= 5 And idx <> 2 Then
        InsertSectionHeading "章标题（无编号、不进目录）", txt
    Else
        InsertSectionHeading "章标题（无编号）", txt
        If THU_UNDERGRADUATE And idx = 2 Then
            Selection.Paragraphs(1).OutlineLevel = wdOutlineLevelBodyText
        End If
    End If
    Exit Sub
EH:
    ReportError "插入其他章节", Err.Number, Err.Description
End Sub

' 插入分节符 + 指定标题样式段落，并选中占位文字
Private Sub InsertSectionHeading(ByVal styleName As String, ByVal placeholder As String)
    Dim p As Paragraph

    If Not StyleExists(styleName) Then
        MsgBox "模板中缺少样式：" & styleName, vbExclamation, "结构"
        Exit Sub
    End If

    Selection.InsertBreak Type:=wdSectionBreakNextPage
    Set p = Selection.Paragraphs(1)
    p.Range.Text = placeholder
    p.Style = ActiveDocument.Styles(styleName)
    p.Range.Select
    Selection.Collapse wdCollapseStart
    Selection.MoveEnd Unit:=wdCharacter, Count:=Len(placeholder)
End Sub

Public Sub ApplyHeadingStyle(control As Object)
    ApplyHeadingStyleByName CStr(control.Tag)
End Sub

Public Function ApplyHeadingStyleByName(ByVal styleName As String) As Boolean
    Dim p As Paragraph

    On Error GoTo EH
    If Not EnsureReady() Then Exit Function
    If Len(styleName) = 0 Then Exit Function
    If Not StyleExists(styleName) Then
        MsgBox "模板中缺少样式：" & styleName, vbExclamation, "结构"
        Exit Function
    End If
    Set p = Selection.Paragraphs(1)
    p.Style = ActiveDocument.Styles(styleName)
    ApplyHeadingStyleByName = True
    Exit Function
EH:
    ReportError "设置标题格式", Err.Number, Err.Description
End Function

Public Sub ApplyBodyStyle(control As Object)
    Dim p As Paragraph
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    ' 对当前选中的全部段落套用正文样式；未选中时处理光标所在段落。
    For Each p In Selection.Paragraphs
        p.Style = ActiveDocument.Styles(wdStyleNormal)
    Next p
    Exit Sub
EH:
    ReportError "设置正文样式", Err.Number, Err.Description
End Sub

Public Sub ApplyEmphasis(control As Object)
    ApplyEmphasisByTag CStr(control.Tag)
End Sub

Public Sub ApplyEmphasisByTag(ByVal tag As String)
    Dim styleId As Variant
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    Select Case Trim$(tag)
        Case "Subtle Emphasis": styleId = wdStyleSubtleEmphasis
        Case "Emphasis": styleId = wdStyleEmphasis
        Case "Strong": styleId = wdStyleStrong
        Case Else
            MsgBox "未知的强调类型：" & tag, vbExclamation, "格式"
            Exit Sub
    End Select
    Selection.Range.Style = styleId
    Exit Sub
EH:
    ReportError "应用强调样式", Err.Number, Err.Description
End Sub

' 测试入口：对选中文字依次应用三种强调样式，并检查字符样式名。
Public Function TestApplyEmphasis() As String
    Dim tags As Variant
    Dim expected As Variant
    Dim p As Paragraph
    Dim r As Range
    Dim i As Long
    Dim actual As String
    On Error GoTo EH
    tags = Array("Subtle Emphasis", "Emphasis", "Strong")
    expected = Array("不明显强调", "强调", "要点")
    For i = LBound(tags) To UBound(tags)
        Set p = ActiveDocument.Paragraphs.Add
        p.Range.InsertBefore "强调测试"
        Set r = p.Range
        r.End = r.Start + 2
        r.Select
        ApplyEmphasisByTag CStr(tags(i))
        actual = PrimaryStyleName(Selection.Range.Style)
        If actual <> CStr(expected(i)) Then
            TestApplyEmphasis = "FAIL|" & tags(i) & "|" & actual
            Exit Function
        End If
    Next i
    TestApplyEmphasis = "ok|3"
    Exit Function
EH:
    TestApplyEmphasis = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Sub ApplyListStyle(control As Object)
    ApplyListStyleByTag CStr(control.Tag)
End Sub

' 对当前选中的全部段落套用列表样式；未选中时处理光标所在段落。
Private Sub ApplyListStyleByTag(ByVal tag As String)
    Dim styleId As Long
    Dim p As Paragraph
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    Select Case tag
        Case "List Number": styleId = wdStyleListNumber
        Case "List Number 2": styleId = wdStyleListNumber2
        Case "List Number 3": styleId = wdStyleListNumber3
        Case "List Number 4": styleId = wdStyleListNumber4
        Case "List Number 5": styleId = wdStyleListNumber5
        Case "List Bullet": styleId = wdStyleListBullet
        Case "List Bullet 2": styleId = wdStyleListBullet2
        Case "List Bullet 3": styleId = wdStyleListBullet3
        Case "List Bullet 4": styleId = wdStyleListBullet4
        Case "List Bullet 5": styleId = wdStyleListBullet5
        Case Else
            MsgBox "未知列表样式：" & tag, vbExclamation, "结构"
            Exit Sub
    End Select
    For Each p In Selection.Paragraphs
        p.Style = ActiveDocument.Styles(styleId)
    Next p
    Exit Sub
EH:
    ReportError "设置列表样式", Err.Number, Err.Description
End Sub

Public Sub InsertUnnumberedSubsection(control As Object)
    Dim p As Paragraph
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    Selection.TypeParagraph
    Set p = Selection.Paragraphs(1)
    p.Range.InsertBefore "章内无编号小节"
    p.Style = ActiveDocument.Styles(STYLE_UNNUM_SUBSECTION)
    p.Range.Select
    Selection.Collapse wdCollapseStart
    Selection.MoveEnd Unit:=wdCharacter, Count:=Len("章内无编号小节")
    Exit Sub
EH:
    ReportError "插入章内无编号小节", Err.Number, Err.Description
End Sub

Public Sub RestartListAtOne(control As Object)
    Dim p As Paragraph
    Dim styleName As String
    Dim lf As ListFormat
    Dim i As Long, n As Long
    Dim lefts() As Single, firsts() As Single
    Dim cuLefts() As Single, cuFirsts() As Single
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    If RestartSubfigureAtSelection() Then Exit Sub
    n = Selection.Paragraphs.Count
    If n = 0 Then Exit Sub
    styleName = PrimaryStyleName(Selection.Paragraphs(1).Style)
    Set lf = Selection.Range.ListFormat
    If lf.ListTemplate Is Nothing Then
        MsgBox "请先把选中段落套用有序或无序列表样式。", vbExclamation, "列表"
        Exit Sub
    End If
    ' Word 的 ApplyListTemplateWithLevel 会把选中段落左缩进重置为 0；
    ' 先保存缩进，重启后恢复，避免“编号归 1 但列表跑到页边”的原生 bug。
    ReDim lefts(1 To n)
    ReDim firsts(1 To n)
    ReDim cuLefts(1 To n)
    ReDim cuFirsts(1 To n)
    For i = 1 To n
        Set p = Selection.Paragraphs(i)
        p.Style = ActiveDocument.Styles(styleName)
        lefts(i) = p.LeftIndent
        firsts(i) = p.FirstLineIndent
        cuLefts(i) = p.CharacterUnitLeftIndent
        cuFirsts(i) = p.CharacterUnitFirstLineIndent
    Next i
    Set lf = Selection.Range.ListFormat
    lf.ApplyListTemplateWithLevel ListTemplate:=lf.ListTemplate, _
        ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList
    For i = 1 To n
        Set p = Selection.Paragraphs(i)
        p.LeftIndent = lefts(i)
        p.FirstLineIndent = firsts(i)
        If cuLefts(i) > -100 And cuLefts(i) < 1000 Then
            p.CharacterUnitLeftIndent = cuLefts(i)
        End If
        If cuFirsts(i) > -100 And cuFirsts(i) < 1000 Then
            p.CharacterUnitFirstLineIndent = cuFirsts(i)
        End If
    Next i
    Exit Sub
EH:
    ReportError "重新开始编号", Err.Number, Err.Description
End Sub

' 继续编号：让选中列表接着上方同模板列表的编号，等价于 Word 的“继续编号”。
Public Sub ContinueListNumbering(control As Object)
    Dim p As Paragraph
    Dim styleName As String
    Dim lf As ListFormat
    Dim i As Long, n As Long
    Dim lefts() As Single, firsts() As Single
    Dim cuLefts() As Single, cuFirsts() As Single
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    n = Selection.Paragraphs.Count
    If n = 0 Then Exit Sub
    styleName = PrimaryStyleName(Selection.Paragraphs(1).Style)
    Set lf = Selection.Range.ListFormat
    If lf.ListTemplate Is Nothing Then
        MsgBox "请先把选中段落套用有序列表样式。", vbExclamation, "列表"
        Exit Sub
    End If
    If lf.ListType = wdListBullet Then
        StatusBar = "项目符号列表没有编号，无需继续。"
        Exit Sub
    End If
    ' 与“重新开始于1”相同：ApplyListTemplateWithLevel 会把选中段落左缩进
    ' 重置为 0；先保存缩进，继续后恢复，避免列表跑到页边。
    ReDim lefts(1 To n)
    ReDim firsts(1 To n)
    ReDim cuLefts(1 To n)
    ReDim cuFirsts(1 To n)
    For i = 1 To n
        Set p = Selection.Paragraphs(i)
        p.Style = ActiveDocument.Styles(styleName)
        lefts(i) = p.LeftIndent
        firsts(i) = p.FirstLineIndent
        cuLefts(i) = p.CharacterUnitLeftIndent
        cuFirsts(i) = p.CharacterUnitFirstLineIndent
    Next i
    Set lf = Selection.Range.ListFormat
    lf.ApplyListTemplateWithLevel ListTemplate:=lf.ListTemplate, _
        ContinuePreviousList:=True, ApplyTo:=wdListApplyToWholeList
    For i = 1 To n
        Set p = Selection.Paragraphs(i)
        p.LeftIndent = lefts(i)
        p.FirstLineIndent = firsts(i)
        If cuLefts(i) > -100 And cuLefts(i) < 1000 Then
            p.CharacterUnitLeftIndent = cuLefts(i)
        End If
        If cuFirsts(i) > -100 And cuFirsts(i) < 1000 Then
            p.CharacterUnitFirstLineIndent = cuFirsts(i)
        End If
    Next i
    StatusBar = "编号已继续。"
    Exit Sub
EH:
    ReportError "继续编号", Err.Number, Err.Description
End Sub

' 把范围内的列表编号值拼成 "1,2,3"，供测试使用。
Private Function ListValuesText(ByVal rng As Range) As String
    Dim p As Paragraph
    Dim s As String
    For Each p In rng.Paragraphs
        s = s & CStr(p.Range.ListFormat.ListValue) & ","
    Next p
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    ListValuesText = s
End Function

' 测试入口：继续编号与重新开始于1互为逆操作。
' 先建列表 A（编号 1,2,3），用正文段落隔开，再建列表 B（默认接着 A，为 4,5,6）；
' 对 B 先“重新开始于1”得到 1,2,3，再“继续编号”应恢复 4,5,6。
Public Function TestListNumberingAdjust() As String
    Dim i As Long
    Dim aStart As Long, bStart As Long, bEnd As Long
    Dim rng As Range
    Dim before As String, restarted As String, continued As String
    On Error GoTo EH
    Selection.EndOf Unit:=wdStory
    Selection.TypeParagraph
    aStart = Selection.Start
    For i = 1 To 3
        Selection.Style = ActiveDocument.Styles(wdStyleListNumber)
        Selection.TypeText "A" & CStr(i)
        Selection.TypeParagraph
    Next i
    Selection.Style = ActiveDocument.Styles(wdStyleNormal)
    Selection.TypeText "分隔"
    Selection.TypeParagraph
    bStart = Selection.Start
    For i = 1 To 3
        Selection.Style = ActiveDocument.Styles(wdStyleListNumber)
        Selection.TypeText "B" & CStr(i)
        Selection.TypeParagraph
    Next i
    bEnd = Selection.Start
    Set rng = ActiveDocument.Range(bStart, bEnd)
    rng.Select
    before = ListValuesText(rng)
    RestartListAtOne Nothing
    Set rng = ActiveDocument.Range(bStart, bEnd)
    rng.Select
    restarted = ListValuesText(rng)
    ContinueListNumbering Nothing
    Set rng = ActiveDocument.Range(bStart, bEnd)
    rng.Select
    continued = ListValuesText(rng)
    TestListNumberingAdjust = "ok|" & before & "|" & restarted & "|" & continued & _
        "|" & CStr(restarted = "1,2,3") & "|" & CStr(continued = before) & _
        "|" & CStr(before <> "1,2,3")
    Exit Function
EH:
    TestListNumberingAdjust = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 缩略语条目：在“符号和缩略语说明”一章中插入或整理“缩略语 + 制表符 + 说明”。
' 空白行 → 写入“缩略语<Tab>说明”；有文字无制表符 → 追加“<Tab>说明”；
' 已有制表符 → 只套用“符号说明条目”样式。三种情况都套用该样式。
Public Sub InsertSymbolEntry(control As Object)
    Dim p As Paragraph
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    If Not StyleExists(STYLE_SYMBOL_ENTRY) Then
        MsgBox "模板中缺少样式：" & STYLE_SYMBOL_ENTRY, vbExclamation, "列表"
        Exit Sub
    End If
    Set p = GetSelectionParagraph()
    ApplySymbolEntryToParagraph p
    StatusBar = "已整理为缩略语条目。"
    Exit Sub
EH:
    ReportError "插入缩略语条目", Err.Number, Err.Description
End Sub

' 由 InsertSymbolEntry 与测试共用：按当前段落内容决定填充方式，再套样式。
Private Sub ApplySymbolEntryToParagraph(ByVal p As Paragraph)
    Dim r As Range
    Dim txt As String
    Set r = SymbolContentRange(p)
    txt = r.Text
    txt = Replace(Replace(Replace(txt, Chr$(13), ""), Chr$(7), ""), Chr$(11), "")
    If InStr(txt, vbTab) > 0 Then
        ' 已有“缩略语<Tab>说明”，只套样式。
    ElseIf Len(Trim$(txt)) = 0 Then
        r.Text = "缩略语" & vbTab & "说明"
    Else
        r.Text = txt & vbTab & "说明"
    End If
    p.Style = ActiveDocument.Styles(STYLE_SYMBOL_ENTRY)
    Set r = SymbolContentRange(p)
    r.Select
    Selection.Collapse wdCollapseEnd
End Sub

' 段落内容范围：去掉段落标记（表格末段可能同时带单元格标记）。
Private Function SymbolContentRange(ByVal p As Paragraph) As Range
    Dim r As Range
    Dim lastChar As String
    Set r = p.Range.Duplicate
    If r.Characters.Count = 0 Then
        Set SymbolContentRange = r
        Exit Function
    End If
    lastChar = r.Characters(r.Characters.Count).Text
    If InStr(lastChar, Chr$(13)) > 0 Or InStr(lastChar, Chr$(7)) > 0 Then
        r.MoveEnd Unit:=wdCharacter, Count:=-1
    End If
    Set SymbolContentRange = r
End Function

' 测试入口：三种情况依次验证（空行、有文字无制表符、已有制表符）。
Public Function TestSymbolEntry() As String
    Dim p As Paragraph
    Dim r As Range
    Dim t1 As String, t2 As String, t3 As String
    Dim okStyle As Boolean
    On Error GoTo EH
    Selection.EndOf Unit:=wdStory
    ' 说明：Paragraphs.Add 的返回值可能指向插入前的末段，统一按索引重取末段。
    ' 情况 1：空行
    Set p = ActiveDocument.Paragraphs.Add
    Set p = ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count)
    Set r = SymbolContentRange(p)
    r.Text = ""
    ApplySymbolEntryToParagraph p
    t1 = Replace(Replace(p.Range.Text, Chr$(13), ""), Chr$(7), "")
    ' 情况 2：已有文字，无制表符
    Set p = ActiveDocument.Paragraphs.Add
    Set p = ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count)
    Set r = SymbolContentRange(p)
    r.Text = "CPU"
    ApplySymbolEntryToParagraph p
    t2 = Replace(Replace(p.Range.Text, Chr$(13), ""), Chr$(7), "")
    ' 情况 3：已有制表符
    Set p = ActiveDocument.Paragraphs.Add
    Set p = ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count)
    Set r = SymbolContentRange(p)
    r.Text = "GPU" & vbTab & "图形处理器"
    ApplySymbolEntryToParagraph p
    t3 = Replace(Replace(p.Range.Text, Chr$(13), ""), Chr$(7), "")
    okStyle = (PrimaryStyleName(p.Style) = STYLE_SYMBOL_ENTRY)
    TestSymbolEntry = "ok|" & t1 & "|" & t2 & "|" & t3 & "|" & CStr(okStyle)
    Exit Function
EH:
    TestSymbolEntry = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function ReferenceStyleForTag(ByVal tag As String) As String
    If LCase$(Trim$(tag)) = "author-year" Then
        ReferenceStyleForTag = STYLE_REF_ENTRY_AY
    Else
        ReferenceStyleForTag = STYLE_REF_ENTRY
    End If
End Function

Public Sub ApplyReferenceList(control As Object)
    Dim p As Paragraph
    Dim styleName As String
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    If Not control Is Nothing Then
        styleName = ReferenceStyleForTag(CStr(control.Tag))
    Else
        styleName = STYLE_REF_ENTRY
    End If
    If Not StyleExists(styleName) Then
        MsgBox "模板中缺少样式：" & styleName, vbExclamation, "列表"
        Exit Sub
    End If
    For Each p In Selection.Paragraphs
        p.Style = ActiveDocument.Styles(styleName)
    Next p
    Exit Sub
EH:
    ReportError "转换为参考文献格式", Err.Number, Err.Description
End Sub

Public Function ApplyParagraphStyleByName(ByVal styleName As String) As Boolean
    On Error GoTo EH
    If Not EnsureReady() Then Exit Function
    If Not StyleExists(styleName) Then
        MsgBox "模板中缺少样式：" & styleName, vbExclamation, "结构"
        Exit Function
    End If
    Selection.Paragraphs(1).Style = ActiveDocument.Styles(styleName)
    ApplyParagraphStyleByName = True
    Exit Function
EH:
    ReportError "设置段落样式", Err.Number, Err.Description
End Function

' 插入图片：嵌入型，套“图片”样式
Public Sub InsertPicture(control As Object)
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    If Dialogs(wdDialogInsertPicture).Show <> -1 Then Exit Sub
    If Selection.InlineShapes.Count = 0 Then Exit Sub
    Selection.Paragraphs(1).Style = ActiveDocument.Styles(STYLE_IMAGE)
    Exit Sub
EH:
    ReportError "插入图片", Err.Number, Err.Description
End Sub

' 插入表格：复用 Word 插入表格对话框，默认三线表
Public Sub InsertTable(control As Object)
    Dim t As Table
    Dim c As Cell
    Dim p As Paragraph
    Dim cellStyle As Style

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    If Dialogs(wdDialogTableInsertTable).Show <> -1 Then Exit Sub
    If Selection.Tables.Count = 0 Then Exit Sub
    Set t = Selection.Tables(1)
    t.Style = ActiveDocument.Styles("三线表")
    t.Rows(1).HeadingFormat = True
    ' 单元格文字统一用“表格文字”（无段前/段后距）
    On Error Resume Next
    Set cellStyle = ActiveDocument.Styles("表格文字")
    If Not cellStyle Is Nothing Then
        For Each c In t.Range.Cells
            c.VerticalAlignment = wdCellAlignVerticalCenter
            For Each p In c.Range.Paragraphs
                p.Style = cellStyle
            Next p
        Next c
    End If
    On Error GoTo EH
    Exit Sub
EH:
    ReportError "插入表格", Err.Number, Err.Description
End Sub



' 行内代码
Public Sub InsertInlineCode(control As Object)
    Dim p As Paragraph

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    Selection.Style = ActiveDocument.Styles("行内代码")
    If Selection.Type = wdSelectionIP Then Selection.TypeText "Hello World!"
    ' 行内代码不引入段前/段后距（字符样式本身不管段落间距）
    Set p = Selection.Paragraphs(1)
    p.SpaceBefore = 0
    p.SpaceAfter = 0
    Exit Sub
EH:
    ReportError "插入行内代码", Err.Number, Err.Description
End Sub

' 行间代码
Public Sub InsertBlockCode(control As Object)
    Dim p As Paragraph
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    Set p = Selection.Paragraphs(1)
    p.Style = ActiveDocument.Styles("行间代码")
    If Len(CleanText(p.Range.Text)) = 0 Then p.Range.InsertBefore "Hello World!"
    Exit Sub
EH:
    ReportError "插入行间代码", Err.Number, Err.Description
End Sub

' 测试入口：逐级设置标题样式
Public Function TestChapterReferenceItems() As String
    Dim p As Paragraph
    Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long, na As Long

    On Error GoTo EH
    For Each p In ActiveDocument.Paragraphs
        Select Case PrimaryStyleName(p.Style)
            Case "标题 1"
                n1 = n1 + 1
            Case "标题 2"
                n2 = n2 + 1
            Case "标题 3"
                n3 = n3 + 1
            Case "标题 4"
                n4 = n4 + 1
            Case "附录标题"
                na = na + 1
        End Select
    Next p
    TestChapterReferenceItems = "ok|" & CStr(n1) & "|" & CStr(n2) & "|" & _
        CStr(n3) & "|" & CStr(n4) & "|" & CStr(na)
    Exit Function
EH:
    TestChapterReferenceItems = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function TestChapterReferenceCode() As String
    TestChapterReferenceCode = "ok|" & ChapterReferenceFieldCode("THU_CH_0001") & _
        "|" & ChapterReferenceFieldCode("THU_CH_0002")
End Function

Public Function TestChapterReferenceSelection() As String
    Dim doc As Document
    Dim p As Paragraph
    Dim rng As Range
    Dim f As Field
    Dim firstIdx As Long
    Dim lastIdx As Long
    Dim i As Long
    Dim bmName As String
    Dim expected As String
    Dim currentText As String
    Dim resultText As String

    On Error GoTo EH
    Set doc = ActiveDocument

    For i = 1 To doc.Paragraphs.Count
        Set p = doc.Paragraphs(i)
        If PrimaryStyleName(p.Style) = "标题 1" Then
            If firstIdx = 0 Then firstIdx = i
            lastIdx = i
        End If
    Next i
    If firstIdx = 0 Then
        TestChapterReferenceSelection = "FAIL|no-heading"
        Exit Function
    End If

    Set p = ChapterParagraphAt(doc.Paragraphs(firstIdx).Range.Start)
    If PrimaryStyleName(p.Style) <> "标题 1" Then
        TestChapterReferenceSelection = "FAIL|resolve"
        Exit Function
    End If
    expected = GetHeadingListString(p)
    bmName = EnsureChapterBookmark(p)

    Set p = doc.Paragraphs(lastIdx)
    currentText = GetHeadingListString(p)

    Set rng = doc.Content
    rng.Collapse wdCollapseEnd
    rng.InsertAfter vbCr
    Set rng = doc.Content
    rng.Collapse wdCollapseEnd
    Set f = doc.Fields.Add(rng, wdFieldEmpty, ChapterReferenceFieldCode(bmName), False)
    f.Update
    resultText = CleanText(f.Result.Text)
    f.Delete

    TestChapterReferenceSelection = "ok|" & resultText & "|" & expected & "|" & _
        currentText & "|" & CStr(ActiveDocument.Bookmarks.Exists(bmName))
    Exit Function
EH:
    TestChapterReferenceSelection = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function TestHeadingStyles() As String
    Dim names As Variant
    Dim i As Long
    Dim p As Paragraph
    Dim n As Long
    names = Array("标题 1", "标题 2", "标题 3", "标题 4", _
                  "附录标题", "附录一级节标题", "附录二级节标题", "附录三级节标题")
    For i = LBound(names) To UBound(names)
        Set p = ActiveDocument.Paragraphs.Add
        p.Range.Select
        If Not ApplyHeadingStyleByName(CStr(names(i))) Then
            TestHeadingStyles = "FAIL|" & CStr(names(i))
            Exit Function
        End If
        If PrimaryStyleName(p.Style) <> CStr(names(i)) Then
            TestHeadingStyles = "FAIL_STYLE|" & CStr(names(i)) & "|" & PrimaryStyleName(p.Style)
            Exit Function
        End If
        n = n + 1
    Next i
    TestHeadingStyles = "ok|" & CStr(n)
End Function

' 章节引用样式：标题 1-4 与附录标题
Public Function IsChapterReferenceStyle(ByVal styleName As String) As Boolean
    Select Case styleName
        Case "标题 1", "标题 2", "标题 3", "标题 4", "附录标题", _
             "附录一级节标题", "附录二级节标题", "附录三级节标题"
            IsChapterReferenceStyle = True
    End Select
End Function

' 章节引用域：指向具体标题书签的 REF 编号域（与 Word 原生交叉引用同一机制）
Public Function ChapterReferenceFieldCode(ByVal bmName As String) As String
    ChapterReferenceFieldCode = "REF " & bmName & " \n \h"
End Function

' 取标题段落对应的书签；没有则新建，保证引用指向这一处标题而不是最近标题
Private Function EnsureChapterBookmark(p As Paragraph) As String
    Dim bm As Bookmark
    Dim r As Range
    Dim bmName As String

    For Each bm In ActiveDocument.Bookmarks
        If Left$(bm.Name, Len(BM_CH_PREFIX)) = BM_CH_PREFIX Then
            If bm.Range.Start >= p.Range.Start And bm.Range.Start < p.Range.End Then
                EnsureChapterBookmark = bm.Name
                Exit Function
            End If
        End If
    Next bm

    bmName = GetUniqueBookmarkName(BM_CH_PREFIX)
    Set r = p.Range.Duplicate
    If r.End > r.Start Then r.MoveEnd wdCharacter, -1
    If r.Start >= r.End Then Set r = p.Range.Duplicate
    ActiveDocument.Bookmarks.Add bmName, r
    EnsureChapterBookmark = bmName
End Function

' 按标题段落的起始位置取回段落（picker 只保存段落起始位置）
Private Function ChapterParagraphAt(ByVal pos As Long) As Paragraph
    Dim p As Paragraph
    Set p = ActiveDocument.Range(pos, pos).Paragraphs(1)
    If p.Range.Start <> pos Then
        Set p = ActiveDocument.Range(pos, pos + 1).Paragraphs(1)
    End If
    Set ChapterParagraphAt = p
End Function

' 引用章节：列出现有标题，选中后插入指向该标题书签的 REF 编号域
Public Sub InsertChapterReference(control As IRibbonControl)
    Dim p As Paragraph
    Dim code As String
    Dim key As String
    Dim pos As Long
    Dim idx As Long
    Dim first As Boolean
    Dim bmNames() As String

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub

    ResetPicker
    For Each p In ActiveDocument.Paragraphs
        If IsChapterReferenceStyle(PrimaryStyleName(p.Style)) Then
            AddPickerItem GetHeadingListString(p) & "　" & CleanText(p.Range.Text), CStr(p.Range.Start)
        End If
    Next p

    If PickerItemCount() = 0 Then
        MsgBox "没有找到章节标题。", vbInformation, "引用章节"
        Exit Sub
    End If
    If Not ShowPicker("引用章节", "选择要引用的章节（可多选）：") Then Exit Sub

    ' 先给选中的标题建好书签，再统一插域，避免插入过程中段落位置变化
    ReDim bmNames(0 To gPickerSelectedCount - 1)
    For idx = 0 To gPickerSelectedCount - 1
        key = gPickerItems(gPickerSelected(idx)).Key
        pos = CLng(key)
        Set p = ChapterParagraphAt(pos)
        bmNames(idx) = EnsureChapterBookmark(p)
    Next idx

    first = True
    For idx = 0 To gPickerSelectedCount - 1
        If Not first Then Selection.TypeText "、"
        code = ChapterReferenceFieldCode(bmNames(idx))
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
                             Text:=code, PreserveFormatting:=False
        Selection.Collapse wdCollapseEnd
        first = False
    Next idx
    Exit Sub
EH:
    ReportError "引用章节", Err.Number, Err.Description
End Sub

' 构建冒烟测试入口：Application.Run("TestSmoke")
Public Function TestSmoke() As String
    Dim doc As Document
    Dim msg As String

    On Error GoTo EH
    If Documents.Count = 0 Then
        TestSmoke = "FAIL|no-document"
        Exit Function
    End If
    Set doc = ActiveDocument
    msg = "OK"
    msg = msg & "|cc=" & CStr(doc.ContentControls.Count)
    msg = msg & "|sections=" & CStr(doc.Sections.Count)
    msg = msg & "|styles=" & CStr(doc.Styles.Count)
    msg = msg & "|formula=" & CStr(StyleExists(STYLE_FORMULA))
    msg = msg & "|citation=" & CStr(StyleExists(STYLE_CITATION))
    msg = msg & "|refentry=" & CStr(StyleExists(STYLE_REF_ENTRY))
    msg = msg & "|macros=" & CStr(Application.Version)
    TestSmoke = msg
    Exit Function
EH:
    TestSmoke = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：修复旧版附录复位域
Public Function TestRepairAppendixResetFields() As String
    Dim p As Paragraph
    Dim f As Field
    Dim lastField As Field
    Dim n As Long

    On Error GoTo EH
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Set p = Selection.Paragraphs(1)
    p.Range.InsertBefore "附录复位域测试："
    Set f = ActiveDocument.Fields.Add(Range:=p.Range, Type:=wdFieldEmpty, _
                                      Text:="SEQ 表 0 \h", PreserveFormatting:=False)
    f.Update
    n = RepairAppendixResetFields(True)
    For Each f In p.Range.Fields
        Set lastField = f
    Next f
    TestRepairAppendixResetFields = "ok|" & CStr(n) & "|" & _
        CStr(InStr(p.Range.Text, "错误") = 0) & "|" & lastField.Code.Text
    Exit Function
EH:
    TestRepairAppendixResetFields = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：结构组插入章节/附录
Public Function TestStructureMacros() As String
    Dim n0 As Long
    Dim n1 As Long

    On Error GoTo EH
    Selection.EndOf Unit:=wdStory
    n0 = ActiveDocument.Sections.Count
    InsertChapter Nothing
    InsertAppendix Nothing
    n1 = ActiveDocument.Sections.Count
    TestStructureMacros = "ok|" & CStr(n1 - n0)
    Exit Function
EH:
    TestStructureMacros = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：正文按钮处理选中多段
Public Function TestBodyStyleSelection() As String
    Dim p As Paragraph
    Dim i As Long
    Dim startPos As Long
    Dim ok As Boolean

    On Error GoTo EH
    Selection.EndOf Unit:=wdStory
    startPos = Selection.Start
    For i = 1 To 3
        Selection.TypeParagraph
        Selection.TypeText "正文测试" & CStr(i)
    Next i
    Selection.SetRange startPos, ActiveDocument.Content.End
    ApplyBodyStyle Nothing
    ok = (Selection.Paragraphs.Count >= 3)
    For Each p In Selection.Paragraphs
        If PrimaryStyleName(p.Style) <> "正文" Then ok = False
    Next p
    TestBodyStyleSelection = "ok|" & CStr(Selection.Paragraphs.Count) & "|" & CStr(ok)
    Exit Function
EH:
    TestBodyStyleSelection = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：代码样式与列表样式
Public Function TestCodeAndList() As String
    Dim p As Paragraph
    Dim inlineOk As Boolean
    Dim blockOk As Boolean
    Dim numOk As Boolean
    Dim bulOk As Boolean
    Dim selOk As Boolean
    Dim inlineStyle As String
    Dim blockStyle As String
    Dim blockText As String
    Dim i As Long
    Dim startPos As Long

    On Error GoTo EH
    Selection.EndOf Unit:=wdStory
    Set p = ActiveDocument.Paragraphs.Add
    p.Range.Text = "print(1)"
    p.Range.Select
    InsertInlineCode Nothing
    inlineStyle = p.Range.Characters(1).Style.NameLocal
    inlineOk = (inlineStyle = "行内代码")

    Selection.EndOf Unit:=wdStory
    Selection.TypeParagraph
    Set p = Selection.Paragraphs(1)
    InsertBlockCode Nothing
    blockStyle = PrimaryStyleName(p.Style)
    blockText = CleanText(p.Range.Text)
    blockOk = (blockStyle = "行间代码") And (InStr(blockText, "Hello World!") > 0)

    Selection.EndOf Unit:=wdStory
    Selection.TypeParagraph
    Set p = Selection.Paragraphs(1)
    p.Range.Text = "有序项"
    p.Style = ActiveDocument.Styles(wdStyleListNumber)
    numOk = (Len(p.Range.ListFormat.ListString) > 0)

    Selection.EndOf Unit:=wdStory
    Selection.TypeParagraph
    Set p = Selection.Paragraphs(1)
    p.Range.Text = "无序项"
    p.Style = ActiveDocument.Styles(wdStyleListBullet)
    bulOk = (Len(p.Range.ListFormat.ListString) > 0)

    ' 选中多个段落，一次点击按钮应全部套用列表样式；首行顶格、续行缩进 2 字符。
    Selection.EndOf Unit:=wdStory
    startPos = Selection.Start
    For i = 1 To 3
        Selection.TypeParagraph
        Selection.TypeText "列表项" & CStr(i)
    Next i
    Selection.SetRange startPos, ActiveDocument.Content.End
    ApplyListStyleByTag "List Number"
    selOk = (Selection.Paragraphs.Count >= 3)
    For Each p In Selection.Paragraphs
        If PrimaryStyleName(p.Style) <> "列表编号" Then selOk = False
        If Len(p.Range.ListFormat.ListString) = 0 Then selOk = False
        If p.Format.CharacterUnitLeftIndent <> 0 Then selOk = False
        If p.Format.CharacterUnitFirstLineIndent <> -2 Then selOk = False
    Next p

    TestCodeAndList = "ok|" & CStr(inlineOk) & "|" & CStr(blockOk) & "|" & _
        CStr(numOk) & "|" & CStr(bulOk) & "|" & CStr(selOk)
    Exit Function
EH:
    TestCodeAndList = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：附录各级标题编号
Public Function TestAppendixNumbering() As String
    Dim names As Variant
    Dim i As Long
    Dim p As Paragraph
    Dim s As String

    On Error GoTo EH
    names = Array("附录标题", "附录一级节标题", "附录二级节标题", "附录三级节标题")
    For i = LBound(names) To UBound(names)
        Set p = ActiveDocument.Paragraphs.Add
        p.Range.Select
        If Not ApplyHeadingStyleByName(CStr(names(i))) Then
            TestAppendixNumbering = "FAIL_STYLE|" & CStr(names(i))
            Exit Function
        End If
        s = s & CStr(names(i)) & "=" & p.Range.ListFormat.ListString & ";"
    Next i
    TestAppendixNumbering = "ok|" & s
    Exit Function
EH:
    TestAppendixNumbering = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：三线表



Public Function TestThreeLineTable() As String
    Dim p As Paragraph
    Dim t As Table

    On Error GoTo EH
    Selection.EndOf Unit:=wdStory
    Selection.TypeParagraph
    Set p = Selection.Paragraphs(1)
    Set t = ActiveDocument.Tables.Add(p.Range, 3, 2)
    t.Style = ActiveDocument.Styles("三线表")
    t.Rows(1).HeadingFormat = True
    TestThreeLineTable = "ok|" & t.Style.NameLocal & "|" & CStr(t.Rows(1).HeadingFormat)
    Exit Function
EH:
    TestThreeLineTable = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：列表样式常量可用
Public Function TestListConstants() As String
    Dim ids As Variant
    On Error GoTo EH
    ids = Array(wdStyleListNumber, wdStyleListNumber2, wdStyleListNumber3, _
                wdStyleListNumber4, wdStyleListNumber5, _
                wdStyleListBullet, wdStyleListBullet2, wdStyleListBullet3, _
                wdStyleListBullet4, wdStyleListBullet5)
    TestListConstants = "ok|" & CStr(UBound(ids) + 1)
    Exit Function
EH:
    TestListConstants = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：参考文献两种格式样式可用
Public Function TestReferenceFormatStyles() As String
    Dim n As String
    Dim y As String
    On Error GoTo EH
    n = ReferenceStyleForTag("numeric")
    y = ReferenceStyleForTag("author-year")
    If Not StyleExists(n) Or Not StyleExists(y) Then
        TestReferenceFormatStyles = "FAIL|missing style"
        Exit Function
    End If
    TestReferenceFormatStyles = "ok|" & n & "|" & y
    Exit Function
EH:
    TestReferenceFormatStyles = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function
