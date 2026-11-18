Attribute VB_Name = "modCaption"
Option Explicit

' ============================================================================
' 题注：正文/附录的图注、表注，续表，图/表引用
' 约定：
'   图2.1　题名 / 表2.1　题名（标签与编号之间不留空格，编号与题名一个全角空格）
'   正文：STYLEREF 1 \s + "." + SEQ 图/表 \* ARABIC \s 1
'   附录：STYLEREF "附录标题" \s + "." + SEQ 附录图/附录表 \* ARABIC \s 1
'   书签只包编号（2.1 / A.1），供 REF 交叉引用。
' ============================================================================

' ---------------------------------------------------------------------------
' 图注
' ---------------------------------------------------------------------------

Public Sub InsertFigureCaption(control As Object, Optional ByVal silent As Boolean = False, Optional ByVal titleIn As String = "")
    Dim p As Paragraph
    Dim cap As Paragraph
    Dim numText As String
    Dim bmName As String
    Dim title As String
    Dim ur As Object

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    Set p = GetSelectionParagraph()
    If p Is Nothing Then
        MsgBox "请把光标放在图片所在段落。", vbExclamation, "图注"
        Exit Sub
    End If

    If silent Then
        title = titleIn
    Else
        title = CaptionTitleFromContext(p)
    End If
    If Len(Trim$(title)) = 0 Then title = "图题"

    If IsFigureCaptionStyle(p.Style.NameLocal) Then
        Set cap = p
    ElseIf HasFigureAt(p) Then
        p.Range.InsertParagraphAfter
        Set cap = p.Next
    Else
        Set cap = p
    End If

    gDirty = False
    BuildCaption cap, "图", STYLE_FIG_CAPTION, "SEQ 图 \* ARABIC \s 1", title
    numText = ExtractCaptionNumber(cap, "图")
    If Len(numText) = 0 Then
        MsgBox "未能读取图号，请点“更新域”后重试。", vbExclamation, "图注"
        Exit Sub
    End If
    bmName = GetUniqueBookmarkName(BM_FIG_PREFIX)
    If Not BookmarkText(bmName, cap.Range, numText) Then
        MsgBox "未能为图号建立引用书签。", vbExclamation, "图注"
        Exit Sub
    End If
    Exit Sub
EH:
    gDirty = False
    ReportError "插入图注", Err.Number, Err.Description
End Sub

' ---------------------------------------------------------------------------
' 子图注
' ---------------------------------------------------------------------------

' 插入子图题注：编号为 (a)、(b)...，不插入主图号，不进入插图清单。
Public Sub InsertSubfigureCaption(control As Object, Optional ByVal silent As Boolean = False, Optional ByVal titleIn As String = "")
    Dim p As Paragraph
    Dim cap As Paragraph
    Dim title As String

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    If Not StyleExists(STYLE_SUBFIG_CAPTION) Then
        MsgBox "当前文档缺少“子图注”样式，请用本模板新建文档。", vbExclamation, "子图注"
        Exit Sub
    End If

    Set p = GetSelectionParagraph()
    If p Is Nothing Then
        MsgBox "请把光标放在图片所在段落。", vbExclamation, "子图注"
        Exit Sub
    End If

    If silent Then
        title = titleIn
    Else
        title = CaptionTitleFromContext(p)
    End If
    If Len(Trim$(title)) = 0 Then title = "子图题"

    If p.Style.NameLocal = STYLE_SUBFIG_CAPTION Then
        Set cap = p
    ElseIf HasFigureAt(p) Then
        p.Range.InsertParagraphAfter
        Set cap = p.Next
    Else
        Set cap = p
    End If

    gDirty = False
    BuildSubfigureCaption cap, "SEQ 子图 \* alphabetic", title
    Exit Sub
EH:
    gDirty = False
    ReportError "插入子图注", Err.Number, Err.Description
End Sub

' 如果当前选中的是子图注，把该子图注的 SEQ 域改为从 1 重新开始。
' 供“列表→重新开始于1”复用，返回 True 表示已处理。
Public Function RestartSubfigureAtSelection() As Boolean
    Dim p As Paragraph
    Dim f As Field

    On Error GoTo EH
    If Selection.Paragraphs.Count = 0 Then Exit Function
    Set p = Selection.Paragraphs(1)
    If p.Style.NameLocal <> STYLE_SUBFIG_CAPTION Then Exit Function
    For Each f In p.Range.Fields
        If InStr(1, f.Code.Text, "SEQ 子图", vbTextCompare) > 0 Then
            f.Code.Text = " SEQ 子图 \r 1 \* alphabetic "
            f.Update
            ActiveDocument.Fields.Update
            RestartSubfigureAtSelection = True
            Exit Function
        End If
    Next f
    Exit Function
EH:
    ReportError "重新开始子图编号", Err.Number, Err.Description
End Function

' 生成“(a)　题名”；字母来自可更新的 SEQ 子图 域。
Private Sub BuildSubfigureCaption(cap As Paragraph, ByVal seqField As String, ByVal title As String)
    cap.Style = ActiveDocument.Styles(STYLE_SUBFIG_CAPTION)
    ClearParagraphKeepingMark cap
    cap.Range.Select
    Selection.Collapse wdCollapseStart
    Selection.TypeText "("
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
                         Text:=seqField, PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
    Selection.TypeText ")" & "　" & title
    cap.Range.Fields.Update
End Sub

' 统计正文里已有的子图注数量。测试前先记下基数，避免把模板自带的示例子图注算进去。
Private Function CountSubfigureCaptions() As Long
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        If p.Style.NameLocal = STYLE_SUBFIG_CAPTION Then
            CountSubfigureCaptions = CountSubfigureCaptions + 1
        End If
    Next p
End Function

' 按顺序取第 n 个子图注，供测试使用。
Private Function FindTestSubCaption(ByVal n As Long) As Paragraph
    Dim p As Paragraph
    Dim k As Long

    For Each p In ActiveDocument.Paragraphs
        If p.Style.NameLocal = STYLE_SUBFIG_CAPTION Then
            k = k + 1
            If k = n Then
                Set FindTestSubCaption = p
                Exit Function
            End If
        End If
    Next p
End Function

' 测试：连续编号与 (a) 格式。
Public Function TestSubfigureCaption() As String
    Dim parts As String
    Dim p As Paragraph
    Dim base As Long

    On Error GoTo EH
    base = CountSubfigureCaptions()
    parts = AddTestSubCaption("甲")
    parts = parts & "|" & AddTestSubCaption("乙")
    parts = parts & "|" & AddTestSubCaption("丙")
    parts = parts & "|" & AddTestSubCaption("丁")
    parts = parts & "|" & AddTestSubCaption("戊")
    parts = parts & "|" & AddTestSubCaption("己")
    Set p = FindTestSubCaption(base + 3)
    p.Range.Select
    Selection.Collapse wdCollapseStart
    RestartListAtOne Nothing
    ActiveDocument.Fields.Update
    parts = parts & "|" & CleanText(p.Range.Text)
    Set p = FindTestSubCaption(base + 4)
    parts = parts & "|" & CleanText(p.Range.Text)
    TestSubfigureCaption = "ok|" & parts
    Exit Function
EH:
    TestSubfigureCaption = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function AddTestSubCaption(ByVal title As String) As String
    Dim p As Paragraph
    Dim cap As Paragraph

    Set p = ActiveDocument.Content.Paragraphs.Add
    p.Range.InsertBefore title
    p.Range.Select
    InsertSubfigureCaption Nothing, True, title
    Set cap = Selection.Paragraphs(1)
    If cap.Range.Fields.Count <> 1 Then
        Err.Raise vbObjectError + 1601, , "子图注应只含 1 个 SEQ 域"
    End If
    If cap.Style.NameLocal <> STYLE_SUBFIG_CAPTION Then
        Err.Raise vbObjectError + 1603, , "子图注样式未应用"
    End If
    If InStr(1, cap.Range.Fields(1).Code.Text, "SEQ 子图", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1602, , "子图注域代码不是 SEQ 子图"
    End If
    AddTestSubCaption = CleanText(cap.Range.Text)
End Function

' ---------------------------------------------------------------------------
' 表注
' ---------------------------------------------------------------------------

Public Sub InsertTableCaption(control As Object, Optional ByVal silent As Boolean = False, Optional ByVal titleIn As String = "")
    Dim t As Table
    Dim p As Paragraph
    Dim cap As Paragraph
    Dim numText As String
    Dim bmName As String
    Dim title As String
    Dim ur As Object

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    Set p = GetSelectionParagraph()
    If p Is Nothing Then Exit Sub

    If silent Then
        title = titleIn
    Else
        title = CaptionTitleFromContext(p)
    End If
    If Len(Trim$(title)) = 0 Then title = "表题"

    If IsTableCaptionStyle(p.Style.NameLocal) Then
        Set cap = p
    ElseIf p.Range.Information(wdWithInTable) Then
        Set t = Selection.Tables(1)
        Set cap = InsertParagraphBeforeTable(t)
        If cap Is Nothing Then
            MsgBox "未能在表格上方插入表注段落。", vbExclamation, "表注"
            Exit Sub
        End If
    Else
        Set cap = p
    End If

    gDirty = False
    BuildCaption cap, "表", STYLE_TAB_CAPTION, "SEQ 表 \* ARABIC \s 1", title
    numText = ExtractCaptionNumber(cap, "表")
    If Len(numText) = 0 Then
        MsgBox "未能读取表号，请点“更新域”后重试。", vbExclamation, "表注"
        Exit Sub
    End If
    bmName = GetUniqueBookmarkName(BM_TAB_PREFIX)
    If Not BookmarkText(bmName, cap.Range, numText) Then
        MsgBox "未能为表号建立引用书签。", vbExclamation, "表注"
        Exit Sub
    End If
    Exit Sub
EH:
    gDirty = False
    ReportError "插入表注", Err.Number, Err.Description
End Sub

' ---------------------------------------------------------------------------
' 续表
' ---------------------------------------------------------------------------

' Public Sub InsertContinuedTableCaption(control As Object, Optional ByVal silent As Boolean = False, Optional ByVal titleIn As String = "")
'     Dim t As Table
'     Dim p As Paragraph
'     Dim cap As Paragraph
'     Dim prevNum As String
'     Dim prevBm As String
'     Dim title As String
' 
'     On Error GoTo EH
'     If Not EnsureReady() Then Exit Sub
'     Set p = GetSelectionParagraph()
'     If p Is Nothing Then Exit Sub
' 
'     If silent Then
'         title = titleIn
'     Else
'         title = CaptionTitleFromContext(p)
'     End If
'     If Len(Trim$(title)) = 0 Then title = "续表题"
' 
'     If p.Range.Information(wdWithInTable) Then
'         ' 光标/选区在表格内：在表格上方插续表题注。
'         Set t = Selection.Tables(1)
'         Set cap = InsertParagraphBeforeTable(t)
'         If cap Is Nothing Then
'             MsgBox "未能在表格上方插入续表题注段落。", vbExclamation, "续表"
'             Exit Sub
'         End If
'     Else
'         ' 与表注一致：选区在表格外时，直接把当前段作为题注段。
'         Set t = FindTableAfterParagraph(p)
'         If t Is Nothing Then
'             MsgBox "请把光标放在表格内，或选中表格上方的题名文字。", vbExclamation, "续表"
'             Exit Sub
'         End If
'         Set cap = p
'     End If
' 
'     prevNum = FindPreviousTableNumber(t, prevBm)
'     If Len(prevNum) = 0 Then
'         MsgBox "没有找到上一张表的表注，无法确定续表编号。" & vbCrLf & _
'                "请先给上一张表插入表注。", vbExclamation, "续表"
'         Exit Sub
'     End If
' 
'     gDirty = False
'     WriteContinuedCaption cap, prevNum, prevBm, title
'     ' 表头重复：把首行标记为标题行，跨页自动重复
'     If t.Rows.Count > 1 Then t.Rows(1).HeadingFormat = True
'     Exit Sub
' EH:
'     ReportError "插入续表", Err.Number, Err.Description
' End Sub

'
' 在段落之后查找最近的表格；没有则返回 Nothing。
Private Function FindTableAfterParagraph(ByVal p As Paragraph) As Table
    Dim t As Table
    Dim best As Table
    Dim bestStart As Long

    For Each t In ActiveDocument.Tables
        If t.Range.Start >= p.Range.End Then
            If best Is Nothing Then
                Set best = t
                bestStart = t.Range.Start
            ElseIf t.Range.Start < bestStart Then
                Set best = t
                bestStart = t.Range.Start
            End If
        End If
    Next t
    Set FindTableAfterParagraph = best
End Function

Private Sub WriteContinuedCaption(cap As Paragraph, ByVal prevNum As String, _
                                  ByVal prevBm As String, ByVal title As String)
    Dim r As Range
    Dim f As Field

    If IsAppendixAt(cap.Range) Then
        cap.Style = ActiveDocument.Styles(STYLE_APP_TAB_CAPTION)
    Else
        cap.Style = ActiveDocument.Styles(STYLE_TAB_CAPTION)
    End If
    ClearParagraphKeepingMark cap
    Set r = cap.Range
    r.Collapse wdCollapseStart
    r.Text = THU_CONTINUED_PREFIX
    If Len(prevBm) > 0 Then
        Set r = ActiveDocument.Range(cap.Range.End - 1, cap.Range.End - 1)
        Set f = ActiveDocument.Fields.Add(Range:=r, Type:=wdFieldEmpty, _
                                          Text:="REF " & prevBm & " \h", _
                                          PreserveFormatting:=False)
        f.Update
        Set r = ActiveDocument.Range(cap.Range.End - 1, cap.Range.End - 1)
        r.InsertAfter THU_CONTINUED_SUFFIX & "　" & title
    Else
        Set r = ActiveDocument.Range(cap.Range.End - 1, cap.Range.End - 1)
        r.InsertAfter prevNum & THU_CONTINUED_SUFFIX & "　" & title
    End If
    cap.Range.Fields.Update
End Sub

' ---------------------------------------------------------------------------
' 图/表引用
' ---------------------------------------------------------------------------

Public Sub InsertFigureReference(control As IRibbonControl)
    InsertNumberedReference "图", BM_FIG_PREFIX, "引用图", STYLE_FIG_CAPTION
End Sub

Public Sub InsertTableReference(control As IRibbonControl)
    InsertNumberedReference "表", BM_TAB_PREFIX, "引用表", STYLE_TAB_CAPTION
End Sub

Private Function EnsureCaptionBookmark(p As Paragraph, ByVal label As String, _
                                       ByVal bmPrefix As String) As String
    Dim numText As String
    Dim bmName As String
    numText = ExtractCaptionNumber(p, label)
    If Len(numText) = 0 Then Exit Function
    bmName = FindBookmarkByText(bmPrefix, numText)
    If Len(bmName) > 0 Then
        EnsureCaptionBookmark = bmName
        Exit Function
    End If
    bmName = GetUniqueBookmarkName(bmPrefix)
    If BookmarkText(bmName, p.Range, numText) Then EnsureCaptionBookmark = bmName
End Function

Private Sub InsertNumberedReference(ByVal prefixText As String, ByVal bmPrefix As String, _
                                    ByVal dlgTitle As String, ByVal styleName As String)
    Dim p As Paragraph
    Dim numText As String
    Dim bmName As String
    Dim preview As String
    Dim k As Long
    Dim idx As Long
    Dim first As Boolean

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub

    ResetPicker
    For Each p In ActiveDocument.Paragraphs
        If (p.Style.NameLocal = styleName) Or _
           (styleName = STYLE_FIG_CAPTION And p.Style.NameLocal = STYLE_APP_FIG_CAPTION) Or _
           (styleName = STYLE_TAB_CAPTION And p.Style.NameLocal = STYLE_APP_TAB_CAPTION) Then
            numText = ExtractCaptionNumber(p, prefixText)
            If Len(numText) > 0 Then
                bmName = FindBookmarkByText(bmPrefix, numText)
                If Len(bmName) = 0 Then bmName = EnsureCaptionBookmark(p, prefixText, bmPrefix)
                If Len(bmName) > 0 Then
                    preview = CleanText(p.Range.Text)
                    If Len(preview) > 60 Then preview = Left$(preview, 60) & "…"
                    AddPickerItem preview, bmName
                End If
            End If
        End If
    Next p

    If PickerItemCount() = 0 Then
        MsgBox "没有找到带编号的" & prefixText & "题注。" & vbCrLf & _
               "请先插入题注或更新域。", vbExclamation, dlgTitle
        Exit Sub
    End If
    If Not ShowPicker(dlgTitle, "勾选要引用的" & prefixText & "（可多选）：") Then Exit Sub

    first = True
    For k = 0 To gPickerSelectedCount - 1
        idx = gPickerSelected(k)
        bmName = gPickerItems(idx).Key
        If Len(bmName) > 0 Then
            If Not first Then Selection.TypeText "、"
            InsertRefField prefixText, bmName
            first = False
        End If
    Next k
    Exit Sub
EH:
    ReportError dlgTitle, Err.Number, Err.Description
End Sub

Public Function NormalizeCaptionNumber(ByVal s As String, ByVal label As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, "（", "(")
    t = Replace(t, "）", ")")
    t = Replace(t, "(", "")
    t = Replace(t, ")", "")
    If Len(label) > 0 Then
        If Left$(t, Len(label)) = label Then t = Mid$(t, Len(label) + 1)
    End If
    t = Replace(t, " ", "")
    t = Replace(t, "　", "")
    NormalizeCaptionNumber = t
End Function

' ---------------------------------------------------------------------------
' 公共工具
' ---------------------------------------------------------------------------

' 段落里是否有插图。嵌入型图片计入 InlineShapes；而 Word 原生的
' wordprocessingShape（本模板的示例插图占位）即使 XML 写成 wp:inline，
' 在对象模型里也只出现在 ShapeRange 中，所以两种都要判断。
Private Function HasFigureAt(ByVal p As Paragraph) As Boolean
    On Error GoTo EH
    If p.Range.InlineShapes.Count > 0 Then
        HasFigureAt = True
    ElseIf p.Range.ShapeRange.Count > 0 Then
        HasFigureAt = True
    End If
    Exit Function
EH:
    HasFigureAt = False
End Function

' 取第一个插图所在段落：先按嵌入型图片找，找不到再按 ShapeRange 找。
Private Function FindFigureParagraph() As Paragraph
    Dim p As Paragraph

    On Error GoTo EH
    If ActiveDocument.InlineShapes.Count > 0 Then
        Set FindFigureParagraph = ActiveDocument.InlineShapes(1).Range.Paragraphs(1)
        Exit Function
    End If
    For Each p In ActiveDocument.Paragraphs
        If p.Range.ShapeRange.Count > 0 Then
            Set FindFigureParagraph = p
            Exit Function
        End If
    Next p
    Exit Function
EH:
    Set FindFigureParagraph = Nothing
End Function

Private Function CaptionTitleFromContext(ByVal p As Paragraph) As String
    Dim t As String

    If Selection.Type <> wdSelectionIP Then t = CleanText(Selection.Text)
    If Len(t) = 0 And Not HasFigureAt(p) Then t = CleanText(p.Range.Text)
    CaptionTitleFromContext = t
End Function

Private Sub BuildCaption(cap As Paragraph, ByVal label As String, ByVal styleName As String, _
                         ByVal seqField As String, ByVal title As String)
    Dim chapterCode As String

    If IsAppendixAt(cap.Range) Then
        chapterCode = "STYLEREF ""附录标题"" \s"
        ' 附录题注用独立 SEQ 标识符：\c "图"/\c "表" 清单只收正文的
        ' SEQ 图 / SEQ 表，附录图表因此不会进入插图清单、附表清单。
        If label = "图" Then seqField = "SEQ 附录图 \* ARABIC \s 1"
        If label = "表" Then seqField = "SEQ 附录表 \* ARABIC \s 1"
        If styleName = STYLE_FIG_CAPTION Then styleName = STYLE_APP_FIG_CAPTION
        If styleName = STYLE_TAB_CAPTION Then styleName = STYLE_APP_TAB_CAPTION
    Else
        chapterCode = "STYLEREF 1 \s"
    End If

    cap.Style = ActiveDocument.Styles(styleName)
    ClearParagraphKeepingMark cap
    cap.Range.Select
    Selection.Collapse wdCollapseStart
    Selection.TypeText label
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
                         Text:=chapterCode, PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
    Selection.TypeText "."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
                         Text:=seqField, PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
    Selection.TypeText "　" & title
    If styleName = STYLE_APP_FIG_CAPTION Or styleName = STYLE_APP_TAB_CAPTION Then RefreshAppendixSequences
    cap.Range.Fields.Update
End Sub

Private Function IsFigureCaptionStyle(ByVal name As String) As Boolean
    IsFigureCaptionStyle = (name = STYLE_FIG_CAPTION Or name = STYLE_APP_FIG_CAPTION)
End Function

Private Function IsTableCaptionStyle(ByVal name As String) As Boolean
    IsTableCaptionStyle = (name = STYLE_TAB_CAPTION Or name = STYLE_APP_TAB_CAPTION)
End Function

' 清空段落内容但保留段落标记
Private Sub ClearParagraphKeepingMark(p As Paragraph)
    Dim r As Range
    Set r = p.Range
    If r.End > r.Start Then r.MoveEnd wdCharacter, -1
    r.Text = ""
End Sub

' 在整张表格前插入一个空段落并返回它（不进入第一个单元格）
Private Function InsertParagraphBeforeTable(t As Table) As Paragraph
    Dim r As Range
    Dim p As Paragraph

    ' 在表格起点前一个字符处插入段落标记；这是在表格外新增段落的可靠位置
    If t.Range.Start > 0 Then
        Set r = ActiveDocument.Range(t.Range.Start - 1, t.Range.Start - 1)
        r.InsertAfter vbCr
        For Each p In ActiveDocument.Paragraphs
            If p.Range.End = t.Range.Start Then
                If Not p.Range.Information(wdWithInTable) Then
                    Set InsertParagraphBeforeTable = p
                    Exit Function
                End If
            End If
        Next p
    End If

    ' 兜底：旧方法
    t.Range.Select
    Selection.Collapse wdCollapseStart
    Selection.InsertParagraphBefore
    Selection.MoveUp Unit:=wdParagraph, Count:=1
    Set InsertParagraphBeforeTable = Selection.Paragraphs(1)
End Function

' 从题注段落取编号，如“图2.1　题名” -> 2.1
Public Function ExtractCaptionNumber(p As Paragraph, ByVal label As String) As String
    Dim txt As String
    Dim pos As Long
    Dim startPos As Long

    txt = CleanText(p.Range.Text)
    If Left$(txt, Len(label)) <> label Then Exit Function
    startPos = Len(label) + 1
    pos = InStr(startPos, txt, "　")
    If pos = 0 Then pos = InStr(startPos, txt, " ")
    If pos = 0 Then pos = Len(txt) + 1
    ExtractCaptionNumber = Trim$(Mid$(txt, startPos, pos - startPos))
End Function

' 找当前表格之前最近的一个表注，返回编号与书签名
Public Function IsContinuedTableCaption(ByVal text As String) As Boolean
    If Left$(text, 2) = "续表" Then
        IsContinuedTableCaption = True
    ElseIf Left$(text, 1) = "表" Then
        IsContinuedTableCaption = (Right$(Split(text, "　")(0), 3) = "（续）")
    End If
End Function

Private Function FindPreviousTableNumber(t As Table, ByRef bmName As String) As String
    Dim p As Paragraph
    Dim prevP As Paragraph
    Dim numText As String

    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= t.Range.Start Then Exit For
        If IsTableCaptionStyle(p.Style.NameLocal) And Not IsContinuedTableCaption(CleanText(p.Range.Text)) Then
            numText = ExtractCaptionNumber(p, "表")
            If Len(numText) > 0 Then
                FindPreviousTableNumber = numText
                Set prevP = p
            End If
        End If
    Next p

    If Len(FindPreviousTableNumber) > 0 Then
        bmName = FindBookmarkByText(BM_TAB_PREFIX, FindPreviousTableNumber)
        If Len(bmName) = 0 And Not prevP Is Nothing Then
            bmName = EnsureCaptionBookmark(prevP, "表", BM_TAB_PREFIX)
        End If
    End If
End Function
' 测试入口：续表引用主表表注（不弹输入框，使用传入题名）。
' Public Function TestContinuedTableCaption() As String
'     Dim t As Table
'     Dim p As Paragraph
'     Dim cap As Paragraph
'     Dim prevNum As String
'     Dim prevBm As String
'     Dim i As Long
' 
'     On Error GoTo EH
'     If ActiveDocument.Tables.Count = 0 Then
'         TestContinuedTableCaption = "NO_TABLE"
'         Exit Function
'     End If
'     For i = 1 To ActiveDocument.Tables.Count
'         Set t = ActiveDocument.Tables(i)
'         prevNum = FindPreviousTableNumber(t, prevBm)
'         If Len(prevNum) > 0 Then Exit For
'     Next i
'     If Len(prevNum) = 0 Then
'         TestContinuedTableCaption = "NO_PREV"
'         Exit Function
'     End If
' 
'     t.Range.Select
'     InsertContinuedTableCaption Nothing, True, "点击路径续表"
'     For Each p In ActiveDocument.Paragraphs
'         If IsTableCaptionStyle(p.Style.NameLocal) Then
'             If InStr(1, p.Range.Text, "点击路径续表", vbTextCompare) > 0 Then
'                 Set cap = p
'                 Exit For
'             End If
'         End If
'     Next p
'     If cap Is Nothing Then
'         TestContinuedTableCaption = "NO_CAPTION"
'         Exit Function
'     End If
'     TestContinuedTableCaption = "ok|" & CleanText(cap.Range.Text)
' 
'     ' 选中表格外的题名文字路径（表注式，不弹输入框）
'     For i = 1 To ActiveDocument.Tables.Count
'         Set t = ActiveDocument.Tables(i)
'         prevNum = FindPreviousTableNumber(t, prevBm)
'         If Len(prevNum) > 0 Then Exit For
'     Next i
'     If Len(prevNum) > 0 Then
'         Set p = InsertParagraphBeforeTable(t)
'         p.Range.InsertBefore "选中路径续表"
'         p.Range.Select
'         Selection.Collapse wdCollapseStart
'         InsertContinuedTableCaption Nothing, True, "选中路径续表"
'         Set cap = Nothing
'         For Each p In ActiveDocument.Paragraphs
'             If IsTableCaptionStyle(p.Style.NameLocal) Then
'                 If InStr(1, p.Range.Text, "选中路径续表", vbTextCompare) > 0 Then
'                     Set cap = p
'                     Exit For
'                 End If
'             End If
'         Next p
'         If Not cap Is Nothing Then
'             TestContinuedTableCaption = TestContinuedTableCaption & "|" & CleanText(cap.Range.Text)
'         End If
'     End If
' 
'     ' 空白行路径：题名留空，直接插入续表题注
'     For i = 1 To ActiveDocument.Tables.Count
'         Set t = ActiveDocument.Tables(i)
'         prevNum = FindPreviousTableNumber(t, prevBm)
'         If Len(prevNum) > 0 Then Exit For
'     Next i
'     If Len(prevNum) > 0 Then
'         Set p = InsertParagraphBeforeTable(t)
'         p.Range.Text = ""
'         p.Range.Select
'         InsertContinuedTableCaption Nothing, True, ""
'         Set cap = Nothing
'         For Each p In ActiveDocument.Paragraphs
'             If IsTableCaptionStyle(p.Style.NameLocal) Then
'                 If InStr(1, p.Range.Text, "续表题", vbTextCompare) > 0 Then
'                     Set cap = p
'                     Exit For
'                 End If
'             End If
'         Next p
'         If Not cap Is Nothing Then
'             TestContinuedTableCaption = TestContinuedTableCaption & "|" & CleanText(cap.Range.Text)
'         End If
'     End If
'     Exit Function
' EH:
'     TestContinuedTableCaption = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
' End Function

' 测试入口：列出表格题注可引用项（含为已有题注补书签）
Public Function TestTableReferenceItems() As String
    Dim p As Paragraph
    Dim numText As String
    Dim bmName As String
    Dim n As Long
    Dim first As String
    On Error GoTo EH
    For Each p In ActiveDocument.Paragraphs
        If IsTableCaptionStyle(p.Style.NameLocal) Then
            numText = ExtractCaptionNumber(p, "表")
            If Len(numText) > 0 Then
                bmName = FindBookmarkByText(BM_TAB_PREFIX, numText)
                If Len(bmName) = 0 Then bmName = EnsureCaptionBookmark(p, "表", BM_TAB_PREFIX)
                If Len(bmName) > 0 Then
                    n = n + 1
                    If Len(first) = 0 Then first = numText
                End If
            End If
        End If
    Next p
    TestTableReferenceItems = "ok|" & CStr(n) & "|" & first
    Exit Function
EH:
    TestTableReferenceItems = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：已有表注补书签后插入 REF 域
Public Function TestTableRefText() As String
    Dim p As Paragraph
    Dim bmName As String
    Dim r As Range
    Dim f As Field
    On Error GoTo EH
    For Each p In ActiveDocument.Paragraphs
        If IsTableCaptionStyle(p.Style.NameLocal) Then
            bmName = EnsureCaptionBookmark(p, "表", BM_TAB_PREFIX)
            If Len(bmName) > 0 Then
                Set p = ActiveDocument.Paragraphs.Add
                p.Range.Text = "引用：表"
                Set r = p.Range
                r.Collapse wdCollapseEnd
                Set f = ActiveDocument.Fields.Add(Range:=r, Type:=wdFieldEmpty, _
                                                  Text:="REF " & bmName & " \h", _
                                                  PreserveFormatting:=False)
                f.Update
                TestTableRefText = "ok|" & CleanText(p.Range.Text) & "|" & bmName
                Exit Function
            End If
        End If
    Next p
    TestTableRefText = "NO_CAPTION"
    Exit Function
EH:
    TestTableRefText = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function



' 测试入口：Application.Run("TestFigureCaption")
Public Function TestFigureCaption() As String
    Dim p As Paragraph
    Dim cap As Paragraph

    On Error GoTo EH
    Set p = FindFigureParagraph()
    If p Is Nothing Then
        TestFigureCaption = "NO_IMAGE"
        Exit Function
    End If
    p.Range.InsertParagraphAfter
    Set cap = p.Next
    BuildCaption cap, "图", STYLE_FIG_CAPTION, "SEQ 图 \* ARABIC \s 1", "测试图题"
    TestFigureCaption = "ok|" & CleanText(cap.Range.Text)
    Exit Function
EH:
    TestFigureCaption = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：Application.Run("TestTableCaption")
Public Function TestTableCaption() As String
    Dim t As Table
    Dim cap As Paragraph

    On Error GoTo EH
    If ActiveDocument.Tables.Count = 0 Then
        TestTableCaption = "NO_TABLE"
        Exit Function
    End If
    Set t = ActiveDocument.Tables(1)
    Set cap = InsertParagraphBeforeTable(t)
    BuildCaption cap, "表", STYLE_TAB_CAPTION, "SEQ 表 \* ARABIC \s 1", "测试表题"
    TestTableCaption = "ok|" & CleanText(cap.Range.Text)
    Exit Function
EH:
    TestTableCaption = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：空白行直接插入图注/表注/子图注（无选中文字）
Public Function TestCaptionBlankLine() As String
    Dim p As Paragraph
    Dim out As String

    On Error GoTo EH
    Selection.EndOf Unit:=wdStory
    Set p = ActiveDocument.Paragraphs.Add
    p.Range.Select
    InsertFigureCaption Nothing, True, ""
    out = "图注=" & CleanText(p.Range.Text)

    Selection.EndOf Unit:=wdStory
    Set p = ActiveDocument.Paragraphs.Add
    p.Range.Select
    InsertTableCaption Nothing, True, ""
    out = out & "|表注=" & CleanText(p.Range.Text)

    Selection.EndOf Unit:=wdStory
    Set p = ActiveDocument.Paragraphs.Add
    p.Range.Select
    InsertSubfigureCaption Nothing, True, ""
    out = out & "|子图注=" & CleanText(p.Range.Text)

    TestCaptionBlankLine = "ok|" & out
    Exit Function
EH:
    TestCaptionBlankLine = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：附录表格题注（编号应带附录字母，如 表A.2）
Public Function TestAppendixTableCaption() As String
    Dim t As Table
    Dim cap As Paragraph

    On Error GoTo EH
    For Each t In ActiveDocument.Tables
        If IsAppendixAt(t.Range) Then
            Set cap = InsertParagraphBeforeTable(t)
            BuildCaption cap, "表", STYLE_TAB_CAPTION, "SEQ 表 \* ARABIC \s 1", "测试附录表题"
            TestAppendixTableCaption = "ok|" & CleanText(cap.Range.Text)
            Exit Function
        End If
    Next t
    TestAppendixTableCaption = "NO_APPENDIX_TABLE"
    Exit Function
EH:
    TestAppendixTableCaption = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：表注段落必须在整表上方（不在表格内）
Public Function TestTableCaptionPosition() As String
    Dim t As Table
    Dim cap As Paragraph
    Dim inTable As Boolean
    Dim beforeTable As Boolean

    On Error GoTo EH
    Set t = ActiveDocument.Tables(ActiveDocument.Tables.Count)
    Set cap = InsertParagraphBeforeTable(t)
    inTable = cap.Range.Information(wdWithInTable)
    beforeTable = (cap.Range.Start < t.Range.Start)
    TestTableCaptionPosition = "ok|" & CStr(inTable) & "|" & CStr(beforeTable)
    Exit Function
EH:
    TestTableCaptionPosition = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function
' 测试入口：题注按钮点击路径（表格）
Public Function TestTableCaptionClick() As String
    Dim t As Table
    Dim cap As Paragraph

    On Error GoTo EH
    Set t = ActiveDocument.Tables(ActiveDocument.Tables.Count)
    t.Range.Select
    InsertTableCaption Nothing, True, "点击路径表题"
    Set cap = t.Range.Paragraphs(1).Previous
    TestTableCaptionClick = "ok|" & CStr(cap.Range.Information(wdWithInTable)) & "|" & CleanText(cap.Range.Text)
    Exit Function
EH:
    TestTableCaptionClick = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function
' 测试入口：图注按钮点击路径
Public Function TestFigureCaptionClick() As String
    Dim p As Paragraph
    Dim cap As Paragraph

    On Error GoTo EH
    Set p = FindFigureParagraph()
    If p Is Nothing Then
        TestFigureCaptionClick = "NO_IMAGE"
        Exit Function
    End If
    p.Range.Select
    InsertFigureCaption Nothing, True, "点击路径图题"
    Set cap = p.Next
    TestFigureCaptionClick = "ok|" & CleanText(cap.Range.Text)
    Exit Function
EH:
    TestFigureCaptionClick = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：图注不依赖图片
Public Function TestFigureCaptionTextOnly() As String
    Dim p As Paragraph
    On Error GoTo EH
    Selection.EndOf Unit:=wdStory
    Set p = ActiveDocument.Paragraphs.Add
    p.Range.Text = "无图图题"
    p.Range.Select
    InsertFigureCaption Nothing, True, "无图图题"
    TestFigureCaptionTextOnly = "ok|" & p.Style.NameLocal & "|" & CleanText(p.Range.Text)
    Exit Function
EH:
    TestFigureCaptionTextOnly = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：表注不依赖表格
Public Function TestTableCaptionTextOnly() As String
    Dim p As Paragraph
    On Error GoTo EH
    Selection.EndOf Unit:=wdStory
    Set p = ActiveDocument.Paragraphs.Add
    p.Range.Text = "无表表题"
    p.Range.Select
    InsertTableCaption Nothing, True, "无表表题"
    TestTableCaptionTextOnly = "ok|" & p.Style.NameLocal & "|" & CleanText(p.Range.Text)
    Exit Function
EH:
    TestTableCaptionTextOnly = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function
