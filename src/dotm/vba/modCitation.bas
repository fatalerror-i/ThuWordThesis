Attribute VB_Name = "modCitation"
Option Explicit

' 引用域代码前缀：QUOTE "THU_CITE:0001;0003"
Public Const CITE_TAG As String = "THU_CITE:"
Private Const CITE_LAST_ID As String = "THU_CITE_LAST_ID"
Private Const CITE_MISSING As String = "引用失效"

' ============================================================================
' 参考文献引用：条目书签 + 自定义属性编号 + “文献引用”字符样式
' 约定：
'   - 每条“参考文献条目”段落绑定一个稳定书签 THU_REF_ENTRY_0001。
'   - 该书签对应的自定义文档属性 _THU_REF_NUM_0001 保存当前显示编号（1、2…）。
'   - 正文引用是一个 QUOTE 域，域代码保存条目 id，如 THU_CITE:0001;0003;0007。
'   - 域结果由 NormalizeCitations 写成合并后的上标序号，如 [1, 3-7]。
'   - 多选插入合并为一个域；连续序号压缩为区间，逗号后留一个空格。
'   - “更新域”末尾会重新调用 NormalizeCitations，F9 后自动恢复编号。
'   - 导出 docx 时把引用域转成静态文字，避免无宏环境按 F9 回退成域代码。
'   - 普通 F9 保留引用显示，模板的更新宏重新核验目标并刷新编号。
'   - 兼容旧版 DOCPROPERTY 引用域。
' ============================================================================

' ---------------------------------------------------------------------------
' 功能区入口
' ---------------------------------------------------------------------------

Public Sub ShowCitationDialog(control As IRibbonControl)
    Dim doc As Document
    Dim p As Paragraph
    Dim i As Long
    Dim num As String
    Dim txt As String
    Dim id As String
    Dim k As Long
    Dim idx As Long
    Dim ids As String

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    If Not StyleExists(STYLE_REF_ENTRY) Then
        MsgBox "当前文档没有“参考文献条目”样式，无法插入文献引用。", vbExclamation, "引用文献"
        Exit Sub
    End If

    RefreshCitationProperties
    ResetPicker
    Set doc = ActiveDocument
    i = 0
    For Each p In doc.Paragraphs
        If IsReferenceEntry(p) Then
            i = i + 1
            num = GetReferenceNumber(p, i)
            id = EnsureEntryBookmark(p)
            txt = CleanText(p.Range.Text)
            If Len(txt) > 80 Then txt = Left$(txt, 80) & "…"
            AddPickerItem "[" & num & "]  " & txt, id
        End If
    Next p

    If PickerItemCount() = 0 Then
        MsgBox "没有找到参考文献条目。请先按“参考文献条目”样式录入文献。", _
               vbExclamation, "引用文献"
        Exit Sub
    End If

    If Not ShowPicker("引用文献", "勾选要引用的文献（可多选，自动合并为上标）：") Then Exit Sub
    For k = 0 To gPickerSelectedCount - 1
        idx = gPickerSelected(k)
        id = gPickerItems(idx).Key
        If Len(id) > 0 Then
            If Len(ids) > 0 Then ids = ids & ";"
            ids = ids & id
        End If
    Next k

    If Len(ids) = 0 Then
        MsgBox "没有插入任何引用。", vbInformation, "引用文献"
        Exit Sub
    End If
    InsertCitationGroup ids
    Exit Sub
EH:
    ReportError "引用文献", Err.Number, Err.Description
End Sub

' 功能区：插入脚注。正文标记自动编号，光标进入脚注区。
Public Sub InsertFootnote(control As IRibbonControl)
    Dim rng As Range
    Dim fn As Footnote

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    Set rng = Selection.Range
    rng.Collapse wdCollapseEnd
    Set fn = ActiveDocument.Footnotes.Add(Range:=rng, Text:="")
    NormalizeFootnoteParagraph fn
    NormalizeFootnoteReference fn
    EnsureFootnoteSpace fn
    Exit Sub
EH:
    ReportError "脚注", Err.Number, Err.Description
End Sub

' 去掉 Word 从正文段落带入的左侧缩进和段后距，回到“脚注文本”样式。
Public Sub NormalizeFootnoteParagraph(ByVal fn As Footnote)
    On Error Resume Next
    With fn.Range.Paragraphs(1).Range.ParagraphFormat
        .LeftIndent = 13.5
        .FirstLineIndent = -13.5
        .CharacterUnitLeftIndent = 0
        .CharacterUnitFirstLineIndent = -1.5
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
End Sub

' 脚注文本中的编号由“footnote reference”样式控制为普通文本；
' 正文中的标记单独设回上标。
Public Sub NormalizeFootnoteReference(ByVal fn As Footnote)
    On Error Resume Next
    fn.Reference.Font.Superscript = True
End Sub

' 脚注编号后补一个 0.5 字符宽的空格（U+2002，半宽空格）。
' Word 会在编号后自动放一个普通空格；把该空格本身替换为 U+2002。
Public Sub EnsureFootnoteSpace(ByVal fn As Footnote)
    Dim par As Range
    Dim sp As Range
    On Error Resume Next
    Set par = fn.Range.Paragraphs(1).Range
    If par.Start + 1 < par.End Then
        Set sp = par.Duplicate
        sp.Start = par.Start + 1
        sp.End = sp.Start + 1
        If sp.Text = " " Or sp.Text = ChrW(&HA0) Or sp.Text = vbTab Then
            sp.Text = ChrW(&H2002)
        ElseIf sp.Text <> ChrW(&H2002) Then
            sp.InsertBefore ChrW(&H2002)
        End If
    End If
End Sub

' 批量修正已有脚注（“更新域”时调用）。
Public Sub NormalizeFootnoteReferences()
    Dim fn As Footnote

    On Error Resume Next
    For Each fn In ActiveDocument.Footnotes
        NormalizeFootnoteReference fn
        EnsureFootnoteSpace fn
    Next fn
End Sub

' 更新全部域之前调用：按当前条目顺序刷新编号属性
Public Sub RefreshCitationProperties()
    Dim doc As Document
    Dim p As Paragraph
    Dim id As String
    Dim num As String
    Dim i As Long
    Dim lastId As Long

    Set doc = ActiveDocument
    i = 0
    For Each p In doc.Paragraphs
        If IsReferenceEntry(p) Then
            i = i + 1
            id = EnsureEntryBookmark(p, lastId)
            If Len(id) > 0 Then
                num = GetReferenceNumber(p, i)
                SetCustomProperty PROP_REF_NUM_PREFIX & id, num
            End If
        End If
    Next p
End Sub

' ---------------------------------------------------------------------------
' 插入引用
' ---------------------------------------------------------------------------

' 在光标处插入一条或多条文献引用（ids 形如 "0001;0003;0007"）。
' 多选合并成一个 QUOTE 域；域结果由 NormalizeCitations 写成上标序号。
Public Sub InsertCitationGroup(ByVal ids As String)
    Dim f As Field
    Dim code As String

    If Len(ids) = 0 Then Exit Sub
    code = "QUOTE " & Chr$(34) & CITE_TAG & ids & Chr$(34)
    Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, _
                                 Text:=code, PreserveFormatting:=False)
    ' 先写占位结果，避免更新前直接显示域代码；随后由 NormalizeCitations 统一改写。
    On Error Resume Next
    f.Result.Text = "[" & ids & "]"
    ApplyCitationStyle f.Result
    On Error GoTo 0
    Selection.Collapse wdCollapseEnd
    NormalizeCitations
End Sub

' ---------------------------------------------------------------------------
' 合并显示：域代码保存条目 id，域结果写成 [1]、[2-3]、[7, 9, 11-13]
' ---------------------------------------------------------------------------

' 整理全文引用域：把域代码里的条目 id 翻译为当前编号，并压缩连续区间。
Public Sub NormalizeCitations()
    Dim doc As Document
    Dim f As Field
    Dim ids As String
    Dim display As String

    Set doc = ActiveDocument
    For Each f In CitationFields(doc)
        ids = CitationFieldIds(f.Code.Text)
        f.Locked = False
        If f.Type = wdFieldQuote Then
            display = CitationDisplay(ids)
        ElseIf CitationTargetExists(doc, ids) Then
            f.Update
            display = f.Result.Text
        Else
            display = "[" & CITE_MISSING & "]"
        End If
        f.Result.Text = display
        ApplyCitationStyle f.Result
        ' Word 的 F9 和打印时更新域只应更新普通域。
        f.Locked = True
    Next f
End Sub

' 导出时仅冻结模板的文献引用，目录、图表等普通域仍可更新。
Public Sub FreezeCitationFields()
    Dim fieldsToFreeze As Collection
    Dim f As Field
    Dim i As Long

    Set fieldsToFreeze = CitationFields(ActiveDocument)
    For i = fieldsToFreeze.Count To 1 Step -1
        Set f = fieldsToFreeze(i)
        f.Locked = False
        f.Unlink
    Next i
End Sub

Public Function CitationFields(ByVal doc As Document) As Collection
    Dim firstStory As Range
    Dim story As Range
    Dim f As Field

    Set CitationFields = New Collection
    For Each firstStory In doc.StoryRanges
        Set story = firstStory
        Do Until story Is Nothing
            For Each f In story.Fields
                If f.Type = wdFieldQuote Or f.Type = wdFieldDocProperty Then
                    If Len(CitationFieldIds(f.Code.Text)) > 0 Then CitationFields.Add f
                End If
            Next f
            Set story = story.NextStoryRange
        Loop
    Next firstStory
End Function

' 识别当前的 QUOTE 引用和旧版 DOCPROPERTY 引用。
Public Function CitationFieldIds(ByVal code As String) As String
    Dim arg As String
    Dim prefix As String
    Dim last As Long

    code = LTrim$(code)
    If UCase$(Left$(code, 6)) = "QUOTE " Then
        arg = LTrim$(Mid$(code, 7))
        prefix = CITE_TAG
    ElseIf UCase$(Left$(code, 12)) = "DOCPROPERTY " Then
        arg = LTrim$(Mid$(code, 13))
        prefix = PROP_REF_NUM_PREFIX
    Else
        Exit Function
    End If
    If Left$(arg, 1) = Chr$(34) Then
        arg = Mid$(arg, 2)
        last = InStr(arg, Chr$(34))
    Else
        last = InStr(arg, " ")
    End If
    If last > 0 Then arg = Left$(arg, last - 1)
    If StrComp(Left$(arg, Len(prefix)), prefix, vbTextCompare) = 0 Then
        CitationFieldIds = Mid$(arg, Len(prefix) + 1)
    End If
End Function

Private Function CitationDisplay(ByVal ids As String) As String
    Dim id As Variant
    Dim num As String
    Dim nums As String
    Dim missing As Boolean

    For Each id In Split(ids, ";")
        num = CitationNumberForId(Trim$(CStr(id)))
        If Len(num) > 0 Then
            nums = nums & num & ";"
        Else
            missing = True
        End If
    Next id
    If Len(nums) > 0 Then nums = CompressNumbers(Left$(nums, Len(nums) - 1))
    If missing Then
        If Len(nums) > 0 Then nums = nums & ", "
        nums = nums & CITE_MISSING
    End If
    CitationDisplay = "[" & nums & "]"
End Function

Private Function CitationNumberForId(ByVal id As String) As String
    Dim raw As String
    Dim p As Paragraph
    Dim target As Long
    Dim entryIndex As Long

    If Not CitationTargetExists(ActiveDocument, id) Then Exit Function
    On Error Resume Next
    raw = ActiveDocument.CustomDocumentProperties(PROP_REF_NUM_PREFIX & id).Value
    On Error GoTo 0
    If IsNumeric(raw) Then
        CitationNumberForId = CStr(CLng(raw))
        Exit Function
    End If

    target = ActiveDocument.Bookmarks(BM_REF_ENTRY_PREFIX & id).Range.Paragraphs(1).Range.Start
    For Each p In ActiveDocument.Paragraphs
        If IsReferenceEntry(p) Then
            entryIndex = entryIndex + 1
            If p.Range.Start = target Then
                CitationNumberForId = GetReferenceNumber(p, entryIndex)
                Exit Function
            End If
        End If
    Next p
End Function

Public Function CitationTargetExists(ByVal doc As Document, ByVal id As String) As Boolean
    Dim bm As Bookmark
    Dim p As Paragraph

    If Len(id) = 0 Then Exit Function
    If Not doc.Bookmarks.Exists(BM_REF_ENTRY_PREFIX & id) Then Exit Function
    Set bm = doc.Bookmarks(BM_REF_ENTRY_PREFIX & id)
    If bm.Range.StoryType <> wdMainTextStory Then Exit Function
    Set p = bm.Range.Paragraphs(1)
    CitationTargetExists = IsReferenceEntry(p) And BookmarkInsideParagraph(bm, p)
End Function

' 把 "2;3" / "7;9;11;12;13" 压缩为 "2-3" / "7, 9, 11-13"（升序、去重）。
Public Function CompressNumbers(ByVal numbers As String) As String
    Dim parts() As String
    Dim vals() As Long
    Dim uniq() As Long
    Dim n As Long
    Dim m As Long
    Dim i As Long
    Dim j As Long
    Dim t As Long
    Dim out As String

    If Len(numbers) = 0 Then Exit Function
    parts = Split(numbers, ";")
    ReDim vals(0 To UBound(parts))
    n = -1
    For i = 0 To UBound(parts)
        If IsNumeric(Trim$(parts(i))) Then
            n = n + 1
            vals(n) = CLng(Trim$(parts(i)))
        End If
    Next i
    If n < 0 Then Exit Function
    ReDim Preserve vals(0 To n)

    ' 升序排序（条目数很少，直接插入排序）
    For i = 1 To n
        t = vals(i)
        j = i - 1
        Do While j >= 0
            If vals(j) <= t Then Exit Do
            vals(j + 1) = vals(j)
            j = j - 1
        Loop
        vals(j + 1) = t
    Next i

    ' 去重
    ReDim uniq(0 To n)
    m = 0
    uniq(0) = vals(0)
    For i = 1 To n
        If vals(i) <> uniq(m) Then
            m = m + 1
            uniq(m) = vals(i)
        End If
    Next i

    ' 合并连续区间
    i = 0
    Do While i <= m
        j = i
        Do While j < m
            If uniq(j + 1) = uniq(j) + 1 Then
                j = j + 1
            Else
                Exit Do
            End If
        Loop
        If Len(out) > 0 Then out = out & ", "
        If j > i Then
            out = out & CStr(uniq(i)) & "-" & CStr(uniq(j))
        Else
            out = out & CStr(uniq(i))
        End If
        i = j + 1
    Loop
    CompressNumbers = out
End Function

Private Sub ApplyCitationStyle(rng As Range)
    On Error Resume Next
    rng.Style = ActiveDocument.Styles(STYLE_CITATION)
    If Err.Number <> 0 Then
        Err.Clear
        rng.Font.Superscript = True
    End If
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' 条目书签与编号
' ---------------------------------------------------------------------------

' 返回条目的稳定 id（书签名后缀，如 0001）
Private Function EnsureEntryBookmark(p As Paragraph, Optional ByRef lastId As Long = 0) As String
    Dim bm As Bookmark
    Dim name As String

    For Each bm In ActiveDocument.Bookmarks
        If Left$(bm.Name, Len(BM_REF_ENTRY_PREFIX)) = BM_REF_ENTRY_PREFIX Then
            If BookmarkInsideParagraph(bm, p) Then
                EnsureEntryBookmark = Mid$(bm.Name, Len(BM_REF_ENTRY_PREFIX) + 1)
                Exit Function
            End If
        End If
    Next bm

    If lastId = 0 Then lastId = LastCitationId(ActiveDocument)
    lastId = lastId + 1
    ActiveDocument.Variables(CITE_LAST_ID).Value = CStr(lastId)
    name = BM_REF_ENTRY_PREFIX & Format$(lastId, "0000")
    ActiveDocument.Bookmarks.Add name, EntryRange(p)
    EnsureEntryBookmark = Mid$(name, Len(BM_REF_ENTRY_PREFIX) + 1)
End Function

' 旧稿没有计数器。扫描现有书签、属性和引用域，保留已删除条目使用过的 ID。
Private Function LastCitationId(ByVal doc As Document) As Long
    Dim counter As Variable
    Dim bm As Bookmark
    Dim prop As DocumentProperty
    Dim f As Field
    Dim id As Variant
    Dim lastId As Long
    Dim hasCounter As Boolean

    For Each counter In doc.Variables
        If StrComp(counter.Name, CITE_LAST_ID, vbTextCompare) = 0 Then
            IncludeCitationId lastId, counter.Value
            hasCounter = True
            Exit For
        End If
    Next counter
    For Each bm In doc.Bookmarks
        If Left$(bm.Name, Len(BM_REF_ENTRY_PREFIX)) = BM_REF_ENTRY_PREFIX Then
            IncludeCitationId lastId, Mid$(bm.Name, Len(BM_REF_ENTRY_PREFIX) + 1)
        End If
    Next bm
    For Each prop In doc.CustomDocumentProperties
        If Left$(prop.Name, Len(PROP_REF_NUM_PREFIX)) = PROP_REF_NUM_PREFIX Then
            IncludeCitationId lastId, Mid$(prop.Name, Len(PROP_REF_NUM_PREFIX) + 1)
        End If
    Next prop
    For Each f In CitationFields(doc)
        For Each id In Split(CitationFieldIds(f.Code.Text), ";")
            IncludeCitationId lastId, Trim$(CStr(id))
        Next id
    Next f
    If Not hasCounter Then doc.Variables.Add Name:=CITE_LAST_ID, Value:=CStr(lastId)
    LastCitationId = lastId
End Function

Private Sub IncludeCitationId(ByRef lastId As Long, ByVal id As String)
    If Len(id) = 0 Or Len(id) > 9 Then Exit Sub
    If id Like "*[!0-9]*" Then Exit Sub
    If CLng(id) > lastId Then lastId = CLng(id)
End Sub

Private Function IsReferenceEntry(ByVal p As Paragraph) As Boolean
    If p.Style.NameLocal <> STYLE_REF_ENTRY Then Exit Function
    IsReferenceEntry = Len(CleanText(p.Range.Text)) > 0
End Function

Private Function EntryRange(p As Paragraph) As Range
    Dim r As Range
    Set r = p.Range.Duplicate
    If r.End > r.Start Then r.MoveEnd wdCharacter, -1
    If r.End <= r.Start Then Set r = p.Range.Duplicate
    Set EntryRange = r
End Function

Private Function BookmarkInsideParagraph(bm As Bookmark, p As Paragraph) As Boolean
    If bm.Range.StoryType <> p.Range.StoryType Then Exit Function
    If bm.Range.Start >= p.Range.Start And bm.Range.End <= p.Range.End _
       And bm.Range.Start < bm.Range.End Then
        BookmarkInsideParagraph = True
    End If
End Function

' 取条目当前编号：优先 Word 列表编号，其次文本 [n]，最后按顺序号
Private Function GetReferenceNumber(p As Paragraph, ByVal fallbackIndex As Long) As String
    Dim s As String
    Dim l As Long
    Dim r As Long

    s = GetHeadingListString(p)
    If Len(s) = 0 Then s = CleanText(p.Range.Text)
    l = InStr(s, "[")
    r = InStr(s, "]")
    If l > 0 And r > l Then
        GetReferenceNumber = Trim$(Mid$(s, l + 1, r - l - 1))
        Exit Function
    End If
    If Len(s) > 0 Then
        ' 列表编号有时返回 "1." / "1"
        s = Replace(s, ".", "")
        s = Replace(s, "]", "")
        s = Replace(s, "[", "")
        If IsNumeric(s) Then
            GetReferenceNumber = CStr(CLng(s))
            Exit Function
        End If
    End If
    GetReferenceNumber = CStr(fallbackIndex)
End Function

Private Function FindEntryIdByNumber(ByVal numberText As String) As String
    Dim doc As Document
    Dim p As Paragraph
    Dim i As Long
    Dim num As String

    Set doc = ActiveDocument
    i = 0
    For Each p In doc.Paragraphs
        If IsReferenceEntry(p) Then
            i = i + 1
            num = GetReferenceNumber(p, i)
            If num = numberText Then
                FindEntryIdByNumber = EnsureEntryBookmark(p)
                Exit Function
            End If
        End If
    Next p
End Function

' ---------------------------------------------------------------------------
' 预览与属性
' ---------------------------------------------------------------------------

Private Function ReferenceEntriesPreview(ByVal maxCount As Long) As String
    Dim doc As Document
    Dim p As Paragraph
    Dim i As Long
    Dim num As String
    Dim txt As String
    Dim out As String

    Set doc = ActiveDocument
    i = 0
    For Each p In doc.Paragraphs
        If IsReferenceEntry(p) Then
            i = i + 1
            num = GetReferenceNumber(p, i)
            txt = CleanText(p.Range.Text)
            If Len(txt) > 60 Then txt = Left$(txt, 60) & "…"
            out = out & "[" & num & "]  " & txt & vbCrLf
            If i >= maxCount Then
                out = out & "……" & vbCrLf
                Exit For
            End If
        End If
    Next p
    ReferenceEntriesPreview = out
End Function

Private Sub SetCustomProperty(ByVal propName As String, ByVal propValue As String)
    Dim props As DocumentProperties
    Dim prop As DocumentProperty

    On Error Resume Next
    Set props = ActiveDocument.CustomDocumentProperties
    Set prop = props(propName)
    If prop Is Nothing Then
        props.Add Name:=propName, LinkToContent:=False, _
                  Type:=msoPropertyTypeString, Value:=propValue
    Else
        prop.Value = propValue
    End If
    On Error GoTo 0
End Sub
' 测试入口：Application.Run("TestCitation")
Public Function TestCitation() As String
    Dim p As Paragraph
    Dim id As String

    On Error GoTo EH
    RefreshCitationProperties
    For Each p In ActiveDocument.Paragraphs
        If IsReferenceEntry(p) Then
            id = EnsureEntryBookmark(p)
            Exit For
        End If
    Next p
    If Len(id) = 0 Then
        TestCitation = "NO_ENTRY"
        Exit Function
    End If
    Selection.EndKey 6
    Selection.TypeParagraph
    Selection.TypeText "引用："
    InsertCitationGroup id
    ActiveDocument.Fields.Update
    NormalizeCitations
    TestCitation = "ok|" & CleanText(ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count).Range.Text)
    Exit Function
EH:
    TestCitation = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 取第 n 条“参考文献条目”的稳定 id（如 0001）；没有条目返回空串。
Private Function EntryIdAt(ByVal wantIndex As Long) As String
    Dim p As Paragraph
    Dim i As Long

    On Error Resume Next
    i = 0
    For Each p In ActiveDocument.Paragraphs
        If IsReferenceEntry(p) Then
            i = i + 1
            If i = wantIndex Then
                EntryIdAt = EnsureEntryBookmark(p)
                Exit Function
            End If
        End If
    Next p
    On Error GoTo 0
End Function

' 末尾追加一段“说明 + 引用”，返回该段去掉控制符后的文字。
Private Function AppendCitation(ByVal caption As String, ByVal ids As String) As String
    Selection.EndKey 6
    Selection.TypeParagraph
    Selection.TypeText caption
    InsertCitationGroup ids
    AppendCitation = CleanText(ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count).Range.Text)
End Function

' 测试入口：Application.Run("TestCitationCompress")
' 纯函数校验：连续区间、逗号分隔、单值、乱序去重、空串。
Public Function TestCitationCompress() As String
    On Error GoTo EH
    TestCitationCompress = "ok|" & CompressNumbers("2;3") & "|" & _
                           CompressNumbers("7;9;11;12;13") & "|" & _
                           CompressNumbers("1") & "|" & _
                           CompressNumbers("3;1;2;2") & "|[" & _
                           CompressNumbers("") & "]"
    Exit Function
EH:
    TestCitationCompress = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：Application.Run("TestCitationMerge")
' 按条目顺序取前 4 条，验证连续、非连续、重复 id 与重复整理。
Public Function TestCitationMerge() As String
    Dim id1 As String
    Dim id2 As String
    Dim id3 As String
    Dim seq As String
    Dim pair As String
    Dim gap As String
    Dim dedupe As String
    Dim again As String

    On Error GoTo EH
    RefreshCitationProperties
    id1 = EntryIdAt(1)
    id2 = EntryIdAt(2)
    id3 = EntryIdAt(3)
    If Len(id1) = 0 Or Len(id2) = 0 Or Len(id3) = 0 Then
        TestCitationMerge = "NO_ENTRY"
        Exit Function
    End If

    seq = AppendCitation("连续引用：", id1 & ";" & id2 & ";" & id3)
    pair = AppendCitation("两条连续：", id2 & ";" & id3)
    gap = AppendCitation("非连续引用：", id1 & ";" & id3)
    dedupe = AppendCitation("重复引用：", id2 & ";" & id2)
    AppendCitation "再整理一次：", id1 & ";" & id2 & ";" & id3
    NormalizeCitations
    again = CleanText(ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count).Range.Text)
    TestCitationMerge = "ok|" & seq & "|" & pair & "|" & gap & "|" & dedupe & "|" & again
    Exit Function
EH:
    TestCitationMerge = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：Application.Run("TestCitationUpdateRoundTrip")
' 模拟 F9：普通更新保留文献编号，模板的更新宏仍能刷新编号。
Public Function TestCitationUpdateRoundTrip() As String
    Dim id As String
    Dim before As String
    Dim raw As String
    Dim after As String

    On Error GoTo EH
    RefreshCitationProperties
    id = EntryIdAt(1)
    If Len(id) = 0 Then
        TestCitationUpdateRoundTrip = "NO_ENTRY"
        Exit Function
    End If

    Selection.EndKey 6
    Selection.TypeParagraph
    Selection.TypeText "更新域测试："
    InsertCitationGroup id
    before = CleanText(ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count).Range.Text)
    ActiveDocument.Fields.Update
    raw = CleanText(ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count).Range.Text)
    NormalizeCitations
    after = CleanText(ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count).Range.Text)
    TestCitationUpdateRoundTrip = "ok|" & before & "|" & raw & "|" & after
    Exit Function
EH:
    TestCitationUpdateRoundTrip = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：Application.Run("TestCitationFreeze")
' 插入引用后冻结域：正文不再有引用域，显示文字保持 [1]（提交版 docx 用）。
Public Function TestCitationFreeze() As String
    Dim id As String
    Dim before As Long
    Dim after As Long
    Dim text As String

    On Error GoTo EH
    RefreshCitationProperties
    id = EntryIdAt(1)
    If Len(id) = 0 Then
        TestCitationFreeze = "NO_ENTRY"
        Exit Function
    End If

    AppendCitation "冻结测试：", id
    before = CountCiteFields()
    FreezeCitationFields
    after = CountCiteFields()
    text = CleanText(ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count).Range.Text)
    TestCitationFreeze = "ok|" & CStr(before) & "|" & CStr(after) & "|" & text
    Exit Function
EH:
    TestCitationFreeze = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 统计正文中 NEW 式引用域（QUOTE THU_CITE）的数量；旧 DOCPROPERTY 域不计入。
Private Function CountCiteFields() As Long
    Dim f As Field
    Dim n As Long

    On Error Resume Next
    For Each f In ActiveDocument.Fields
        If InStr(1, f.Code.Text, CITE_TAG, vbTextCompare) > 0 Then
            n = n + 1
        End If
    Next f
    On Error GoTo 0
    CountCiteFields = n
End Function

' 测试入口：Application.Run("TestFootnoteInsert")
Public Function TestFootnoteInsert() As String
    Dim before As Long
    Dim rng As Range
    Dim fn As Footnote

    On Error GoTo EH
    before = ActiveDocument.Footnotes.Count
    Set rng = ActiveDocument.Content
    rng.Collapse wdCollapseEnd
    Set fn = ActiveDocument.Footnotes.Add(Range:=rng, Text:="")
    NormalizeFootnoteParagraph fn
    NormalizeFootnoteReference fn
    EnsureFootnoteSpace fn
    fn.Range.InsertAfter "测试脚注"
    TestFootnoteInsert = "ok|" & CStr(ActiveDocument.Footnotes.Count = before + 1) & _
                         "|" & CStr(fn.Reference.Font.Superscript = True) & _
                         "|" & CStr(ActiveDocument.Footnotes.Count) & _
                         "|" & CStr(Left$(fn.Range.Text, 1) = ChrW(&H2002))
    Exit Function
EH:
    TestFootnoteInsert = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function
