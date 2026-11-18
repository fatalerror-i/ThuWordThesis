Attribute VB_Name = "modCheck"
Option Explicit

' ============================================================================
' 引用检查：只核对题注与交叉引用，且只覆盖正文与附录。
'   1) 图片是否有对应图注 / 子图题注；2) 表格是否有对应表注；
'   3) 题注编号是否以域格式存在，图表题注是否可被引用；
'   4) 交叉引用与文献引用的目标是否存在。
' 不检查正文文字、标题字数、封面控件与版式数值等内容项。
' 性能：文档级集合（Fields / Bookmarks / Shapes / Tables）各枚举一次，
'       段落只遍历一次且只读三个属性，关联全部在内存数组内完成，
'       不会随段落数反复来回调用 Word 对象模型。
' ============================================================================

Public Const CHECK_TITLE As String = "引用检查"

Public Sub RunTemplateCheck(control As IRibbonControl)
    Dim issues As Collection
    Dim notes As Collection

    On Error GoTo EH
    If Documents.Count = 0 Then
        MsgBox "请先打开一篇论文文档。", vbInformation, CHECK_TITLE
        Exit Sub
    End If
    Set issues = New Collection
    Set notes = New Collection
    RunCheckCore issues, notes
    ShowReport issues, notes
    Exit Sub
EH:
    MsgBox "检查未完成。" & vbCrLf & Err.Description, vbExclamation, CHECK_TITLE
End Sub

' ---------------------------------------------------------------------------
' 检查主流程：阶段 0 建索引，阶段 1 单次段落遍历，阶段 2 纯内存关联
' ---------------------------------------------------------------------------

Private Sub RunCheckCore(ByRef issues As Collection, ByRef notes As Collection)
    Dim doc As Document
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim p As Paragraph
    Dim f As Field
    Dim bm As Bookmark
    Dim sh As Shape
    Dim ils As InlineShape
    Dim prop As DocumentProperty
    Dim code As String
    Dim seqId As String
    Dim pos As Long
    Dim bodyIdx As Long
    Dim bodyStart As Long
    Dim starts() As Long
    Dim styles() As String
    Dim texts() As String
    Dim seqOk() As Boolean
    Dim hasImage() As Boolean
    Dim bmAt() As String
    Dim bmExists As Object
    Dim refTargets As Object
    Dim propExists As Object
    Dim refList As Collection
    Dim propList As Collection

    Set doc = ActiveDocument
    Set bmExists = CreateObject("Scripting.Dictionary")
    Set refTargets = CreateObject("Scripting.Dictionary")
    Set propExists = CreateObject("Scripting.Dictionary")
    Set refList = New Collection
    Set propList = New Collection

    n = doc.Paragraphs.Count
    If n = 0 Then Exit Sub
    ReDim starts(1 To n)
    ReDim styles(1 To n)
    ReDim texts(1 To n)
    ReDim seqOk(1 To n)
    ReDim hasImage(1 To n)
    ReDim bmAt(1 To n)

    i = 0
    For Each p In doc.Paragraphs
        i = i + 1
        starts(i) = p.Range.Start
        styles(i) = PrimaryStyleName(p.Style)
        texts(i) = CleanText(p.Range.Text)
    Next p

    ' 正文 / 附录起点：首个「标题 1」章标题。它位于封面、目录与名单页之后，
    ' 又早于第一章内任何题注或图片；若以题注自身当分界，用户改写或删除
    ' 早期题注会令分界后移，使前面的表格与图片漏检。
    ' 找不到章标题时回退到首个题注/图片段，仍找不到才从全文第一段开始。
    bodyIdx = 0
    For i = 1 To n
        If styles(i) = "标题 1" Then
            bodyIdx = i
            Exit For
        End If
    Next i
    If bodyIdx = 0 Then
        For i = 1 To n
            If IsCaptionStyle(styles(i)) Or styles(i) = STYLE_IMAGE _
               Or styles(i) = STYLE_SUBFIG_CAPTION Then
                bodyIdx = i
                Exit For
            End If
        Next i
    End If
    If bodyIdx = 0 Then bodyIdx = 1
    bodyStart = starts(bodyIdx)

    For Each f In doc.Fields
        If f.Code.StoryType = wdMainTextStory Then
            code = Trim$(f.Code.Text)
            pos = f.Code.Start
            j = ParaIndexAt(starts, n, pos)
            If Left$(UCase$(code), 4) = "SEQ " Then
                If j >= 1 And j <= n Then
                    seqId = FieldArgument(code)
                    If styles(j) = STYLE_SUBFIG_CAPTION Then
                        If seqId = "子图" Then seqOk(j) = True
                    ElseIf InStr(1, code, "\* ARABIC", vbTextCompare) > 0 Then
                        seqOk(j) = True
                    End If
                End If
            ElseIf Left$(UCase$(code), 4) = "REF " Then
                refList.Add Array(FieldArgument(UCase$(code)), f.Result.Text)
                If j >= 1 And j <= n Then
                    If IsTabCaptionStyle(styles(j)) And IsContinuedTableCaption(texts(j)) Then
                        seqOk(j) = True
                    End If
                End If
            ElseIf Left$(UCase$(code), 11) = "DOCPROPERTY" And Len(CitationFieldIds(code)) = 0 Then
                propList.Add FieldArgument(UCase$(code))
            End If
        End If
    Next f

    For Each bm In doc.Bookmarks
        If bm.Range.StoryType = wdMainTextStory Then
            bmExists(UCase$(bm.Name)) = True
            If Left$(bm.Name, Len(BM_FIG_PREFIX)) = BM_FIG_PREFIX _
               Or Left$(bm.Name, Len(BM_TAB_PREFIX)) = BM_TAB_PREFIX Then
                j = ParaIndexAt(starts, n, bm.Range.Start)
                If j >= 1 And j <= n Then bmAt(j) = bm.Name
            End If
        End If
    Next bm

    On Error Resume Next
    For Each prop In doc.CustomDocumentProperties
        propExists(UCase$(prop.Name)) = True
    Next prop
    On Error GoTo 0

    For Each sh In doc.Shapes
        If sh.Anchor.StoryType = wdMainTextStory And sh.Anchor.Start >= bodyStart Then
            j = ParaIndexAt(starts, n, sh.Anchor.Start)
            If j >= 1 And j <= n Then hasImage(j) = True
        End If
    Next sh
    For Each ils In doc.InlineShapes
        If ils.Range.StoryType = wdMainTextStory And ils.Range.Start >= bodyStart Then
            j = ParaIndexAt(starts, n, ils.Range.Start)
            If j >= 1 And j <= n Then hasImage(j) = True
        End If
    Next ils

    CheckImages issues, n, styles, hasImage, bodyIdx
    CheckTables issues, doc, n, starts, styles, bodyStart
    CheckCaptionFields issues, n, styles, texts, seqOk, bmAt
    CheckReferences issues, doc, refList, propList, bmExists, refTargets, propExists
    CheckCitationTargets issues, doc
    CheckUnreferenced notes, bmExists, refTargets
End Sub

Private Sub CheckCitationTargets(ByRef issues As Collection, ByVal doc As Document)
    Dim f As Field
    Dim id As Variant
    Dim target As String

    For Each f In CitationFields(doc)
        For Each id In Split(CitationFieldIds(f.Code.Text), ";")
            target = Trim$(CStr(id))
            If Not CitationTargetExists(doc, target) Then
                AddIssue issues, "引用", "文献引用的条目已删除或不再是参考文献条目（ID " & target & "）。"
            End If
        Next id
    Next f
End Sub

' ---------------------------------------------------------------------------
' 两类题注配对：图片 -> 图注/子图注链，表格 -> 表注
' ---------------------------------------------------------------------------

Private Sub CheckImages(ByRef issues As Collection, ByVal n As Long, _
                        ByRef styles() As String, ByRef hasImage() As Boolean, _
                        ByVal bodyIdx As Long)
    Dim i As Long
    Dim j As Long
    Dim imgCount As Long
    Dim hasCaption As Boolean

    i = bodyIdx
    Do While i <= n
        If hasImage(i) Or styles(i) = STYLE_IMAGE Then
            imgCount = 0
            j = i
            Do While j <= n
                If hasImage(j) Or styles(j) = STYLE_IMAGE Then
                    imgCount = imgCount + 1
                    j = j + 1
                ElseIf styles(j) = STYLE_SUBFIG_CAPTION Then
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop
            hasCaption = False
            If j <= n Then hasCaption = IsFigCaptionStyle(styles(j))
            If hasCaption Then
                i = j + 1
            Else
                AddIssue issues, "图注", CStr(imgCount) & " 张图片缺少图注：第 " & CStr(i) & " 段附近"
                i = j
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub

Private Sub CheckTables(ByRef issues As Collection, ByVal doc As Document, ByVal n As Long, _
                        ByRef starts() As Long, ByRef styles() As String, _
                        ByVal bodyStart As Long)
    Dim t As Table
    Dim j As Long
    Dim cap As Long
    Dim hasCaption As Boolean

    For Each t In doc.Tables
        If t.Range.Start >= bodyStart And t.Range.StoryType = wdMainTextStory Then
            j = ParaIndexAt(starts, n, t.Range.Start)
            cap = j - 1
            hasCaption = False
            If cap >= 1 Then hasCaption = IsTabCaptionStyle(styles(cap))
            If Not hasCaption Then
                AddIssue issues, "表注", "表格缺少对应表注：第 " & CStr(j) & " 段"
            End If
        End If
    Next t
End Sub

Private Sub CheckCaptionFields(ByRef issues As Collection, ByVal n As Long, _
                               ByRef styles() As String, ByRef texts() As String, _
                               ByRef seqOk() As Boolean, ByRef bmAt() As String)
    Dim i As Long
    Dim category As String
    For i = 1 To n
        If styles(i) = STYLE_SUBFIG_CAPTION Then
            If Not seqOk(i) Then
                AddIssue issues, "图注", "子图注缺少编号域：第 " & CStr(i) & " 段"
            End If
        ElseIf IsCaptionStyle(styles(i)) Then
            category = "图注"
            If IsTabCaptionStyle(styles(i)) Then category = "表注"
            If Not seqOk(i) Then
                AddIssue issues, category, category & "编号缺少域：第 " & CStr(i) & " 段"
            End If
            If Not IsContinuedTableCaption(texts(i)) And Len(bmAt(i)) = 0 Then
                AddIssue issues, category, "缺少可引用的书签：第 " & CStr(i) & " 段 " & Left$(texts(i), 24)
            End If
        End If
    Next i
End Sub

' ---------------------------------------------------------------------------
' 引用目标是否存在，以及图表题注是否未被引用（后者只是提示）
' ---------------------------------------------------------------------------

Private Sub CheckReferences(ByRef issues As Collection, ByVal doc As Document, ByRef refList As Collection, _
                            ByRef propList As Collection, ByRef bmExists As Object, _
                            ByRef refTargets As Object, ByRef propExists As Object)
    Dim v As Variant
    Dim bmName As String
    Dim resultText As String
    Dim propName As String

    For Each v In refList
        bmName = CStr(v(0))
        resultText = CStr(v(1))
        If Len(bmName) = 0 Then
            AddIssue issues, "引用", "REF 域缺少书签名。"
        Else
            refTargets(UCase$(bmName)) = True
            If Not bmExists.Exists(UCase$(bmName)) Then
                AddIssue issues, "引用", "REF 指向的书签不存在：" & bmName
            ElseIf InStr(1, resultText, "错误", vbBinaryCompare) > 0 Then
                AddIssue issues, "引用", "REF 域显示错误：" & bmName
            End If
        End If
    Next v

    For Each v In propList
        propName = CStr(v)
        If Len(propName) > 0 Then
            If Not propExists.Exists(propName) Then
                If Not BuiltInPropertyExists(doc, propName) Then
                    AddIssue issues, "引用", "DOCPROPERTY 指向的属性不存在：" & propName
                End If
            End If
        End If
    Next v
End Sub

Private Sub CheckUnreferenced(ByRef notes As Collection, ByRef bmExists As Object, _
                              ByRef refTargets As Object)
    Dim k As Variant
    Dim bmName As String
    For Each k In bmExists.Keys
        bmName = CStr(k)
        If Left$(bmName, Len(BM_FIG_PREFIX)) = BM_FIG_PREFIX _
           Or Left$(bmName, Len(BM_TAB_PREFIX)) = BM_TAB_PREFIX Then
            If Not refTargets.Exists(bmName) Then
                notes.Add "图表题注未被引用：" & bmName
            End If
        End If
    Next k
End Sub

' ---------------------------------------------------------------------------
' 小工具
' ---------------------------------------------------------------------------

Private Function IsCaptionStyle(ByVal s As String) As Boolean
    IsCaptionStyle = (s = STYLE_FIG_CAPTION Or s = STYLE_APP_FIG_CAPTION _
                      Or s = STYLE_TAB_CAPTION Or s = STYLE_APP_TAB_CAPTION)
End Function

Private Function IsFigCaptionStyle(ByVal s As String) As Boolean
    IsFigCaptionStyle = (s = STYLE_FIG_CAPTION Or s = STYLE_APP_FIG_CAPTION)
End Function

Private Function IsTabCaptionStyle(ByVal s As String) As Boolean
    IsTabCaptionStyle = (s = STYLE_TAB_CAPTION Or s = STYLE_APP_TAB_CAPTION)
End Function

' 二分查找最后一个 Start <= pos 的段落序号；pos 落在正文之前时返回 0。
Private Function ParaIndexAt(ByRef starts() As Long, ByVal n As Long, ByVal pos As Long) As Long
    Dim lo As Long
    Dim hi As Long
    Dim mid As Long
    lo = 1
    hi = n
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If starts(mid) <= pos Then
            ParaIndexAt = mid
            lo = mid + 1
        Else
            hi = mid - 1
        End If
    Loop
End Function

Private Function FieldArgument(ByVal code As String) As String
    Dim arg As String
    Dim stopAt As Long
    code = Replace(Replace(code, vbTab, " "), vbCr, "")
    stopAt = InStr(code, " ")
    If stopAt = 0 Then Exit Function
    arg = LTrim$(Mid$(code, stopAt + 1))
    If Left$(arg, 1) = Chr$(34) Then
        arg = Mid$(arg, 2)
        stopAt = InStr(arg, Chr$(34))
    Else
        stopAt = InStr(arg, " ")
    End If
    If stopAt > 0 Then arg = Left$(arg, stopAt - 1)
    FieldArgument = arg
End Function

Private Function BuiltInPropertyExists(ByVal doc As Document, ByVal name As String) As Boolean
    Dim prop As DocumentProperty
    ' Word 会在枚举内置属性时初始化部分文档属性。只查看域实际用到的名称。
    On Error Resume Next
    Set prop = doc.BuiltInDocumentProperties(name)
    BuiltInPropertyExists = Not (prop Is Nothing)
End Function

Private Sub AddIssue(ByRef issues As Collection, ByVal category As String, ByVal msg As String)
    issues.Add "[" & category & "] " & msg
End Sub

' ---------------------------------------------------------------------------
' 分组报告：硬错误单列，未引用的图表题注只作提示，不计为错误
' ---------------------------------------------------------------------------

Private Sub ShowReport(ByRef issues As Collection, ByRef notes As Collection)
    Dim dialog As New frmReferenceCheck
    dialog.LoadReport ReportGroups(issues, notes), issues.Count, notes.Count
    dialog.Show vbModal
End Sub

Private Function ReportGroups(ByRef issues As Collection, ByRef notes As Collection) As Variant
    ReportGroups = Array(ReportSection("图注（含子图注）", issues, "[图注] "), _
                         ReportSection("表注", issues, "[表注] "), _
                         ReportSection("引用", issues, "[引用] "), _
                         ReportSection("未引用提示", notes))
End Function

Private Function ReportSection(ByVal title As String, ByRef entries As Collection, _
                               Optional ByVal prefix As String = "") As Variant
    Dim item As Variant
    Dim count As Long
    Dim lines As String
    For Each item In entries
        If Left$(CStr(item), Len(prefix)) = prefix Then
            count = count + 1
            lines = lines & vbCrLf & CStr(count) & ". " & Mid$(CStr(item), Len(prefix) + 1)
        End If
    Next item
    If count = 0 Then lines = vbCrLf & "无"
    ReportSection = Array(title, count, Mid$(lines, 3))
End Function

Private Function Summarize(ByRef issues As Collection, ByRef notes As Collection) As String
    Dim item As Variant
    Dim nFigure As Long
    Dim nTable As Long
    Dim nRef As Long
    For Each item In issues
        Select Case Left$(CStr(item), 4)
            Case "[图注]": nFigure = nFigure + 1
            Case "[表注]": nTable = nTable + 1
            Case "[引用]": nRef = nRef + 1
        End Select
    Next item
    Summarize = "图注=" & CStr(nFigure) & "|表注=" & CStr(nTable) _
                & "|引用=" & CStr(nRef) & "|提示=" & CStr(notes.Count)
End Function

' ---------------------------------------------------------------------------
' 测试入口
' ---------------------------------------------------------------------------

Public Function TestTemplateCheck() As String
    Dim issues As Collection
    Dim notes As Collection

    On Error GoTo EH
    Set issues = New Collection
    Set notes = New Collection
    RunCheckCore issues, notes
    TestTemplateCheck = Summarize(issues, notes)
    Exit Function
EH:
    TestTemplateCheck = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function TestCheckAll() As String
    TestCheckAll = TestTemplateCheck()
End Function

Public Function TestReferenceCheckReport() As String
    Dim issues As New Collection
    Dim notes As New Collection
    Dim group As Variant
    Dim title As String

    RunCheckCore issues, notes
    For Each group In ReportGroups(issues, notes)
        title = CStr(group(0))
        If title = "未引用提示" Then title = title & "（不计为错误）"
        TestReferenceCheckReport = TestReferenceCheckReport _
            & title & " · " & CStr(group(1)) & " 项" & vbCrLf & CStr(group(2)) & vbCrLf
    Next group
End Function

Public Function TestReferenceDialog() As String
    Dim issues As New Collection
    Dim notes As New Collection
    Dim dialog As New frmReferenceCheck
    Dim details As MSForms.TextBox
    Dim tabs As MSForms.MultiPage

    AddIssue issues, "表注", "检查示例"
    dialog.LoadReport ReportGroups(issues, notes), issues.Count, notes.Count
    Set tabs = dialog.mpResults
    Set details = tabs.Pages(tabs.Value).Controls("txtDetails")
    TestReferenceDialog = CStr(tabs.Pages.Count) & "|" & tabs.Pages(tabs.Value).Caption _
        & "|" & details.Text & "|" & CStr(details.Locked)
    Unload dialog
End Function

' 在文档末尾一次性追加 extraParas 段，再测一次完整检查的耗时。
Public Function TestCheckLong(ByVal extraParas As Long) As String
    Dim parts() As String
    Dim k As Long
    Dim t0 As Double
    Dim elapsed As Double
    Dim issues As Collection
    Dim notes As Collection

    On Error GoTo EH
    If extraParas < 1 Then extraParas = 1
    If extraParas > 20000 Then extraParas = 20000
    ReDim parts(1 To extraParas)
    For k = 1 To extraParas
        parts(k) = "占位段落 " & CStr(k)
    Next k
    ActiveDocument.Content.InsertAfter vbCr & Join(parts, vbCr)

    Set issues = New Collection
    Set notes = New Collection
    t0 = Timer
    RunCheckCore issues, notes
    elapsed = Timer - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    TestCheckLong = "elapsed=" & Replace$(CStr(Round(elapsed, 3)), ",", ".") _
                    & "|" & Summarize(issues, notes) _
                    & "|paras=" & CStr(ActiveDocument.Paragraphs.Count)
    Exit Function
EH:
    TestCheckLong = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function
