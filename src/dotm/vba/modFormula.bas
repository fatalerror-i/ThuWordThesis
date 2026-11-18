Attribute VB_Name = "modFormula"
Option Explicit

' ============================================================================
' 公式：应用“公式”样式 + Word 内置 #() 编号 + 书签/引用
' 说明：#() 内的域能力随 Word 版本不同；先按域插入，失败则退回纯文本编号。
'       两条路径都保留 #() 的“公式居中、编号右对齐”版式，不使用制表位/表格。
' ============================================================================

' 编号与引用：只枚举“有编号”的行间公式段（段落样式为“公式”、含 m:oMathPara，
' 且含 SEQ/STYLEREF 编号域）；“无编号”行间公式同样是显示公式（保留 oMathPara），
' 但不生成编号域和书签，因此不会进入编号刷新和引用列表。
' 编号直接读 #() 内的域结果，不再逐段遍历全文，成本随公式数量增长（性能优化，2026-11）。


Public Sub RunNormalizeFormula(control As Object, Optional ByVal silent As Boolean = False)
    Dim p As Paragraph
    Dim om As Object
    Dim ur As Object
    Dim ok As Boolean
    Dim hadText As Boolean
    Dim hadMath As Long

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub

    Set p = GetSelectionParagraph()
    If p Is Nothing Then
        If Not silent Then MsgBox "请把光标放在公式行，或选中一个公式。", vbExclamation, "公式"
        Exit Sub
    End If

    hadText = HasOtherText(p)
    hadMath = p.Range.OMaths.Count
    ' 用户手工输入的行内 OMath 先转显示模式；无编号公式本身就是显示公式。
    If hadMath = 1 Then EnsureDisplayMath p

    ' 三态：有文字→行内占位；已有单个行间公式→规范化；空白行→带编号占位
    If hadText Then
        InsertInlineFormulaPlaceholder
        Exit Sub
    End If
    If hadMath > 1 Then
        If Not silent Then MsgBox "当前段落有多个公式，无法自动判断编号对象。" & vbCrLf & _
               "请把光标放到单个公式所在行。", vbExclamation, "公式"
        Exit Sub
    End If
    gDirty = False
    If hadMath = 1 Then
        Set om = p.Range.OMaths(1)
        NormalizeOne p, om, ok
    Else
        InsertDisplayFormulaPlaceholder p, ok
    End If
    If Not ur Is Nothing Then FinishUndo ur, ok

    If Not silent Then
        If ok Then
            MsgBox "公式已应用“公式”样式，并生成编号。" & vbCrLf & _
                   "如编号未显示，请点“更新域”。", vbInformation, "公式"
        Else
            MsgBox "公式未能完整规范化，请检查公式内容后重试。", vbExclamation, "公式"
        End If
    End If
    Exit Sub
EH:
    FinishUndo ur, False
    If Not silent Then ReportError "规范化公式", Err.Number, Err.Description
End Sub

' Ribbon：行内公式。文字行内插入占位并调字体；行内公式内只调字体；
' 空白行或行间公式内复用行间公式逻辑。
Public Sub InsertInlineFormula(control As Object)
    Dim p As Paragraph
    Dim om As Object

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    Set p = GetSelectionParagraph()
    If p Is Nothing Then Exit Sub

    If Selection.Range.OMaths.Count > 0 Then
        Set om = Selection.Range.OMaths(1)
        If IsDisplayFormulaMath(om) Then
            RunNormalizeFormula control, True
        Else
            AdjustInlineFormulaFont om
        End If
        Exit Sub
    End If

    ' 选中文字时，把选中内容转换成行内公式，不能插入占位符覆盖原文。
    If Selection.Type <> wdSelectionIP Then
        If Len(CleanText(Selection.Text)) > 0 Then
            ActiveDocument.OMaths.Add Selection.Range
            If p.Range.OMaths.Count > 0 Then
                Set om = p.Range.OMaths(p.Range.OMaths.Count)
                om.BuildUp
                AdjustInlineFormulaFont om
            End If
            Exit Sub
        End If
    End If

    If HasOtherText(p) Then
        InsertInlineFormulaPlaceholder
        If p.Range.OMaths.Count > 0 Then
            Set om = p.Range.OMaths(p.Range.OMaths.Count)
            AdjustInlineFormulaFont om
        End If
    Else
        ' 空白行或行间公式：复用行间公式功能
        RunNormalizeFormula control, True
    End If
    Exit Sub
EH:
    If Not gSilentClicks Then ReportError "行内公式", Err.Number, Err.Description
End Sub

' Ribbon：行间公式。空白行/已有行间公式直接处理；文字行内则在其后新建一段。
Public Sub InsertDisplayFormula(control As Object)
    Dim p As Paragraph

    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    Set p = GetSelectionParagraph()
    If p Is Nothing Then Exit Sub

    If HasOtherText(p) Then
        p.Range.Select
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
    End If
    RunNormalizeFormula control, True
    Exit Sub
EH:
    If Not gSilentClicks Then ReportError "行间公式", Err.Number, Err.Description
End Sub

' Ribbon：无编号行间公式。只套“公式”样式并居中，不生成 #() 编号和书签。
Public Sub InsertDisplayFormulaPlain(control As Object)
    Dim p As Paragraph
    Dim prev As Boolean

    prev = True
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub
    Set p = GetSelectionParagraph()
    If p Is Nothing Then Exit Sub
    If p.Range.OMaths.Count > 1 Then
        MsgBox "当前段落有多个公式，无法自动处理。", vbExclamation, "无编号行间公式"
        Exit Sub
    End If

    prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    If HasOtherText(p) Then
        p.Range.Select
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
        Set p = GetSelectionParagraph()
    End If
    ApplyPlainDisplayFormula p
    Application.ScreenUpdating = prev
    Application.StatusBar = "已应用无编号行间公式。"
    Exit Sub
EH:
    Application.ScreenUpdating = prev
    If Not gSilentClicks Then ReportError "无编号行间公式", Err.Number, Err.Description
End Sub

' 把段落变成无编号行间公式：套“公式”样式、居中，保留显示公式（oMathPara），
' 但去掉已有 #() 编号、编号域和 THU_EQ_ 书签。
Private Sub ApplyPlainDisplayFormula(p As Paragraph)
    Dim om As Object
    Dim txt As String
    Dim rng As Range

    On Error GoTo EH

    If p.Range.OMaths.Count = 0 Then
        ' 空白段：套样式后在段内建显示公式。OMaths.Add 会生成空位占位符，不含编号。
        p.Range.Select
        Selection.Collapse wdCollapseStart
        p.Style = ActiveDocument.Styles(STYLE_FORMULA)
        Set rng = p.Range.Duplicate
        If rng.Characters.Count > 0 Then rng.MoveEnd wdCharacter, -1
        ' OMaths.Add 的返回值在部分版本是 Range，不能直接 BuildUp；按段落重新取 OMath。
        ActiveDocument.OMaths.Add rng
        If p.Range.OMaths.Count = 0 Then Exit Sub
    Else
        ' 已有公式：线性化后去掉外层容器和 #() 编号，再重建为显示公式。
        Set om = p.Range.OMaths(1)
        om.Linearize
        txt = PlainFormulaText(om.Range.Text)
        If Len(txt) = 0 Then txt = "□"
        om.Range.Text = txt
        om.BuildUp
    End If

    Set om = p.Range.OMaths(1)
    ' 无编号公式仍是显示公式；显式设为显示模式，避免停在行内模式。
    On Error Resume Next
    om.Type = wdOMathDisplay
    On Error GoTo EH

    DeleteFormulaBookmarksIn p
    DeleteFormulaNumberFields p

    With p
        .Style = ActiveDocument.Styles(STYLE_FORMULA)
        .Format.Alignment = wdAlignParagraphCenter
        .Format.LeftIndent = 0
        .Format.RightIndent = 0
        .Format.FirstLineIndent = 0
        .Format.TabStops.ClearAll
    End With
    p.Range.OMaths(1).Range.Select
    Exit Sub
EH:
    If Not gSilentClicks Then ReportError "无编号行间公式", Err.Number, Err.Description
End Sub

' 去掉线性公式末尾的 #(...) 编号（含内部的 STYLEREF/SEQ 域文本）。
Private Function StripDisplayNumber(ByVal s As String) As String
    Dim res As String
    Dim hashPos As Long
    Dim openPos As Long
    Dim closePos As Long
    Dim depth As Long
    Dim k As Long

    res = s
    Do
        hashPos = InStrRev(res, "#")
        If hashPos = 0 Then Exit Do
        openPos = InStr(hashPos + 1, res, "(")
        If openPos = 0 Then Exit Do
        depth = 0
        closePos = 0
        For k = openPos To Len(res)
            Select Case Mid$(res, k, 1)
                Case "("
                    depth = depth + 1
                Case ")"
                    depth = depth - 1
                    If depth = 0 Then
                        closePos = k
                        Exit For
                    End If
            End Select
        Next k
        If closePos = 0 Then Exit Do
        res = RTrim$(Left$(res, hashPos - 1) & Mid$(res, closePos + 1))
    Loop
    StripDisplayNumber = res
End Function

' 把带编号的线性公式文本变成无编号公式内容：去掉外层 █( ) 容器和末尾 #(...) 编号。
Private Function PlainFormulaText(ByVal s As String) As String
    Dim res As String

    res = Trim$(s)
    If Left$(res, 2) = ChrW(&H2588) & "(" Then
        res = Mid$(res, 3)
        If Right$(res, 1) = ")" Then res = Left$(res, Len(res) - 1)
        res = Trim$(res)
    End If
    PlainFormulaText = Trim$(StripDisplayNumber(res))
End Function

' 删除段落内的 SEQ/STYLEREF 编号域（无编号公式不应保留编号域）。
Private Sub DeleteFormulaNumberFields(p As Paragraph)
    Dim i As Long
    Dim code As String

    On Error Resume Next
    For i = p.Range.Fields.Count To 1 Step -1
        code = UCase$(p.Range.Fields(i).Code.Text)
        If InStr(code, "SEQ") > 0 Or InStr(code, "STYLEREF") > 0 Then p.Range.Fields(i).Delete
    Next i
    On Error GoTo 0
End Sub

' 删除落在段落 p 内的公式书签（THU_EQ_ 前缀）。
Private Sub DeleteFormulaBookmarksIn(p As Paragraph)
    Dim bm As Bookmark
    Dim names() As String
    Dim n As Long
    Dim i As Long

    On Error Resume Next
    n = 0
    ReDim names(0)
    For Each bm In ActiveDocument.Bookmarks
        If Left$(bm.Name, Len(BM_EQ_PREFIX)) = BM_EQ_PREFIX Then
            If bm.Range.Start >= p.Range.Start And bm.Range.Start < p.Range.End Then
                ReDim Preserve names(n)
                names(n) = bm.Name
                n = n + 1
            End If
        End If
    Next bm
    For i = 0 To n - 1
        If ActiveDocument.Bookmarks.Exists(names(i)) Then ActiveDocument.Bookmarks(names(i)).Delete
    Next i
    On Error GoTo 0
End Sub

Private Function CountFormulaBookmarksIn(p As Paragraph) As Long
    Dim bm As Bookmark
    Dim n As Long

    For Each bm In ActiveDocument.Bookmarks
        If Left$(bm.Name, Len(BM_EQ_PREFIX)) = BM_EQ_PREFIX Then
            If bm.Range.Start >= p.Range.Start And bm.Range.Start < p.Range.End Then n = n + 1
        End If
    Next bm
    CountFormulaBookmarksIn = n
End Function

' 行间公式的 OMML 外层是 m:oMathPara；行内公式没有这个包裹。
Private Function IsDisplayFormulaMath(om As Object) As Boolean
    On Error Resume Next
    IsDisplayFormulaMath = (InStr(1, om.Range.WordOpenXML, "oMathPara", vbTextCompare) > 0)
    On Error GoTo 0
End Function

' 行内公式统一套“行内公式”字符样式；样式缺失时才退回直接字体。
Private Sub AdjustInlineFormulaFont(om As Object)
    Dim applied As Boolean

    On Error Resume Next
    om.Range.Style = ActiveDocument.Styles(STYLE_INLINE_FORMULA)
    applied = (Err.Number = 0)
    Err.Clear
    If Not applied Then
        om.Range.Font.Name = "XITS Math"
        om.Range.Font.NameAscii = "XITS Math"
        om.Range.Font.NameFarEast = "XITS Math"
        om.Range.Font.Size = 12
        om.Range.Font.SizeBi = 12
    End If
    On Error GoTo 0
End Sub

' 当前段落除公式外是否还有文字（用于三态判断）
Private Function HasOtherText(p As Paragraph) As Boolean
    Dim om As Object
    Dim r As Range
    Dim s As String
    If p.Range.OMaths.Count = 0 Then
        HasOtherText = (Len(CleanText(p.Range.Text)) > 0)
        Exit Function
    End If
    Set om = p.Range.OMaths(1)
    Set r = p.Range.Duplicate
    r.End = om.Range.Start
    s = r.Text
    Set r = p.Range.Duplicate
    r.Start = om.Range.End
    s = s & r.Text
    ' 去掉 #() 编号残留，避免把已有编号误判为正文
    s = CleanText(s)
    If s Like "#*" Then s = ""
    If s Like "#" & "[0-9.]*" Then s = ""
    HasOtherText = (Len(s) > 0)
End Function

' 行内公式占位符：不套“公式”段落样式、不加编号
Private Sub InsertInlineFormulaPlaceholder()
    Dim r As Range
    Dim om As Object
    Selection.TypeText "□"
    Set r = Selection.Range
    r.MoveStart wdCharacter, -1
    ActiveDocument.OMaths.Add r
    If ActiveDocument.OMaths.Count > 0 Then
        Set om = ActiveDocument.OMaths(ActiveDocument.OMaths.Count)
        om.BuildUp
    End If
    Selection.Collapse wdCollapseEnd
End Sub

' 空白行：插入带 #() 编号的独立公式占位符
Private Sub InsertDisplayFormulaPlaceholder(p As Paragraph, ByRef ok As Boolean)
    Dim r As Range
    Dim om As Object

    ok = False
    On Error GoTo EH
    p.Range.Select
    Selection.Collapse wdCollapseStart
    Selection.TypeText "□"
    Set r = Selection.Range
    r.MoveStart wdCharacter, -1
    ActiveDocument.OMaths.Add r
    Set om = p.Range.OMaths(1)
    om.BuildUp
    Set om = p.Range.OMaths(1)
    FinalizeFormula p, om, ok
    Exit Sub
EH:
    ok = False
End Sub

' 供“更新域”调用：只枚举“公式”段刷新编号书签；域编号交给 Word 更新。
' 用 Find 按样式枚举，成本与公式数量成正比，不再逐段遍历全文。
Public Sub RefreshFormulaNumbers(Optional ByVal refreshAppendices As Boolean = True)
    Dim rng As Range
    Dim p As Paragraph
    Dim prevUpdating As Boolean
    Dim bmName As String
    Dim lastStart As Long

    prevUpdating = True
    On Error Resume Next
    prevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    If refreshAppendices Then RefreshAppendixSequences
    lastStart = -1
    Set rng = FormulaFindRange()
    Do While rng.Find.Execute
        Set p = rng.Paragraphs(1)
        ' 文末公式段：Find 可能再次命中同一段，避免死循环。
        If p.Range.Start = lastStart Then Exit Do
        lastStart = p.Range.Start
        If IsFormulaParagraph(p) Then
            bmName = FormulaBookmarkName(p)
            p.Range.Fields.Update
            BindFormulaBookmark p, bmName
        End If
        rng.SetRange p.Range.End, p.Range.End
    Loop
    Application.ScreenUpdating = prevUpdating
    On Error GoTo 0
End Sub

Private Sub NormalizeOne(p As Paragraph, om As Object, ByRef ok As Boolean)
    Dim txt As String
    Dim numText As String
    Dim bmName As String

    ok = False
    On Error GoTo EH
    bmName = FormulaBookmarkName(p)

    ' 1) 线性化，去掉 Word 公式编号并保留公式主体
    om.Linearize
    txt = om.Range.Text
    txt = StripTrailingNumber(txt)
    om.Range.Text = txt
    om.BuildUp

    Set om = p.Range.OMaths(1)
    FinalizeFormula p, om, ok, bmName
    Exit Sub
EH:
    ' 调用方负责撤销；这里只把 ok 置 False
    ok = False
End Sub

' 套“公式”样式、居中、清制表位、加 #() 编号与编号书签
Private Sub FinalizeFormula(p As Paragraph, om As Object, ByRef ok As Boolean, _
                            Optional ByVal bmName As String = "")
    Dim useFields As Boolean
    Dim bApp As Boolean
    Dim prevUpdating As Boolean

    ok = False
    prevUpdating = True
    On Error GoTo EH
    prevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    p.Style = ActiveDocument.Styles(STYLE_FORMULA)
    p.Alignment = wdAlignParagraphCenter
    p.TabStops.ClearAll
    p.FirstLineIndent = 0

    ' 附录判定只按样式查“附录标题”段，不再遍历全文。
    bApp = IsAppendixFormula(p)

    useFields = True
    On Error Resume Next
    BuildNumberWithFields p, om, bApp
    If Err.Number <> 0 Then
        Err.Clear
        useFields = False
    End If
    On Error GoTo EH

    If Not useFields Then
        BuildNumberAsText p, om
    End If

    ' BuildUp 可能重建行间公式所在的段落对象。选区仍在该公式内。
    Set p = Selection.Paragraphs(1)
    p.Range.Fields.Update
    bmName = BindFormulaBookmark(p, bmName)
    If Len(bmName) = 0 Then
        Err.Raise vbObjectError + 1002, "modFormula", "未能为公式编号建立书签"
    End If
    If bApp Then
        RefreshAppendixSequences
        ActiveDocument.Bookmarks(bmName).Range.Paragraphs(1).Range.Fields.Update
    End If
    Application.ScreenUpdating = prevUpdating
    ok = True
    Exit Sub
EH:
    On Error Resume Next
    Application.ScreenUpdating = prevUpdating
    ok = False
End Sub


Private Sub BuildNumberWithFields(p As Paragraph, om As Object, ByVal bApp As Boolean)
    Dim chapterCode As String
    Dim r As Range

    If bApp Then
        chapterCode = "STYLEREF ""附录标题"" \s"
    Else
        chapterCode = "STYLEREF 1 \s"
    End If

    om.Linearize
    om.Range.Select
    Selection.Collapse wdCollapseEnd
    Selection.TypeText "#("
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
                         Text:=chapterCode, PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
    Selection.TypeText "."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
                         Text:="SEQ 式 \* ARABIC \s 1", PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
    Selection.TypeText ")"
    om.BuildUp
End Sub

Private Sub BuildNumberAsText(p As Paragraph, om As Object)
    Dim txt As String
    Dim numText As String

    numText = GetFormulaNumberText(p)
    om.Linearize
    txt = StripTrailingNumber(om.Range.Text)
    om.Range.Text = txt & "#(" & numText & ")"
    om.BuildUp
End Sub

' 仅清理带 # 标记的 Word 公式编号；普通括号属于公式主体。
' 注意：om.Range.Text 可能是 "█(公式#(2.1) )"；#() 不一定在最后一个字符，
' 所以要先定位 #(...) 片段，只删除该片段，保留 OMML 的定界符。
Private Function StripTrailingNumber(ByVal s As String) As String
    Dim t As String
    Dim hashPos As Long
    Dim openPos As Long
    Dim closePos As Long
    Dim inner As String
    Dim cand As String
    Dim i As Long
    Dim ch As String

    t = s
    If Len(t) = 0 Then Exit Function

    hashPos = InStrRev(t, "#")
    If hashPos > 0 Then
        ' 1) #(2.1)
        openPos = InStr(hashPos, t, "(")
        If openPos = hashPos + 1 Then
            closePos = InStr(openPos, t, ")")
            If closePos > openPos Then
                inner = Mid$(t, openPos + 1, closePos - openPos - 1)
                If IsNumberLike(inner) And HasDigit(inner) Then
                    StripTrailingNumber = Left$(t, hashPos - 1) & Mid$(t, closePos + 1)
                    Exit Function
                End If
            End If
        End If

        ' 2) #2.1（# 后可能是换行/空格）
        i = hashPos + 1
        Do While i <= Len(t)
            ch = Mid$(t, i, 1)
            If ch = " " Or ch = vbCr Or ch = vbLf Or ch = vbTab Then
                i = i + 1
            Else
                Exit Do
            End If
        Loop
        cand = ""
        Do While i <= Len(t)
            ch = Mid$(t, i, 1)
            If IsNumberChar(ch) Then
                cand = cand & ch
                i = i + 1
            Else
                Exit Do
            End If
        Loop
        If Len(cand) > 0 And IsNumberLike(cand) And HasDigit(cand) Then
            StripTrailingNumber = Left$(t, hashPos - 1) & Mid$(t, i)
            Exit Function
        End If
    End If

    StripTrailingNumber = t
End Function

Private Function IsNumberChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    If (ch >= "0" And ch <= "9") Or _
       (ch >= "A" And ch <= "Z") Or _
       (ch >= "a" And ch <= "z") Or _
       ch = "." Or ch = "-" Then
        IsNumberChar = True
    End If
End Function

Private Function HasDigit(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            HasDigit = True
            Exit Function
        End If
    Next i
End Function

' 括号内是否像公式编号：数字/字母/点/横线，长度不超过 8
Private Function IsNumberLike(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String

    s = Trim$(s)
    If Len(s) = 0 Or Len(s) > 8 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or _
                (ch >= "A" And ch <= "Z") Or _
                (ch >= "a" And ch <= "z") Or _
                ch = "." Or ch = "-" Or ch = " ") Then
            Exit Function
        End If
    Next i
    IsNumberLike = True
End Function

' 取末尾的 (2.1)/(A.1)；没有则空串
Public Function ExtractTrailingParenText(ByVal s As String) As String
    Dim pos As Long
    Dim closePos As Long

    closePos = InStrRev(s, ")")
    If closePos <> Len(s) Then Exit Function
    pos = InStrRev(s, "(")
    If pos = 0 Or pos >= closePos Then Exit Function
    If Not IsNumberLike(Mid$(s, pos + 1, closePos - pos - 1)) Then Exit Function
    ExtractTrailingParenText = Mid$(s, pos, closePos - pos + 1)
End Function

' ---------------------------------------------------------------------------
' 编号读取：域结果 / 附录判定 / 按样式查找公式段
' ---------------------------------------------------------------------------

' 从 #() 内的 STYLEREF/SEQ 域结果取“章号.章内序号”；读不到返回空串。
Private Function FormulaFieldNumText(p As Paragraph) As String
    Dim fld As Field
    Dim code As String
    Dim chapter As String
    Dim seqText As String

    For Each fld In p.Range.Fields
        code = UCase$(fld.Code.Text)
        If InStr(code, "STYLEREF") > 0 Then
            chapter = Trim$(CleanText(fld.Result.Text))
        ElseIf InStr(code, "SEQ") > 0 Then
            seqText = Trim$(CleanText(fld.Result.Text))
        End If
    Next fld
    If Len(chapter) > 0 And Len(seqText) > 0 Then
        FormulaFieldNumText = chapter & "." & seqText
    End If
End Function

Private Function FormulaBookmarkName(p As Paragraph) As String
    Dim bm As Bookmark
    For Each bm In p.Range.Bookmarks
        If Left$(bm.Name, Len(BM_EQ_PREFIX)) = BM_EQ_PREFIX And Not bm.Empty Then
            FormulaBookmarkName = bm.Name
            Exit Function
        End If
    Next bm
End Function

Private Function BindFormulaBookmark(p As Paragraph, ByVal bmName As String) As String
    Dim f As Field
    Dim chapter As Range
    Dim seq As Range

    For Each f In p.Range.Fields
        If Left$(UCase$(Trim$(f.Code.Text)), 9) = "STYLEREF " Then
            Set chapter = f.Result
        ElseIf Left$(Trim$(f.Code.Text), 5) = "SEQ 式" Then
            Set seq = f.Result
        End If
    Next f
    If chapter Is Nothing Or seq Is Nothing Then Exit Function
    If Len(bmName) = 0 Then bmName = NewFormulaBookmarkName()
    chapter.End = seq.End
    ActiveDocument.Bookmarks.Add bmName, chapter
    BindFormulaBookmark = bmName
End Function

Private Function NewFormulaBookmarkName() As String
    Const COUNTER_NAME As String = "THU_EQ_LAST_ID"
    Dim doc As Document
    Dim counterValue As String
    Dim f As Field
    Dim parts As Variant
    Dim refCode As String
    Dim n As Long
    Dim candidate As String
    Dim reserved As Object

    Set doc = ActiveDocument
    On Error Resume Next
    counterValue = doc.Variables(COUNTER_NAME).Value
    On Error GoTo 0
    Set reserved = CreateObject("Scripting.Dictionary")
    reserved.CompareMode = vbTextCompare
    ' 复制过来的 REF 也要保留目标名，即使当前文档没有对应书签。
    For Each f In doc.Fields
        refCode = ""
        ' Word 更新嵌套域后，枚举器偶尔仍返回已删除的旧域对象。
        On Error Resume Next
        refCode = Trim$(Replace(f.Code.Text, vbTab, " "))
        On Error GoTo 0
        If Left$(UCase$(refCode), 4) = "REF " Then
            parts = Split(Replace(Trim$(Mid$(refCode, 5)), """", ""), " ")
            If UBound(parts) >= 0 Then reserved(parts(0)) = True
        End If
    Next f
    n = CLng(Val(counterValue))
    Do
        n = n + 1
        candidate = BM_EQ_PREFIX & "N" & CStr(n)
    Loop While doc.Bookmarks.Exists(candidate) Or reserved.Exists(candidate)
    If Len(counterValue) = 0 Then
        doc.Variables.Add COUNTER_NAME, CStr(n)
    Else
        doc.Variables(COUNTER_NAME).Value = CStr(n)
    End If
    NewFormulaBookmarkName = candidate
End Function

Private Function IsAppendixFormula(p As Paragraph) As Boolean
    IsAppendixFormula = IsAppendixAt(p.Range)
End Function

' 是否为带编号的行间公式段：段落样式“公式”，恰有一个 OMath，
' 且该 OMath 是独立行间公式（oMathPara）。行内公式不算。
' 把行内 OMath 转为显示模式（oMathPara），供有编号流程使用。
Private Sub EnsureDisplayMath(p As Paragraph)
    If p.Range.OMaths.Count <> 1 Then Exit Sub
    On Error Resume Next
    If InStr(p.Range.WordOpenXML, "oMathPara") = 0 Then
        p.Range.OMaths(1).Type = wdOMathDisplay
    End If
    On Error GoTo 0
End Sub

' 有编号的行间公式段：样式“公式”、含 m:oMathPara，且含 SEQ/STYLEREF 编号域。
' 无编号公式同样是显示公式（有 oMathPara），但没有编号域，因此排除。
Private Function IsFormulaParagraph(p As Paragraph) As Boolean
    Dim x As String

    On Error GoTo 0
    If p.Style.NameLocal <> STYLE_FORMULA Then Exit Function
    On Error GoTo Fallback
    x = p.Range.WordOpenXML
    If InStr(x, "oMathPara") = 0 Then Exit Function
    IsFormulaParagraph = HasFormulaNumberField(p)
    Exit Function
Fallback:
    IsFormulaParagraph = True
End Function

' 段落内是否有公式编号域（SEQ 式 / STYLEREF）。
Private Function HasFormulaNumberField(p As Paragraph) As Boolean
    Dim i As Long
    Dim code As String

    On Error Resume Next
    For i = 1 To p.Range.Fields.Count
        code = UCase$(p.Range.Fields(i).Code.Text)
        If InStr(code, "SEQ") > 0 Or InStr(code, "STYLEREF") > 0 Then
            HasFormulaNumberField = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

' 返回一个已配置为“按样式查找行间公式段”的 Range；调用方用 Find.Execute 推进。
Private Function FormulaFindRange() As Range
    Dim rng As Range

    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Style = STYLE_FORMULA
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
    End With
    Set FormulaFindRange = rng
End Function

' 纯文本路径的编号：优先读 #() 内的域结果；读不到时回退到逐段计算（仅兜底）。
Public Function GetFormulaNumberText(p As Paragraph) As String
    Dim t As String

    t = FormulaFieldNumText(p)
    If Len(t) > 0 Then
        GetFormulaNumberText = "(" & t & ")"
    Else
        GetFormulaNumberText = GetFormulaNumberTextSlow(p)
    End If
End Function

' 兜底实现：按当前章与本章公式序号逐段计算（仅在读不到域结果时使用）。
Private Function GetFormulaNumberTextSlow(p As Paragraph) As String
    Dim q As Paragraph
    Dim h As Paragraph
    Dim chapter As String
    Dim count As Long
    Dim inCurrent As Boolean

    Set h = PreviousSectionHeading(p)
    If Not h Is Nothing Then
        chapter = GetHeadingNumberText(h)
    End If
    If Len(chapter) = 0 Then chapter = "1"

    count = 0
    inCurrent = (h Is Nothing)
    For Each q In ActiveDocument.Paragraphs
        If q.Range.Start >= p.Range.Start Then Exit For
        If IsChapterHeading(q) Or IsAppendixHeading(q) Then
            If h Is Nothing Then
                inCurrent = True
            Else
                If q.Range.Start >= h.Range.Start Then inCurrent = True
                If q.Range.Start > h.Range.Start Then Exit For
            End If
        End If
        If inCurrent And IsFormulaParagraph(q) Then
            count = count + 1
        End If
    Next q

    GetFormulaNumberTextSlow = "(" & chapter & "." & CStr(count + 1) & ")"
End Function

Private Function PreviousSectionHeading(p As Paragraph) As Paragraph
    Dim q As Paragraph
    Dim found As Paragraph

    For Each q In ActiveDocument.Paragraphs
        If q.Range.Start >= p.Range.Start Then Exit For
        If IsChapterHeading(q) Or IsAppendixHeading(q) Then Set found = q
    Next q
    Set PreviousSectionHeading = found
End Function

' ---------------------------------------------------------------------------
' 公式引用：按编号找书签，插入 REF 域
' ---------------------------------------------------------------------------

' 光标若在 OMath 内，移到段落末尾，避免把 REF 域插进公式。
Public Sub EnsureOutsideFormula()
    Dim p As Paragraph
    If Selection.Range.OMaths.Count = 0 Then Exit Sub
    Set p = Selection.Paragraphs(1)
    If p.Range.End > p.Range.Start Then
        Selection.SetRange p.Range.End - 1, p.Range.End - 1
    End If
End Sub

Public Sub RunInsertFormulaReference(control As IRibbonControl)
    Dim rng As Range
    Dim p As Paragraph
    Dim bmName As String
    Dim numText As String
    Dim preview As String
    Dim k As Long
    Dim idx As Long
    Dim first As Boolean
    Dim prevUpdating As Boolean
    Dim errNo As Long
    Dim errDesc As String
    Dim lastStart As Long

    prevUpdating = True
    On Error GoTo EH
    If Not EnsureReady() Then Exit Sub

    ' 先刷新编号，使列表与原生交叉引用一致；
    ' RefreshFormulaNumbers 只枚举公式段，成本随公式数量增长。
    RefreshFormulaNumbers

    prevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    ResetPicker
    lastStart = -1
    Set rng = FormulaFindRange()
    Do While rng.Find.Execute
        Set p = rng.Paragraphs(1)
        If p.Range.Start = lastStart Then Exit Do
        lastStart = p.Range.Start
        If IsFormulaParagraph(p) Then
            bmName = FormulaBookmarkName(p)
            numText = FormulaFieldNumText(p)
            If Len(bmName) > 0 And Len(numText) > 0 Then
                preview = CleanText(p.Range.Text)
                If Len(preview) > 60 Then preview = Left$(preview, 60) & "…"
                AddPickerItem numText & "  " & preview, bmName
            End If
        End If
        rng.SetRange p.Range.End, p.Range.End
    Loop
    Application.ScreenUpdating = prevUpdating

    If PickerItemCount() = 0 Then
        MsgBox "没有找到已编号的公式。" & vbCrLf & _
               "请先对公式点“公式”生成编号，或点“更新域”。", vbExclamation, "引用公式"
        Exit Sub
    End If
    If Not ShowPicker("引用公式", "勾选要引用的公式（可多选）：") Then Exit Sub

    EnsureOutsideFormula
    first = True
    For k = 0 To gPickerSelectedCount - 1
        idx = gPickerSelected(k)
        bmName = gPickerItems(idx).Key
        If Len(bmName) > 0 Then
            If Not first Then Selection.TypeText "、"
            InsertFormulaRefField bmName
            first = False
        End If
    Next k
    Exit Sub
EH:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = prevUpdating
    ReportError "引用公式", errNo, errDesc
End Sub

Public Function NormalizeParenNumber(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, "（", "(")
    t = Replace(t, "）", ")")
    If Left$(t, 1) <> "(" Then t = "(" & t
    If Right$(t, 1) <> ")" Then t = t & ")"
    NormalizeParenNumber = t
End Function

Public Function FindBookmarkByText(ByVal prefix As String, ByVal text As String) As String
    Dim bm As Bookmark
    Dim s As String

    For Each bm In ActiveDocument.Bookmarks
        If Left$(bm.Name, Len(prefix)) = prefix Then
            s = CleanText(bm.Range.Text)
            If s = text Then
                FindBookmarkByText = bm.Name
                Exit Function
            End If
        End If
    Next bm
End Function

Public Sub InsertRefField(ByVal prefixText As String, ByVal bmName As String)
    If Len(prefixText) > 0 Then Selection.TypeText prefixText
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
                         Text:="REF " & bmName & " \h", PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
End Sub

' 公式引用插入可更新的 REF 域；域结果字体为 XITS Math。
Public Sub InsertFormulaRefField(ByVal bmName As String)
    Dim f As Field
    Dim r As Range
    Dim code As String

    code = "REF " & bmName & " \h \* MERGEFORMAT"
    Selection.Font.Name = "XITS Math"
    Selection.TypeText "式("
    Set f = Selection.Fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, _
                                 Text:=code, PreserveFormatting:=False)
    f.Update
    On Error Resume Next
    f.Result.Font.Name = "XITS Math"
    f.Result.Font.NameAscii = "XITS Math"
    f.Result.Font.NameFarEast = "XITS Math"
    f.Code.Font.Name = "XITS Math"
    On Error GoTo 0
    Set r = ActiveDocument.Range(f.Result.End, f.Result.End)
    r.InsertAfter ")"
End Sub

' 兼容旧测试入口：改为域引用。
Public Sub InsertFormulaRefText(ByVal bmName As String)
    InsertFormulaRefField bmName
End Sub


' 测试入口：构建期已带 #() 编号的公式，应能建立书签
Public Function TestFormulaReferenceExisting() As String
    Dim rng As Range
    Dim p As Paragraph
    Dim bmName As String
    Dim idx As Long
    Dim lastStart As Long

    On Error GoTo EH
    RefreshFormulaNumbers
    idx = 0
    lastStart = -1
    Set rng = FormulaFindRange()
    Do While rng.Find.Execute
        Set p = rng.Paragraphs(1)
        If p.Range.Start = lastStart Then Exit Do
        lastStart = p.Range.Start
        If IsFormulaParagraph(p) Then
            idx = idx + 1
            bmName = FormulaBookmarkName(p)
            If ActiveDocument.Bookmarks.Exists(bmName) Then
                TestFormulaReferenceExisting = "ok|" & CStr(idx) & "|" & bmName & "|" & _
                    CleanText(ActiveDocument.Bookmarks(bmName).Range.Text)
                Exit Function
            End If
        End If
        rng.SetRange p.Range.End, p.Range.End
    Loop
    TestFormulaReferenceExisting = "NO_BM"
    Exit Function
EH:
    TestFormulaReferenceExisting = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：公式引用是否为普通文本（OMath 数应为 0）
Public Function TestFormulaRefText() As String
    Dim r As String
    Dim p As Paragraph
    Dim f As Field
    Dim fontName As String
    r = TestFormulaReference()
    Set p = ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count)
    If p.Range.Fields.Count > 0 Then
        Set f = p.Range.Fields(1)
        fontName = f.Result.Font.Name
    End If
    TestFormulaRefText = r & "|fields=" & CStr(p.Range.Fields.Count) & _
                         "|font=" & fontName
End Function

' 测试入口：Application.Run("TestFormulaNormalize")
Public Function TestFormulaNormalize() As String
    Dim p As Paragraph
    Dim om As Object
    Dim ok As Boolean

    On Error GoTo EH
    For Each p In ActiveDocument.Paragraphs
        If p.Style.NameLocal = STYLE_FORMULA And p.Range.OMaths.Count = 1 Then
            Set om = p.Range.OMaths(1)
            NormalizeOne p, om, ok
            Set om = p.Range.OMaths(1)
            TestFormulaNormalize = "ok=" & CStr(ok) & "|text=" & _
                Replace(Replace(CleanText(p.Range.Text), vbCr, ""), vbLf, "")
            Exit Function
        End If
    Next p
    TestFormulaNormalize = "NO_FORMULA"
    Exit Function
EH:
    TestFormulaNormalize = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：Application.Run("TestFormulaReference")
Public Function TestFormulaReference() As String
    Dim b As Bookmark
    Dim bm As String

    On Error GoTo EH
    For Each b In ActiveDocument.Bookmarks
        If Left$(b.Name, Len(BM_EQ_PREFIX)) = BM_EQ_PREFIX Then
            bm = b.Name
            Exit For
        End If
    Next b
    If Len(bm) = 0 Then
        TestFormulaReference = "NO_BM"
        Exit Function
    End If
    Selection.EndKey 6
    Selection.TypeParagraph
    Selection.TypeText "引用："
    InsertFormulaRefField bm
    TestFormulaReference = "ok|" & CleanText(Selection.Paragraphs(1).Range.Text)
    ActiveDocument.Fields.Update
    Exit Function
EH:
    TestFormulaReference = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：两个公式按钮（行内插入/调字体、空白行回退行间、行间新段）
Public Function TestFormulaButtons() As String
    Dim doc As Document
    Dim p As Paragraph
    Dim om As Object
    Dim r As Range
    Dim inlineOk As Boolean
    Dim adjustOk As Boolean
    Dim blankOk As Boolean
    Dim displayOk As Boolean
    Dim stage As String

    On Error GoTo EH
    Set doc = ActiveDocument

    stage = "inline_insert"
    Set p = doc.Content.Paragraphs.Add
    p.Range.InsertBefore "文字"
    Set r = p.Range
    r.End = r.Start + 2
    r.Select
    InsertInlineFormula Nothing
    inlineOk = (p.Range.OMaths.Count = 1) And (InStr(p.Range.Text, "#") = 0)
    If p.Range.OMaths.Count = 1 Then
        Set om = p.Range.OMaths(1)
        inlineOk = inlineOk And (InStr(om.Range.Text, "文字") > 0)
        inlineOk = inlineOk And (om.Range.Style.NameLocal = STYLE_INLINE_FORMULA)
        inlineOk = inlineOk And (InStr(om.Range.Text, "□") = 0)
    End If

    stage = "inline_adjust"
    If p.Range.OMaths.Count = 1 Then
        Set om = p.Range.OMaths(1)
        om.Range.Select
        InsertInlineFormula Nothing
        adjustOk = (p.Range.OMaths.Count = 1) And (om.Range.Style.NameLocal = STYLE_INLINE_FORMULA)
    End If

    stage = "blank_fallback"
    Set p = doc.Content.Paragraphs.Add
    p.Range.Select
    InsertInlineFormula Nothing
    blankOk = (p.Range.OMaths.Count = 1) And (InStr(p.Range.Text, "#") > 0)

    stage = "display_after_text"
    Set p = doc.Content.Paragraphs.Add
    p.Range.InsertBefore "文字"
    p.Range.Select
    Selection.Collapse wdCollapseEnd
    InsertDisplayFormula Nothing
    displayOk = (p.Next.Range.OMaths.Count = 1) And (InStr(p.Next.Range.Text, "#") > 0)

    TestFormulaButtons = "ok|" & CStr(inlineOk) & "|" & CStr(adjustOk) & _
                         "|" & CStr(blankOk) & "|" & CStr(displayOk)
    Exit Function
EH:
    TestFormulaButtons = "FAIL|" & stage & "|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：模拟 Ribbon“公式”点击
Public Function TestFormulaClickPath() As String
    Dim p As Paragraph

    On Error GoTo EH
    For Each p In ActiveDocument.Paragraphs
        If p.Style.NameLocal = STYLE_FORMULA Then
            p.Range.Select
            gSilentClicks = True
            NormalizeFormula Nothing
            gSilentClicks = False
            TestFormulaClickPath = "ok"
            Exit Function
        End If
    Next p
    TestFormulaClickPath = "NO_FORMULA"
    Exit Function
EH:
    gSilentClicks = False
    TestFormulaClickPath = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：单独验证 OMaths.Add 是否可用
Public Function TestOMathAdd() As String
    Dim doc As Document
    Dim p As Paragraph
    Dim om As Object
    Dim stage As String

    On Error GoTo EH
    Set doc = ActiveDocument
    Set p = doc.Content.Paragraphs.Add
    p.Range.Select
    Selection.Collapse wdCollapseStart
    stage = "add"
    doc.OMaths.Add Selection.Range
    stage = "get"
    Set om = p.Range.OMaths(1)
    stage = "text"
    om.Range.Text = "□"
    stage = "build"
    om.BuildUp
    TestOMathAdd = "ok|" & CStr(p.Range.OMaths.Count)
    Exit Function
EH:
    TestOMathAdd = "FAIL|" & stage & "|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：公式按钮三态（有文字→行内；已有公式→编号；空白行→编号占位）
Public Function TestFormulaThreeStates() As String
    Dim doc As Document
    Dim p As Paragraph
    Dim om As Object
    Dim inlineOk As Boolean
    Dim formulaOk As Boolean
    Dim blankOk As Boolean
    Dim stage As String

    On Error GoTo EH
    Set doc = ActiveDocument

    ' 1) 有文字：应插入行内占位，不加编号
    stage = "inline_setup"
    Set p = doc.Content.Paragraphs.Add
    p.Range.Text = "文字"
    p.Range.Select
    RunNormalizeFormula Nothing, True
    inlineOk = (p.Range.OMaths.Count = 1) And (InStr(p.Range.Text, "#") = 0)

    ' 2) 已有行间公式：应加 #() 编号
    Set p = doc.Content.Paragraphs.Add
    p.Range.Select
    Selection.Collapse wdCollapseStart
    doc.OMaths.Add Selection.Range
    Set om = p.Range.OMaths(1)
    om.Range.Text = "x=1"
    om.BuildUp
    p.Range.Select
    RunNormalizeFormula Nothing, True
    formulaOk = (p.Range.OMaths.Count = 1) And (InStr(p.Range.Text, "#") > 0)

    ' 3) 空白行：应插入带 #() 编号的占位公式
    Set p = doc.Content.Paragraphs.Add
    p.Range.Select
    RunNormalizeFormula Nothing, True
    blankOk = (p.Range.OMaths.Count = 1) And (InStr(p.Range.Text, "#") > 0)

    TestFormulaThreeStates = "ok|" & CStr(inlineOk) & "|" & CStr(formulaOk) & "|" & CStr(blankOk)
    Exit Function
EH:
    TestFormulaThreeStates = "FAIL|" & stage & "|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：行间公式菜单的“无编号 / 有编号 / 再转无编号”三条路径
Public Function TestDisplayFormulaModes() As String
    Dim doc As Document
    Dim p As Paragraph
    Dim plainDisplay As Boolean
    Dim plainFields As Long
    Dim plainBookmarks As Long
    Dim upgradedMath As Boolean
    Dim upgradedBookmarks As Long
    Dim downgradedDisplay As Boolean
    Dim downgradedFields As Long
    Dim downgradedBookmarks As Long

    On Error GoTo EH
    Set doc = ActiveDocument

    ' 1) 无编号：空白段。无编号公式仍是显示公式（有 oMathPara），但没有编号域/书签。
    Set p = doc.Content.Paragraphs.Add
    p.Range.Select
    InsertDisplayFormulaPlain Nothing
    Set p = Selection.Paragraphs(1)
    plainDisplay = (InStr(p.Range.WordOpenXML, "oMathPara") > 0)
    plainFields = p.Range.Fields.Count
    plainBookmarks = CountFormulaBookmarksIn(p)

    ' 2) 无编号 → 有编号
    p.Range.Select
    InsertDisplayFormula Nothing
    Set p = Selection.Paragraphs(1)
    upgradedMath = (InStr(p.Range.WordOpenXML, "oMathPara") > 0)
    upgradedBookmarks = CountFormulaBookmarksIn(p)

    ' 3) 有编号 → 无编号（保留显示公式，去掉编号域/书签）
    p.Range.Select
    InsertDisplayFormulaPlain Nothing
    Set p = Selection.Paragraphs(1)
    downgradedDisplay = (InStr(p.Range.WordOpenXML, "oMathPara") > 0)
    downgradedFields = p.Range.Fields.Count
    downgradedBookmarks = CountFormulaBookmarksIn(p)

    TestDisplayFormulaModes = "ok|" & CStr(plainDisplay) & "|" & CStr(plainFields) & "|" & _
        CStr(plainBookmarks) & "|" & CStr(upgradedMath) & "|" & CStr(upgradedBookmarks > 0) & "|" & _
        CStr(downgradedDisplay) & "|" & CStr(downgradedFields) & "|" & CStr(downgradedBookmarks)
    Exit Function
EH:
    TestDisplayFormulaModes = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' 测试入口：附录内公式的章号取附录字母，引用域回指同一编号
Public Function TestAppendixFormulaNumbering() As String
    Dim doc As Document
    Dim rng As Range
    Dim q As Paragraph
    Dim p As Paragraph
    Dim fld As Field
    Dim num As String
    Dim codes As String
    Dim bmName As String
    Dim bmText As String
    Dim refText As String
    Dim stage As String

    On Error GoTo EH
    Set doc = ActiveDocument

    ' 1) 找附录标题，没有就不测。
    stage = "find_appendix"
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Style = "附录标题"
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
    End With
    If Not rng.Find.Execute Then
        TestAppendixFormulaNumbering = "NO_APPENDIX"
        Exit Function
    End If
    Set q = rng.Paragraphs(1)

    ' 2) 在附录标题之后新起一段，再在该段插入带编号的行间公式。
    stage = "insert"
    Set rng = doc.Range(q.Range.End, q.Range.End)
    rng.Select
    Selection.TypeParagraph
    ' 回退到新建的空段：TypeParagraph 后光标停在后一段开头。
    Selection.Paragraphs(1).Previous.Range.Select
    RunNormalizeFormula Nothing, True
    Set p = Selection.Paragraphs(1)
    num = FormulaFieldNumText(p)
    If Len(num) = 0 Then
        TestAppendixFormulaNumbering = "NO_NUMBER"
        Exit Function
    End If

    ' 3) 章号域必须是 STYLEREF 附录标题，不能沿用 STYLEREF 1。
    stage = "code"
    For Each fld In p.Range.Fields
        codes = codes & UCase$(CleanText(fld.Code.Text)) & "|"
    Next fld

    ' 4) 编号应已建好书签。
    stage = "bookmark"
    bmName = FormulaBookmarkName(p)
    If doc.Bookmarks.Exists(bmName) Then
        bmText = CleanText(doc.Bookmarks(bmName).Range.Text)
    End If

    ' 5) 插一条公式引用，REF 域结果应与编号一致。
    stage = "reference"
    doc.Content.Paragraphs.Add.Range.Select
    InsertFormulaRefField bmName
    Set fld = Selection.Paragraphs(1).Range.Fields(1)
    fld.Update
    refText = CleanText(fld.Result.Text)

    TestAppendixFormulaNumbering = "ok|" & num & "|" & _
        CStr(InStr(codes, "附录标题") > 0) & "|" & bmText & "|" & refText
    Exit Function
EH:
    TestAppendixFormulaNumbering = "FAIL|" & stage & "|" & CStr(Err.Number) & "|" & Err.Description
End Function
