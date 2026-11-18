Attribute VB_Name = "modExport"
Option Explicit

Private Declare PtrSafe Function MoveFileExW Lib "kernel32" ( _
    ByVal existingName As LongPtr, ByVal newName As LongPtr, ByVal flags As Long) As Long
Private Const MOVEFILE_REPLACE_EXISTING As Long = 1

Public LastExportError As String

Public Function GetLastExportError() As String
    GetLastExportError = LastExportError
End Function

Public Sub ExportDocx(control As IRibbonControl)
    ExportDocxFile False
End Sub

Public Sub ExportPdf(control As IRibbonControl)
    ExportPdfFile False
End Sub

' silent=True 用于自动化，调用方应先保存写作稿。
Public Function ExportDocxFile(ByVal silent As Boolean) As String
    ExportDocxFile = ExportThesisFile(False, silent)
End Function

Public Function ExportPdfFile(ByVal silent As Boolean) As String
    ExportPdfFile = ExportThesisFile(True, silent)
End Function

Private Function ExportThesisFile(ByVal asPdf As Boolean, ByVal silent As Boolean) As String
    Dim doc As Document
    Dim tmpDoc As Document
    Dim fso As Object
    Dim exportPath As String
    Dim tmpPath As String
    Dim stagedDocx As String
    Dim title As String
    Dim oldAlerts As WdAlertLevel
    Dim oldUpdating As Boolean
    Dim restoreState As Boolean
    Dim ownsCopy As Boolean
    Dim errNo As Long
    Dim errText As String

    LastExportError = ""
    On Error GoTo EH
    Set doc = ActiveDocument
    If asPdf Then
        title = "导出 PDF"
    Else
        title = "导出 DOCX"
    End If

    If Len(doc.Path) = 0 Then
        LastExportError = "请先保存论文，再导出。"
        If Not silent Then MsgBox LastExportError, vbExclamation, title
        Exit Function
    End If

    If asPdf Then
        exportPath = StripExtension(doc.FullName) & ".pdf"
    Else
        exportPath = StripExtension(doc.FullName) & "_提交版.docx"
    End If
    If FileExists(exportPath) And Not silent Then
        If MsgBox("文件已存在，是否覆盖？" & vbCrLf & exportPath, _
                  vbQuestion + vbYesNo, title) <> vbYes Then Exit Function
    End If
    If Not doc.Saved Then
        If silent Then
            LastExportError = "请先保存写作稿的修改，再导出。"
            Exit Function
        End If
        If MsgBox("写作稿有未保存的修改。是否先保存，再导出？", _
                  vbQuestion + vbYesNo, title) <> vbYes Then Exit Function
        doc.Save
        If Not doc.Saved Then Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    tmpPath = fso.BuildPath(doc.Path, fso.GetTempName() & "." & fso.GetExtensionName(doc.Name))
    fso.CopyFile doc.FullName, tmpPath, False
    ownsCopy = True

    oldAlerts = Application.DisplayAlerts
    oldUpdating = Application.ScreenUpdating
    restoreState = True
    Application.DisplayAlerts = wdAlertsNone
    Application.ScreenUpdating = False
    Set tmpDoc = Documents.Open(FileName:=tmpPath, ReadOnly:=False, AddToRecentFiles:=False)
    tmpDoc.Activate
    UpdateAllFieldsNow
    FreezeCitationFields

    If asPdf Then
        tmpDoc.ExportAsFixedFormat OutputFileName:=exportPath, ExportFormat:=wdExportFormatPDF, _
                                   OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
                                   Range:=wdExportAllDocument, Item:=wdExportDocumentContent, _
                                   IncludeDocProps:=True, CreateBookmarks:=wdExportCreateHeadingBookmarks, _
                                   DocStructureTags:=True
    Else
        If tmpDoc.Type = wdTypeDocument Then
            ' 写作稿通过关联模板加载宏。提交版保留已更新的格式和编号。
            tmpDoc.UpdateStylesOnOpen = False
            tmpDoc.AttachedTemplate = NormalTemplate.FullName
        End If
        ' 关闭 Word 持有的文件后再移除内嵌功能区，清理成功才替换提交版。
        stagedDocx = fso.BuildPath(doc.Path, fso.GetTempName() & ".docx")
        tmpDoc.SaveAs2 FileName:=stagedDocx, FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
        tmpDoc.Close SaveChanges:=wdDoNotSaveChanges
        Set tmpDoc = Nothing
        RemoveEmbeddedRibbon stagedDocx
        If MoveFileExW(StrPtr(stagedDocx), StrPtr(exportPath), MOVEFILE_REPLACE_EXISTING) = 0 Then
            Err.Raise vbObjectError + 1701, "ExportThesisFile", "无法写入提交版，请检查文件是否被占用。"
        End If
    End If
    ExportThesisFile = exportPath

CleanUp:
    On Error Resume Next
    If Not tmpDoc Is Nothing Then tmpDoc.Close SaveChanges:=wdDoNotSaveChanges
    If Len(stagedDocx) > 0 Then
        If FileExists(stagedDocx) Then Kill stagedDocx
    End If
    If ownsCopy Then Kill tmpPath
    If Not doc Is Nothing Then doc.Activate
    If restoreState Then
        Application.ScreenUpdating = oldUpdating
        Application.DisplayAlerts = oldAlerts
    End If
    On Error GoTo 0
    If errNo <> 0 Then
        If Not silent Then ReportError title, errNo, errText
    ElseIf Len(ExportThesisFile) > 0 And Not silent Then
        MsgBox "已导出：" & vbCrLf & exportPath, vbInformation, title
    End If
    Exit Function
EH:
    errNo = Err.Number
    errText = Err.Description
    LastExportError = CStr(errNo) & "|" & errText
    Resume CleanUp
End Function

Private Sub RemoveEmbeddedRibbon(ByVal path As String)
    ' 使用 Windows 自带的 ZIP 接口。仅修改已经关闭的导出暂存文件。
    Dim script As String
    Dim payload() As Byte
    Dim encoded As Object
    Dim shell As Object
    Dim status As Long

    script = "$ErrorActionPreference = 'Stop'" & vbCrLf _
        & "$path = '" & Replace(path, "'", "''") & "'" & vbCrLf _
        & "Add-Type -AssemblyName System.IO.Compression.FileSystem" & vbCrLf _
        & "$zip = [IO.Compression.ZipFile]::Open($path, 'Update')" & vbCrLf _
        & "function Read-XmlPart($name) {" & vbCrLf _
        & "  $stream = $zip.GetEntry($name).Open()" & vbCrLf _
        & "  try { $xml = New-Object Xml.XmlDocument; $xml.Load($stream); return ,$xml }" & vbCrLf _
        & "  finally { $stream.Dispose() }" & vbCrLf _
        & "}" & vbCrLf _
        & "function Write-XmlPart($name, $xml) {" & vbCrLf _
        & "  $zip.GetEntry($name).Delete()" & vbCrLf _
        & "  $stream = $zip.CreateEntry($name).Open()" & vbCrLf _
        & "  try { $xml.Save($stream) } finally { $stream.Dispose() }" & vbCrLf _
        & "}" & vbCrLf _
        & "try {" & vbCrLf _
        & "  $rels = Read-XmlPart '_rels/.rels'" & vbCrLf _
        & "  $links = @($rels.DocumentElement.ChildNodes | Where-Object { $_.Type -like '*/ui/extensibility' })" & vbCrLf _
        & "  $parts = @($links | ForEach-Object { $_.Target.TrimStart('/') })" & vbCrLf _
        & "  $entries = @($zip.Entries | Where-Object { $_.FullName -like 'customUI/*' -or $_.FullName -in $parts })" & vbCrLf _
        & "  if ($links.Count -or $entries.Count) {" & vbCrLf _
        & "    $parts += @($entries | ForEach-Object { $_.FullName })" & vbCrLf _
        & "    foreach ($link in $links) { [void]$rels.DocumentElement.RemoveChild($link) }" & vbCrLf _
        & "    Write-XmlPart '_rels/.rels' $rels" & vbCrLf
    script = script _
        & "    $types = Read-XmlPart '[Content_Types].xml'" & vbCrLf _
        & "    foreach ($entry in @($types.DocumentElement.ChildNodes)) {" & vbCrLf _
        & "      if ($entry.PartName -and $entry.PartName.TrimStart('/') -in $parts) {" & vbCrLf _
        & "        [void]$types.DocumentElement.RemoveChild($entry)" & vbCrLf _
        & "      }" & vbCrLf _
        & "    }" & vbCrLf _
        & "    Write-XmlPart '[Content_Types].xml' $types" & vbCrLf _
        & "    foreach ($entry in $entries) { $entry.Delete() }" & vbCrLf _
        & "  }" & vbCrLf _
        & "} finally { $zip.Dispose() }"

    ' PowerShell 的 EncodedCommand 使用 UTF-16LE，VBA 的字符串内存格式与之相同。
    payload = script
    Set encoded = CreateObject("MSXML2.DOMDocument.6.0").createElement("base64")
    encoded.DataType = "bin.base64"
    encoded.nodeTypedValue = payload
    Set shell = CreateObject("WScript.Shell")
    status = shell.Run("""" & Environ$("SystemRoot") _
        & "\System32\WindowsPowerShell\v1.0\powershell.exe"" -NoProfile -NonInteractive -EncodedCommand " _
        & Replace(Replace(encoded.Text, vbCr, ""), vbLf, ""), 0, True)
    If status <> 0 Then
        Err.Raise vbObjectError + 1702, "RemoveEmbeddedRibbon", "无法清理提交版内嵌的功能区。"
    End If
End Sub

Public Function StripExtension(ByVal path As String) As String
    Dim pos As Long
    pos = InStrRev(path, ".")
    If pos > 0 Then
        StripExtension = Left$(path, pos - 1)
    Else
        StripExtension = path
    End If
End Function

Public Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(path)) > 0)
    On Error GoTo 0
End Function
