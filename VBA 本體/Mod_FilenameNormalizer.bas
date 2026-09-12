Option Explicit

' ==========================================================
' MODULE:       Mod_FilenameNormalizer (標準模組)
' PURPOSE:      應急工具。讓使用者手動指定任意資料夾，預覽並執行
'               Windows自動編號格式「檔名 (數字)」轉換為系統慣用
'               「檔名_數字」底線格式，並把 - 複製 刪除，不需要跑完整全面排版。
' EXPORTS:      RunNormalizeTool
' IMPORTS:      Mod_Utils（ConvertParenSuffixToUnderscore, BuildSafeWindowsPath）, cls_Config
' FORBIDDEN:    僅處理檔名字串格式，嚴禁更動任何檔案內容；
'               嚴禁覆寫已存在的同名目標檔案，撞名一律跳過並記錄。
' DEPENDENCIES: Windows FileSystemObject (FSO), Application.FileDialog
' VERSION:      1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_FilenameNormalizer"
Private Const MODULE_VERSION As String = "1.0.0"
' ------------------------------------------------------------------------------
' [入口]：RunNormalizeTool
' 流程：選擇資料夾 → Dry Run預覽（僅列出會被改名的清單，不動手）→
'      使用者確認 → 正式執行改名
' ------------------------------------------------------------------------------
Public Sub 修改流水號格式()
    Dim targetFolder As String
    targetFolder = PickFolder()
    If targetFolder = "" Then Exit Sub  ' 使用者取消

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim colFiles As New Collection
    Call GatherFilesRecursiveFlat(fso, targetFolder, colFiles)
    
    ' 【V2.3.0新增】這支工具目前不吃 cls_ExecutionContext，本身就是應急獨立工具，
    ' 需要用到 config.CopySuffixPatterns 時直接自己建一個
    Dim config As New cls_Config

    ' --- Dry Run：只列出會被改名的清單，不動手 ---
    Dim previewList As String, previewCount As Long
    Dim varFile As Variant, filePath As String, baseName As String, newBaseName As String

    For Each varFile In colFiles
        filePath = CStr(varFile)
        baseName = fso.GetBaseName(filePath)
        newBaseName = Mod_Utils.StripKnownCopySuffix(baseName, config.CopySuffixPatterns)
        newBaseName = Mod_Utils.ConvertParenSuffixToUnderscore(newBaseName)
        If newBaseName <> baseName Then
            previewCount = previewCount + 1
            If previewCount <= 20 Then previewList = previewList & baseName & "  →  " & newBaseName & vbCrLf
        End If
    Next varFile

    If previewCount = 0 Then
        MsgBox "掃描完成，沒有找到符合Windows自動編號格式的檔案，不需要任何動作。", vbInformation, "檔名正規化工具"
        Set fso = Nothing
        Exit Sub
    End If

    Dim previewMsg As String
    previewMsg = "掃描到 " & previewCount & " 個檔案符合Windows自動編號格式，預覽（最多顯示20筆）：" & vbCrLf & vbCrLf & previewList
    If previewCount > 20 Then previewMsg = previewMsg & "……以及其餘 " & (previewCount - 20) & " 筆" & vbCrLf
    previewMsg = previewMsg & vbCrLf & "確定要執行改名嗎？"

    If MsgBox(previewMsg, vbYesNo + vbQuestion, "檔名正規化工具 - Dry Run預覽") <> vbYes Then
        Set fso = Nothing
        Exit Sub
    End If

    ' --- 正式執行：單一檔案錯誤隔離，撞名絕不覆寫，跳過並計數 ---
    Dim renamedCount As Long, skippedCount As Long, errorCount As Long

    For Each varFile In colFiles
        On Error GoTo FileErrorHandler
        filePath = CStr(varFile)
        baseName = fso.GetBaseName(filePath)
        newBaseName = Mod_Utils.StripKnownCopySuffix(baseName, config.CopySuffixPatterns)
        newBaseName = Mod_Utils.ConvertParenSuffixToUnderscore(newBaseName)

        If newBaseName <> baseName Then
            Dim ext As String: ext = fso.GetExtensionName(filePath)
            Dim newPath As String
            newPath = Mod_Utils.BuildSafeWindowsPath(fso, fso.GetParentFolderName(filePath), newBaseName & "." & ext)

            If fso.FileExists(newPath) Then
                skippedCount = skippedCount + 1
            Else
                fso.MoveFile filePath, newPath
                renamedCount = renamedCount + 1
            End If
        End If

FileNextItem:
    Next varFile
    On Error GoTo 0

    MsgBox "執行完成。" & vbCrLf & "已改名：" & renamedCount & " 個" & vbCrLf & _
           "撞名跳過：" & skippedCount & " 個" & vbCrLf & "錯誤：" & errorCount & " 個", _
           vbInformation, "檔名正規化工具"

    Set fso = Nothing
    Exit Sub

FileErrorHandler:
    errorCount = errorCount + 1
    Err.Clear
    Resume FileNextItem
End Sub

Private Function PickFolder() As String
    PickFolder = ""
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "選擇要正規化檔名的資料夾"
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
            If Right(PickFolder, 1) <> "\" Then PickFolder = PickFolder & "\"
        End If
    End With
End Function

Private Sub GatherFilesRecursiveFlat(ByVal fso As Object, ByVal currentPath As String, ByRef sourceFiles As Collection)
    Dim folder As Object, subFolder As Object, file As Object
    On Error Resume Next
    Set folder = fso.GetFolder(currentPath)
    If folder Is Nothing Then Exit Sub
    For Each file In folder.Files: sourceFiles.Add file.Path: Next file
    For Each subFolder In folder.SubFolders: Call GatherFilesRecursiveFlat(fso, subFolder.Path, sourceFiles): Next subFolder
    On Error GoTo 0
End Sub
