Option Explicit
Option Private Module

' ==========================================================
' MODULE:       Mod_EmptyFolderCleaner (標準模組)
' PURPOSE:      應急工具。掃描整理夾，找出真正空的子資料夾（無檔案、
'               無子子資料夾）並清除，保留系統自身管理的功能性資料夾
'               （縮圖快取、暫存區等）不受影響。
' EXPORTS:      清除空資料夾
' IMPORTS:      Mod_Utils, cls_Config
' FORBIDDEN:    僅刪除確認完全空白的資料夾；系統保護資料夾一律排除，
'               不論是否為空
' VERSION:      1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_EmptyFolderCleaner"
Private Const MODULE_VERSION As String = "1.0.0"

' ------------------------------------------------------------------------------
' [入口]：清除空資料夾
' 範圍：僅限 config.IOOrganizeFolder（圖片整理夾），不觸碰其他資料夾。
' 由下往上遞迴：先確認最深層的子資料夾是否為空，刪除後可能讓上一層
' 也變空，因此需要由下往上處理，一次掃描即可處理多層巢狀空資料夾。
' ------------------------------------------------------------------------------
Public Sub 清除空資料夾()
    Dim config As New cls_Config
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(config.IOOrganizeFolder) Then
        MsgBox "圖片整理夾不存在，未執行任何操作", vbInformation, "空資料夾清理工具"
        Set fso = Nothing
        Exit Sub
    End If

    Dim removedCount As Long, previewList As String
    Call ScanAndRemoveEmptyFolders(fso, config, config.IOOrganizeFolder, removedCount, previewList, True)

    If removedCount = 0 Then
        MsgBox "掃描完成，沒有找到空資料夾。", vbInformation, "空資料夾清理工具"
        Set fso = Nothing
        Exit Sub
    End If

    Dim previewMsg As String
    previewMsg = "掃描到 " & removedCount & " 個空資料夾，預覽（最多顯示20筆）：" & vbCrLf & vbCrLf & previewList
    previewMsg = previewMsg & vbCrLf & "確定要清除嗎？"

    If MsgBox(previewMsg, vbYesNo + vbQuestion, "空資料夾清理工具 - Dry Run預覽") <> vbYes Then
        Set fso = Nothing
        Exit Sub
    End If

    Dim actualRemoved As Long, actualPreview As String
    Call ScanAndRemoveEmptyFolders(fso, config, config.IOOrganizeFolder, actualRemoved, actualPreview, False)

    MsgBox "執行完成，已清除 " & actualRemoved & " 個空資料夾。", vbInformation, "空資料夾清理工具"
    Set fso = Nothing
End Sub

' ------------------------------------------------------------------------------
' [輔助引擎]：ScanAndRemoveEmptyFolders
' dryRun=True時只計數、不動手；False時才真正刪除。
' 由下往上遞迴（先處理子資料夾，再檢查自己是否因此變空）。
' 系統保護資料夾（縮圖快取、暫存區等）一律排除，不論是否為空。
' ------------------------------------------------------------------------------
Public Sub ScanAndRemoveEmptyFolders(ByVal fso As Object, ByVal config As cls_Config, ByVal currentPath As String, _
                                        ByRef removedCount As Long, ByRef previewList As String, ByVal dryRun As Boolean)
    On Error Resume Next
    Dim folder As Object: Set folder = fso.GetFolder(currentPath)
    If folder Is Nothing Then Exit Sub
    On Error GoTo 0

   ' 【V2-091修正】路徑字串比對前，兩邊都統一補上結尾反斜線——
   ' FSO的Folder.Path屬性結尾沒有反斜線，但cls_Config的資料夾路徑
   ' 屬性結尾都有反斜線，格式不一致導致UCase比對永遠不相等，
   ' 保護清單不管列什麼內容都形同虛設，這才是「不管哪個資料夾都被清空」的真正原因
   Dim normalizedCurrentPath As String: normalizedCurrentPath = currentPath
   If Right(normalizedCurrentPath, 1) <> "\" Then normalizedCurrentPath = normalizedCurrentPath & "\"

    Dim isProtected As Boolean: isProtected = False
    Dim managedPath As Variant
    For Each managedPath In Mod_Actions.GetAllManagedFolderPaths(config)
       Dim normalizedManagedPath As String: normalizedManagedPath = CStr(managedPath)
       If Right(normalizedManagedPath, 1) <> "\" Then normalizedManagedPath = normalizedManagedPath & "\"
       If UCase(normalizedCurrentPath) = UCase(normalizedManagedPath) Then
            isProtected = True
            Exit For
        End If
    Next managedPath

    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        Call ScanAndRemoveEmptyFolders(fso, config, subFolder.Path, removedCount, previewList, dryRun)
    Next subFolder
    
    If isProtected Then Exit Sub

    ' 遞迴處理完子層之後，重新確認自己現在是不是真的空了
    On Error Resume Next
    Dim isEmpty As Boolean
    isEmpty = (folder.Files.count = 0 And folder.SubFolders.count = 0)
    On Error GoTo 0

    If isEmpty Then
        removedCount = removedCount + 1
        If removedCount <= 20 Then previewList = previewList & currentPath & vbCrLf
        If Not dryRun Then
            On Error Resume Next
            fso.DeleteFolder currentPath
            On Error GoTo 0
        End If
    End If
End Sub
