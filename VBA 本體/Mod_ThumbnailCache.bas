Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_ThumbnailCache (標準模組)
' PURPOSE: 專職負責畫布嵌入用縮圖的快取建立、查詢與同步清理。
'          利用旗幟命名（內容雜湊衍生）本身的唯一性，快取有效性
'          不需額外比對修改時間，快取檔名等於旗幟名稱即視為有效。
' EXPORTS: GetOrCreateCachedPreview, CleanOrphanedCache
' IMPORTS: cls_Config, Mod_Utils, Mod_Actions
' FORBIDDEN: 嚴禁在此撰寫任何 UI 彈窗或業務邏輯判斷（如身分比對、
'            正名決策），僅負責「給我一個原始檔案，回我一份可以
'            直接拿去 AddPicture 的縮圖路徑」這一件事。
' DEPENDENCIES: Windows Scripting Host (FileSystemObject), WIA
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_ThumbnailCache"
Private Const MODULE_VERSION As String = "1.0.0"

Public Function GetOrCreateCachedPreview(ByVal fso As Object, ByVal config As cls_Config, ByVal sourceFilePath As String) As String
    Dim identifierName As String: identifierName = fso.GetFileName(sourceFilePath)
    Dim cachePath As String: cachePath = Mod_Utils.BuildSafeWindowsPath(fso, config.IOThumbnailCacheFolder, identifierName)
    
    If fso.FileExists(cachePath) Then
        GetOrCreateCachedPreview = cachePath
        Exit Function
    End If
    
    Call Mod_Utils.OptimizeAndCopyImage(sourceFilePath, cachePath, fso, config.ThumbnailCacheDimension)

     ' 【補漏，Grok審閱發現】OptimizeAndCopyImage極端情況下可能連內部陽春複製
     ' 都失敗，卻不會讓這裡知道。改為實際確認檔案存在才回傳快取路徑；
     ' 真的失敗則退回直接使用原始檔案，讓AddPicture至少有東西可嵌入
     ' （代價是這種極端情況下縮圖優化會失效，但比讓呼叫端拿到必定失敗的路徑安全）
     If fso.FileExists(cachePath) Then
         GetOrCreateCachedPreview = cachePath
     Else
         GetOrCreateCachedPreview = sourceFilePath
     End If

End Function

Public Sub CleanOrphanedCache(ByVal fso As Object, ByVal config As cls_Config)
    If Not fso.FolderExists(config.IOThumbnailCacheFolder) Then Exit Sub
    
     ' 先遞迴掃一次整理夾,建立「目前真正存在的檔名」清單,不限根目錄
    Dim colValidFiles As New Collection
    Call Mod_Actions.GatherFilesRecursive(fso, config.IOOrganizeFolder, colValidFiles, config)
    
     ' 【補漏，Grok審閱發現】GetOrCreateCachedPreview對救援檔案(isRescueFile=True)
     ' 同樣會建立快取，此前這裡只掃整理夾，救援檔案的快取名稱永遠不在有效清單裡，
     ' 每次全面排版都被誤判孤兒刪除、下次又要重新壓縮一次——非資料遺失，但是無效重工
     If fso.FolderExists(config.IORescueFolder) Then
         Call GatherRescueFilesForCache(fso, config.IORescueFolder, colValidFiles, config)
     End If
    
    Dim dictValidNames As Object
    Set dictValidNames = CreateObject("Scripting.Dictionary"): dictValidNames.CompareMode = 1
    Dim vFile As Variant
    For Each vFile In colValidFiles
        dictValidNames(fso.GetFileName(CStr(vFile))) = True
    Next vFile
    
    Dim cacheFile As Object
    On Error Resume Next
    For Each cacheFile In fso.GetFolder(config.IOThumbnailCacheFolder).Files
        If Not dictValidNames.Exists(cacheFile.Name) Then
            fso.DeleteFile cacheFile.Path, True
        End If
    Next cacheFile
    On Error GoTo 0
End Sub


' ------------------------------------------------------------------------------
' [輔助引擎]：GatherRescueFilesForCache
' 目的：專供CleanOrphanedCache掃描救援夾使用，排除已歸位／淘汰子夾——
'      這兩個子夾內容已是塵埃落定的封存區，不該被視為「有效」快取來源，
'      與cmd_Deduplicate、Mod_StateScanner既有的排除邏輯一致
' ------------------------------------------------------------------------------
Private Sub GatherRescueFilesForCache(ByVal fso As Object, ByVal currentPath As String, ByRef sourceFiles As Collection, ByVal config As cls_Config)
    Dim folder As Object, subFolder As Object, file As Object
    If UCase(currentPath) = UCase(config.IORescueDowncastFolder) Or UCase(currentPath) = UCase(config.IORescueDiscardFolder) Then Exit Sub
    On Error Resume Next
    Set folder = fso.GetFolder(currentPath)
    If folder Is Nothing Then Exit Sub
    For Each file In folder.Files: sourceFiles.Add file.Path: Next file
    For Each subFolder In folder.SubFolders: Call GatherRescueFilesForCache(fso, subFolder.Path, sourceFiles, config): Next subFolder
    On Error GoTo 0
End Sub
