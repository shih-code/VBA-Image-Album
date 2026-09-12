Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_SystemAdmin (標準模組)
' PURPOSE: 系統管理與格式化引擎。實裝危險區絕對熔斷、事前精準預報、
'          扁平化備份快照，與「逐檔複製 -> 核對 -> 刪除」之三步物理防線。
' EXPORTS: 執行格式化, 執行系統緊急重啟, 開圖片母夾, 開匯出成果, 開歸檔附件, 開雜物箱
' IMPORTS: cls_ExecutionContext, cls_Config, cls_StringLibrary, Mod_Utils, Mod_UIMessenger, Mod_Actions, Mod_Rules, Mod_UIReset
' FORBIDDEN: 嚴禁使用 FSO.MoveFolder 進行備份。嚴禁在此模組私自越權修改全域 ScreenUpdating。
' DEPENDENCIES: Microsoft Scripting Runtime (FileSystemObject), Shell
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_SystemAdmin"
Private Const MODULE_VERSION As String = "1.0.0"

' ------------------------------------------------------------------------------
' 程式名稱：執行格式化
' 功用與目的：銷毀現有畫布，並根據二元軌道選項決定是否將實體檔案送入歷史扁平備份庫。
' ------------------------------------------------------------------------------
Public Sub 執行格式化(ByVal Context As cls_ExecutionContext)
    Dim fso As Object
    Dim config As cls_Config
    Dim strLib As cls_StringLibrary
    Dim formatMode As String
    Dim backupFolderName As String
    Dim hasLockedFiles As Boolean
    Dim totalBackupFiles As Long
    Dim wsUI As Worksheet
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set strLib = New cls_StringLibrary
    Set config = Context.config
    
    backupFolderName = ""
    hasLockedFiles = False
    totalBackupFiles = 0
    
    ' ==========================================================================
    ' 1. 危險區絕對熔斷防線 (最高物理防線)
    ' 為什麼要這樣設計：防止使用者在桌面、下載區或根目錄不小心按下格式化，導致系統核心或整個桌面被吞噬。
    ' ==========================================================================
    If Mod_Rules.IsDangerousPath(config.BaseFolder) Then
        Call Mod_UIMessenger.ShowError(strLib.ErrDangerZoneFormatBlocked, strLib.TitleFormatBlocked)
        Set fso = Nothing: Set strLib = Nothing: Set config = Nothing
        Exit Sub
    End If
    
    ' ==========================================================================
    ' 2. 讀取操作台二元軌道與事前精準預報
    ' ==========================================================================
    If Not DetermineFormatModeAndConfirm(config, fso, strLib, formatMode, totalBackupFiles) Then
        Set fso = Nothing: Set strLib = Nothing: Set config = Nothing
        Exit Sub
    End If
    
    Context.LogStore.Record "TX_START", "FORMAT", "使用者確認執行格式化 | 模式：" & formatMode & IIf(formatMode = "保留備份", " | 預計搬遷 " & totalBackupFiles & " 個檔案", "")
    
    ' 凍結核心事件以防干擾 (ScreenUpdating 保留，由 Mod_Main 或呼叫端控制)
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' ==========================================================================
    ' 3. 【呼叫清除模組】：安全銷毀 Excel 畫布
    ' ==========================================================================
    Call Mod_Actions.ResetWorkspaceToZero(config.TargetWB, config)
    Context.LogStore.Record "TX_INFO", "FORMAT", "畫布銷毀完成"
    
    ' ==========================================================================
    ' 4. 【三步物理防線與扁平化備份】：安全搬遷實體檔案
    ' ==========================================================================
    Call PerformBackupAndCleanup(Context, fso, config, strLib, formatMode, backupFolderName, hasLockedFiles)
    
    ' ==========================================================================
    ' 5. 從零開始重新建立操作介面與資料夾，確保環境純淨
    ' ==========================================================================
    Call Mod_Actions.InitializeSystemFolders(config)
    Call Mod_UIReset.重建操作介面(True, True)
    
    ' 同步看板大母夾名稱
    Set wsUI = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameUI)
    If Not wsUI Is Nothing Then
        Call Mod_Utils.SafeUnprotect(wsUI, config.SecurityPassword)
        wsUI.Range(config.CellProjName).Value = fso.GetFileName(config.BaseFolder)
        Call Mod_Utils.SafeProtect(wsUI, config.SecurityPassword)
    End If
    
    ' 解除底層事件凍結
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' ==========================================================================
    ' 6. 輸出完成訊息
    ' ==========================================================================
    Call ShowFormatCompletionReport(strLib, formatMode, backupFolderName, hasLockedFiles)
    Context.LogStore.Record "TX_COMMIT_FS", "FORMAT", "格式化作業全部完成"
    
    Application.StatusBar = False
    Set fso = Nothing: Set strLib = Nothing: Set config = Nothing: Set wsUI = Nothing
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：DetermineFormatModeAndConfirm
' 功用與目的：讀取操作台格式化軌道設定（僅重置工作表／保留備份），統計預計搬遷
'             檔案數，組出對應提示訊息並跳出嚴重警告確認彈窗。
' 回傳值：True 代表使用者按下「是」，呼叫端應繼續執行；False 代表使用者反悔。
' ------------------------------------------------------------------------------
Private Function DetermineFormatModeAndConfirm(ByVal config As cls_Config, ByVal fso As Object, ByVal strLib As cls_StringLibrary, ByRef formatMode As String, ByRef totalBackupFiles As Long) As Boolean
    Dim wsUI As Worksheet
    Dim ans As VbMsgBoxResult
    
    Set wsUI = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameUI)
    If Not wsUI Is Nothing Then
        formatMode = Trim(CStr(wsUI.Range(config.CellDangerConsole).Value))
    Else
        formatMode = config.FormatBehavior
    End If
    
    ' 【防禦收斂】：只允許兩種合法狀態，遇到舊版或被亂改的文字，強制自癒為預設最安全的保留備份
    If formatMode <> "僅重置工作表" And formatMode <> "保留備份" Then formatMode = "保留備份"
    
    ' 預先統計即將被送入備份區的實體檔案數量
    totalBackupFiles = 0
    If formatMode = "保留備份" Then
        Call CountFilesRecursive(fso, config.IOTargetFolder, totalBackupFiles)
        Call CountFilesRecursive(fso, config.IOAttachmentFolder, totalBackupFiles)
        Call CountFilesRecursive(fso, config.IOExportFolder, totalBackupFiles)
    End If
    
    Dim confirmMsg As String
    If formatMode = "僅重置工作表" Then
        confirmMsg = strLib.MsgFormatConfirmResetOnly
    Else
        confirmMsg = strLib.MsgFormatConfirmBackup(totalBackupFiles)
    End If
    
    ' 使用 AskCritical 彈出帶有嚴重警告圖示的彈窗
    ans = Mod_UIMessenger.AskCritical(confirmMsg, strLib.TitleFormatConfirm)
    
    DetermineFormatModeAndConfirm = (ans = vbYes)
End Function

' ------------------------------------------------------------------------------
' 程式名稱：PerformBackupAndCleanup
' 功用與目的：依格式化模式執行三個目標資料夾（整理夾/附件夾/匯出夾）的扁平化
'             備份搬遷，並於搬遷完成後由下而上拆除殘留空資料夾，記錄備份結果。
' ------------------------------------------------------------------------------
Private Sub PerformBackupAndCleanup(ByVal Context As cls_ExecutionContext, ByVal fso As Object, ByVal config As cls_Config, ByVal strLib As cls_StringLibrary, ByVal formatMode As String, ByRef backupFolderName As String, ByRef hasLockedFiles As Boolean)
    Dim targetDir As String, attachDir As String, exportDir As String
    targetDir = config.IOTargetFolder
    attachDir = config.IOAttachmentFolder
    exportDir = config.IOExportFolder
    
    If formatMode = "僅重置工作表" Then
        backupFolderName = "無 (檔案留存於原地)"
    ElseIf formatMode = "保留備份" Then
        Dim backupRoot As String, backupSubFolder As String
        backupRoot = config.IOBackupSnapshotFolder
        Call Mod_Utils.EnsureFolderExists(fso, backupRoot)
        
        Dim bFolderName As String: bFolderName = Format(Now, "yyyy-mm-dd_hh-mm-ss")
        backupSubFolder = fso.BuildPath(backupRoot, bFolderName)
        Call Mod_Utils.EnsureFolderExists(fso, backupSubFolder)
        
        Application.StatusBar = strLib.StatusArchivingBackup
        DoEvents
        
        ' 呼叫深度扁平化備份引擎 (含刪除原始檔案)
        If fso.FolderExists(targetDir) Then Call BackupFlattenedRecursive(fso, targetDir, backupSubFolder, hasLockedFiles)
        If fso.FolderExists(attachDir) Then Call BackupFlattenedRecursive(fso, attachDir, backupSubFolder, hasLockedFiles)
        If fso.FolderExists(exportDir) Then Call BackupFlattenedRecursive(fso, exportDir, backupSubFolder, hasLockedFiles)
        
        ' 備份完畢後，由於檔案已被物理刪除，原地的空資料夾由下而上徹底拆除自癒
        Application.StatusBar = strLib.StatusCleaningEmptyFolders
        DoEvents
        If fso.FolderExists(targetDir) Then Call Mod_Utils.CleanEmptyFoldersBottomUp(fso, targetDir, targetDir, config)
        If fso.FolderExists(attachDir) Then Call Mod_Utils.CleanEmptyFoldersBottomUp(fso, attachDir, attachDir, config)
        If fso.FolderExists(exportDir) Then Call Mod_Utils.CleanEmptyFoldersBottomUp(fso, exportDir, exportDir, config)
        
        backupFolderName = "封存備份\快照備份\" & bFolderName
    End If
    
    If formatMode = "保留備份" Then
       Context.LogStore.Record "TX_INFO", "FORMAT", "實體檔案備份完成 | 目的地：" & backupFolderName & " | 是否有檔案因被鎖定而跳過：" & IIf(hasLockedFiles, "是", "否")
       If hasLockedFiles Then
           Context.LogStore.Record "TX_ERROR", "FORMAT", "備份過程中偵測到鎖定檔案，部分檔案可能未被搬遷，原地保留"
       End If
   End If
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：ShowFormatCompletionReport
' 功用與目的：組合格式化完成報告文字（保留原地/搬遷位置/鎖定檔案警告），
'             並跳出完成提示彈窗。
' ------------------------------------------------------------------------------
Private Sub ShowFormatCompletionReport(ByVal strLib As cls_StringLibrary, ByVal formatMode As String, ByVal backupFolderName As String, ByVal hasLockedFiles As Boolean)
    Dim finalReportMessage As String
    finalReportMessage = strLib.InfoFormatBaseSuccess
    
    If formatMode = "僅重置工作表" Then
        finalReportMessage = finalReportMessage & vbCrLf & strLib.InfoFilesKeptInPlace
    Else
        finalReportMessage = finalReportMessage & vbCrLf & strLib.InfoFilesMovedTo(backupFolderName)
    End If
    
    If hasLockedFiles Then
        finalReportMessage = finalReportMessage & vbCrLf & vbCrLf & strLib.WarnLockedFilesSkipped
    End If
    
    Call Mod_UIMessenger.ShowInfo(finalReportMessage, strLib.TitleFormatComplete)
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：BackupFlattenedRecursive (扁平化三步備份引擎)
' 功用與目的：深入資料夾，依據副檔名將檔案壓平分類至「圖片、歸檔附錄、雜物箱」，
'             並執行「1. 複製 -> 2. 核對大小 -> 3. 刪除原檔」，降低防毒軟體鎖定造成的資料遺失風險。
' ------------------------------------------------------------------------------
Private Sub BackupFlattenedRecursive(ByVal fso As Object, ByVal srcFolder As String, ByVal backupSnapshotFolder As String, ByRef outHasLockedFiles As Boolean)
    Dim fld As Object, subFld As Object, fileObj As Object
    Dim destFolder As String, destPath As String, finalName As String
    Dim ext As String, cat As String
    
    If Not fso.FolderExists(srcFolder) Then Exit Sub
    Set fld = fso.GetFolder(srcFolder)
    
    For Each fileObj In fld.Files
        ext = LCase(fso.GetExtensionName(fileObj.Name))
        cat = Mod_Rules.GetFileCategory(ext)
        
        ' 依據副檔名分類，決定壓平後的平級子夾
        If cat = "IMAGE" Then
            destFolder = fso.BuildPath(backupSnapshotFolder, "圖片")
        ElseIf cat = "UNKNOWN" Then
            destFolder = fso.BuildPath(backupSnapshotFolder, "雜物箱")
        Else
            destFolder = fso.BuildPath(backupSnapshotFolder, "歸檔附錄")
        End If
        
        ' 確保扁平子夾存在
        Call Mod_Utils.EnsureFolderExists(fso, destFolder)
        
        ' 處理撞名：因為壓平了所有層級，同名檔案機率提升，調用 Unique 命名以策安全
        finalName = Mod_Utils.GenerateUniqueFileName(fso, destFolder, fileObj.Name)
        destPath = Mod_Utils.BuildSafeWindowsPath(fso, destFolder, finalName)
        
        On Error Resume Next
        ' 第一步：嘗試複製檔案
        fso.CopyFile fileObj.Path, destPath, True
        
        If Err.Number = 0 Then
            ' 第二步：雙重物理核對，確保不是空殼或複製不完全
            If fso.FileExists(destPath) Then
                If fso.GetFile(destPath).Size = fileObj.Size Then
                    ' 第三步：核對通過，刪除原始檔案
                    Call Mod_Utils.DeletePhysicalFile(fso, fileObj.Path)
                Else
                    outHasLockedFiles = True ' 大小不符，標記異常
                End If
            Else
                outHasLockedFiles = True ' 實體未成功落地
            End If
        Else
            outHasLockedFiles = True ' 複製過程即被鎖定
        End If
        Err.Clear
        On Error GoTo 0
    Next fileObj
    
    ' 繼續往下鑽深層資料夾，但目標位置永遠是扁平的 backupSnapshotFolder
    For Each subFld In fld.SubFolders
        Call BackupFlattenedRecursive(fso, subFld.Path, backupSnapshotFolder, outHasLockedFiles)
    Next subFld
    
    Set fld = Nothing: Set subFld = Nothing: Set fileObj = Nothing
End Sub

' ------------------------------------------------------------------------------
' 內部輔助函數：遞迴估算檔案數量
' ------------------------------------------------------------------------------
Private Sub CountFilesRecursive(ByVal fso As Object, ByVal targetPath As String, ByRef totalCount As Long)
    On Error Resume Next
    If fso.FolderExists(targetPath) Then
        Dim fld As Object
        Set fld = fso.GetFolder(targetPath)
        totalCount = totalCount + fld.Files.count
        Set fld = Nothing
        
        Dim subFld As Object
        For Each subFld In fso.GetFolder(targetPath).SubFolders
            Call CountFilesRecursive(fso, subFld.Path, totalCount)
        Next subFld
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：執行系統緊急重啟 (合併 UI 還原)
' 功能描述：修復 Excel 核心環境變數，強制重繪 UI 介面並貫徹三大防禦鎖。
' ------------------------------------------------------------------------------
Public Sub 執行系統緊急重啟()
    ' 不碰 ScreenUpdating，避免在某些極端報錯中卡死黑畫面
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    Dim config As cls_Config
    Dim strLib As New cls_StringLibrary
    Set config = New cls_Config
    
    Call Mod_UIReset.重建操作介面(False, True)
    Call Mod_UIMessenger.ShowInfo(strLib.InfoSysRestarted, strLib.TitleRestartSuccess)
    
    Set config = Nothing: Set strLib = Nothing
End Sub

' ------------------------------------------------------------------------------
' 實體資料夾快速導航器 (已對齊平級三柱路徑常數)
' ------------------------------------------------------------------------------
Public Sub 開圖片母夾()
    Dim config As cls_Config: Set config = New cls_Config
   
    If Mod_Rules.IsDangerousPath(config.BaseFolder) Then
        Dim strLib As New cls_StringLibrary
        Call Mod_UIMessenger.ShowError(strLib.PromptDangerousZone(""), strLib.TitleDefend)
        Set config = Nothing: Exit Sub
    End If
    
    Call Mod_Actions.InitializeSystemFolders(config)
    On Error Resume Next: Shell "explorer.exe """ & config.IOTargetFolder & """", vbNormalFocus: On Error GoTo 0
    Set config = Nothing
End Sub

Public Sub 開匯出成果()
    Dim config As cls_Config: Set config = New cls_Config
    If Mod_Rules.IsDangerousPath(config.BaseFolder) Then
        Dim strLib As New cls_StringLibrary
        Call Mod_UIMessenger.ShowError(strLib.PromptDangerousZone(""), strLib.TitleDefend)
        Set config = Nothing: Exit Sub
    End If
    Call Mod_Actions.InitializeSystemFolders(config)
    On Error Resume Next: Shell "explorer.exe """ & config.IOExportFolder & """", vbNormalFocus: On Error GoTo 0
    Set config = Nothing
End Sub

Public Sub 開歸檔附件()
    Dim config As cls_Config: Set config = New cls_Config
    If Mod_Rules.IsDangerousPath(config.BaseFolder) Then
        Dim strLib As New cls_StringLibrary
        Call Mod_UIMessenger.ShowError(strLib.PromptDangerousZone(""), strLib.TitleDefend)
        Set config = Nothing: Exit Sub
    End If
    Call Mod_Actions.InitializeSystemFolders(config)
    On Error Resume Next: Shell "explorer.exe """ & config.IOAttachmentFolder & """", vbNormalFocus: On Error GoTo 0
    Set config = Nothing
End Sub

Public Sub 開雜物箱()
    Dim config As cls_Config: Set config = New cls_Config
    If Mod_Rules.IsDangerousPath(config.BaseFolder) Then
        Dim strLib As New cls_StringLibrary
        Call Mod_UIMessenger.ShowError(strLib.PromptDangerousZone(""), strLib.TitleDefend)
        Set config = Nothing: Exit Sub
    End If
    Call Mod_Actions.InitializeSystemFolders(config)
    On Error Resume Next: Shell "explorer.exe """ & config.IOQuarantineFolder & """", vbNormalFocus: On Error GoTo 0
    Set config = Nothing
End Sub
