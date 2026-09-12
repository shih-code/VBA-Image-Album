Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_Pipeline (標準模組)
' PURPOSE: 全域管線生命週期控管模組。
'          統一控管 UX 流程，集中實施環境封鎖、開機引導、
'          最高危險區全域熔斷、事前精準預報 (Pre-Flight Check) 與全流程防禦與收尾機制。
' EXPORTS: BootSystem, Run
' IMPORTS: cls_Config, cls_ExecutionContext, cls_StringLibrary, Mod_Actions, Mod_Rules, Mod_Utils, Mod_UIMessenger, Mod_UIToggles, cls_DAGEngine, Mod_CommandFactory, Mod_StateScanner, Mod_ImportCollisionGate, Mod_UIReset
' FORBIDDEN: 1. 嚴禁在此撰寫實體檔案搬移或幾何排版邏輯 (交給底層工人)。
'            2. 嚴禁在此調用 FileDialog (UI 挑選檔案應切還給 Mod_Main)。
' DEPENDENCIES: Microsoft Scripting Runtime (FileSystemObject)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_Pipeline"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
' 區塊 1：開機引導中樞 (System Boot Sequence)
' 目的：接管 ThisWorkbook_Open 的呼叫，執行四象限真值表判定與開機自癒。
' ==============================================================================
Public Sub BootSystem(ByVal Context As cls_ExecutionContext)
    Dim config As cls_Config: Set config = Context.config
    Dim strLib As cls_StringLibrary: Set strLib = New cls_StringLibrary
    Dim actualName As String, memoryName As String
    Dim quad As Integer


   On Error GoTo BootErrorHandler

    ' 【畫面凍結】：凍結畫面，避免掃描與切換分頁時產生閃爍
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' [探測]：呼叫 Actions 取得四象限狀態碼
    quad = Mod_Actions.GetBootQuadrant(config, actualName, memoryName)

    ' [決策]：統一進行 UI 對話與調度
    Select Case quad
        Case -1 ' 【危險區熔斷】
            Call LockSystemUI(config)
            Application.ScreenUpdating = True
            Call Mod_UIMessenger.ShowError(strLib.PromptDangerousZone(""), strLib.TitleDefend)
            
        Case 1 ' 【象限 1：叢集健康，名稱吻合】 (不彈窗開機)
            If memoryName = "" Then Call Mod_Actions.UpdateMemoryPoint(ThisWorkbook, actualName, config)
            Call Mod_Actions.ExecuteStandardBoot(config, actualName)
            
        Case 2 ' 【象限 2：叢集健康，名稱被外部更改】
            Call Mod_Actions.UpdateMemoryPoint(ThisWorkbook, actualName, config)
            Call Mod_Actions.ExecuteStandardBoot(config, actualName)
            Application.ScreenUpdating = True
            Call Mod_UIMessenger.ShowInfo(strLib.InfoProjectNameSynced(actualName), strLib.TitleSysPrompt)
            
        Case 3 ' 【象限 3：叢集殘缺，名稱吻合 (判定為環境誤刪)】
            Application.ScreenUpdating = True ' 必須開燈才能彈窗對話
            If Mod_UIMessenger.AskQuestion(strLib.PromptEnvAutoHeal, "環境修復") = vbYes Then
                Application.ScreenUpdating = False
                Call Mod_Actions.InitializeSystemFolders(config)
                Call Mod_Actions.ExecuteProactiveRescueAndSnapshot(config) ' 觸發緊急快照與救援
                Call Mod_Actions.ExecuteStandardBoot(config, actualName)
            Else
                Call Mod_UIMessenger.ShowWarning("系統環境未修復，部分功能將無法運作。", "警告")
            End If
            
        Case 4 ' 【象限 4：叢集殘缺，名稱不同 (判定為全新專案)】
            Application.ScreenUpdating = True
            If Mod_UIMessenger.AskQuestion("偵測到全新專案目錄！" & vbCrLf & "是否在此建構基礎資料夾環境並初始化專案？", "新案初始化") = vbYes Then
                Application.ScreenUpdating = False
                Call Mod_Actions.InitializeSystemFolders(config)
                Call Mod_Actions.UpdateMemoryPoint(ThisWorkbook, actualName, config)
                Call Mod_Actions.ExecuteStandardBoot(config, actualName)
            End If
    End Select

    ' 【開機日誌封存檢查】：僅在非危險區時執行，避免環境尚未確認安全時進行任何檔案 I/O
    If quad <> -1 Then
        If Not Context.LogStore Is Nothing Then Context.LogStore.ArchiveAndTrim
       ' 【補漏】CleanSystemTempFiles此前已寫好卻從未被任何地方呼叫過，
       ' 是死程式碼——%TEMP%底下的app_preview_*.jpg殘骸從未被自動清理過
       Dim fsoCleanup As Object: Set fsoCleanup = CreateObject("Scripting.FileSystemObject")
       Call Mod_Utils.CleanSystemTempFiles(fsoCleanup)
       Set fsoCleanup = Nothing
        
    End If

BootCleanup:
   ' 開機完成，解除畫面凍結（正常結束與例外結束都會走到這裡）
   Application.EnableEvents = True
   Application.ScreenUpdating = True
   Set strLib = Nothing
   Exit Sub

BootErrorHandler:
   ' 開機中途崩潰時，至少要讓使用者看得到訊息、看得到畫面，而不是卡死的空白 Excel
   Dim bootErrDesc As String: bootErrDesc = Err.Description
   Err.Clear
   Resume BootCleanup
End Sub

' ------------------------------------------------------------------------------
' 內部輔助：UI 鎖定 (專供危險區熔斷使用)
' ------------------------------------------------------------------------------
Private Sub LockSystemUI(ByVal config As cls_Config)
    Dim wsUI As Worksheet: Set wsUI = Mod_Utils.GetSheetSafe(ThisWorkbook, config.SheetNameUI)
    If Not wsUI Is Nothing Then
        Call Mod_Utils.SafeUnprotect(wsUI, config.SecurityPassword)
        wsUI.Range(config.CellProjName).Value = "【系統鎖定：請移至安全資料夾】"
        Call Mod_Utils.SafeProtect(wsUI, config.SecurityPassword)
    End If
End Sub

' ==============================================================================
' 區塊 2：全域管線生命週期控管模組 (Main Pipeline Runner)
' 目的：控制 IMPORT / REBUILD / EXPORT 的完整執行流與防禦網。
' ==============================================================================
Public Sub Run(ByVal opMode As String, ByVal cmdNames As Variant, ByVal Context As cls_ExecutionContext)
    Dim config As cls_Config, LogStore As Object
    Dim wsUI As Worksheet, strLib As cls_StringLibrary, fso As Object
    Dim startTime As Double, elapsed As Double
    Dim pipelineAborted As Boolean: pipelineAborted = False

    ' 全域錯誤攔截網：任何崩潰皆導流至 Cleanup 進行復原處理
    On Error GoTo PipelineErrorHandler
    startTime = Timer

    Set strLib = New cls_StringLibrary
    Set config = Context.config
    Set LogStore = Context.LogStore
    Context.OperationMode = opMode

    ' 確保前端焦點強制聚焦在「操作介面」
    Set wsUI = Mod_Utils.GetSheetSafe(ThisWorkbook, config.SheetNameUI)
    If Not wsUI Is Nothing Then wsUI.Activate
    
    ' 【畫面凍結】：開局立刻凍結畫面，阻止一切解鎖與掃描造成的閃爍
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ==========================================================================
    ' 【第一至四關卡：危險區熔斷／環境健檢／硬碟空間查核／Pre-Flight事前預報】
    ' ==========================================================================
    If Not RunPreExecutionGates(opMode, Context, config, strLib, fso) Then
        pipelineAborted = True
        GoTo PipelineCleanup
    End If

    ' ==========================================================================
    ' 5. 解除保護鎖 (供掃描器讀取)
    ' ==========================================================================
    Call UnlockWorkbookForScan(config)

    ' ==========================================================================
    ' 5.5 【名稱衝突偵測】：僅在匯入模式，蒐集即將匯入的檔案是否撞到等待救援中的紀錄，
    '     真正的詢問延後到 cmd_Rebuild 執行時統一處理
    ' ==========================================================================
    Call DetectImportCollisionsIfNeeded(opMode, fso, config, Context)

    ' ==========================================================================
    ' 6. 【MVC 核心】：呼叫狀態掃描器 (Scanner)
    ' ==========================================================================
    Call Mod_StateScanner.ScanCurrentState(Context)

    ' ==========================================================================
    ' 【遺失圖片處理決策】
    ' ==========================================================================
    Call ResolveMissingShapesDecision(config, Context)

    ' ==========================================================================
    ' 7. 全域應用程式環境封鎖 (對話結束，進入批次處理模式)
    ' ==========================================================================
    Call FreezeForBatchExecution

    ' ==========================================================================
    ' 8. DAG 引擎組裝與執行
    ' ==========================================================================
    Call AssembleAndExecuteDAG(config, LogStore, cmdNames, Context)

    ' ==========================================================================
    ' 9. 後製視圖修復：色彩標籤重整與活頁簿物理拓撲重排
    ' ==========================================================================
    Call FinalizeViewState(config)

    ' ==========================================================================
    ' 9.5 【局部重排】：跑到這裡代表本次 REBUILD 沒有在任何關卡中途被攔下或放棄，
    '     只有這個前提成立，才真正執行鎖定，避免排版被中止但鎖定狀態卻先改變
    ' ==========================================================================
    If opMode = "REBUILD" And Context.RequestLayoutLockAfterRebuild Then
        Call Mod_UIToggles.ApplyLayoutLock(config)
    End If


' ==============================================================================
' 收尾與安全閉環
' ==============================================================================
PipelineCleanup:
    On Error Resume Next
    elapsed = Round(Timer - startTime, 2)
    
    ' 1. 【視角回正】：在解除畫面凍結與彈窗前，先將視角切回操作介面
    If Not Context Is Nothing Then
        Call Mod_UIToggles.EnforceStandardSecurityState(Context)
    End If
    
    ' 2. 【解除畫面凍結】：讓畫面平順地亮起，此時背景已經是乾淨的操作台
    Application.ScreenUpdating = True
    DoEvents ' 強迫作業系統立刻刷新螢幕 (消滅白畫面)
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' 3. 【結果回報】：背景已鎖定，結果彈窗會顯示在操作介面之上
    Call ReportPipelineResult(opMode, Context, strLib, elapsed, pipelineAborted)
    
    ' 4. 【原子級釋放】：最後一步才清空所有記憶體指針，根除洩漏隱患
    Set fso = Nothing: Set config = Nothing
    Set wsUI = Nothing: Set strLib = Nothing: Set LogStore = Nothing
    On Error GoTo 0
    Exit Sub

' ─── 異常處理 ───
PipelineErrorHandler:
    ' 萬一發生半路當機，優先確保畫面不卡死，隨後拋出異常
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If strLib Is Nothing Then Set strLib = New cls_StringLibrary
    Call Mod_UIMessenger.ShowError(strLib.ErrPipeline(Err.Description), strLib.TitleSysWarning)
    pipelineAborted = True
    Resume PipelineCleanup
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：RunPreExecutionGates
' 功用與目的：依序執行四道事前關卡（危險區熔斷／環境健檢／硬碟空間查核／
'             Pre-Flight事前預報），任一關卡未通過即中止，不繼續往下執行。
' 回傳值：True 代表四道關卡全數通過；False 代表其中一關未通過，呼叫端應中止。
' ------------------------------------------------------------------------------
Private Function RunPreExecutionGates(ByVal opMode As String, ByVal Context As cls_ExecutionContext, ByVal config As cls_Config, ByVal strLib As cls_StringLibrary, ByRef fso As Object) As Boolean
    RunPreExecutionGates = False

    ' ==========================================================================
    ' 【第一關卡：最高危險區全域熔斷】
    ' ==========================================================================
    If Not CheckDangerousZoneGate(config, strLib) Then Exit Function

    ' ==========================================================================
    ' 【第二關卡：實體環境無條件查核與修復】
    ' ==========================================================================
    If Not CheckEnvironmentHealthGate(config, strLib) Then Exit Function

    ' ==========================================================================
    ' 【第三關卡：硬碟空間 Fail-Fast 查核】
    ' ==========================================================================
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not CheckDiskSpaceGate(fso, config, strLib) Then Exit Function

    ' ==========================================================================
    ' 【第四關卡：事前精準預報 (Pre-Flight Check)】
    ' ==========================================================================
    If Not Context.isSilent Then
        Application.ScreenUpdating = True ' 預報需要讓使用者看到
        If Not ExecutePreFlightCheck(opMode, Context, fso, strLib) Then Exit Function
        Application.ScreenUpdating = False
    End If

    RunPreExecutionGates = True
End Function

' ------------------------------------------------------------------------------
' 程式名稱：ReportPipelineResult
' 功用與目的：依作業模式（IMPORT/REBUILD）與執行結果，組合並彈出對應的
'             結果報告訊息；靜默模式或作業被中止時比照原邏輯不彈窗。
' ------------------------------------------------------------------------------
Private Sub ReportPipelineResult(ByVal opMode As String, ByVal Context As cls_ExecutionContext, ByVal strLib As cls_StringLibrary, ByVal elapsed As Double, ByVal pipelineAborted As Boolean)
    If Not Context Is Nothing And Not strLib Is Nothing Then
        If Not Context.isSilent Then
            If opMode = "IMPORT" Then
                If Context.SessionImported > 0 Or Context.SessionFailed > 0 Or Context.SessionRescued > 0 Then
                    Beep
                    Call Mod_UIMessenger.ShowInfo(strLib.ReportImportComplete(Context.SessionImported, Context.SessionRescued, Context.SessionFailed, Context.SessionFailedReasons, elapsed), strLib.TitleReport)
                End If
            ElseIf opMode = "REBUILD" And Not pipelineAborted Then
                Beep
                Call Mod_UIMessenger.ShowInfo(strLib.InfoRebuildCompleteWithRescue(Context.SessionRescued, Context.SessionFailed, elapsed), strLib.TitleSysPrompt)
            End If
        End If
    End If
End Sub

' --------------------------------------------------------------------------
' 【第一關卡】最高危險區全域熔斷。True=可繼續，False=已熔斷並顯示錯誤
' --------------------------------------------------------------------------
Private Function CheckDangerousZoneGate(ByVal config As cls_Config, ByVal strLib As cls_StringLibrary) As Boolean
    CheckDangerousZoneGate = True
    If Mod_Rules.IsDangerousPath(config.BaseFolder) Then
        Call LockSystemUI(config)
        Application.ScreenUpdating = True
        Call Mod_UIMessenger.ShowError(strLib.PromptDangerousZone(""), strLib.TitleDefend)
        CheckDangerousZoneGate = False
    End If
End Function

' --------------------------------------------------------------------------
' 【第二關卡】實體環境無條件查核與修復。True=可繼續，False=使用者拒絕或修復失敗
' --------------------------------------------------------------------------
Private Function CheckEnvironmentHealthGate(ByVal config As cls_Config, ByVal strLib As cls_StringLibrary) As Boolean
    CheckEnvironmentHealthGate = True
    If Not Mod_Actions.IsEnvironmentHealthy(config) Then
        Application.ScreenUpdating = True ' 必須開燈才能詢問
        If Mod_UIMessenger.AskQuestion(strLib.PromptEnvAutoHeal, "環境補全與修復") = vbYes Then
            Application.ScreenUpdating = False
            Call Mod_Actions.InitializeSystemFolders(config)
            Dim t As Single: t = Timer
            Do While Abs(Timer - t) < 0.5: DoEvents: Loop ' 等待 MFT 刷新
            If Not Mod_Actions.IsEnvironmentHealthy(config) Then
                Application.ScreenUpdating = True
                Call Mod_UIMessenger.ShowError("建立失敗，請檢查磁碟權限或防毒軟體設定。", "錯誤")
                CheckEnvironmentHealthGate = False
            End If
        Else
            CheckEnvironmentHealthGate = False
        End If
    End If
End Function

' --------------------------------------------------------------------------
' 【第三關卡】硬碟空間 Fail-Fast 查核。True=可繼續，False=空間不足
' --------------------------------------------------------------------------
Private Function CheckDiskSpaceGate(ByVal fso As Object, ByVal config As cls_Config, ByVal strLib As cls_StringLibrary) As Boolean
    CheckDiskSpaceGate = True
    If Not Mod_Rules.IsDiskSpaceAvailable(fso, config.BaseFolder, 1) Then
        Application.ScreenUpdating = True
        Call Mod_UIMessenger.ShowError(strLib.ErrSpaceLimit, strLib.TitleSysPrompt)
        CheckDiskSpaceGate = False
    End If
End Function

' --------------------------------------------------------------------------
' 解除保護鎖，供掃描器讀取。永遠執行，不參與流程中止判斷
' --------------------------------------------------------------------------
Private Sub UnlockWorkbookForScan(ByVal config As cls_Config)
    Dim tmpWs As Worksheet
    On Error Resume Next
    config.TargetWB.Unprotect Password:=config.SecurityPassword
    For Each tmpWs In config.TargetWB.Worksheets
        Call Mod_Utils.SafeUnprotect(tmpWs, config.SecurityPassword)
    Next tmpWs
    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------
' 名稱衝突偵測：僅在匯入模式蒐集衝突，真正詢問延後到 cmd_Rebuild 統一處理
' --------------------------------------------------------------------------
Private Sub DetectImportCollisionsIfNeeded(ByVal opMode As String, ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext)
    If opMode = "IMPORT" Then
        Dim wsCatPreCheck As Worksheet
        Set wsCatPreCheck = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameCatalog)
        If Not wsCatPreCheck Is Nothing Then
            Call Mod_ImportCollisionGate.GatherConflicts(fso, config, Context, wsCatPreCheck, opMode)
        End If
    End If
End Sub

' --------------------------------------------------------------------------
' 遺失圖片處理決策：統計遺失類型，詢問使用者是否救援，設定 Context.UserWantsRescue
' --------------------------------------------------------------------------
Private Sub ResolveMissingShapesDecision(ByVal config As cls_Config, ByVal Context As cls_ExecutionContext)
    If Context.MissingShapes.count > 0 Then
        Dim missingStandard As Long: missingStandard = 0
        Dim missingRescue As Long: missingRescue = 0
        Dim key As Variant
        
        ' 統計遺失類型
        For Each key In Context.MissingShapes.Keys
            If InStr(1, CStr(key), config.SheetNameRescue) > 0 Then
                missingRescue = missingRescue + 1
            Else
                missingStandard = missingStandard + 1
            End If
        Next key
        
        Dim msgRescue As String, titleRescue As String
        If missingStandard > 0 Then
            titleRescue = "發現未歸檔圖片"
            msgRescue = "【系統自動偵測：發現未歸檔圖片】" & vbCrLf & vbCrLf & _
                        "畫布上有 " & missingStandard & " 張圖片，但在整理夾中找不到實體檔案。" & vbCrLf & _
                        "是否將圖片匯出至『救援夾』保留，並加入待還原清單？" & vbCrLf & _
                        "（若選『否』，將在排版時直接清除該筆記錄）"
        Else
            titleRescue = "救援備份遺失警告"
            msgRescue = "【嚴重警報：救援備份實體遭外部物理抹除！】" & vbCrLf & vbCrLf & _
                        "偵測到有 " & missingRescue & " 張已在遺失清單上的圖片，其硬碟備份檔再次消失。" & vbCrLf & _
                        "是否再次嘗試救援？將從畫布重新擷取圖片、補回救援夾，並保留原始比對記錄。" & vbCrLf & _
                        "（若選『否』，將移除此筆記錄並清除畫布上的殘影）"
        End If
                    
        Application.ScreenUpdating = True ' 開燈詢問
        If Mod_UIMessenger.AskQuestion(msgRescue, titleRescue) = vbYes Then
            Context.UserWantsRescue = True
        Else
            Context.UserWantsRescue = False
        End If
    End If
End Sub

' --------------------------------------------------------------------------
' 全域應用程式環境封鎖：對話結束，進入批次處理模式
' --------------------------------------------------------------------------
Private Sub FreezeForBatchExecution()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

' --------------------------------------------------------------------------
' DAG 引擎組裝與執行：建立引擎、加入指令、驅動全面開工
' --------------------------------------------------------------------------
Private Sub AssembleAndExecuteDAG(ByVal config As cls_Config, ByVal LogStore As Object, ByVal cmdNames As Variant, ByVal Context As cls_ExecutionContext)
    Dim engine As Object, cmd As Variant
    Set engine = New cls_DAGEngine
    engine.Initialize config, LogStore

    For Each cmd In cmdNames
        engine.AddCommand Mod_CommandFactory.Create(CStr(cmd))
    Next cmd

    ' 驅動引擎，全面開工 (工人僅依賴 Context 行事)
    engine.ExecuteAll Context
End Sub

' --------------------------------------------------------------------------
' 後製視圖修復：色彩標籤重整與活頁簿物理拓撲重排
' --------------------------------------------------------------------------
Private Sub FinalizeViewState(ByVal config As cls_Config)
    Call Mod_Actions.UpdateMemoryPoint(ThisWorkbook, config.UIProjectName, config)
    Call Mod_UIReset.SortWorksheetsStrictly(config)
End Sub

' ------------------------------------------------------------------------------
' 內部輔助函數：事前精準預報 (Pre-Flight Check)
' ------------------------------------------------------------------------------
Private Function ExecutePreFlightCheck(ByVal opMode As String, ByVal Context As cls_ExecutionContext, ByVal fso As Object, ByVal strLib As cls_StringLibrary) As Boolean
    Dim msg As String, totalFiles As Long, totalSizeBytes As Double, itm As Variant
    ExecutePreFlightCheck = True
    
    If opMode = "IMPORT" Then
        Application.StatusBar = "正在估算預計處理的檔案數量與容量..."
        totalFiles = 0: totalSizeBytes = 0
        For Each itm In Context.SelectedFiles
            Call CalculateFileStats(fso, CStr(itm), totalFiles, totalSizeBytes)
        Next itm
        Application.StatusBar = False
        
        msg = strLib.PromptPreFlightImport(totalFiles, (totalSizeBytes / 1048576))
        If Mod_UIMessenger.AskQuestion(msg, "匯入作業預報") = vbNo Then ExecutePreFlightCheck = False
        
    ElseIf opMode = "REBUILD" Then
         msg = "即將重組全域目錄與排版。系統將重新掛載畫布，並執行環境清理與自癒檢查，是否執行？"
         If Context.RequestLayoutLockAfterRebuild Then
             msg = msg & vbCrLf & vbCrLf & "【局部重排】完成後將鎖定寬度/間距/排版模式三格。"
         End If
        If Mod_UIMessenger.AskQuestion(msg, "排版與清理預報") = vbNo Then ExecutePreFlightCheck = False
    End If
End Function

' ------------------------------------------------------------------------------
' 內部輔助函數：計算檔案大小與數量
' ------------------------------------------------------------------------------
Private Sub CalculateFileStats(ByVal fso As Object, ByVal itemPath As String, ByRef tCount As Long, ByRef tSize As Double)
    If fso.FolderExists(itemPath) Then
        Dim fld As Object, fileObj As Object, subFld As Object
        Set fld = fso.GetFolder(itemPath)
        For Each fileObj In fld.Files
            tCount = tCount + 1: tSize = tSize + fileObj.Size
        Next fileObj
        For Each subFld In fld.SubFolders
            Call CalculateFileStats(fso, subFld.Path, tCount, tSize)
        Next subFld
    ElseIf fso.FileExists(itemPath) Then
        tCount = tCount + 1: tSize = tSize + fso.GetFile(itemPath).Size
    End If
End Sub
