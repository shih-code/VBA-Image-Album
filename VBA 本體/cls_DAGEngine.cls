Option Explicit

' ==========================================================
' MODULE: cls_DAGEngine (類別模組)
' PURPOSE: 流水線任務調度引擎。負責工人註冊、依賴鏈有向無環圖 (DAG) 解析、死鎖安全攔截。
' EXPORTS: Initialize, AddCommand, ExecuteAll
' IMPORTS: ICommand, cls_Config, cls_Log
' FORBIDDEN: 嚴禁在此模組直接呼叫任何 MsgBox 或實體硬碟/畫布操作，必須透過拋出 Err 讓 Pipeline 捕捉，或交接給具體的 Command 執行。
' DEPENDENCIES: 引用 Microsoft Scripting Runtime (Dictionary)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "cls_DAGEngine"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
' 【底層私有變數】
' ==============================================================================
Private m_Commands As Object
Private m_CompletedStates As Object
Private m_Config As cls_Config
Private m_LogStore As cls_Log ' 【型別升級】：對齊實體類別，獲取編譯器保護

' ------------------------------------------------------------------------------
' 程式名稱：Class_Initialize
' 功用與目的：初始化引擎內部的任務與狀態容器。
' 為什麼要這樣設計：使用強型態字典維持高速字串雜湊比對，設定 CompareMode 為文字比對，防止任務名稱因大小寫差異而撞名。
' ------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_Commands = CreateObject("Scripting.Dictionary")
    Set m_CompletedStates = CreateObject("Scripting.Dictionary")
    m_Commands.CompareMode = 1
    m_CompletedStates.CompareMode = 1
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：Class_Terminate
' 功用與目的：當引擎生命週期結束時，強制回收所有記憶體資源。
' 為什麼要這樣設計：防止記憶體洩漏 (Memory Leak)，確保 Excel 運行長期穩定。
' ------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set m_Commands = Nothing
    Set m_CompletedStates = Nothing
    Set m_Config = Nothing
    Set m_LogStore = Nothing
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：Initialize
' 功用與目的：將全域靜態設定與日誌庫注入到引擎中。
' 為什麼要這樣設計：嚴格遵守「依賴注入 (DI)」，拒絕引擎內部私自 New 出資源，確保全系統共用同一個 SSOT。
' ------------------------------------------------------------------------------
Public Sub Initialize(ByVal config As cls_Config, ByVal LogStore As cls_Log)
    Set m_Config = config
    Set m_LogStore = LogStore
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：AddCommand
' 功用與目的：將具體的管線任務（工人）註冊進引擎的候選名單中。
' 參數說明：cmd (ICommand) - 實作了 ICommand 合約的命令物件。
' ------------------------------------------------------------------------------
Public Sub AddCommand(ByVal cmd As ICommand)
    ' 為什麼要這樣判斷：防禦空物件傳入導致崩潰。
    If cmd Is Nothing Then Exit Sub
    
    ' 為什麼要這樣判斷：捍衛冪等性 (Idempotency)。防止同一個任務被重複註冊兩次。
    If Not m_Commands.Exists(cmd.CommandName) Then
        m_Commands.Add cmd.CommandName, cmd
    End If
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：ExecuteAll
' 功用與目的：啟動核心調度迴圈，依據任務的依賴關係 (Dependencies) 依序派發執行。
' 參數說明：Context (cls_ExecutionContext) - 當次執行的動態狀態物件。
' ------------------------------------------------------------------------------
Public Sub ExecuteAll(ByVal Context As cls_ExecutionContext)
    ' --------------------------------------------------------------------------
    ' [第一步：啟動前防禦校驗]
    ' --------------------------------------------------------------------------
    ' 為什麼要這樣判斷：落實 Fail-Fast (快速失效)。如果外部忘記呼叫 Initialize 就啟動引擎，立刻報警並中斷。
    If m_Config Is Nothing Or m_LogStore Is Nothing Then
        Err.Raise 9991, MODULE_NAME & ".ExecuteAll", _
            "[系統報警] 模組：" & MODULE_NAME & " | 函數：ExecuteAll | 錯誤碼：9991 | 描述：核心引擎尚未完成 Initialize 變數綁定"
        Exit Sub
    End If

    ' 清空歷史完工狀態字典，重置引擎發動機，確保可以被重複呼叫
    m_CompletedStates.RemoveAll

    Dim cmd As ICommand
    Dim key As Variant
    Dim loopCounter As Long
    Dim executedInThisLoop As Boolean
    Dim alarmString As String
    
    loopCounter = 0
    
    ' --------------------------------------------------------------------------
    ' [第二步：DAG 拓撲解析核心調度迴圈]
    ' 功用與目的：只要還有任務未完工，就持續在迴圈中尋找「符合執行條件」的任務並派發。
    ' --------------------------------------------------------------------------
    Do While m_CompletedStates.count < m_Commands.count
        executedInThisLoop = False
        loopCounter = loopCounter + 1
        
        ' 為什麼要這樣判斷 (死鎖防禦線 A)：嚴防執行緒咬死。向靜態組態索取迭代上限值，超過立即熔斷保護。
        If loopCounter > m_Config.MaxDagLoopCountDAGEngine Then
            alarmString = "[系統報警] 模組：" & MODULE_NAME & " | 函數：ExecuteAll | 錯誤碼：9992 | 描述：管線遭遇未知循環相依，超出最大迭代限制熔斷 | 當前關鍵參數：LoopCounter=" & loopCounter
            Call m_LogStore.Record("ERROR", "ExecuteAll", alarmString)
            Err.Raise 9992, MODULE_NAME & ".ExecuteAll", alarmString
            GoTo ExecuteAllExit
        End If
        
        ' 有向無環圖 (DAG) 順序解碼掃描
        For Each key In m_Commands.Keys
            Set cmd = m_Commands(key)
            
            ' 為什麼要這樣判斷：過濾掉已經執行完畢的節點，專注於尚未開工的任務。
            If Not m_CompletedStates.Exists(cmd.CommandName) Then
                
                ' 為什麼要這樣判斷：審查該任務的前置條件是否全部達成。
                If CanExecute(cmd) Then
                    executedInThisLoop = True
                    
                    On Error GoTo CommandErrorHandler
                    ' 移交動態上下文，驅動具體的命令物件執行業務
                    cmd.Execute Context
                    On Error GoTo 0
                    
CommandLoopResume:
                    ' 為什麼要這樣判斷 (完工標記防呆)：若子命令內部發生降級跳過且未主動回報，由引擎強制將其寫入成功清單，維持管線推進。
                    If Not m_CompletedStates.Exists(cmd.CommandName) Then
                        m_CompletedStates.Add cmd.CommandName, True
                    End If
                End If
            End If
        Next key
        
        ' 為什麼要這樣判斷 (死鎖防禦線 B)：整輪迴圈空轉且未全員完工，必定是遇到了「循環相依 (A等B，B等A)」或「缺失依賴 (依賴了不存在的任務)」。
        If Not executedInThisLoop And m_CompletedStates.count < m_Commands.count Then
            Dim pendingCount As Long
            pendingCount = m_Commands.count - m_CompletedStates.count
            alarmString = "[系統報警] 模組：" & MODULE_NAME & " | 函數：ExecuteAll | 錯誤碼：9993 | 描述：管線依賴鏈發生循環死鎖或前置任務遺失，安全攔截 | 當前關鍵參數：PendingCount=" & pendingCount
            
            Call m_LogStore.Record("ERROR", "ExecuteAll", alarmString)
            
            ' 為什麼要這樣設計：強制有序填補失效標記，徹底解凍 Excel 進程避免無窮迴圈死機。
            For Each key In m_Commands.Keys
                If Not m_CompletedStates.Exists(CStr(key)) Then
                    m_CompletedStates.Add CStr(key), False
                End If
            Next key
            
            ' 拋出結構化異常，由上層管線 (Mod_Pipeline) 捕捉，徹底拔除引擎內部越權呼叫 UI 的惡習。
            Err.Raise 9993, MODULE_NAME & ".ExecuteAll", alarmString
            Exit Do
        End If
    Loop
    
ExecuteAllExit:
    Set cmd = Nothing
    Exit Sub

' ------------------------------------------------------------------------------
' 3. 管線熔斷與異常降級容錯中樞
' ------------------------------------------------------------------------------
CommandErrorHandler:
    Dim savedErrNum As Long: savedErrNum = Err.Number
    Dim savedErrDesc As String: savedErrDesc = Err.Description
    Err.Clear
    
    ' 為什麼要這樣判斷：依照合約審視該任務是否為「關鍵核心節點 (IsCritical)」。
    If cmd.IsCritical Then
        ' 分流 A：關鍵核心組件崩潰 -> 立即啟動環境解凍閉環，中斷全盤管線
        On Error Resume Next
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        On Error GoTo 0
        
        alarmString = "[系統報警] 模組：" & MODULE_NAME & " | 函數：ExecuteAll | 錯誤碼：" & savedErrNum & " | 描述：關鍵流水線組件 [" & cmd.CommandName & "] 崩潰中斷，核心熔斷 | 當前關鍵參數：" & savedErrDesc
        Call m_LogStore.Record("ERROR", cmd.CommandName, alarmString)
        
        ' 向上層管線 (Mod_Pipeline) 丟出強型態結構化報警單
        Err.Raise savedErrNum, MODULE_NAME & ".ExecuteAll", alarmString
        GoTo ExecuteAllExit
    Else
        ' 分流 B：非關鍵組件報錯 -> 修正為標記 False。
        ' 為什麼要這樣設計：確保依賴此故障點的後續子任務不會被矇騙，而是會在 CanExecute 階段被安全跳過，避免連鎖崩潰。
        If Not m_CompletedStates.Exists(cmd.CommandName) Then
            m_CompletedStates.Add cmd.CommandName, False
        End If
        
        ' 發布降級容錯運行日誌，允許管線其餘無關分支繼續推進
        Dim warningMsg As String
        warningMsg = "非關鍵錯誤已略過，管線執行降級容錯推進(來源節點: " & cmd.CommandName & "): " & savedErrDesc
        Call m_LogStore.Record("WARNING", cmd.CommandName, warningMsg)
        
        Resume CommandLoopResume
    End If
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：CanExecute
' 功用與目的：判定指定的任務是否具備執行資格（所有前置依賴是否皆已成功完工）。
' ------------------------------------------------------------------------------
Private Function CanExecute(ByVal cmd As ICommand) As Boolean
    Dim deps As Variant
    Dim i As Long
    
    CanExecute = False
    deps = cmd.Dependencies
    
    If isEmpty(deps) Then
        CanExecute = True
        Exit Function
    End If
    
    On Error Resume Next
    Dim testBound As Long
    testBound = LBound(deps)
    If Err.Number <> 0 Then
        Err.Clear
        CanExecute = True
        Exit Function
    End If
    On Error GoTo 0
    
    For i = LBound(deps) To UBound(deps)
        ' 【V4.2.7 修復：略過缺失依賴 (Ghost Dependency)】
        ' 只有當依賴的任務「存在於這次的排程清單 (m_Commands)」時，才需要等它。
        ' 解決了單獨執行 REBUILD (跳過 IMPORT) 時導致的 PendingCount=2 死鎖。
        If m_Commands.Exists(deps(i)) Then
            If Not m_CompletedStates.Exists(deps(i)) Then Exit Function
            If m_CompletedStates(deps(i)) = False Then Exit Function
        End If
    Next i
    
    CanExecute = True
End Function
