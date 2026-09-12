Option Explicit

' ==========================================================
' MODULE: cls_ExecutionContext (類別模組)
' PURPOSE: 強型態資料傳輸物件 (DTO)。負責在主控端與命令物件間傳遞執行期狀態、布林旗標、
'          執行統計，以及全域生命週期中的 HASH 狀態記錄與遺失追蹤。
' EXPORTS: Initialize, LoadUIState 及各種唯讀狀態屬性 (Getters) 與統計/狀態屬性 (Let/Set)
' IMPORTS: cls_Config
' FORBIDDEN: 嚴禁將此模組的屬性當作全域變數使用，必須透過 ICommand.Execute 參數傳遞。
' DEPENDENCIES: Microsoft Scripting Runtime (Dictionary)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "cls_ExecutionContext"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
' 1. 內部變數宣告區
' ==============================================================================
Private m_SelectedFiles As Collection
Private m_OperationMode As String
Private m_Config As cls_Config
Private m_LogStore As Object

' 統計與追蹤
Private m_SessionImported As Long
Private m_SessionFailed As Long
Private m_SessionRescued As Long
Private m_SessionFailedReasons As String

' 運行期狀態旗標 (由 UI 介面動態解碼驅動)
Private m_IsDebugMode As Boolean
Private m_IsAutoLayout As Boolean
Private m_IsSingleMode As Boolean
Private m_IsMergeMode As Boolean
Private m_IsFolderMergeMode As Boolean
Private m_IsCatalogAppend As Boolean
Private m_IsFormatErase As Boolean

' 主被動解耦旗標
Private m_IsSilent As Boolean

' 動態操作參數承載區
Private m_UIProjName As String
Private m_UIImgWidth As Double
Private m_UIImgSpacing As Double
Private m_UINamingMode As String
Private m_UISheetPrefix As String
Private m_UIPdfImgName As String
Private m_UIPdfCatName As String
Private m_UIPdfGrid As Integer
Private m_UIExportFilter As String
Private m_UIExportRange As String


' 全域狀態與雜湊記錄 (State Machine & Roster)
Private m_DictGlobalHash As Object      ' 實體檔案全域指紋庫 [HASH -> 實體路徑/名稱] （正常圖片表，對應架構圖 HASH:A）
Private m_DictRescueHash As Object      ' 救援圖片指紋庫 [HASH -> 實體路徑/名稱]（救援圖片表，對應架構圖 HASH:B）
 Private m_DictPathToHash As Object      ' 【V2-047新增】圖片雜湊正向索引 [實體路徑 -> HASH]，
                                                                          ' 供 cmd_Deduplicate 重用 ScanCurrentState 已算好的雜湊，避免同一次管線內重複計算
Private m_MissingShapes As Object       ' 畫布遺失名單 [Shape名稱 -> 工作表名稱]
Private m_RestoreCandidates As Object          ' 歷史待救援清單 [原HASH -> 暫時名稱/狀態]
Private m_UserWantsRescue As Boolean    ' 記錄管線彈窗後，使用者的最終決策
Private m_RequestLayoutLockAfterRebuild As Boolean    ' 局部重排：本次 REBUILD 成功跑完後是否要順便上鎖
Private m_NameCollisionDecisions As Object  '一次性撞名詢問機制

Private m_PendingConflicts As Object   ' 累積尚未詢問的衝突：[原始檔案路徑] -> [目錄識別名稱]


' ==============================================================================
' 2. 類別建構與解構機制
' ==============================================================================
Private Sub Class_Initialize()
    Set m_SelectedFiles = New Collection
    m_OperationMode = ""
    
    m_SessionImported = 0
    m_SessionFailed = 0
    m_SessionRescued = 0
    m_SessionFailedReasons = ""
    
    m_IsDebugMode = False
    m_IsAutoLayout = True
    m_IsSingleMode = False
    m_IsMergeMode = False
    m_IsFolderMergeMode = False
    m_IsCatalogAppend = False
    m_IsFormatErase = False
    m_IsSilent = False
    
    ' 初始化全域狀態字典，並統一設定為不區分大小寫的文字比對模式 (CompareMode = 1)
    Set m_DictGlobalHash = CreateObject("Scripting.Dictionary")
    m_DictGlobalHash.CompareMode = 1
    Set m_DictRescueHash = CreateObject("Scripting.Dictionary")
    m_DictRescueHash.CompareMode = 1
    
     Set m_DictPathToHash = CreateObject("Scripting.Dictionary")
     m_DictPathToHash.CompareMode = 1
    
    Set m_MissingShapes = CreateObject("Scripting.Dictionary")
    m_MissingShapes.CompareMode = 1
    
    Set m_RestoreCandidates = CreateObject("Scripting.Dictionary")
    m_RestoreCandidates.CompareMode = 1
    
    Set m_NameCollisionDecisions = CreateObject("Scripting.Dictionary")   '一次性撞名機制
    m_NameCollisionDecisions.CompareMode = 1
    
    Set m_PendingConflicts = CreateObject("Scripting.Dictionary")
    m_PendingConflicts.CompareMode = 1
    
    m_UserWantsRescue = False
    m_RequestLayoutLockAfterRebuild = False
End Sub

Private Sub Class_Terminate()
    Set m_SelectedFiles = Nothing
    Set m_Config = Nothing
    Set m_LogStore = Nothing
    Set m_DictGlobalHash = Nothing
    Set m_DictRescueHash = Nothing
    Set m_DictPathToHash = Nothing
    Set m_MissingShapes = Nothing
    Set m_RestoreCandidates = Nothing
End Sub

' ==============================================================================
' 3. 核心依賴綁定與 UI 狀態解碼引擎
' ==============================================================================
Public Sub Initialize(ByVal injectedConfig As cls_Config, ByVal injectedLogStore As Object)
    If injectedConfig Is Nothing Then
        Err.Raise 9971, MODULE_NAME & ".Initialize", _
            "[系統報警] 模組：" & MODULE_NAME & " | 錯誤碼：9971 | 描述：傳入的 Config 靜態組態為空，上下文拒絕啟動"
    End If
    
    Set m_Config = injectedConfig
    Set m_LogStore = injectedLogStore
End Sub

Public Sub LoadUIState(ByVal wsUI As Worksheet)
    If wsUI Is Nothing Then
        Err.Raise 9972, MODULE_NAME & ".LoadUIState", _
            "[系統報警] 模組：" & MODULE_NAME & " | 錯誤碼：9972 | 描述：傳入的 UI 工作表實體為空，無法解碼狀態旗標"
    End If
    
    If m_Config Is Nothing Then
        Err.Raise 9973, MODULE_NAME & ".LoadUIState", _
            "[系統報警] 模組：" & MODULE_NAME & " | 錯誤碼：9973 | 描述：尚未執行 Initialize 綁定 Config，拒絕盲目解碼"
    End If
    
    On Error GoTo DecodeError
    
    ' --- 布林旗標解碼區 ---
    m_IsAutoLayout = (Trim(CStr(wsUI.Range(m_Config.CellAutoLayoutStateConsole).Value)) = "是")
    m_IsDebugMode = (Trim(CStr(wsUI.Range(m_Config.CellDebugStateFlagSet).Value)) = "是")
    
    Dim valDisplay As String
    valDisplay = Trim(CStr(wsUI.Range(m_Config.CellDisplayModeFlagSet).Value))
    m_IsSingleMode = (valDisplay = "單一工作表")
    m_IsMergeMode = (valDisplay = "同名合併工作表")
    m_IsFolderMergeMode = (valDisplay = "依資料夾合併")
    
    m_IsCatalogAppend = (Trim(CStr(wsUI.Range(m_Config.CellCatalogModeFlagSet).Value)) = "新增")
    
    Dim valFormat As String
    valFormat = Trim(CStr(wsUI.Range(m_Config.CellDangerConsole).Value))
    m_IsFormatErase = (valFormat = "僅重置工作表")
    
    ' --- 動態操作參數捕獲區 ---
    m_UIProjName = Trim(CStr(wsUI.Range(m_Config.CellProjName).Value))
    m_UIImgWidth = Val(wsUI.Range(m_Config.CellImgWidth).Value)
    m_UIImgSpacing = Val(wsUI.Range(m_Config.CellImgSpacing).Value)
    m_UINamingMode = Trim(CStr(wsUI.Range(m_Config.CellNamingMode).Value))
    m_UISheetPrefix = Trim(CStr(wsUI.Range(m_Config.CellSheetPrefix).Value))
    m_UIPdfImgName = Trim(CStr(wsUI.Range(m_Config.CellPdfImgName).Value))
    m_UIPdfCatName = Trim(CStr(wsUI.Range(m_Config.CellPdfCatName).Value))
    m_UIPdfGrid = Val(wsUI.Range(m_Config.CellPdfGrid).Value)
    m_UIExportFilter = Trim(CStr(wsUI.Range(m_Config.CellExportFilter).Value))
    m_UIExportRange = Trim(CStr(wsUI.Range(m_Config.CellExportRange).Value))
    
    On Error GoTo 0
    Exit Sub

DecodeError:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Err.Raise 9974, MODULE_NAME & ".LoadUIState", _
        "[系統報警] 模組：" & MODULE_NAME & " | 錯誤碼：9974 | 描述：UI 狀態解碼失敗，操作台介面可能遭到嚴重破壞 (" & errDesc & ")"
End Sub

' ==============================================================================
' 4. 基礎屬性與全域旗標封裝區
' ==============================================================================
Public Property Get SelectedFiles() As Collection: Set SelectedFiles = m_SelectedFiles: End Property
Public Property Set SelectedFiles(ByVal vNewValue As Collection): Set m_SelectedFiles = vNewValue: End Property

Public Property Get OperationMode() As String: OperationMode = m_OperationMode: End Property
Public Property Let OperationMode(ByVal vNewValue As String): m_OperationMode = vNewValue: End Property

Public Property Get config() As cls_Config: Set config = m_Config: End Property
Public Property Get LogStore() As Object: Set LogStore = m_LogStore: End Property

Public Property Get isSilent() As Boolean: isSilent = m_IsSilent: End Property
Public Property Let isSilent(ByVal vNewValue As Boolean): m_IsSilent = vNewValue: End Property

' ==============================================================================
' 5. 執行統計屬性區
' ==============================================================================
Public Property Get SessionImported() As Long: SessionImported = m_SessionImported: End Property
Public Property Let SessionImported(ByVal vNewValue As Long): m_SessionImported = vNewValue: End Property

Public Property Get SessionFailed() As Long: SessionFailed = m_SessionFailed: End Property
Public Property Let SessionFailed(ByVal vNewValue As Long): m_SessionFailed = vNewValue: End Property

Public Property Get SessionRescued() As Long: SessionRescued = m_SessionRescued: End Property
Public Property Let SessionRescued(ByVal vNewValue As Long): m_SessionRescued = vNewValue: End Property

Public Property Get SessionFailedReasons() As String: SessionFailedReasons = m_SessionFailedReasons: End Property
Public Property Let SessionFailedReasons(ByVal vNewValue As String): m_SessionFailedReasons = vNewValue: End Property


'一次性撞名詢問機制
Public Property Get NameCollisionDecisions() As Object: Set NameCollisionDecisions = m_NameCollisionDecisions: End Property
Public Property Set NameCollisionDecisions(ByVal v As Object): Set m_NameCollisionDecisions = v: End Property
Public Property Get PendingConflicts() As Object: Set PendingConflicts = m_PendingConflicts: End Property

' ==============================================================================
' 6. 狀態旗標與動態參數屬性區 (外界僅能讀取狀態，絕對不可逆向竄改)
' ==============================================================================
Public Property Get IsDebugMode() As Boolean: IsDebugMode = m_IsDebugMode: End Property
Public Property Get IsAutoLayout() As Boolean: IsAutoLayout = m_IsAutoLayout: End Property
Public Property Get IsSingleMode() As Boolean: IsSingleMode = m_IsSingleMode: End Property
Public Property Get IsMergeMode() As Boolean: IsMergeMode = m_IsMergeMode: End Property
Public Property Get IsFolderMergeMode() As Boolean: IsFolderMergeMode = m_IsFolderMergeMode: End Property
Public Property Get isCatalogAppend() As Boolean: isCatalogAppend = m_IsCatalogAppend: End Property
Public Property Get IsFormatErase() As Boolean: IsFormatErase = m_IsFormatErase: End Property

Public Property Get DynamicProjName() As String: DynamicProjName = m_UIProjName: End Property
Public Property Get DynamicImgWidth() As Double: DynamicImgWidth = m_UIImgWidth: End Property
Public Property Get DynamicImgSpacing() As Double: DynamicImgSpacing = m_UIImgSpacing: End Property
Public Property Get DynamicNamingMode() As String: DynamicNamingMode = m_UINamingMode: End Property
Public Property Get DynamicSheetPrefix() As String: DynamicSheetPrefix = m_UISheetPrefix: End Property
Public Property Get DynamicPdfImgName() As String: DynamicPdfImgName = m_UIPdfImgName: End Property
Public Property Get DynamicPdfCatName() As String: DynamicPdfCatName = m_UIPdfCatName: End Property
Public Property Get DynamicPdfGrid() As Integer: DynamicPdfGrid = m_UIPdfGrid: End Property
Public Property Get DynamicExportFilter() As String: DynamicExportFilter = m_UIExportFilter: End Property
Public Property Get DynamicExportRange() As String: DynamicExportRange = m_UIExportRange: End Property

' ==============================================================================
' 7. 全域狀態與雜湊記錄區 (State Machine & Hash Roster)
' ==============================================================================
Public Property Get DictGlobalHash() As Object: Set DictGlobalHash = m_DictGlobalHash: End Property
Public Property Get DictRescueHash() As Object: Set DictRescueHash = m_DictRescueHash: End Property
Public Property Get DictPathToHash() As Object: Set DictPathToHash = m_DictPathToHash: End Property
Public Property Get MissingShapes() As Object: Set MissingShapes = m_MissingShapes: End Property
Public Property Get RestoreCandidates() As Object: Set RestoreCandidates = m_RestoreCandidates: End Property

Public Property Get UserWantsRescue() As Boolean: UserWantsRescue = m_UserWantsRescue: End Property
Public Property Let UserWantsRescue(ByVal vNewValue As Boolean): m_UserWantsRescue = vNewValue: End Property
Public Property Get RequestLayoutLockAfterRebuild() As Boolean: RequestLayoutLockAfterRebuild = m_RequestLayoutLockAfterRebuild: End Property
Public Property Let RequestLayoutLockAfterRebuild(ByVal vNewValue As Boolean): m_RequestLayoutLockAfterRebuild = vNewValue: End Property
