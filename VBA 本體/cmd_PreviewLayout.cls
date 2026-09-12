Option Explicit
Implements ICommand

' ==========================================================
' MODULE: cmd_PreviewLayout (類別模組)
' PURPOSE: 預覽排版流程命令。針對所有合法的圖片分頁，執行就地微調排版。
' EXPORTS: ICommand_Execute 等介面合約
' IMPORTS: cls_ExecutionContext, Mod_Rules, Mod_LayoutEngine, cls_Config
' FORBIDDEN: 1. 嚴禁直接操作或覆蓋實體檔案。
'            2. 嚴禁在此調用 Application.StatusBar 或 DoEvents (說話權統一抽離至管線層)。
' DEPENDENCIES: 無外部特殊依賴
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "cmd_PreviewLayout"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
' ICommand 介面實作區
' ==============================================================================
Private Property Get ICommand_IsCritical() As Boolean
    ICommand_IsCritical = False ' 預覽排版屬非關鍵任務，單頁失敗可略過不熔斷管線
End Property

Private Property Get ICommand_CommandName() As String
    ICommand_CommandName = "PREVIEW_LAYOUT"
End Property

Private Property Get ICommand_Dependencies() As Variant
    ICommand_Dependencies = Array() ' 獨立就地排版，無相依任務
End Property

' ==============================================================================
' 核心執行邏輯 (僅透過 LogStore 回報進度，不使用 UI 元件)
' ==============================================================================
Private Sub ICommand_Execute(ByVal Context As cls_ExecutionContext)
    Dim wsLoop As Worksheet
    Dim errNum As Long
    Dim errDesc As String
    
    Dim config As cls_Config
    Dim LogStore As Object
    Dim totalSheets As Long
    Dim currSheet As Long
    
    On Error GoTo PreviewErrorHandler

    ' 1. 從 Context 取得環境依賴
    Set config = Context.config
    Set LogStore = Context.LogStore
    
    If config Is Nothing Or LogStore Is Nothing Then
        Err.Raise 9991, MODULE_NAME & ".Execute", _
            "[系統報警] 執行期上下文 (Context) 遺失核心參照，管線拒絕執行"
    End If

    LogStore.Record "TX_START", "PREVIEW_LAYOUT", "開始執行就地預覽排版"

    ' --------------------------------------------------------------------------
    ' 2. 進度統計：計算需要排版的合法分頁總數
    ' --------------------------------------------------------------------------
    totalSheets = 0
    For Each wsLoop In config.TargetWB.Worksheets
        If Mod_Rules.IsStandardImgSheet(wsLoop.Name, config) Or wsLoop.Name = config.SheetNameAllPics Or wsLoop.Name = config.SheetNameRescue Then
            totalSheets = totalSheets + 1
        End If
    Next wsLoop

    ' --------------------------------------------------------------------------
    ' 3. 巡迴工作表執行排版 (不使用 StatusBar，改用日誌回報進度)
    ' --------------------------------------------------------------------------
    currSheet = 0
    For Each wsLoop In config.TargetWB.Worksheets
        If Mod_Rules.IsStandardImgSheet(wsLoop.Name, config) Or wsLoop.Name = config.SheetNameAllPics Or wsLoop.Name = config.SheetNameRescue Then
            currSheet = currSheet + 1
            
            ' 【效能優化】：移除 StatusBar 與 DoEvents，改用標準日誌流提報進度
            If currSheet Mod 2 = 0 Or currSheet = totalSheets Then
                LogStore.Record "TX_PROGRESS", "PREVIEW_LAYOUT", "正在算繪畫布幾何排版 (" & currSheet & " / " & totalSheets & ") -> [" & wsLoop.Name & "]"
            End If
            
            ' 呼叫核心幾何瀑布流引擎，執行就地重組
            Call Mod_LayoutEngine.ExecuteMasonryLayout(wsLoop, Context)
        End If
    Next wsLoop

    LogStore.Record "TX_COMMIT", "PREVIEW_LAYOUT", "預覽排版畫布算繪完畢"

PreviewCleanup:
    Set config = Nothing
    Set LogStore = Nothing
    Set wsLoop = Nothing
    Exit Sub

PreviewErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    
    If Not LogStore Is Nothing Then
        LogStore.Record "TX_ERROR", "PREVIEW_LAYOUT", "代碼: " & errNum & " | 描述: " & errDesc
    End If
    
    Set config = Nothing
    Set LogStore = Nothing
    Set wsLoop = Nothing
    
    ' 向上層 (DAG) 拋出結構化異常
    Err.Raise errNum, MODULE_NAME & ".Execute", "預覽排版管線發生異常：" & errDesc
End Sub
