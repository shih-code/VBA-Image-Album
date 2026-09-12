Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_UIToggles (標準模組)
' PURPOSE: UI 狀態切換模組。負責局部切換操作台按鈕的視覺外觀，與設定工作表/活頁簿防護屬性。
' EXPORTS: ToggleConsoleLock, ToggleCatalogLock, ToggleAutoLayout, ToggleGlobalLock, ToggleLayoutLock, SyncAllButtonStates 等
' IMPORTS: cls_Config, Mod_Utils, Mod_UIMessenger, cls_StringLibrary, cls_Log
' FORBIDDEN: 嚴禁在此撰寫任何與實體檔案或業務邏輯相關的程式碼，僅限操作 Excel 介面渲染。
' DEPENDENCIES: 無外部特殊依賴
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_UIToggles"
Private Const MODULE_VERSION As String = "1.0.0"

' 儲存視角用的自訂型態
Private Type ViewState
    ws As Worksheet
    row As Long
    col As Long
End Type

' ------------------------------------------------------------------------------
' 程式名稱：ToggleConsoleLock
' 功能描述：切換操作台防護鎖定，並決定是否顯示危險設定(第 22 列)。
' ------------------------------------------------------------------------------
Public Sub ToggleConsoleLock()
    Dim config As cls_Config
    Dim wsUI As Worksheet
    Dim strLib As New cls_StringLibrary
    Dim vs As ViewState
    Dim isNowLocked As Boolean
    
    On Error GoTo ToggleErrorHandler
    vs = SaveViewState()
    Application.ScreenUpdating = False

    Set config = New cls_Config
    Set wsUI = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameUI)
    If wsUI Is Nothing Then GoTo ToggleCleanup

    If wsUI.ProtectContents Then
        wsUI.Unprotect Password:=config.SecurityPassword
        isNowLocked = False
    Else
        wsUI.Protect Password:=config.SecurityPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True
        isNowLocked = True
    End If

    If isNowLocked Then
        wsUI.Unprotect Password:=config.SecurityPassword
        Call UpdateButtonVisual(wsUI, config.btnLockUI, strLib.BtnTextUIUnlock, config.UIColorLockRed, config.ColorWhite)
        wsUI.Rows("22:22").Hidden = True
        wsUI.Range("A21").Value = strLib.HintFormatLocked
        wsUI.Range("A21").Font.Color = config.ColorTextWarning
        wsUI.Protect Password:=config.SecurityPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True
    Else
        Call UpdateButtonVisual(wsUI, config.btnLockUI, strLib.BtnTextUILock, config.UIColorPassGreen, config.ColorTextMain)
        wsUI.Rows("22:22").Hidden = False
        wsUI.Range("A21").Value = strLib.HintFormatUnlocked
        wsUI.Range("A21").Font.Color = config.ColorTextSuccess
    End If

ToggleCleanup:
    Call RestoreViewState(vs)
    Set config = Nothing: Set wsUI = Nothing: Set strLib = Nothing
    Exit Sub
    
ToggleErrorHandler:
    ' 【精準修正】：拔除指鹿為馬的 ErrUIParamsInvalid，改為正確的切換失敗訊息
    Call Mod_UIMessenger.ShowError(strLib.ErrToggleFailed("操作台鎖", Err.Description), strLib.TitleSysWarning)
    Resume ToggleCleanup
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：ToggleCatalogLock
' 功能描述：切換圖片目錄防護鎖定。
' ------------------------------------------------------------------------------
Public Sub ToggleCatalogLock()
    Dim config As cls_Config
    Dim wsUI As Worksheet
    Dim wsCatalog As Worksheet
    Dim strLib As New cls_StringLibrary
    Dim vs As ViewState
    Dim isNowLocked As Boolean
    Dim wasUIProtected As Boolean
    
    Set config = New cls_Config
    Set wsCatalog = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameCatalog)
    
    If wsCatalog Is Nothing Then
        Call Mod_UIMessenger.ShowWarning(strLib.ErrCatNotFound, strLib.TitleSysPrompt)
        Exit Sub
    End If
    
    On Error GoTo ToggleErrorHandler
    vs = SaveViewState()
    Application.ScreenUpdating = False
    
    If wsCatalog.ProtectContents Then
        wsCatalog.Unprotect Password:=config.SecurityPassword
        wsCatalog.Range("A1").Value = strLib.CatTitleUnlocked
        wsCatalog.Range("A1").Font.Color = config.ColorTextWarning
        isNowLocked = False
    Else
        wsCatalog.Range("A1").Value = strLib.CatTitleLocked
        wsCatalog.Range("A1").Font.Color = config.ColorWhite
        wsCatalog.Protect Password:=config.SecurityPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True
        isNowLocked = True
    End If
    
    Set wsUI = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameUI)
    If Not wsUI Is Nothing Then
        wasUIProtected = wsUI.ProtectContents
        If wasUIProtected Then wsUI.Unprotect Password:=config.SecurityPassword
        
        If isNowLocked Then
            Call UpdateButtonVisual(wsUI, config.btnLockDir, strLib.BtnTextCatUnlock, config.UIColorLockRed, config.ColorWhite)
        Else
            Call UpdateButtonVisual(wsUI, config.btnLockDir, strLib.BtnTextCatLock, config.UIColorPassGreen, config.ColorTextMain)
        End If
        
        If wasUIProtected Then wsUI.Protect Password:=config.SecurityPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If

ToggleCleanup:
    Call RestoreViewState(vs)
    Set config = Nothing: Set wsCatalog = Nothing: Set wsUI = Nothing: Set strLib = Nothing
    Exit Sub
    
ToggleErrorHandler:
    Call Mod_UIMessenger.ShowError(strLib.ErrToggleFailed("圖目錄鎖", Err.Description), strLib.TitleSysWarning)
    Resume ToggleCleanup
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：ToggleAutoLayout
' 功能描述：切換操作台上的自動排版狀態開關。
' ------------------------------------------------------------------------------
Public Sub ToggleAutoLayout()
    Dim config As cls_Config
    Dim wsUI As Worksheet
    Dim strLib As New cls_StringLibrary
    Dim vs As ViewState
    Dim wasUIProtected As Boolean
    Dim isAuto As Boolean
    
    On Error GoTo ToggleErrorHandler
    vs = SaveViewState()
    Application.ScreenUpdating = False
    
    Set config = New cls_Config
    Set wsUI = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameUI)
    If wsUI Is Nothing Then GoTo ToggleCleanup
    
    wasUIProtected = wsUI.ProtectContents
    If wasUIProtected Then wsUI.Unprotect Password:=config.SecurityPassword
    
    ' 【精準修正】：拔除魔法字串 Z1，對齊 config SSOT 錨點
    If Trim(CStr(wsUI.Range(config.CellAutoLayoutStateConsole).Value)) = "是" Then
        wsUI.Range(config.CellAutoLayoutStateConsole).Value = "否"
        isAuto = False
    Else
        wsUI.Range(config.CellAutoLayoutStateConsole).Value = "是"
        isAuto = True
    End If
    
    If isAuto Then
        Call UpdateButtonVisual(wsUI, config.BtnAutoLayout, strLib.BtnTextAutoOn, config.UIColorPassGreen, config.ColorTextMain)
    Else
        Call UpdateButtonVisual(wsUI, config.BtnAutoLayout, strLib.BtnTextAutoOff, config.UIColorLockRed, config.ColorWhite)
    End If
    
    If wasUIProtected Then wsUI.Protect Password:=config.SecurityPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True

ToggleCleanup:
    Call RestoreViewState(vs)
    Set config = Nothing: Set wsUI = Nothing: Set strLib = Nothing
    Exit Sub
    
ToggleErrorHandler:
    Call Mod_UIMessenger.ShowError(strLib.ErrToggleFailed("自動排版開關", Err.Description), strLib.TitleSysWarning)
    Resume ToggleCleanup
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：ToggleGlobalLock
' 功能描述：切換活頁簿與所有工作表的最高權限鎖定。
' ------------------------------------------------------------------------------
Public Sub ToggleGlobalLock()
    Dim config As cls_Config
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsUI As Worksheet
    Dim strLib As New cls_StringLibrary
    Dim vs As ViewState
    Dim isCurrentlyLocked As Boolean
    Dim isNowLocked As Boolean
    
    On Error GoTo ToggleErrorHandler
    vs = SaveViewState()
    Application.ScreenUpdating = False
    
    Set config = New cls_Config
    Set wb = config.TargetWB
    Set wsUI = Mod_Utils.GetSheetSafe(wb, config.SheetNameUI)
    
    isCurrentlyLocked = wb.ProtectStructure

    If isCurrentlyLocked Then
        On Error Resume Next
        wb.Unprotect Password:=config.SecurityPassword
        On Error GoTo ToggleErrorHandler
        
        If wb.ProtectStructure Then
            ' 【精準修正】：對齊解鎖失敗台詞
            Call Mod_UIMessenger.ShowError(strLib.ErrUnlockFailed, strLib.TitleSysWarning)
            GoTo ToggleCleanup
        End If
        
        For Each ws In wb.Worksheets
            Call Mod_Utils.SafeUnprotect(ws, config.SecurityPassword)
        Next ws
        
        isNowLocked = False
        
        If Not wsUI Is Nothing Then
            Call UpdateButtonVisual(wsUI, config.btnGlobalLock, strLib.BtnTextGlobalLock, config.UIColorPassGreen, config.ColorTextMain)
        End If
        
        Call Mod_UIMessenger.ShowInfo(strLib.WarnGlobalUnlocked, "權限釋放警告")
            
    Else
        If Not wsUI Is Nothing Then Call Mod_Utils.SafeUnprotect(wsUI, config.SecurityPassword)
        
        If Not wsUI Is Nothing Then
            Call UpdateButtonVisual(wsUI, config.btnGlobalLock, strLib.BtnTextGlobalUnlock, config.UIColorLockRed, config.ColorWhite)
        End If
        
        For Each ws In wb.Worksheets
            If Left(ws.Name, Len(config.PrefixCustom)) <> config.PrefixCustom Then
                Call Mod_Utils.SafeProtect(ws, config.SecurityPassword)
            End If
        Next ws
        
        On Error Resume Next
        wb.Protect Password:=config.SecurityPassword, Structure:=True, Windows:=False
        On Error GoTo ToggleErrorHandler
        
        If Not wb.ProtectStructure Then
            ' 【精準修正】：對齊上鎖被攔截台詞
            Call Mod_UIMessenger.ShowError(strLib.ErrLockIntercepted, strLib.TitleSysWarning)
            GoTo ToggleCleanup
        End If
        
        isNowLocked = True
    End If

ToggleCleanup:
    Call RestoreViewState(vs)
    Set config = Nothing: Set strLib = Nothing
    Exit Sub
    
ToggleErrorHandler:
    Call Mod_UIMessenger.ShowError(strLib.ErrToggleFailed("全域鎖", Err.Description), strLib.TitleSysWarning)
    Resume ToggleCleanup
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：ApplyLayoutLock
' 功能描述：實際執行「局部重排」上鎖動作（鎖定 B6:B8、改色）。
'           不含確認彈窗——確認已在呼叫端 REBUILD 的 Pre-Flight Check 完成，
'           只由 Mod_Pipeline.Run 在管線成功跑完後呼叫。
' ------------------------------------------------------------------------------
Public Sub ApplyLayoutLock(ByVal config As cls_Config)
    Dim wsUI As Worksheet
    Dim wasUIProtected As Boolean

    Set wsUI = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameUI)
    If wsUI Is Nothing Then Exit Sub

    wasUIProtected = wsUI.ProtectContents
    If wasUIProtected Then wsUI.Unprotect Password:=config.SecurityPassword

    wsUI.Range(config.CellImgWidth & ":" & config.CellDisplayModeFlagSet).Locked = True
    wsUI.Range(config.CellImgWidth & ":" & config.CellDisplayModeFlagSet).Interior.Color = config.UIColorLockRed

    If wasUIProtected Then wsUI.Protect Password:=config.SecurityPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True
    Set wsUI = Nothing
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：ToggleLayoutLock
' 功能描述：切換「局部重排」鎖定狀態。上鎖不自己彈窗，交給 REBUILD 自己的
'           Pre-Flight Check 一併詢問（見 Mod_Pipeline.ExecutePreFlightCheck）；
'           解鎖為低風險動作，不觸發排版、不需確認。
' ------------------------------------------------------------------------------
' 回傳 True 代表「使用者剛把它從關切換成開」，呼叫端需要自己接著觸發一次
' 帶鎖定意圖的全面排版；回傳 False 代表本次呼叫已經處理完畢，不需要額外動作
Public Function ToggleLayoutLock() As Boolean
    ToggleLayoutLock = False
    Dim config As cls_Config
    Dim wsUI As Worksheet
    Dim strLib As New cls_StringLibrary
    Dim vs As ViewState
    Dim isNowLocked As Boolean
    Dim wasUIProtected As Boolean

    On Error GoTo ToggleErrorHandler
    vs = SaveViewState()

    Set config = New cls_Config
    Set wsUI = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameUI)
    If wsUI Is Nothing Then GoTo ToggleCleanup

    isNowLocked = CBool(wsUI.Range(config.CellImgWidth).Locked)

    If Not isNowLocked Then
        ' 不再自己呼叫 Mod_Main／Mod_Pipeline（會造成循環依賴，見
        ' ACDS_DependencyGraph.md 附錄四），只回傳訊號，交給呼叫端處理
        ToggleLayoutLock = True
    Else
        Application.ScreenUpdating = False
        wasUIProtected = wsUI.ProtectContents
        If wasUIProtected Then wsUI.Unprotect Password:=config.SecurityPassword

        wsUI.Range(config.CellImgWidth & ":" & config.CellDisplayModeFlagSet).Locked = False
        wsUI.Range(config.CellImgWidth & ":" & config.CellDisplayModeFlagSet).Interior.Color = config.UIColorPassGreen
        Call UpdateButtonVisual(wsUI, config.BtnLayoutLock, strLib.BtnTextLayoutOff, config.UIColorLockRed, config.ColorWhite)

        If wasUIProtected Then wsUI.Protect Password:=config.SecurityPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If

ToggleCleanup:
    Call RestoreViewState(vs)
    Set config = Nothing: Set wsUI = Nothing: Set strLib = Nothing
    Exit Function

ToggleErrorHandler:
    Call Mod_UIMessenger.ShowError(strLib.ErrToggleFailed("局部重排鎖", Err.Description), strLib.TitleSysWarning)
    Resume ToggleCleanup
End Function


' ==============================================================================
' 內部輔助工具區
' ==============================================================================

Private Sub UpdateButtonVisual(ws As Worksheet, btnName As String, txt As String, bgColor As Long, fgColor As Long)
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes(btnName)
    If Not shp Is Nothing Then
        shp.TextFrame.Characters.Text = txt
        shp.Fill.ForeColor.RGB = bgColor
        shp.TextFrame.Characters.Font.Color = fgColor
    End If
    On Error GoTo 0
End Sub

Private Function SaveViewState() As ViewState
    On Error Resume Next
    Set SaveViewState.ws = ActiveSheet
    SaveViewState.row = ActiveWindow.ScrollRow
    SaveViewState.col = ActiveWindow.ScrollColumn
    On Error GoTo 0
End Function

Private Sub RestoreViewState(vs As ViewState)
    On Error Resume Next
    If Not vs.ws Is Nothing Then vs.ws.Activate
    Application.ScreenUpdating = True
    ActiveWindow.ScrollRow = vs.row
    ActiveWindow.ScrollColumn = vs.col
    On Error GoTo 0
End Sub

' ==============================================================================
' 【核心擴充】：介面狀態全域對齊引擎 (解決按鈕外觀與實際鎖定狀態不同步的問題)
' ==============================================================================

' ------------------------------------------------------------------------------
' 程式名稱：SyncAllButtonStates
' 功用與目的：強制讀取活頁簿與工作表的「真實物理防護狀態」，並一次性重繪操作台上的所有按鈕顏色與文字。
' 為什麼要這樣設計：解耦「業務邏輯」與「介面渲染」。管線結束後只需呼叫此函數，就能確保 UI 100% 反映真實狀態。
' ------------------------------------------------------------------------------
Public Sub SyncAllButtonStates(Optional ByVal config As cls_Config = Nothing, Optional ByVal strLib As cls_StringLibrary = Nothing)
    Dim wsUI As Worksheet
    Dim wsCatalog As Worksheet
    Dim wasUIProtected As Boolean
    Dim bConfigOwn As Boolean, bStrOwn As Boolean
    
    On Error GoTo SyncErr
    
    ' 為什麼要這樣判斷：支援依賴注入，若外部未傳入則自行宣告，減少記憶體浪費。
    If config Is Nothing Then
        Set config = New cls_Config
        bConfigOwn = True
    End If
    If strLib Is Nothing Then
        Set strLib = New cls_StringLibrary
        bStrOwn = True
    End If
    
    Set wsUI = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameUI)
    If wsUI Is Nothing Then GoTo SyncCleanup
    
    wasUIProtected = wsUI.ProtectContents
    If wasUIProtected Then wsUI.Unprotect Password:=config.SecurityPassword
    
    ' --------------------------------------------------------------------------
    ' 對齊 1：全域鎖 (依賴活頁簿結構鎖狀態)
    ' --------------------------------------------------------------------------
    If config.TargetWB.ProtectStructure Then
        Call UpdateButtonVisual(wsUI, config.btnGlobalLock, strLib.BtnTextGlobalUnlock, config.UIColorLockRed, config.ColorWhite)
    Else
        Call UpdateButtonVisual(wsUI, config.btnGlobalLock, strLib.BtnTextGlobalLock, config.UIColorPassGreen, config.ColorTextMain)
    End If
    
    ' --------------------------------------------------------------------------
    ' 對齊 2：操作台鎖 (依賴 UI 工作表防護狀態)
    ' --------------------------------------------------------------------------
    If wasUIProtected Then
        Call UpdateButtonVisual(wsUI, config.btnLockUI, strLib.BtnTextUIUnlock, config.UIColorLockRed, config.ColorWhite)
        wsUI.Rows("22:22").Hidden = True
        wsUI.Range("A21").Value = strLib.HintFormatLocked
        wsUI.Range("A21").Font.Color = config.ColorTextWarning
    Else
        Call UpdateButtonVisual(wsUI, config.btnLockUI, strLib.BtnTextUILock, config.UIColorPassGreen, config.ColorTextMain)
        wsUI.Rows("22:22").Hidden = False
        wsUI.Range("A21").Value = strLib.HintFormatUnlocked
        wsUI.Range("A21").Font.Color = config.ColorTextSuccess
    End If
    
    ' --------------------------------------------------------------------------
    ' 對齊 3：圖目錄鎖 (依賴 目錄工作表防護狀態)
    ' --------------------------------------------------------------------------
    Set wsCatalog = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameCatalog)
    If Not wsCatalog Is Nothing Then
        If wsCatalog.ProtectContents Then
            Call UpdateButtonVisual(wsUI, config.btnLockDir, strLib.BtnTextCatUnlock, config.UIColorLockRed, config.ColorWhite)
        Else
            Call UpdateButtonVisual(wsUI, config.btnLockDir, strLib.BtnTextCatLock, config.UIColorPassGreen, config.ColorTextMain)
        End If
    End If
    
    ' --------------------------------------------------------------------------
    ' 對齊 4：自動排版旗標 (依賴 Z1 儲存格字串)
    ' --------------------------------------------------------------------------
    If Trim(CStr(wsUI.Range(config.CellAutoLayoutStateConsole).Value)) = "是" Then
        Call UpdateButtonVisual(wsUI, config.BtnAutoLayout, strLib.BtnTextAutoOn, config.UIColorPassGreen, config.ColorTextMain)
    Else
        Call UpdateButtonVisual(wsUI, config.BtnAutoLayout, strLib.BtnTextAutoOff, config.UIColorLockRed, config.ColorWhite)
    End If

    ' --------------------------------------------------------------------------
    ' 對齊 5：LOG可見度按鈕 (依賴 系統日誌工作表 Visible 屬性)
    ' --------------------------------------------------------------------------
    Dim logWsSync As Worksheet
    On Error Resume Next
    Set logWsSync = config.TargetWB.Worksheets(config.SheetSysLog)
    On Error GoTo SyncErr
    If Not logWsSync Is Nothing Then
        If logWsSync.Visible = xlSheetVisible Then
            Call UpdateButtonVisual(wsUI, config.BtnToggleLog, strLib.BtnTextLogHide, config.UIColorPassGreen, config.ColorTextMain)
        Else
            Call UpdateButtonVisual(wsUI, config.BtnToggleLog, strLib.BtnTextLogShow, config.UIColorLockRed, config.ColorWhite)
        End If
    End If

    ' --------------------------------------------------------------------------
    ' 對齊 6：局部重排鎖 (依賴 B6:B8 儲存格 Locked 屬性)
    ' --------------------------------------------------------------------------
    Dim isLayoutLockedSync As Boolean
    On Error Resume Next
    isLayoutLockedSync = CBool(wsUI.Range(config.CellImgWidth).Locked)
    On Error GoTo SyncErr
    
     If isLayoutLockedSync Then
        Call UpdateButtonVisual(wsUI, config.BtnLayoutLock, strLib.BtnTextLayoutOn, config.UIColorPassGreen, config.ColorTextMain)
     Else
       Call UpdateButtonVisual(wsUI, config.BtnLayoutLock, strLib.BtnTextLayoutOff, config.UIColorLockRed, config.ColorWhite)
     End If

    If wasUIProtected Then wsUI.Protect Password:=config.SecurityPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True

SyncCleanup:
    If bConfigOwn Then Set config = Nothing
    If bStrOwn Then Set strLib = Nothing
    Set wsUI = Nothing
    Set wsCatalog = Nothing
    Exit Sub
    
SyncErr:
    Err.Clear
    Resume SyncCleanup
End Sub


' ==============================================================================
' 程式名稱：EnforceStandardSecurityState (支援有無執行期狀態物件兩種呼叫情境)
' 功用與目的：支援 Optional Context 參數。
'            無論從管線退場（有 Context）或系統開機（無 Context），皆能自動判斷情境並統一執行鎖定。
' ==============================================================================
Public Sub EnforceStandardSecurityState(Optional ByVal Context As cls_ExecutionContext = Nothing)
    Dim config As cls_Config
    Dim ws As Worksheet, wsUI As Worksheet
    Dim opMode As String
    
    ' ──────────────────────────────────────────────────────────
    ' 【情境判斷】：若有傳入 Context 則依其狀態對齊；未傳入時（如開機階段）則預設為 BOOT 模式
    ' ──────────────────────────────────────────────────────────
    If Not Context Is Nothing Then
        Set config = Context.config
        opMode = UCase(Trim(Context.OperationMode))
    Else
        Set config = New cls_Config
        opMode = "BOOT" ' 自動降維為開機定錨模式
    End If
    
    Set wsUI = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameUI)
       
    ' 2. 統一套用所有分頁防護鎖 (包含 UI 與 Manual)
    For Each ws In config.TargetWB.Worksheets
        If Left(ws.Name, Len(config.PrefixCustom)) <> config.PrefixCustom Then
            Call Mod_Utils.SafeProtect(ws, config.SecurityPassword)
        End If
    Next ws
    
    ' 3. 套用活頁簿全域結構鎖
    On Error Resume Next
    config.TargetWB.Protect Password:=config.SecurityPassword, Structure:=True, Windows:=False
    On Error GoTo 0
    
    ' 4. 視角物理定錨：永遠扯回操作介面 A1
    If Not wsUI Is Nothing Then
        On Error Resume Next
        wsUI.Activate
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        wsUI.Range("A1").Select
        
        ' 【新增緩衝】：讓作業系統有時間處理鏡頭切換，排空繪圖殘渣
        DoEvents
        
        On Error GoTo 0
    End If
    
    ' 5. 驅動 ViewSyncer，讓前台按鈕外觀與實體鎖頭 100% 絕對對齊
    Call SyncAllButtonStates(config)
    
    Set config = Nothing: Set wsUI = Nothing
End Sub


Public Sub ToggleLogVisibility()
    Dim config As cls_Config
    Dim wsUI As Worksheet
    Dim strLib As New cls_StringLibrary
    Dim vs As ViewState
    Dim LogStore As cls_Log
    Dim isNowVisible As Boolean
    Dim wasUIProtected As Boolean
    Dim isDebug As Boolean
    
    On Error GoTo ToggleErrorHandler
    vs = SaveViewState()
    Application.ScreenUpdating = False
    
    Set config = New cls_Config
    Set wsUI = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameUI)
    If wsUI Is Nothing Then GoTo ToggleCleanup
    
    isDebug = False
    If Trim(CStr(wsUI.Range(config.CellDebugStateFlagSet).Value)) = "是" Then isDebug = True
    Set LogStore = New cls_Log
    Call LogStore.Initialize(config, config.TargetWB, isDebug)
    
    wasUIProtected = wsUI.ProtectContents
    If wasUIProtected Then wsUI.Unprotect Password:=config.SecurityPassword
    
    Dim logWs As Worksheet
    On Error Resume Next
    Set logWs = config.TargetWB.Worksheets(config.SheetSysLog)
    On Error GoTo ToggleErrorHandler
    
    If Not logWs Is Nothing Then
        If logWs.Visible = xlSheetVisible Then
            Call LogStore.HideLogSheet
            isNowVisible = False
        Else
            Call LogStore.RevealLogSheet
            isNowVisible = True
        End If
    End If
    
    If isNowVisible Then
        Call UpdateButtonVisual(wsUI, config.BtnToggleLog, strLib.BtnTextLogHide, config.UIColorPassGreen, config.ColorTextMain)
    Else
        Call UpdateButtonVisual(wsUI, config.BtnToggleLog, strLib.BtnTextLogShow, config.UIColorLockRed, config.ColorWhite)
    End If

ToggleCleanup:

    On Error Resume Next
    If wasUIProtected Then wsUI.Protect Password:=config.SecurityPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True
    On Error GoTo 0

    Call RestoreViewState(vs)
    Set config = Nothing: Set wsUI = Nothing: Set strLib = Nothing: Set LogStore = Nothing
    Exit Sub
    
ToggleErrorHandler:
    Call Mod_UIMessenger.ShowError(strLib.ErrToggleFailed("LOG顯示切換", Err.Description), strLib.TitleSysWarning)
    Resume ToggleCleanup
End Sub
