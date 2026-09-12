Option Explicit

' ==========================================================
' MODULE: 操作介面 (Excel 物件: 工作表)
' PURPOSE: 工作表事件監聽 (View 層)。
' EXPORTS: Worksheet_SelectionChange, Worksheet_Change
' IMPORTS: cls_Config, cls_StringLibrary, Mod_Main, Mod_UIMessenger, Mod_UIToggles, Mod_Utils
' FORBIDDEN: 嚴禁在此撰寫任何實體檔案搬遷、刪除或複雜業務邏輯，僅限「監聽與通報」。
' DEPENDENCIES: 綁定於「操作介面」工作表
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "工作表1"     ' 緊急重置時，指向哪張工作表，需額外調整
Private Const MODULE_VERSION As String = "1.0.0"

' 靜態變數：全域追蹤使用者在操作介面的最後落點
Private m_LastActiveCell As Range

' ------------------------------------------------------------------------------
' 程式名稱：Worksheet_SelectionChange
' 功用與目的：輕量化監聽，紀錄使用者最後選取的儲存格，供後續重繪後視角歸位。
' ------------------------------------------------------------------------------
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim config As cls_Config: Set config = New cls_Config
    Dim targetRange As String: targetRange = config.CellProjName & ":" & config.CellDangerConsole
    If Not Intersect(Target, Me.Range(targetRange)) Is Nothing Then
        If Target.CountLarge = 1 Then Set m_LastActiveCell = Target
    End If
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：Worksheet_Change
' 功用與目的：核心事件監聽。捕捉參數變更並精準分發給 Mod_Main 路由。
' ------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim labelText As String, autoSwitch As String
    Dim strLib As cls_StringLibrary, config As cls_Config
    
    On Error GoTo ChangeErrorHandler
    Set config = New cls_Config
    
    ' 【防禦線】：拒絕批次修改與非目標區域的事件，防止迴圈卡死
    If Target.CountLarge > 1 Then Exit Sub
    If Intersect(Target, Me.Range(config.CellProjName & ":" & config.CellDangerConsole)) Is Nothing Then Exit Sub

    ' ──────────────────────────────────────────────────────────────────────────
    ' 立即凍結畫面，阻止任何背景排版導致的視角跳動與閃爍
    ' ──────────────────────────────────────────────────────────────────────────
    'Application.ScreenUpdating = False
    
    Set m_LastActiveCell = Target
    Set strLib = New cls_StringLibrary

    ' [管線一]：專案名稱主動變更管線 (唯讀看板防禦)
    If Not Intersect(Target, Me.Range(config.CellProjName)) Is Nothing Then
        Application.EnableEvents = False
        Call Mod_Main.UI_觸發專案更名(Trim(CStr(Target.Value)))
        Application.EnableEvents = True
        GoTo ChangeCleanup
    End If

    ' [管線二]：格式化安全驗證核對 (對齊二元安全軌道)
    If Not Intersect(Target, Me.Range(config.CellDangerConsole)) Is Nothing Then
        If Trim(CStr(Target.Value)) <> "保留備份" And Trim(CStr(Target.Value)) <> "僅重置工作表" Then
            Application.EnableEvents = False
            ' 【防禦修正】：強制改值前先解鎖，避免觸發保護衝突
            Call Mod_Utils.SafeUnprotect(Me, config.SecurityPassword)
            Target.Value = "保留備份"
            Call Mod_Utils.SafeProtect(Me, config.SecurityPassword)
            Application.EnableEvents = True
        End If
        GoTo ChangeCleanup
    End If
    
    ' [管線三]：精準自動排版路由器 (對齊硬體契約)
    autoSwitch = Trim(CStr(Me.Range(config.CellAutoLayoutStateConsole).Value))
    If autoSwitch <> "是" Then GoTo ChangeCleanup
    
    ' 透過 A 欄的標籤名稱，判斷使用者修改了哪一個參數
    labelText = Trim(Replace(Replace(CStr(Me.Cells(Target.row, 1).Value), "：", ""), ":", ""))
    
    If labelText = strLib.LabelImgWidthConsole Or labelText = strLib.LabelImgSpacingConsole Then
        Application.EnableEvents = False
        Call Mod_Main.UI_切換正式排版
        GoTo ChangeCleanup
    End If
    
    If labelText = strLib.LabelLayoutModeConsole Then
        Application.EnableEvents = False
        Call Mod_Main.UI_全面排版入口
        GoTo ChangeCleanup
    End If

ChangeCleanup:
    On Error Resume Next
    ' ==========================================================================
    ' 【核心要求 1：視角固定於操作介面】
    ' 強制將畫面與游標鎖回操作介面原位，且這件事必須在 ScreenUpdating = True 前完成
    ' ==========================================================================
    Me.Activate
    If Not m_LastActiveCell Is Nothing Then
        m_LastActiveCell.Select
    Else
        Me.Range(config.CellProjName).Select
    End If
    
    ' ==========================================================================
    ' 【核心要求 2：對準鎖頭狀態】
    ' 強制呼叫 UI 狀態機，確保按鈕顏色與活頁簿真實的密碼保護狀態 100% 吻合
    ' ==========================================================================
    Call Mod_UIToggles.SyncAllButtonStates(config, strLib)
    
    Set strLib = Nothing: Set config = Nothing
    
    ' ──────────────────────────────────────────────────────────────────────────
    ' 所有底層排版、移動與狀態對齊皆已結束，此時才解除安全防線與畫面凍結
    ' ──────────────────────────────────────────────────────────────────────────
    Application.EnableEvents = True
    'Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub

ChangeErrorHandler:
    ' 發生預期外崩潰時，強制進入 Cleanup 確保 EnableEvents 被解開
    Err.Clear
    Resume ChangeCleanup
End Sub
