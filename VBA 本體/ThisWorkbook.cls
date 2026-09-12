Option Explicit
' ==========================================================
' MODULE: ThisWorkbook (Excel 物件: ThisWorkbook)
' PURPOSE: 活頁簿生命週期事件監聽器 (Trigger 層)。
' EXPORTS: Workbook_Open, Workbook_BeforeClose
' IMPORTS: Mod_Main
' FORBIDDEN: 絕對禁止在此撰寫任何業務邏輯，僅限觸發路由。
' DEPENDENCIES: 綁定於活頁簿本身
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "ThisWorkbook"
Private Const MODULE_VERSION As String = "1.0.0"

Private Sub Workbook_Open()
    ' 檔案開啟的第一秒，立刻將指揮權移交給 Controller (Mod_Main)
    Call Mod_Main.SYS_系統開機引導
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 關閉檔案前，強制還原 Excel 應用程式環境變數，避免因非預期崩潰導致的介面鎖死
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    On Error GoTo 0
End Sub
