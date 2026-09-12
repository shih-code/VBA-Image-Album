Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_Main (標準模組)
' PURPOSE: 系統總機與按鈕轉接中樞。絕對不負責任何底層數據運算，專職路由分發。
' EXPORTS: 開始作業, UI_全面排版入口, 操作台格式化, 匯出圖片PDF, SYS_系統開機引導, 局部重排鎖定切換 等
' IMPORTS: Mod_Pipeline, Mod_SystemAdmin, Mod_Export, Mod_UIReset, Mod_UIToggles, cls_StringLibrary, cls_ExecutionContext, cls_Config, Mod_Utils, Mod_UIMessenger, cls_Log
' FORBIDDEN: 1. 嚴禁在此撰寫任何實體檔案操作或 HASH 運算。
'            2. 嚴禁出現任何硬編碼的中文台詞 (應交由 strLib)。
'            3. 嚴禁在此越權控管防護鎖狀態 (已移交 Pipeline 與 UIToggles)。
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_Main"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
' 1. 核心管線調度區 (對接 DAG 引擎與 Pipeline)
' ==============================================================================
Public Sub 開始作業(Optional ByVal Dummy As Byte = 0)
    Dim ctx As cls_ExecutionContext
    Dim strLib As cls_StringLibrary
    Dim fd As FileDialog
    Dim varFile As Variant
    Dim answer As VbMsgBoxResult
    
    Set ctx = BuildCurrentContext()
    If ctx Is Nothing Then Exit Sub ' 若檢驗不通過則安全攔截
    
    Set strLib = New cls_StringLibrary
 
    If Mod_Rules.IsDangerousPath(ctx.config.BaseFolder) Then
        Call Mod_UIMessenger.ShowError(strLib.PromptDangerousZone(""), strLib.TitleDefend)
        Set ctx = Nothing: Set strLib = Nothing: Exit Sub
    End If
    
    ' ==========================================================================
    ' 【匯入來源選擇】：實施檔案與資料夾雙軌分流
    ' ==========================================================================
    answer = MsgBox(strLib.PromptImportMode, vbYesNoCancel + vbQuestion, strLib.TitleImportMode)
                    
    If answer = vbCancel Then GoTo ImportExit
    
    If answer = vbYes Then
        ' ----------------------------------------------------------------------
        ' 軌道 A：單檔挑選模式 (實裝三段式下拉清單過濾器)
        ' ----------------------------------------------------------------------
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .Title = strLib.TitleDialogFilePicker
            .AllowMultiSelect = True
            .Filters.Clear
            .Filters.Add "1. 圖片檔案 (JPG, PNG, GIF...)", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.tif;*.tiff;*.webp"
            .Filters.Add "2. 歸檔附件 (PDF, DOCX, XLSX, MP4...)", "*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.txt;*.mp4;*.zip;*.rar"
            .Filters.Add "3. 所有檔案 (*.*)", "*.*"
            .FilterIndex = 1 ' 預設幫使用者鎖定在「圖片檔案」
            
            If .Show <> -1 Then GoTo ImportExit
            For Each varFile In .SelectedItems
                ctx.SelectedFiles.Add CStr(varFile)
            Next varFile
        End With
        
    ElseIf answer = vbNo Then
        ' ----------------------------------------------------------------------
        ' 軌道 B：資料夾挑選模式
        ' ----------------------------------------------------------------------
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        With fd
            .Title = strLib.TitleDialogFolderPicker
            If .Show <> -1 Then GoTo ImportExit
            ' FolderPicker 只會回傳一個路徑，直接加入待處理清單
            ctx.SelectedFiles.Add CStr(.SelectedItems(1))
        End With
    End If
    
    ' 啟動管線，將待處理清單交給 Pipeline 執行
    Call Mod_Pipeline.Run("IMPORT", Array("IMPORT", "RESCUE", "DEDUPLICATE", "REBUILD"), ctx)
    
ImportExit:
    Set ctx = Nothing: Set strLib = Nothing: Set fd = Nothing
End Sub

Public Sub UI_切換正式排版(Optional ByVal Dummy As Byte = 0)
    Dim ctx As cls_ExecutionContext: Set ctx = BuildCurrentContext()
    If ctx Is Nothing Then Exit Sub
    Call Mod_Pipeline.Run("PREVIEW_LAYOUT", Array("PREVIEW_LAYOUT"), ctx)
    Set ctx = Nothing
End Sub

Public Sub UI_全面排版入口(Optional ByVal Dummy As Byte = 0)
    Dim ctx As cls_ExecutionContext: Set ctx = BuildCurrentContext()
    If ctx Is Nothing Then Exit Sub
    Call Mod_Pipeline.Run("REBUILD", Array("RESCUE", "DEDUPLICATE", "REBUILD"), ctx)
    Set ctx = Nothing
End Sub

' ==============================================================================
' 2. 執行期狀態物件建構與 UI 防呆檢核
' ==============================================================================
Public Function BuildCurrentContext(Optional ByVal Dummy As Byte = 0) As cls_ExecutionContext
    Dim config As cls_Config, evStore As cls_Log, ctx As cls_ExecutionContext
    Dim wsUI As Worksheet, isDebug As Boolean
    
    Set config = New cls_Config
    Set wsUI = Mod_Utils.GetSheetSafe(ThisWorkbook, config.SheetNameUI)
    
    ' 發動 UI 防呆檢核，若失敗則退回 Nothing 中斷後續管線
    If Not VerifyUISettingsSafe(wsUI, config) Then
        Set BuildCurrentContext = Nothing: Set config = Nothing: Set wsUI = Nothing
        Exit Function
    End If
    
    isDebug = False
    If Not wsUI Is Nothing Then If Trim(CStr(wsUI.Range(config.CellDebugStateFlagSet).Value)) = "是" Then isDebug = True
    
    Set evStore = New cls_Log
    Call evStore.Initialize(config, ThisWorkbook, isDebug)
    
    Set ctx = New cls_ExecutionContext
    Call ctx.Initialize(config, evStore)
    If Not wsUI Is Nothing Then Call ctx.LoadUIState(wsUI)
    
    Set BuildCurrentContext = ctx
End Function

Private Function VerifyUISettingsSafe(ByVal ws As Worksheet, ByVal config As cls_Config) As Boolean
    Dim strLib As New cls_StringLibrary
    VerifyUISettingsSafe = False
    If ws Is Nothing Then
        ' UI 工作表尚不存在（全新專案或核心骨架遺失），交由開機流程的象限判斷處理，這裡不視為錯誤
        VerifyUISettingsSafe = True
        Exit Function
    End If
    
    Dim wVal As Variant, sVal As Variant
    wVal = ws.Range(config.CellImgWidth).Value: sVal = ws.Range(config.CellImgSpacing).Value
    
    If Not IsNumeric(wVal) Or Not IsNumeric(sVal) Then
        Call Mod_UIMessenger.ShowError(strLib.ErrUIParamsInvalid, strLib.TitleSysPrompt)
        Exit Function
    End If
    
    If CDbl(wVal) < config.MinImgWidth Then
        Call Mod_UIMessenger.ShowInfo(strLib.InfoWidthCorrected(config.MinImgWidth, True), strLib.TitleSysPrompt)
        Call Mod_Utils.SafeUnprotect(ws, config.SecurityPassword)
        Application.EnableEvents = False: ws.Range(config.CellImgWidth).Value = config.MinImgWidth: Application.EnableEvents = True
        Call Mod_Utils.SafeProtect(ws, config.SecurityPassword)
    ElseIf CDbl(wVal) > config.MaxImgWidth Then
        Call Mod_UIMessenger.ShowInfo(strLib.InfoWidthCorrected(config.MaxImgWidth, False), strLib.TitleSysPrompt)
        Call Mod_Utils.SafeUnprotect(ws, config.SecurityPassword)
        Application.EnableEvents = False: ws.Range(config.CellImgWidth).Value = config.MaxImgWidth: Application.EnableEvents = True
        Call Mod_Utils.SafeProtect(ws, config.SecurityPassword)
    End If
    
    If CDbl(sVal) < config.MinImgSpacing Then
        Call Mod_UIMessenger.ShowInfo(strLib.InfoSpaceCorrected(config.MinImgSpacing, True), strLib.TitleSysPrompt)
        Call Mod_Utils.SafeUnprotect(ws, config.SecurityPassword)
        Application.EnableEvents = False: ws.Range(config.CellImgSpacing).Value = config.MinImgSpacing: Application.EnableEvents = True
        Call Mod_Utils.SafeProtect(ws, config.SecurityPassword)
    ElseIf CDbl(sVal) > config.MaxImgSpacing Then
        Call Mod_UIMessenger.ShowInfo(strLib.InfoSpaceCorrected(config.MaxImgSpacing, False), strLib.TitleSysPrompt)
        Call Mod_Utils.SafeUnprotect(ws, config.SecurityPassword)
        Application.EnableEvents = False: ws.Range(config.CellImgSpacing).Value = config.MaxImgSpacing: Application.EnableEvents = True
        Call Mod_Utils.SafeProtect(ws, config.SecurityPassword)
    End If
    VerifyUISettingsSafe = True
End Function

' ==============================================================================
' 3. 系統防護、格式化與環境開機轉接區
' ==============================================================================
Public Sub 操作台恢復預設()
    Dim ctx As cls_ExecutionContext: Set ctx = BuildCurrentContext()
    Application.EnableEvents = False: Application.ScreenUpdating = False
    
    ' 重建介面並帶入執行期狀態
    Call Mod_UIReset.重建操作介面(True, True, ctx)
    
    Application.ScreenUpdating = True: Application.EnableEvents = True
    Set ctx = Nothing
End Sub

Public Sub 操作台格式化(Optional ByVal Dummy As Byte = 0)
    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True
    
    Dim ctx As cls_ExecutionContext: Set ctx = BuildCurrentContext()
    If ctx Is Nothing Then isRunning = False: Exit Sub
    
    Dim strLib As New cls_StringLibrary
    On Error GoTo FormatErrorHandler
    Application.ScreenUpdating = False
    
    ' 執行實體格式化
    Call Mod_SystemAdmin.執行格式化(ctx)
    
    ' 格式化收尾：記錄操作模式，執行強制全域封鎖
    ctx.OperationMode = "FORMAT"
    Call Mod_UIToggles.EnforceStandardSecurityState(ctx)
    Call Mod_UIReset.GoToUI

FormatCleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Set ctx = Nothing
    isRunning = False
    Exit Sub
FormatErrorHandler:
    Call Mod_UIMessenger.ShowError(strLib.ErrFormatStartupFailed(Err.Description), strLib.TitleSysPrompt)
    Resume FormatCleanup
End Sub

Public Sub 系統緊急重啟(Optional ByVal Dummy As Byte = 0)
    Dim ctx As cls_ExecutionContext: Set ctx = BuildCurrentContext()
    Application.ScreenUpdating = False
    Call Mod_SystemAdmin.執行系統緊急重啟
    
    ctx.OperationMode = "RESTART"
    Call Mod_UIToggles.EnforceStandardSecurityState(ctx)
    Application.ScreenUpdating = True
    Set ctx = Nothing
End Sub

' ------------------------------------------------------------------------------
' [系統開機]：對接 Pipeline 開機判斷邏輯
' ------------------------------------------------------------------------------
Public Sub SYS_系統開機引導(Optional ByVal Dummy As Byte = 0)
    #If Mac Then
            Dim strLib As New cls_StringLibrary
            MsgBox strLib.ErrMacOS, vbCritical: ThisWorkbook.Close SaveChanges:=False: Exit Sub
    #End If
    
    Dim ctx As cls_ExecutionContext
    Set ctx = BuildCurrentContext()
    
    If Not ctx Is Nothing Then
        ' 將開機流程交給 Mod_Pipeline 處理
        Call Mod_Pipeline.BootSystem(ctx)
    End If
    
    ' 開機完成：以 "BOOT" 模式統一執行全域上鎖
    Call Mod_UIToggles.EnforceStandardSecurityState
    Set ctx = Nothing
End Sub

' ==============================================================================
' 4. 單一按鈕微操作路由 (UI Toggles & Exports)
' ==============================================================================
Public Sub 操作台鎖定設定(Optional ByVal Dummy As Byte = 0): Application.ScreenUpdating = False: Call Mod_UIToggles.ToggleConsoleLock: Call Mod_UIReset.GoToUI: Call Mod_UIToggles.SyncAllButtonStates: Application.ScreenUpdating = True: End Sub
Public Sub 目錄鎖定設定(Optional ByVal Dummy As Byte = 0): Application.ScreenUpdating = False: Call Mod_UIToggles.ToggleCatalogLock: Call Mod_UIReset.GoToUI: Call Mod_UIToggles.SyncAllButtonStates: Application.ScreenUpdating = True: End Sub
Public Sub 自動排版開關設定(Optional ByVal Dummy As Byte = 0): Application.ScreenUpdating = False: Call Mod_UIToggles.ToggleAutoLayout: Call Mod_UIReset.GoToUI: Call Mod_UIToggles.SyncAllButtonStates: Application.ScreenUpdating = True: End Sub
Public Sub 工作表鎖定切換(Optional ByVal Dummy As Byte = 0): Application.ScreenUpdating = False: Call Mod_UIToggles.ToggleGlobalLock: Call Mod_UIReset.GoToUI: Call Mod_UIToggles.SyncAllButtonStates: Application.ScreenUpdating = True: End Sub
Public Sub LOG顯示切換設定(Optional ByVal Dummy As Byte = 0): Application.ScreenUpdating = False: Call Mod_UIToggles.ToggleLogVisibility: Call Mod_UIReset.GoToUI: Call Mod_UIToggles.SyncAllButtonStates: Application.ScreenUpdating = True: End Sub
Public Sub 局部重排鎖定切換(Optional ByVal Dummy As Byte = 0)
    Application.ScreenUpdating = False
    If Mod_UIToggles.ToggleLayoutLock() Then
        ' 收到「要上鎖」訊號，由 Router 層自己觸發全面排版，
        ' 跟 Mod_UIToggles 完全無關，不產生反向依賴
        Dim ctx As cls_ExecutionContext: Set ctx = BuildCurrentContext()
        If Not ctx Is Nothing Then
            ctx.RequestLayoutLockAfterRebuild = True
            Call Mod_Pipeline.Run("REBUILD", Array("RESCUE", "DEDUPLICATE", "REBUILD"), ctx)
        End If
    End If
    Call Mod_UIReset.GoToUI
    Call Mod_UIToggles.SyncAllButtonStates
    Application.ScreenUpdating = True
End Sub

Public Sub 匯出圖片PDF(Optional ByVal Dummy As Byte = 0): Dim ctx As cls_ExecutionContext: Set ctx = BuildCurrentContext(): Call Mod_Export.執行匯出圖片PDF(ctx): Set ctx = Nothing: End Sub
Public Sub 匯出橫式目錄PDF(Optional ByVal Dummy As Byte = 0): Dim ctx As cls_ExecutionContext: Set ctx = BuildCurrentContext(): Call Mod_Export.執行匯出橫式目錄PDF(ctx): Set ctx = Nothing: End Sub
Public Sub 匯出附件目錄PDF(Optional ByVal Dummy As Byte = 0): Dim ctx As cls_ExecutionContext: Set ctx = BuildCurrentContext(): Call Mod_Export.執行匯出附件目錄PDF(ctx): Set ctx = Nothing: End Sub

Public Sub UI_匯出當前畫布(Optional ByVal Dummy As Byte = 0)
    Dim config As cls_Config: Set config = New cls_Config
    Dim strLib As New cls_StringLibrary
    If Not Mod_Rules.IsStandardImgSheet(ActiveSheet.Name, config) And ActiveSheet.Name <> config.SheetNameAllPics And ActiveSheet.Name <> config.SheetNameRescue Then
        If Left(ActiveSheet.Name, Len(config.PrefixCustom)) <> config.PrefixCustom Then
            Call Mod_UIMessenger.ShowWarning(strLib.ErrExportUnsupportedSheet, strLib.TitleExportAbort)
            Set config = Nothing: Exit Sub
        End If
    End If
    Call Mod_Export.執行單一畫布匯出PDF(ActiveSheet, config)
    Set config = Nothing
End Sub

' ------------------------------------------------------------------------------
' [專案改名防禦]：純 UI 提示，不涉及實體操作，由總機直接發布
' ------------------------------------------------------------------------------
Public Sub UI_觸發專案更名(ByVal rawInputName As String)
    Dim strLib As cls_StringLibrary: Set strLib = New cls_StringLibrary
    Call Mod_UIMessenger.ShowWarning(strLib.PromptRenameRestriction, strLib.TitleDefend)
    Set strLib = Nothing
End Sub



' ==========================================================
' AUTHOR: 鄭詩樺 (Cheng, Shih-Hua)
' LICENSE - CODE (本活頁簿內全部 VBA 原始碼適用)：MIT License
' Copyright (c) 2026 鄭詩樺 (Cheng, Shih-Hua)
' ==========================================================
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.
' ==========================================================
' LICENSE - DOCUMENTATION（隨附之使用說明書、規格書等文件適用）：
' CC BY-NC-SA 4.0（姓名標示-非商業性-相同方式分享）
'   - 姓名標示（BY）：轉載或改作須標明原作者
'   - 非商業性（NC）：不得作為商業用途
'   - 相同方式分享（SA）：改作後之衍生文件須採用相同授權條款釋出
'   完整條款：https://creativecommons.org/licenses/by-nc-sa/4.0/deed.zh-hant
' ==========================================================
