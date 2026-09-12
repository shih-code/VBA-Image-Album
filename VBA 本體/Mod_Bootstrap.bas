Option Explicit
' ==========================================================
' MODULE: Mod_Bootstrap
' PURPOSE: 系統骨架首次建置與急救入口。
'          專供「裸匯入 VBA 模組到空白活頁簿」或核心骨架
'          (UI/說明書工作表) 完全遺失、Workbook_Open 事件
'          無法自動觸發時，使用者手動執行以重建整套系統骨架。
' EXPORTS: 主動初始化
' IMPORTS: cls_Config, cls_StringLibrary, Mod_Rules, Mod_Actions, Mod_UIMessenger, cls_Log, Mod_UIReset, Mod_Utils
' FORBIDDEN: 1. 嚴禁在此撰寫匯入/排版/匯出等業務邏輯，僅限首次骨架建置，
'               其餘一律委由 Mod_Pipeline / Mod_Export 處理。
'            2. 本模組內的 Public 進入點嚴禁加上任何參數（含 Optional），
'               必須保持零參數以確保留在 Alt+F8 巨集清單內可被手動找到，
'               這是本模組存在的唯一理由，違反即失去急救入口的意義。
' DEPENDENCIES: Windows Scripting Host (FileSystemObject)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_Bootstrap"
Private Const MODULE_VERSION As String = "1.0.0"

' ------------------------------------------------------------------------------
' [手動急救入口]：主動初始化
' 目的：不依賴 Mod_Main.BuildCurrentContext（該函式在 UI 表不存在時原本會
'       回傳 Nothing），刻意保持零參數，確保一定留在巨集清單內可被找到。
' ------------------------------------------------------------------------------
Public Sub 主動初始化()
    Dim config As cls_Config: Set config = New cls_Config
    Dim strLib As New cls_StringLibrary
    
    If Mod_Rules.IsDangerousPath(config.BaseFolder) Then
        Call Mod_UIMessenger.ShowError(strLib.PromptDangerousZone(""), strLib.TitleDefend)
        Set config = Nothing: Set strLib = Nothing: Exit Sub
    End If
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim actualName As String: actualName = fso.GetFileName(ThisWorkbook.FullName)
    If actualName = "" Then actualName = "未命名專案"
    
    Call Mod_Actions.InitializeSystemFolders(config)
    Call Mod_Actions.UpdateMemoryPoint(ThisWorkbook, actualName, config)
    Call Mod_Actions.ExecuteStandardBoot(config, actualName)
    
   ' 骨架建立完成，UI 工作表此刻應已存在，清掉不屬於本系統的預設空白表（如 Excel 新檔案自帶的「工作表1」）
    Call Mod_Actions.CleanUpOldSheets(config, False)
    
    
    Call Mod_UIMessenger.ShowInfo("系統骨架已建立完成，包含資料夾結構與操作介面。", "初始化完成")
    Set config = Nothing: Set strLib = Nothing: Set fso = Nothing
End Sub

Public Sub 主動重繪UI()
    Dim config As cls_Config: Set config = New cls_Config
    Dim strLib As New cls_StringLibrary
    If Mod_Rules.IsDangerousPath(config.BaseFolder) Then
        Call Mod_UIMessenger.ShowError(strLib.PromptDangerousZone(""), strLib.TitleDefend)
        Set config = Nothing: Set strLib = Nothing: Exit Sub
    End If
    
    Call Mod_UIReset.重建操作介面(True, True)
    
   ' 骨架建立完成，UI 工作表此刻應已存在，清掉不屬於本系統的預設空白表（如 Excel 新檔案自帶的「工作表1」）
    Call Mod_Actions.CleanUpOldSheets(config, False)

End Sub

' ------------------------------------------------------------------------------
' 程式名稱：強制統一密碼
' 功能描述：嘗試利用候選密碼清單解鎖活頁簿與工作表，並統一重新套用由 Config 指定的系統密碼。
' ------------------------------------------------------------------------------
Public Sub 強制統一密碼()
    Dim config As cls_Config
    Dim strLib As New cls_StringLibrary
    Dim ws As Worksheet
    Dim newPass As String
    Dim passCandidates As Variant
    Dim p As Variant
    Dim isWbUnlocked As Boolean
    Dim hadFailure As Boolean: hadFailure = False
    Dim failedSheetList As String
    
    Set config = New cls_Config
    newPass = config.SecurityPassword
    passCandidates = Array("", "曾用過的密碼候選", newPass)  '建議每次換密碼，都把上一個舊密碼寫上來
    Application.ScreenUpdating = False
    
    isWbUnlocked = False
    If Not ThisWorkbook.ProtectStructure And Not ThisWorkbook.ProtectWindows Then
        isWbUnlocked = True
    Else
        For Each p In passCandidates
            On Error Resume Next
            ThisWorkbook.Unprotect Password:=CStr(p)
            On Error GoTo 0
            If Not ThisWorkbook.ProtectStructure And Not ThisWorkbook.ProtectWindows Then
                isWbUnlocked = True
                Exit For
            End If
        Next p
    End If
    
    If isWbUnlocked Then
        ThisWorkbook.Protect Password:=newPass, Structure:=True, Windows:=False
    Else
        hadFailure = True
        Call Mod_UIMessenger.ShowError(strLib.ErrUnlockFailed, "解鎖失敗")
    End If
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
            For Each p In passCandidates
                On Error Resume Next
                ws.Unprotect Password:=CStr(p)
                On Error GoTo 0
                If Not (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios) Then Exit For
            Next p
            If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
                hadFailure = True
                failedSheetList = failedSheetList & vbCrLf & "?" & ws.Name
            End If
        End If
        Call Mod_Utils.SafeProtect(ws, newPass)
    Next ws
    
    Application.ScreenUpdating = True
    If hadFailure Then
        Dim failMsg As String: failMsg = "部分工作表無法用候選密碼統一，仍維持原密碼："
        If failedSheetList <> "" Then failMsg = failMsg & failedSheetList
        Call Mod_UIMessenger.ShowWarning(failMsg, "統一密碼未完全成功")
    Else
        Call Mod_UIMessenger.ShowInfo(strLib.InfoUnlockSuccess, "執行完成")
    End If
    Set config = Nothing
    Set strLib = Nothing
End Sub


' ------------------------------------------------------------------------------
' 程式名稱：封存並清除系統日誌
' 功能描述：手動觸發日誌封存機制，將超出保留列數的舊日誌搬移至封存備份資料夾。
' ------------------------------------------------------------------------------
 Public Sub 封存並清除系統日誌()
     Dim config As cls_Config: Set config = New cls_Config
     Dim LogStore As New cls_Log
     Call LogStore.Initialize(config, config.TargetWB, False)

     Call LogStore.ArchiveAndTrim
     Call Mod_UIMessenger.ShowInfo("系統日誌封存檢查已完成。", "執行完成")
     Set config = Nothing: Set LogStore = Nothing
 End Sub
