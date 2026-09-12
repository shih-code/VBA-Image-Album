Option Explicit
Option Private Module
' ==========================================================
' MODULE: Mod_CommandFactory (標準模組)
' PURPOSE: 指令工廠。負責依據傳入的指令名稱字串，動態實例化並產出對應的 ICommand 命令實體物件。
' EXPORTS: Create
' IMPORTS: ICommand, cmd_Deduplicate, cmd_Import, cmd_Rescue, cmd_Rebuild, cmd_PreviewLayout
' FORBIDDEN: 嚴禁在此模組撰寫任何實體業務邏輯，僅允許進行純粹的類別實例化與分發 (Switch/Case)。
' DEPENDENCIES: 系統中所有的具體命令類別模組
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_CommandFactory"
Private Const MODULE_VERSION As String = "1.0.0"

' ------------------------------------------------------------------------------
' 函數名稱：Create
' 功用與目的：接收字串指令，回傳實作了 ICommand 介面的具體物件。
' 為什麼要這樣設計：實現「控制反轉 (IoC)」。呼叫端 (Mod_Pipeline) 不需知道物件是怎麼建立的，只要跟工廠要，就能拿到符合介面合約的命令物件，大幅降低系統耦合度。
' ------------------------------------------------------------------------------
Public Function Create(ByVal cmdName As String) As ICommand
    ' --------------------------------------------------------------------------
    ' [第一步：傳入參數清理與防禦]
    ' 功用與目的：將傳入的字串轉為大寫並去除空白，防範人為輸入失誤。
    ' --------------------------------------------------------------------------
    Dim cleanName As String
    cleanName = UCase(Trim(cmdName))
    
    ' 為什麼要這樣判斷：Fail-Fast (快速失敗) 原則。如果連名字都沒有，立刻報警並中斷，絕不產出空物件去污染後續的管線。
    If cleanName = "" Then
        Dim emptyAlarm As String
        emptyAlarm = "[系統報警] 模組：" & MODULE_NAME & " | 函數：Create | 錯誤碼：9997 | 描述：傳入的命令名稱為空，工廠拒絕生產 | 參數：cmdName=" & cmdName
        Err.Raise 9997, MODULE_NAME & ".Create", emptyAlarm
        Exit Function
    End If
    
    ' --------------------------------------------------------------------------
    ' [第二步：動態工廠實體分發]
    ' 功用與目的：根據清洗後的名稱，動態 New 出對應的實體類別。
    ' --------------------------------------------------------------------------
    Dim newCmd As ICommand
    
    Select Case cleanName
        Case "DEDUPLICATE"
            Set newCmd = New cmd_Deduplicate
            
        Case "IMPORT"
            Set newCmd = New cmd_Import
            
        Case "RESCUE"
            Set newCmd = New cmd_Rescue
            
        Case "REBUILD"
            Set newCmd = New cmd_Rebuild
            
        Case "PREVIEW_LAYOUT"
            Set newCmd = New cmd_PreviewLayout
            
        Case Else
            ' 為什麼要這樣判斷：防範拼字錯誤或呼叫了尚未開發的指令。此處輸出標準結構化報警單，強制中斷進程。
            Dim errAlarm As String
            errAlarm = "[系統報警] 模組：" & MODULE_NAME & " | 函數：Create | 錯誤碼：9998 | 描述：未知的命令名稱，工廠無法解析對應的類別物件 | 參數：cmdName=" & cmdName
            Err.Raise 9998, MODULE_NAME & ".Create", errAlarm
            Exit Function
    End Select
    
    ' --------------------------------------------------------------------------
    ' [第三步：輸出實體移交]
    ' 功用與目的：將造好的命令物件回傳給呼叫端。
    ' --------------------------------------------------------------------------
    Set Create = newCmd
End Function
