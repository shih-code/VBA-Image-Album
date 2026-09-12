Option Explicit

' ==========================================================
' MODULE: cls_Log (類別模組)
' PURPOSE: 事件日誌儲存器。負責將系統執行過程中的狀態與錯誤訊息，無聲寫入 VeryHidden 的日誌工作表中。
' EXPORTS: Initialize, Record, ArchiveAndTrim, RevealLogSheet, HideLogSheet
' IMPORTS: cls_Config, cls_StringLibrary
' FORBIDDEN: 嚴禁在此模組呼叫 MsgBox 或依賴外部 Utils，必須保持獨立運作，不依賴外部模組。
' DEPENDENCIES: 無外部特殊依賴
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "cls_Log"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
' 【底層私有變數】
' ==============================================================================
Private m_BoundWB As Workbook
Private m_LogSheet As Worksheet
Private m_Config As cls_Config
Private m_IsDebugMode As Boolean
Private m_StrLib As cls_StringLibrary

' ------------------------------------------------------------------------------
' 程式名稱：Class_Initialize
' 功用與目的：初始化物件時，一次性載入字串文庫。
' 為什麼要這樣設計：物件快取優化。由於日誌寫入 (Record) 是高頻呼叫的動作，在初始化時綁定字串庫，能徹底杜絕重複 New 物件的效能開銷。
' ------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_StrLib = New cls_StringLibrary
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：Class_Terminate
' 功用與目的：物件銷毀時釋放記憶體。
' ------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set m_BoundWB = Nothing
    Set m_LogSheet = Nothing
    Set m_Config = Nothing
    Set m_StrLib = Nothing
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：Initialize
' 功用與目的：接收外部傳入的動態與靜態設定，綁定日誌的輸出靶機 (TargetWB)。
' 參數說明：config (靜態設定), TargetWB (目標活頁簿), isDebug (除錯模式開關)
' ------------------------------------------------------------------------------
Public Sub Initialize(ByVal config As cls_Config, ByVal TargetWB As Workbook, ByVal isDebug As Boolean)
    If config Is Nothing Then
        Err.Raise 9981, MODULE_NAME & ".Initialize", _
            "[系統報警] 模組：" & MODULE_NAME & " | 函數：Initialize | 錯誤碼：9981 | 描述：初始化失敗，傳入的靜態組態 Config 為空"
        Exit Sub
    End If
    
    Set m_Config = config
    
    If TargetWB Is Nothing Then
        Set m_BoundWB = ThisWorkbook
    Else
        Set m_BoundWB = TargetWB
    End If
    
    m_IsDebugMode = isDebug
    
    ' 【效能修正】：每次重新 Initialize（代表新的一輪管線執行），強制清空快取的日誌表參照，
    ' 確保不會誤用上一輪執行留下的過期參照（例如活頁簿被重新開啟過）
    Set m_LogSheet = Nothing
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：Record
' 功用與目的：核心合約。將事件層級、來源與描述，無聲無息地寫入隱藏的系統日誌表中。
' 參數說明：eventType (如 ERROR/INFO), sourceCommand (來源函數名稱), customDesc (詳細報警單)
' 效能設計：日誌表的「尋找」與「解鎖」只在 m_LogSheet 快取失效時才執行一次，
'          同一輪管線執行期間的連續寫入直接複用已解鎖狀態，避免每筆日誌都重複
'          觸發 Worksheet.Unprotect/Protect 這類昂貴操作，尤其在工作表數量多、
'          日誌呼叫頻繁時，效能差異會被大量放大。
' ------------------------------------------------------------------------------
Public Sub Record(ByVal eventType As String, ByVal sourceCommand As String, Optional ByVal customDesc As String = "")
    ' --------------------------------------------------------------------------
    ' [第一步：環境檢查與參數防禦]
    ' --------------------------------------------------------------------------
    If m_BoundWB Is Nothing Then Set m_BoundWB = ThisWorkbook
    
    ' 為什麼要這樣判斷：若尚未調用 Initialize 綁定組態，代表系統處於危險狀態。拒絕盲目猜測密碼，直接拋出結構化異常。
    If m_Config Is Nothing Then
        Err.Raise 9982, MODULE_NAME & ".Record", _
            "[系統報警] 模組：" & MODULE_NAME & " | 函數：Record | 錯誤碼：9982 | 描述：未配置系統組態，拒絕猜測初始密碼，熔斷以保護活頁簿結構安全"
        Exit Sub
    End If
    
    Dim sysPwd As String
    sysPwd = m_Config.SecurityPassword
    
    On Error GoTo LogStoreErrorHandler
    
    ' --------------------------------------------------------------------------
    ' [第二步：日誌表偵測與解鎖 (僅在快取失效時執行)]
    ' --------------------------------------------------------------------------
Call ValidateLogSheetCache
    
    If m_LogSheet Is Nothing Then
        Dim wasStructureProtected As Boolean
        wasStructureProtected = m_BoundWB.ProtectStructure
        
        Dim ws As Worksheet
        For Each ws In m_BoundWB.Worksheets
            If ws.Name = m_Config.SheetSysLog Then
                Set m_LogSheet = ws
                Exit For
            End If
        Next ws
        
        ' 為什麼要這樣判斷：若日誌表不存在，代表可能被使用者刪除。立即自動重建並上鎖。
        If m_LogSheet Is Nothing Then
            If wasStructureProtected Then
                m_BoundWB.Unprotect Password:=sysPwd
                
                If m_BoundWB.ProtectStructure Then
                    Err.Raise 9983, MODULE_NAME & ".Record", m_StrLib.ErrUnlockFailed
                End If
            End If
            
            Set m_LogSheet = m_BoundWB.Worksheets.Add(After:=m_BoundWB.Sheets(m_BoundWB.Sheets.count))
            m_LogSheet.Name = m_Config.SheetSysLog
            m_LogSheet.Visible = xlSheetVeryHidden
            
            m_LogSheet.Range("A1:D1").Value = m_StrLib.GetLogHeaders()
            m_LogSheet.Protect Password:=sysPwd
            
            If wasStructureProtected Then
                m_BoundWB.Protect Password:=sysPwd, Structure:=True, Windows:=False
            End If
        End If
        
        ' 解鎖日誌表，之後同一輪連續寫入都不用重複解鎖上鎖
        If m_LogSheet.ProtectContents Then
            m_LogSheet.Unprotect Password:=sysPwd
        End If
    End If
    
    ' --------------------------------------------------------------------------
    ' [第三步：寫入日誌]
    ' --------------------------------------------------------------------------
    Dim nextRow As Long
    nextRow = m_LogSheet.Cells(m_LogSheet.Rows.count, 1).End(xlUp).row + 1
    
    m_LogSheet.Cells(nextRow, 1).Value = Format(Now, "yyyy/mm/dd hh:mm:ss")
    m_LogSheet.Cells(nextRow, 2).Value = UCase(Trim(eventType))
    m_LogSheet.Cells(nextRow, 3).Value = Trim(sourceCommand)
    
    If Trim(customDesc) <> "" Then
        m_LogSheet.Cells(nextRow, 4).Value = Trim(customDesc)
    Else
        m_LogSheet.Cells(nextRow, 4).Value = m_StrLib.MsgLogSuccessLogStore
    End If
    
    Exit Sub

' ------------------------------------------------------------------------------
' 4. 錯誤處理與保護鎖還原
' ------------------------------------------------------------------------------
LogStoreErrorHandler:
    Dim savedErrNum As Long: savedErrNum = Err.Number
    Dim savedErrDesc As String: savedErrDesc = Err.Description
    Err.Clear
    
    ' 為什麼要這樣判斷：若開啟除錯模式，將錯誤輸出到即時運算視窗，避免此模組的錯誤被忽略。
    If m_IsDebugMode Then
        Debug.Print "[" & MODULE_NAME & "] [Record] 關鍵崩潰! 錯誤碼: " & savedErrNum & " | 描述: " & savedErrDesc
    End If
    
    ' 出錯時清空快取，強制下次呼叫重新掃描與修復，避免快取著一個已經失效或損毀的參照
    Set m_LogSheet = Nothing
End Sub


' ------------------------------------------------------------------------------
' 程式名稱：ArchiveAndTrim
' 功用與目的：日誌保留期限管理。當日誌列數超過門檻時，將較舊的資料封存為 CSV
'             至 IOLogArchiveFolder，僅保留最新一批於原表，避免無限增長。
' 為什麼要這樣設計：比照系統既有「複製優先於刪除」原則——先確認 CSV 寫入成功，
'             才刪除原表舊列；任何一步失敗，寧可維持原表不動，也不留下半套結果。
' 參數說明：thresholdRows (觸發封存的列數門檻，預設3000)，keepRows (封存後保留的最新列數，預設500)
' ------------------------------------------------------------------------------
Public Sub ArchiveAndTrim(Optional ByVal thresholdRows As Long = 3000, Optional ByVal keepRows As Long = 500)
    On Error GoTo ArchiveErrorHandler

    If m_BoundWB Is Nothing Then Set m_BoundWB = ThisWorkbook
    If m_Config Is Nothing Then Exit Sub  ' 未初始化，安靜跳過，不影響開機主流程

    Dim sysPwd As String: sysPwd = m_Config.SecurityPassword

    Call LocateLogSheetIfNeeded
    If m_LogSheet Is Nothing Then Exit Sub  ' 日誌表尚不存在，無需封存

    Dim wasProtected As Boolean: wasProtected = m_LogSheet.ProtectContents
    If wasProtected Then m_LogSheet.Unprotect Password:=sysPwd

    Dim lastRow As Long
    lastRow = m_LogSheet.Cells(m_LogSheet.Rows.count, 1).End(xlUp).row
    Dim dataRowCount As Long: dataRowCount = lastRow - 1  ' 扣除標頭列

    If dataRowCount > thresholdRows Then
        Dim archiveEndRow As Long
        archiveEndRow = lastRow - keepRows

        If archiveEndRow >= 2 Then
            Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
            Call EnsureArchiveFolderExists(fso, m_Config.IOLogArchiveFolder)

            Dim archivePath As String
            archivePath = m_Config.IOLogArchiveFolder & "SysLog_封存_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"

            Dim ts As Object: Set ts = fso.CreateTextFile(archivePath, True, True)
            ts.WriteLine """時間""" & "," & """層級""" & "," & """來源""" & "," & """描述"""

            Dim r As Long
            For r = 2 To archiveEndRow
                ts.WriteLine CsvField(m_LogSheet.Cells(r, 1).Value) & "," & _
                             CsvField(m_LogSheet.Cells(r, 2).Value) & "," & _
                             CsvField(m_LogSheet.Cells(r, 3).Value) & "," & _
                             CsvField(m_LogSheet.Cells(r, 4).Value)
            Next r
            ts.Close
            Set ts = Nothing

            ' 【複製優先於刪除】：CSV 已確認寫入完成，才刪除原表舊列
            m_LogSheet.Rows(2 & ":" & archiveEndRow).Delete
            Set fso = Nothing
        End If
    End If

ArchiveCleanup:
    On Error Resume Next
    If Not m_LogSheet Is Nothing Then m_LogSheet.Protect Password:=sysPwd
    On Error GoTo 0
    Exit Sub

ArchiveErrorHandler:
    If m_IsDebugMode Then
        Debug.Print "[" & MODULE_NAME & "] [ArchiveAndTrim] 封存失敗: " & Err.Description
    End If
    Resume ArchiveCleanup
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：RevealLogSheet / HideLogSheet
' 功用與目的：日誌可見度切換。供操作台按鈕呼叫，讓不熟悉 VBE 的使用者也能查看／收起日誌，
'             不需要另外教學怎麼手動取消隱藏工作表。
' ------------------------------------------------------------------------------
Public Sub RevealLogSheet()
    Call LocateLogSheetIfNeeded
    If m_LogSheet Is Nothing Then Exit Sub
    If m_Config Is Nothing Then Exit Sub

    On Error GoTo RevealCleanup
    Dim wasWbStructLocked As Boolean
    wasWbStructLocked = m_BoundWB.ProtectStructure
    If wasWbStructLocked Then m_BoundWB.Unprotect Password:=m_Config.SecurityPassword

    m_LogSheet.Visible = xlSheetVisible

RevealCleanup:
    If wasWbStructLocked Then m_BoundWB.Protect Password:=m_Config.SecurityPassword, Structure:=True, Windows:=False
End Sub

Public Sub HideLogSheet()
    Call LocateLogSheetIfNeeded
    If m_LogSheet Is Nothing Then Exit Sub
    If m_Config Is Nothing Then Exit Sub

    On Error GoTo HideCleanup
    Dim wasWbStructLocked As Boolean
    wasWbStructLocked = m_BoundWB.ProtectStructure
    If wasWbStructLocked Then m_BoundWB.Unprotect Password:=m_Config.SecurityPassword

    m_LogSheet.Visible = xlSheetVeryHidden

HideCleanup:
    If wasWbStructLocked Then m_BoundWB.Protect Password:=m_Config.SecurityPassword, Structure:=True, Windows:=False
End Sub

' ==============================================================================
' 內部輔助（新增）
' ==============================================================================
Private Sub LocateLogSheetIfNeeded()
    Call ValidateLogSheetCache
    If Not m_LogSheet Is Nothing Then Exit Sub
    If m_BoundWB Is Nothing Then Set m_BoundWB = ThisWorkbook
    If m_Config Is Nothing Then Exit Sub
    Dim ws As Worksheet
    For Each ws In m_BoundWB.Worksheets
        If ws.Name = m_Config.SheetSysLog Then
            Set m_LogSheet = ws
            Exit For
        End If
    Next ws
End Sub

Private Function CsvField(ByVal v As Variant) As String
    Dim s As String: s = CStr(v)
    s = Replace(s, """", """""")
    CsvField = """" & s & """"
End Function

Private Sub EnsureArchiveFolderExists(ByVal fso As Object, ByVal folderPath As String)
    Dim p As String: p = folderPath
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
    Dim parts() As String: parts = Split(p, "\")
    Dim cur As String: cur = parts(0) & "\"
    Dim i As Integer
    For i = 1 To UBound(parts)
        cur = cur & parts(i) & "\"
        If Not fso.FolderExists(cur) Then fso.CreateFolder cur
    Next i
End Sub

' ------------------------------------------------------------------------------
' [共用自癒探測]：ValidateLogSheetCache
' 目的：VBA物件變數在其指向的工作表被使用者手動刪除後，不會自動變回Nothing，
'      單純的「Is Nothing」判斷測不出「曾經抓到、但現在已失效」的情況。
'      主動驗證快取參照是否仍然有效，偵測到失效立刻設回Nothing，讓所有
'      呼叫端（Record寫入、Reveal/Hide顯示切換）共用同一套自癒判斷，
'      不再各自維護一份、導致其中一邊漏掉（見Grok審閱、V2-076）
' ------------------------------------------------------------------------------
Private Sub ValidateLogSheetCache()
    If m_LogSheet Is Nothing Then Exit Sub
    Dim logSheetStillAlive As Boolean: logSheetStillAlive = True
    On Error Resume Next
    Dim probeName As String: probeName = m_LogSheet.Name
    If Err.Number <> 0 Then logSheetStillAlive = False
    Err.Clear
    On Error GoTo 0
    If Not logSheetStillAlive Then Set m_LogSheet = Nothing
End Sub
