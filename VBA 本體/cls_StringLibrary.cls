Option Explicit

' ==========================================================
' MODULE: cls_StringLibrary (類別模組)
' PURPOSE: 全系統唯一的中文文字庫 (SSOT)。集中管理對話框文字、UI 標籤、狀態列與日誌模板。
' EXPORTS: 各式字串屬性 (Getters) 與參數化字串函數
' IMPORTS: 無
' FORBIDDEN: 嚴禁在此撰寫任何業務邏輯，僅允許純字串回傳或變數拼接。包含任何格式化與計算。
' DEPENDENCIES: 無
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "cls_StringLibrary"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
'  【UI 狀態與欄位標籤】
' ==============================================================================
Public Property Get HintFormatLocked() As String: HintFormatLocked = "※ 格式化進階設定已鎖定，解開右側【操作台鎖】以展開。": End Property
Public Property Get HintFormatUnlocked() As String: HintFormatUnlocked = "※ 鎖定已解除：請謹慎操作下方之格式化行為。": End Property

Public Property Get CatTitleUnlocked() As String: CatTitleUnlocked = "[ 圖片自動化歸檔目錄清單 ] ── 手動編輯模式已開啟": End Property
Public Property Get CatTitleLocked() As String: CatTitleLocked = "IMAGE ALBUM | 歸檔目錄中心": End Property

Public Property Get BtnTextUILock() As String: BtnTextUILock = "【操作台鎖】" & vbCrLf & "上鎖": End Property
Public Property Get BtnTextUIUnlock() As String: BtnTextUIUnlock = "【操作台鎖】" & vbCrLf & "解鎖": End Property
Public Property Get BtnTextCatLock() As String: BtnTextCatLock = "【圖目錄鎖】" & vbCrLf & "上鎖": End Property
Public Property Get BtnTextCatUnlock() As String: BtnTextCatUnlock = "【圖目錄鎖】" & vbCrLf & "解鎖": End Property
Public Property Get BtnTextAutoOn() As String: BtnTextAutoOn = "【自動排版】" & vbCrLf & "開啟": End Property
Public Property Get BtnTextAutoOff() As String: BtnTextAutoOff = "【自動排版】" & vbCrLf & "關閉": End Property
Public Property Get BtnTextGlobalLock() As String: BtnTextGlobalLock = "【全域鎖】" & vbCrLf & "上鎖": End Property
Public Property Get BtnTextGlobalUnlock() As String: BtnTextGlobalUnlock = "【全域鎖】" & vbCrLf & "解鎖": End Property
Public Property Get BtnTextLogShow() As String: BtnTextLogShow = "【系統日誌】" & vbCrLf & "顯示": End Property
Public Property Get BtnTextLogHide() As String: BtnTextLogHide = "【系統日誌】" & vbCrLf & "隱藏": End Property
Public Property Get BtnTextLayoutOn() As String: BtnTextLayoutOn = "【局部重排】" & vbCrLf & "開啟": End Property
Public Property Get BtnTextLayoutOff() As String: BtnTextLayoutOff = "【局部重排】" & vbCrLf & "關閉": End Property

Public Property Get LabelImgWidthConsole() As String: LabelImgWidthConsole = "圖片寬度 (px)": End Property
Public Property Get LabelImgSpacingConsole() As String: LabelImgSpacingConsole = "圖片間距 (px)": End Property
Public Property Get LabelLayoutModeConsole() As String: LabelLayoutModeConsole = "排版模式": End Property

Public Property Get BtnTextExportSheet() As String: BtnTextExportSheet = "匯出本頁 PDF": End Property
Public Property Get BtnTextGoHome() As String: BtnTextGoHome = "返回控制台": End Property

' ==============================================================================
'  【左下角狀態列 (StatusBar) 動態提示】
' ==============================================================================
Public Property Get StatusReady() As String: StatusReady = "就緒": End Property
Public Property Get StatusScanning() As String: StatusScanning = "正在進行深度遞迴掃描，請稍候...": End Property
Public Property Get StatusDeduplicating() As String: StatusDeduplicating = "正在計算檔案指紋與比對重複項目...": End Property
Public Property Get StatusFormatCleaning() As String: StatusFormatCleaning = "正在執行實體檔案安全封存與畫布清理...": End Property
Public Function StatusImporting(ByVal curr As Long, ByVal total As Long) As String: StatusImporting = "正在匯入並分流資產 (" & curr & " / " & total & ")...": End Function
Public Function StatusExporting(ByVal curr As Long, ByVal total As Long) As String: StatusExporting = "正在靜默匯出 PDF 報告 (" & curr & " / " & total & ")...": End Function
Public Function StatusRebuilding(ByVal curr As Long, ByVal total As Long) As String: StatusRebuilding = "正在執行分頁幾何排版 (" & curr & " / " & total & ")...": End Function
Public Property Get StatusCleaningEmptyFolders() As String: StatusCleaningEmptyFolders = "正在執行沙盒底部空資料夾清理 (由下而上)...": End Property
Public Property Get StatusArchivingBackup() As String: StatusArchivingBackup = "正在執行備份與完整性檢查...": End Property
Public Property Get StatusSyncingView() As String: StatusSyncingView = "作業完成，正在同步操作面板與記憶體狀態...": End Property

' ==============================================================================
'  【系統對話框標題】
' ==============================================================================
Public Property Get TitleSysWarning() As String: TitleSysWarning = "系統警告": End Property
Public Property Get TitleSysPrompt() As String: TitleSysPrompt = "系統提示": End Property
Public Property Get TitleReport() As String: TitleReport = "作業完成報告": End Property
Public Property Get TitleComplete() As String: TitleComplete = "執行完成": End Property
Public Property Get TitleImportMode() As String: TitleImportMode = "匯入模式選擇": End Property
Public Property Get TitleEnvDefend() As String: TitleEnvDefend = "環境異常處理": End Property
Public Property Get TitleDefendSuccess() As String: TitleDefendSuccess = "防護機制啟動成功": End Property
Public Property Get TitleDefend() As String: TitleDefend = "安全防護提示": End Property
Public Property Get TitleDialogFilePicker() As String: TitleDialogFilePicker = "請選擇要匯入的檔案 (支援多選與多種格式)": End Property
Public Property Get TitleDialogFolderPicker() As String: TitleDialogFolderPicker = "請選擇要匯入的資料夾 (系統將自動掃描底下所有子資料夾)": End Property
Public Property Get TitleExportAbort() As String: TitleExportAbort = "匯出中止": End Property                 ' 【新增】
Public Property Get TitleFormatBlocked() As String: TitleFormatBlocked = "格式化防禦攔截": End Property         ' 【新增】
Public Property Get TitleFormatConfirm() As String: TitleFormatConfirm = "系統重置與格式化": End Property       ' 【新增】
Public Property Get TitleFormatComplete() As String: TitleFormatComplete = "格式化完成": End Property           ' 【新增】
Public Property Get TitleRestartSuccess() As String: TitleRestartSuccess = "緊急重啟成功": End Property         ' 【新增】

' ==============================================================================
'  【事前精準預報 (Pre-Flight Check) 模板】
' ==============================================================================
Public Function PromptPreFlightImport(ByVal totalFiles As Long, ByVal totalMB As Double) As String
    PromptPreFlightImport = "【匯入前置確認】" & vbCrLf & vbCrLf & _
          "即將處理 [ " & totalFiles & " ] 個檔案，" & vbCrLf & _
          "總容量約為 [ " & Format(totalMB, "0.00") & " MB ]。" & vbCrLf & vbCrLf & _
          "這需要一點時間，是否確定執行？"
End Function

Public Property Get PromptPreFlightRebuild() As String
    PromptPreFlightRebuild = "【全面排版前置確認】" & vbCrLf & vbCrLf & _
          "即將重組全域目錄與排版。系統將重新掛載所有合法畫布，並剃除無效圖片。" & vbCrLf & _
          "此作業將精準對齊當前實體硬碟之真實狀態。" & vbCrLf & vbCrLf & _
          "是否繼續執行？"
End Property

' ==============================================================================
'  【動態字典產生器】
' ==============================================================================
Public Function GetManualDict(ByVal fontName As String) As Variant
    GetManualDict = Array( _
        Array("[ 開始作業 ]", "啟動匯入程序。支援資料夾全選匯入與手動挑選多張檔案。"), _
        Array("[ 預覽排版 ]", "形狀微調預覽，僅調整圖片寬高，不異動結構與目錄。"), _
        Array("[ 全面排版 ]", "結構重組引擎。依照最新設定重新分裝工作表結構與刷新目錄。"), _
        Array("[ 匯出圖片 ]", "產生靠上對齊、" & fontName & "、淺灰底標籤的 PDF 歸檔報告。"), _
        Array("[ 匯出目錄 ]", "自動過濾技術路徑，產生 A4 橫向、防縮放欄寬的目錄 PDF。") _
    )
End Function

Public Function GetLogHeaders() As Variant
    GetLogHeaders = Array("時間戳記", "事件類型", "來源節點", "狀態描述")
End Function

' ==============================================================================
'  【全域警告與錯誤訊息】
' ==============================================================================
Public Property Get WarnGlobalUnlocked() As String
    WarnGlobalUnlocked = "【系統狀態提示】" & vbCrLf & _
        "全域鎖定已解除。您現在可以自由新增、刪除、移動工作表，或修改儲存格資料。" & vbCrLf & vbCrLf & _
        "請注意以下架構規範：" & vbCrLf & _
        "1. 請勿變更【操作介面】與【系統說明書】的工作表名稱與順序。" & vbCrLf & _
        "2. 若建立新的自訂工作表，名稱請務必以「###」開頭，以免執行排版時被系統清理。"
End Property

Public Property Get ErrCatNotFound() As String: ErrCatNotFound = "系統目前沒有圖片目錄，請先執行排版作業。": End Property
Public Property Get ErrSpaceLimit() As String: ErrSpaceLimit = "【空間不足】剩餘磁碟空間不足以存放歸檔檔案，作業已中斷。": End Property
Public Property Get ErrMacOS() As String: ErrMacOS = "本系統依賴 Windows 原生 API 與檔案系統，不支援 Mac 作業系統。" & vbCrLf & "系統將安全退出。": End Property
Public Property Get ErrUIParamsInvalid() As String: ErrUIParamsInvalid = "【參數錯誤】圖片寬度與間距必須為純數字，請修正後再試。": End Property

Public Property Get ErrDangerZoneFormatBlocked() As String                                                     ' 【新增】
    ErrDangerZoneFormatBlocked = "【系統鎖定：請移至安全資料夾】" & vbCrLf & _
        "系統偵測到您目前位於高風險區域 (如桌面、下載區或硬碟根目錄)。" & vbCrLf & _
        "為防止誤刪您的系統核心檔案，格式化與重置作業已被強制攔截。" & vbCrLf & _
        "請將本專案資料夾移至獨立的磁碟或安全路徑後再試。"
End Property

Public Property Get ErrExportUnsupportedSheet() As String                                                      ' 【新增】
    ErrExportUnsupportedSheet = "目前所在的分頁不支援單頁匯出，請至圖片畫布執行。"
End Property

Public Function ErrFormatStartupFailed(ByVal errDesc As String) As String                                      ' 【新增】
    ErrFormatStartupFailed = "格式化啟動過程發生異常：" & errDesc
End Function

' ==============================================================================
'  【系統核心與管線錯誤模板】
' ==============================================================================
Public Property Get MsgLogSuccessLogStore() As String: MsgLogSuccessLogStore = "管線節點執行成功，實體變更已動態提交": End Property

Public Function MsgMaxLoopExceeded(ByVal alarm As String) As String
    MsgMaxLoopExceeded = "管線遭遇異常導致超出最大迭代限制，已安全攔截。" & vbCrLf & alarm
End Function

Public Function MsgDeadlockDetected(ByVal pendingCount As Long, ByVal alarm As String) As String
    MsgDeadlockDetected = "管線依賴鏈發生循環死鎖或前置依賴遺失，剩餘 " & pendingCount & " 個任務無法解析，安全機制已啟動。" & vbCrLf & alarm
End Function

Public Function MsgCriticalComponentCrash(ByVal cmdName As String, ByVal errNum As Long, ByVal errDesc As String, ByVal alarm As String) As String
    MsgCriticalComponentCrash = "影像自動化流水線遭遇關鍵組件異常，已中斷執行。" & vbCrLf & _
                                "異常組件：" & cmdName & vbCrLf & _
                                "錯誤代碼：" & errNum & vbCrLf & _
                                "異常描述：" & errDesc & vbCrLf & vbCrLf & alarm
End Function

Public Function MsgNonCriticalBypass(ByVal errDesc As String) As String
    MsgNonCriticalBypass = "非關鍵錯誤已略過，管線執行降級容錯推進: " & errDesc
End Function

Public Function ErrPipeline(ByVal errDesc As String) As String: ErrPipeline = "管線異常：" & errDesc: End Function
Public Function ErrLayoutPreview(ByVal errDesc As String) As String: ErrLayoutPreview = "就地排版失敗：" & errDesc: End Function
Public Function ErrLayoutFull(ByVal errDesc As String) As String: ErrLayoutFull = "全面排版失敗：" & errDesc: End Function

' ==============================================================================
'  【提示、更正與確認對話框模板】
' ==============================================================================
 Public Function InfoWidthCorrected(ByVal correctedValue As Double, ByVal isTooLow As Boolean) As String
     If isTooLow Then
         InfoWidthCorrected = "【設定更正】圖片寬度低於下限。系統已主動更正為 " & correctedValue & " px，以防止排版異常。"
     Else
         InfoWidthCorrected = "【設定更正】圖片寬度超過上限。系統已主動更正為 " & correctedValue & " px，以防止排版異常。"
     End If
 End Function

 Public Function InfoSpaceCorrected(ByVal correctedValue As Double, ByVal isTooLow As Boolean) As String
     If isTooLow Then
         InfoSpaceCorrected = "【設定更正】圖片間距低於下限（不可為負數）。系統已主動更正為 " & correctedValue & " px。"
     Else
         InfoSpaceCorrected = "【設定更正】圖片間距超過上限。系統已主動更正為 " & correctedValue & " px。"
     End If
 End Function

Public Function InfoProjectNameSynced(ByVal newName As String) As String
    InfoProjectNameSynced = "偵測到外部大資料夾名稱變更。" & vbCrLf & "系統已自動同步更新 B3 看板名稱至：[" & newName & "]，以保持 SSOT 對齊。"
End Function

Public Property Get PromptImportMode() As String
    PromptImportMode = "請問您要匯入【個別檔案】還是【整個資料夾】？" & vbCrLf & vbCrLf & _
                       "選 [是 (Y)]：挑選個別檔案 (支援多選與下拉格式過濾)" & vbCrLf & _
                       "選 [否 (N)]：挑選整個資料夾 (掃描底下所有檔案與結構)"
End Property

Public Function PromptDangerousZone(ByVal sandboxName As String) As String
    PromptDangerousZone = "【危險路徑警告】" & vbCrLf & vbCrLf & _
                          "系統偵測到目前運行於風險區域（如：桌面/下載/磁碟根目錄）。" & vbCrLf & _
                          "為保護硬碟資料，系統拒絕在此建構三柱目錄。" & vbCrLf & vbCrLf & _
                          "請將本檔案移動到專屬的專案資料夾後再重新開啟！"
End Function

Public Property Get PromptEnvAutoHeal() As String
    PromptEnvAutoHeal = "【系統環境補全提醒】" & vbCrLf & vbCrLf & _
                        "系統發現本 Excel 檔案周圍缺乏完整的三柱資料夾結構。" & vbCrLf & _
                        "請問是否要在這裡 (原地) 自動為您補全【圖片母夾】、【歸檔附錄】與【雜物垃圾】？"
End Property

Public Property Get PromptRenameRestriction() As String                                                       ' 【新增】
    PromptRenameRestriction = "【專案改名限制提示】" & vbCrLf & _
        "受限於 Windows 作業系統的檔案安全鎖定機制，無法在此處直接更改大母夾名稱。" & vbCrLf & vbCrLf & _
        "如果您需要更改專案名稱，請：" & vbCrLf & _
        "1. 儲存並完全關閉本 Excel 檔案。" & vbCrLf & _
        "2. 在 Windows 資料夾中，手動對大母夾重新命名。" & vbCrLf & _
        "3. 重新開啟檔案，系統將自動同步並更新看板名稱！"
End Property

' ==============================================================================
'  【系統管理與格式化專用提示】
' ==============================================================================
Public Function MsgFormatConfirmBackup(ByVal fileCount As Long) As String
    MsgFormatConfirmBackup = "警告：操作台即將重置！" & vbCrLf & _
                 "預計將有 [ " & fileCount & " ] 個實體檔案被移入『備份快照』軟性垃圾桶。" & vbCrLf & vbCrLf & _
                 "請確認是否執行？"
End Function

Public Property Get MsgFormatConfirmResetOnly() As String
    MsgFormatConfirmResetOnly = "警告：操作台與畫布即將重置！" & vbCrLf & _
                 "此模式【僅清除 Excel 內容】，實體檔案將保留於原地不動。" & vbCrLf & vbCrLf & _
                 "請確認是否執行？"
End Property

Public Property Get InfoFormatBaseSuccess() As String
    InfoFormatBaseSuccess = "系統重置與畫布銷毀作業已完成。"
End Property

Public Property Get InfoFilesKeptInPlace() As String: InfoFilesKeptInPlace = "■ 實體檔案已安全保留。": End Property    ' 【新增】

Public Function InfoFilesMovedTo(ByVal folderName As String) As String                                          ' 【新增】
    InfoFilesMovedTo = "■ 實體檔案已移至：[" & folderName & "]"
End Function

Public Property Get WarnLockedFilesSkipped() As String                                                          ' 【新增】
    WarnLockedFilesSkipped = "【注意】部分硬碟檔案正被開啟或遭防毒軟體鎖定，系統已自動略過搬遷以防當機，該檔案仍留存於原處。"
End Property
' =====================================================
' 檔案挑選器相關
' =====================================================
Public Property Get TitleSelectFiles() As String
    TitleSelectFiles = "請選取要匯入的圖片或資料夾 (支援多選)"
End Property

' ==============================================================================
'  【動態報告模板】
' ==============================================================================

 ' 【V2-059新增】REBUILD路徑（全面排版按鈕）原本完全不帶救援數字，
 ' 只有IMPORT路徑（開始作業按鈕）的ReportImportComplete才會顯示。
 Public Function InfoRebuildCompleteWithRescue(ByVal rescued As Long, ByVal failed As Long, ByVal elapsed As Double) As String
     Dim msg As String
     msg = "全面排版與系統重整完畢。耗時 " & elapsed & " 秒。"
     If rescued > 0 Or failed > 0 Then
         msg = msg & vbCrLf & vbCrLf & _
               "────────────────────────" & vbCrLf & _
               "■ 救援成功：" & rescued & " 張" & vbCrLf & _
               "■ 救援失敗：" & failed & " 張"
         If failed > 0 Then
             msg = msg & vbCrLf & vbCrLf & _
                   "失敗項目已保留於原工作表，下次全面排版將自動嘗試救援這 " & failed & " 張圖片。"
         End If
     End If
     InfoRebuildCompleteWithRescue = msg
 End Function

Public Function ReportImportComplete(ByVal imported As Long, ByVal rescued As Long, ByVal failed As Long, ByVal failReasons As String, ByVal elapsed As Double) As String
    ReportImportComplete = "【自動化歸檔管線執行完畢】" & vbCrLf & _
                           "────────────────────────" & vbCrLf & _
                           "■ 成功匯入：" & imported & " 張" & vbCrLf & _
                           "■ 系統救援：" & rescued & " 張" & vbCrLf & _
                           "■ 分流/隔離：" & failed & " 個 (明細: " & Left(failReasons, 60) & "...)" & vbCrLf & _
                           "■ 總耗時  ：" & elapsed & " 秒" & vbCrLf & _
                           "────────────────────────" & vbCrLf & _
                           "詳細記錄已更新於中控台與隔離清單。"
End Function

' ==========================================================
' 匯出模式選擇
' ==========================================================

' 標題：選擇 PDF 圖片來源模式
Public Property Get TitleSelectExportMode() As String
    TitleSelectExportMode = "選擇 PDF 圖片來源模式"
End Property

' 內文：請問是否優先使用【工作表畫布】上的圖片進行匯出？
Public Property Get PromptSelectExportMode() As String
    PromptSelectExportMode = "請問是否優先使用【工作表畫布】上的圖片進行匯出？" & vbCrLf & vbCrLf & _
                           "【是】畫布優先（適用於工作表上已裁切/標註之圖片，失敗自動降流）" & vbCrLf & _
                           "【否】資料夾直匯（讀取硬碟原始高畫質檔）"
End Property



' ==============================================================================
'  【匯出引擎提示模板】
' ==============================================================================
Public Property Get ErrNoCatalog() As String: ErrNoCatalog = "找不到圖片目錄或尚未生成，請先執行全面排版作業。": End Property
Public Property Get ErrNoCatRecord() As String: ErrNoCatRecord = "目錄中尚無資料可供匯出。": End Property
Public Property Get ErrNoValidPicToExport() As String: ErrNoValidPicToExport = "沒有找到符合篩選條件的圖片工作表，或系統中尚無圖片可供匯出。": End Property

Public Property Get InfoExportCatSuccess() As String: InfoExportCatSuccess = "目錄報告 PDF 已成功匯出至 [匯出成果\目錄] 資料夾。": End Property
Public Function InfoBatchExportSuccess(ByVal pdfPath As String) As String: InfoBatchExportSuccess = "影像 PDF 報告匯出成功。" & vbCrLf & "路徑：" & pdfPath: End Function
Public Function InfoSingleExportSuccess(ByVal wsName As String) As String: InfoSingleExportSuccess = "【" & wsName & "】PDF 單頁匯出成功。": End Function

Public Function ErrFileLocked(ByVal fileName As String) As String: ErrFileLocked = "無法產生報告。檔案 [" & fileName & "] 正被其他程式開啟或鎖定，請關閉後重試。": End Function
Public Function ErrExportPic(ByVal errDesc As String) As String: ErrExportPic = "【批次匯出失敗】錯誤描述：" & errDesc: End Function
Public Function ErrSingleExport(ByVal errDesc As String) As String: ErrSingleExport = "【單頁匯出失敗】錯誤描述：" & errDesc: End Function
Public Function ErrExportCat(ByVal errDesc As String) As String: ErrExportCat = "【目錄匯出失敗】錯誤描述：" & errDesc: End Function

' ==============================================================================
'  【UI 狀態切換與安全鎖定錯誤】
' ==============================================================================
Public Property Get ErrUnlockFailed() As String: ErrUnlockFailed = "無法解除活頁簿保護，密碼驗證失敗。": End Property
Public Property Get ErrPwdMismatch() As String: ErrPwdMismatch = "【解鎖失敗】密碼不符，請至「校閱 -> 保護活頁簿」手動解鎖。": End Property
Public Property Get ErrLockIntercepted() As String: ErrLockIntercepted = "【上鎖失敗】系統發生未知攔截，防禦程序未能完成。": End Property
Public Function ErrUIRenderFailed(ByVal errDesc As String) As String: ErrUIRenderFailed = "中控台渲染失敗：" & errDesc: End Function
Public Function ErrToggleFailed(ByVal actionName As String, ByVal errDesc As String) As String: ErrToggleFailed = "切換 [" & actionName & "] 時發生錯誤：" & errDesc: End Function

Public Property Get InfoUIRendered() As String: InfoUIRendered = "中控台重組成功。" & vbCrLf & "■ 系統已恢復預設防禦鎖定狀態。": End Property
Public Property Get InfoSysRestarted() As String: InfoSysRestarted = "系統環境變數已重置，操作介面已重新載入，防護鎖定已恢復預設狀態。": End Property
Public Property Get InfoUnlockSuccess() As String: InfoUnlockSuccess = "系統解鎖與防護更新完畢。" & vbCrLf & "全系統實體鎖已統一同步。": End Property

Public Property Get MsgLogPrefixApplied() As String: MsgLogPrefixApplied = "自訂前綴已於落地沙盒階段套用，嚴格保護來源檔案。": End Property
