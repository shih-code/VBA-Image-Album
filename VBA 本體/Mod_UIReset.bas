Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_UIReset (標準模組)
' PURPOSE: 介面重繪與狀態控制模組。負責動態重繪系統說明書與操作介面，
'          維護工作表的標準物理排序，並確保表格自然展開與視覺一致性。
' EXPORTS: 重建操作介面, SortWorksheetsStrictly, UI_重整匯出選單, RenderSheetNavButtons, GoToUI, GoToManual
' IMPORTS: cls_Config, cls_StringLibrary, Mod_Utils, Mod_UIMessenger, Mod_UIToggles, cls_Log
' FORBIDDEN: 嚴禁在此撰寫任何與實體檔案或業務運算相關的程式碼，100% 專注於 View 層渲染。
' DEPENDENCIES: 無外部特殊依賴
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_UIReset"
Private Const MODULE_VERSION As String = "1.0.0"

' ------------------------------------------------------------------------------
' 程式名稱：重建操作介面
' 功用與目的：全域重繪引擎。負責清除舊畫布，依據最新設定與權限，
'             重新渲染控制面板所有按鈕，並對齊狀態。
' ------------------------------------------------------------------------------
Public Sub 重建操作介面(Optional isSilent As Boolean = False, _
                        Optional forceDefaultLocks As Boolean = False, _
                        Optional ByVal Context As cls_ExecutionContext = Nothing)
                        
    Dim wb As Workbook, config As cls_Config
    Dim wsM As Worksheet, wsU As Worksheet, catWs As Worksheet, attWs As Worksheet
    Dim strLib As cls_StringLibrary
    Dim curWs As Worksheet, curRow As Long, curCol As Long
    Dim origEvents As Boolean, origScreen As Boolean, isStateSaved As Boolean
    Dim oldGlobalLock As Boolean, oldUILock As Boolean, oldCatLock As Boolean, oldLayoutLock As Boolean
    Dim hadRenderError As Boolean: hadRenderError = False
    
    ' 優先從既有的執行期狀態物件取得 SSOT 設定，維持輕量化
    If Not Context Is Nothing Then
        Set config = Context.config
    Else
        Set config = New cls_Config
    End If
    
    Set wb = config.TargetWB
    Set strLib = New cls_StringLibrary
    isStateSaved = False
    
    On Error GoTo RenderErrorHandler
    
    ' 視角定錨暫存
    Set curWs = ActiveSheet
    curRow = ActiveWindow.ScrollRow
    curCol = ActiveWindow.ScrollColumn
    
    origEvents = Application.EnableEvents
    origScreen = Application.ScreenUpdating
    isStateSaved = True
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
        
    ' 【自癒前置】：活頁簿結構若被鎖，必須先解開，GetOrCreateSheet 才有辦法新增缺失的核心分頁
    Dim wasWbStructLocked As Boolean: wasWbStructLocked = wb.ProtectStructure
    If wasWbStructLocked Then wb.Unprotect Password:=config.SecurityPassword

    Set wsU = Mod_Utils.GetOrCreateSheet(wb, config.SheetNameUI)
    Set wsM = Mod_Utils.GetOrCreateSheet(wb, config.SheetNameManual)
    Set catWs = Mod_Utils.GetSheetSafe(wb, config.SheetNameCatalog)
    
    If wsU Is Nothing Or wsM Is Nothing Then
        Call Mod_UIMessenger.ShowError("【致命異常】系統核心畫布 (UI/Manual) 遺失！" & vbCrLf & _
                                                                            vbCrLf & _
                                                                            "● 排除方式：" & vbCrLf & _
                                                                            "　　1. 檢查「系統說明書」、「操作介面」是否完整。" & vbCrLf & _
                                                                            "　　2. 新增空白工作表，並「重新命名」為遺失工作表。" & vbCrLf & _
                                                                            "　　3. 開啟 F8 點選救援巨集。", strLib.TitleDefend)
        GoTo RenderCleanup
    End If
    
    ' 【自癒延伸】SYS_METADATA跟LOG_SYSTEM比照UI/Manual納入自癒前置，
    ' 讓「重建操作介面」成為真正的全系統骨架自癒入口，不需要使用者記住
    ' 「哪張表壞了要按哪個按鈕」，觸發這一個動作就能一併檢查修復
    Dim wsSysMeta As Worksheet
    Set wsSysMeta = Mod_Utils.GetOrCreateSheet(wb, config.SheetSysMeta)

    Dim uiLogStore As cls_Log
    If Not Context Is Nothing Then
        Set uiLogStore = Context.LogStore
    Else
        Set uiLogStore = New cls_Log
        Call uiLogStore.Initialize(config, wb, False)
    End If
    uiLogStore.Record "TX_INFO", "UIRESET", "重建操作介面觸發系統骨架自癒檢查"

    ' 鎖定狀態記存
    If forceDefaultLocks Then
        oldGlobalLock = True: oldUILock = True: oldCatLock = True
    Else
        oldGlobalLock = wasWbStructLocked
        oldUILock = wsU.ProtectContents
        If Not catWs Is Nothing Then oldCatLock = catWs.ProtectContents Else oldCatLock = True
    End If
    
     ' 【局部重排鎖】：儲存格層級的 Locked 屬性不受工作表保護狀態影響，
    ' 需獨立記存於此，才能在 RenderControlPanelFields 重繪時正確還原
    If forceDefaultLocks Then
        oldLayoutLock = False
    Else
        On Error Resume Next
        oldLayoutLock = wsU.Range(config.CellImgWidth).Locked
        On Error GoTo 0
    End If
    
    ' 解鎖重繪
    If wsU.ProtectContents Then wsU.Unprotect Password:=config.SecurityPassword
    If wsM.ProtectContents Then wsM.Unprotect Password:=config.SecurityPassword
    If Not catWs Is Nothing Then If catWs.ProtectContents Then catWs.Unprotect Password:=config.SecurityPassword
    
    If wsM.Index <> 1 Then wsM.Move Before:=wb.Worksheets(1)
    If wsU.Index <> 2 Then wsU.Move Before:=wb.Worksheets(2)
    
    Call SortWorksheetsStrictly(config)
    Call RenderManualSheet(wsM, config)
    Call RenderControlPanel(wsU, wb, config, oldUILock, oldGlobalLock, oldLayoutLock)
    ' 刷新前台右側的動態抽屜清單
    Call RefreshExportDrawer(wsU, config)
    
    ' ──────────────────────────────────────────────────────────
    ' 【鎖頭對齊修正】：重整渲染 UI 結束，依據引據全面回歸安全防線
    ' ──────────────────────────────────────────────────────────
    If forceDefaultLocks Then
        ' 既然是強制恢復預設，直接呼叫剛才確認過的安全鎖定機制，執行「全域強制上鎖」
        Call Mod_UIToggles.EnforceStandardSecurityState(Context)
    Else
        ' 常規局部重繪，還原剛才紀錄的舊鎖頭狀態
        On Error Resume Next
        If oldUILock Then Call Mod_Utils.SafeProtect(wsU, config.SecurityPassword)
        Call Mod_Utils.SafeProtect(wsM, config.SecurityPassword)
        If oldCatLock Then
            If Not catWs Is Nothing Then Call Mod_Utils.SafeProtect(catWs, config.SecurityPassword)
            If Not attWs Is Nothing Then Call Mod_Utils.SafeProtect(attWs, config.SecurityPassword)
        End If
        If oldGlobalLock Then wb.Protect Password:=config.SecurityPassword, Structure:=True, Windows:=False
        Call Mod_UIToggles.SyncAllButtonStates(config, strLib)
    End If

RenderCleanup:
    On Error Resume Next
    If isStateSaved Then
        If Not curWs Is Nothing Then curWs.Activate
        ActiveWindow.ScrollRow = curRow
        ActiveWindow.ScrollColumn = curCol
        Application.EnableEvents = origEvents
        Application.ScreenUpdating = origScreen
    End If
    
    If Not isSilent And Not hadRenderError Then Call Mod_UIMessenger.ShowInfo(strLib.InfoUIRendered, strLib.TitleSysPrompt)
    Set strLib = Nothing: Set config = Nothing: Set wsM = Nothing: Set wsU = Nothing: Set catWs = Nothing: Set curWs = Nothing: Set wb = Nothing
    Exit Sub

RenderErrorHandler:
    Call Mod_UIMessenger.ShowError(strLib.ErrUIRenderFailed(Err.Description), strLib.TitleSysWarning)
    hadRenderError = True
    Err.Clear
    Resume RenderCleanup
End Sub
' ------------------------------------------------------------------------------
' 程式名稱：SortWorksheetsStrictly
' 功用與目的：強制對齊系統的物理分頁順序（手冊 -> UI -> 隔離 -> 目錄 -> 自訂 -> 圖片）。
' ------------------------------------------------------------------------------
Public Sub SortWorksheetsStrictly(ByVal config As cls_Config)
    Dim wb As Workbook, ws As Worksheet
    Dim oldCats() As String, imgSheets() As String, customSheets() As String
    Dim catCount As Long, imgCount As Long, customCount As Long
    Dim i As Long, j As Long, tmp As String, coreOrder As Variant, wasProtected As Boolean

    Set wb = config.TargetWB
    wasProtected = wb.ProtectStructure
    If wasProtected Then wb.Unprotect Password:=config.SecurityPassword
    
    catCount = 0: imgCount = 0: customCount = 0

    ' 將不同類型的工作表收編到陣列中，以便後續進行排序與搬移
    For Each ws In wb.Worksheets
        If InStr(ws.Name, config.PrefixOldCatalog) = 1 Then
            ReDim Preserve oldCats(0 To catCount): oldCats(catCount) = ws.Name: catCount = catCount + 1
        ElseIf Left(ws.Name, Len(config.PrefixCustom)) = config.PrefixCustom Then
            ReDim Preserve customSheets(0 To customCount): customSheets(customCount) = ws.Name: customCount = customCount + 1
        ElseIf IsImageSheetForSort(ws.Name, config) Then
            ReDim Preserve imgSheets(0 To imgCount): imgSheets(imgCount) = ws.Name: imgCount = imgCount + 1
        End If
    Next ws

    Application.DisplayAlerts = False
    
    ' 氣泡排序法：對舊目錄進行排序並強制清理只保留兩份
    If catCount > 1 Then
        For i = 0 To catCount - 2
            For j = i + 1 To catCount - 1
                If oldCats(i) > oldCats(j) Then tmp = oldCats(i): oldCats(i) = oldCats(j): oldCats(j) = tmp
            Next j
        Next i
    End If
    If catCount > 2 Then
        While catCount > 2
            Set ws = Mod_Utils.GetSheetSafe(wb, oldCats(0))
            If Not ws Is Nothing Then ws.Delete
            For i = 0 To catCount - 2
                oldCats(i) = oldCats(i + 1)
            Next i
            catCount = catCount - 1
        Wend
    End If

    ' 氣泡排序法：自訂分頁排序
    If customCount > 1 Then
        For i = 0 To customCount - 2
            For j = i + 1 To customCount - 1
                If customSheets(i) > customSheets(j) Then tmp = customSheets(i): customSheets(i) = customSheets(j): customSheets(j) = tmp
            Next j
        Next i
    End If

    ' 氣泡排序法：動態圖片分頁排序
    If imgCount > 1 Then
        For i = 0 To imgCount - 2
            For j = i + 1 To imgCount - 1
                If imgSheets(i) > imgSheets(j) Then tmp = imgSheets(i): imgSheets(i) = imgSheets(j): imgSheets(j) = tmp
            Next j
        Next i
    End If

    On Error Resume Next
    ' 核心基座排序：嚴格遵循規格書順序
    coreOrder = Array(config.SheetNameManual, config.SheetNameUI, config.SheetNameQuarantine, config.SheetNameCatalog, config.SheetNameAttachment, config.SheetNameRescue)
    
    For i = LBound(coreOrder) To UBound(coreOrder)
        Set ws = Mod_Utils.GetSheetSafe(wb, CStr(coreOrder(i)))
        If Not ws Is Nothing Then ws.Move After:=wb.Sheets(wb.Sheets.count)
    Next i

    If customCount > 0 Then
        For i = 0 To customCount - 1
            Set ws = Mod_Utils.GetSheetSafe(wb, customSheets(i))
            If Not ws Is Nothing Then ws.Move After:=wb.Sheets(wb.Sheets.count)
        Next i
    End If

    If imgCount > 0 Then
        For i = 0 To imgCount - 1
            Set ws = Mod_Utils.GetSheetSafe(wb, imgSheets(i))
            If Not ws Is Nothing Then ws.Move After:=wb.Sheets(wb.Sheets.count)
        Next i
    End If

    If catCount > 0 Then
        For i = 0 To catCount - 1
            Set ws = Mod_Utils.GetSheetSafe(wb, oldCats(i))
            If Not ws Is Nothing Then ws.Move After:=wb.Sheets(wb.Sheets.count)
        Next i
    End If
    Application.DisplayAlerts = True
    
    If wasProtected Then wb.Protect Password:=config.SecurityPassword, Structure:=True, Windows:=False
    On Error GoTo 0
End Sub

Private Function IsImageSheetForSort(ByVal sName As String, ByVal config As cls_Config) As Boolean
    IsImageSheetForSort = False
    If Left(sName, Len(config.PrefixCustom)) = config.PrefixCustom Then Exit Function
    If InStr(sName, config.PrefixOldCatalog) > 0 Then Exit Function
    
    On Error Resume Next
    If sName = config.SheetNameManual Or sName = config.SheetNameUI Or sName = config.SheetNameQuarantine Or sName = config.SheetNameCatalog Or sName = config.SheetNameAttachment Or sName = config.SheetNameRescue Then Exit Function
    On Error GoTo 0
    
    If sName = config.SheetSysMeta Or sName = config.SheetSysLog Or sName = config.SheetSysTempSettings Then Exit Function
    IsImageSheetForSort = True
End Function

' ------------------------------------------------------------------------------
' 程式名稱：RenderManualSheet
' 功用與目的：重繪系統說明書，並確保說明文字自然展開且包含完整的新手提示。
' ------------------------------------------------------------------------------
Private Sub RenderManualSheet(ws As Worksheet, config As cls_Config)
    Dim strLib As New cls_StringLibrary
    Dim dict As Variant, i As Integer, btnToUI As Shape, shp As Shape, r As Variant, rowOffset As Long, subArr As Variant
    
    dict = strLib.GetManualDict(config.FontMain)
    
    ws.Cells.Clear
    ws.Cells.Font.Name = config.FontMain: ws.Cells.Font.Size = config.FontSizeBase: ws.Cells.Interior.Color = config.ColorWhite
    For Each shp In ws.Shapes
        If shp.Type = msoAutoShape Or shp.Type = msoTextBox Then On Error Resume Next: shp.Delete: On Error GoTo 0
    Next shp
    
    With ws
        .Columns("A:A").ColumnWidth = 32: .Columns("B:B").ColumnWidth = 70: .Columns("A:B").HorizontalAlignment = xlLeft
        .Range("A1").Value = config.AppTitleManual
        .Range("A1").Font.Size = config.FontSizeTitle: .Range("A1").Font.Bold = True: .Range("A1").Font.Color = config.ColorTitle
        .Range("A2:B2").Interior.Color = config.ThemeManual: .Range("A2:B2").Font.Color = config.ColorWhite
        .Range("A2").Value = "■ 使用說明書": .Range("A2").Font.Bold = True
        .Range("A3").Value = "1. 圖片整理："
        .Range("B3").Value = "支援 JPG, JPEG, PNG, BMP, JFIF, WEBP, TIFF 格式，自動進入圖片整理夾與排版畫布。"
        .Range("A4").Value = "2. 附件管理："
        .Range("B4").Value = "支援 PDF、Word、Excel、PowerPoint、純文字/Markdown、影音、壓縮檔，依格式自動分類存放至歸檔附錄；不在此清單內的格式會被送往隔離區，不會遺失但需自行確認處理。"
        .Range("A5").Value = "3. 開工存檔："
        .Range("B5").Value = "全新檔案請先按 [Ctrl + S] 存檔，否則 PDF 無法生成。"
        .Range("A6").Value = "4. 關閉舊報："
        .Range("B6").Value = "匯出前請先關閉同名的舊 PDF 檔案，避免系統路徑鎖死。"
        .Range("A7").Value = "5. 進階擴充："
        ' 【UX 擴充】：補齊自訂表單圖章注入的說明，消滅使用者困惑
        .Range("B7").Value = "若需新增個人工作表且免於被系統排版刪除，請將分頁名稱開頭加上「" & config.PrefixCustom & "」。剛建立時不會有導航按鈕，執行【全面排版】後系統將為其自動注入；此類分頁永遠保持解鎖，不受系統鎖定狀態影響。"
        
        .Range("A9:B9").Interior.Color = config.ThemeExp: .Range("A9:B9").Font.Color = config.ColorWhite
        .Range("A9").Value = "■ 物理環境與圖片宣告": .Range("A9").Font.Bold = True
        .Range("A10").Value = "‧ 系統環境建議："
        .Range("B10").Value = "建議使用 64 位元版本之 Excel，以獲取最佳記憶體調度效能。"
        .Range("A11").Value = "‧ 圖片壓縮："
        .Range("B11").Value = "系統內建 WIA 壓縮，匯入大型圖片時將自動進行等比縮放。"
        
       ' 【V2.2.1新增】資料安全與備份建議——即使系統有多重保護，仍建議使用者
       ' 養成核對數量、額外備份的習慣，見ADR對應圖片消失問題修復系列
       .Range("A13:B13").Interior.Color = config.ThemeExp: .Range("A13:B13").Font.Color = config.ColorWhite
       .Range("A13").Value = "■ 資料安全與備份建議": .Range("A13").Font.Bold = True
       .Range("A14").Value = "‧ 大量作業後："
       .Range("B14").Value = "建議核對目錄表最上方的統計數字，是否跟妳預期的張數吻合，尤其上千張規模作業後。"
       .Range("A15").Value = "‧ 重要圖片備份："
       .Range("B15").Value = "無法重新取得的圖片，請額外自行匯出一份到系統以外的地方，不要只依賴系統畫布顯示。"
       .Range("A16").Value = "‧ 自訂備份表："
       .Range("B16").Value = "工作表名稱開頭加上「" & config.PrefixCustom & "」即可自建純自用的備份記錄表，保證不受任何自動化流程影響。"
       .Range("A17").Value = "‧ 消失時如何查證："
       .Range("B17").Value = "系統日誌搜尋「CleanInvalidShapes判定孤兒並刪除」，可查到具體工作表、名稱與判斷依據。"

       .Range("A19:B19").Interior.Color = config.ThemeIO: .Range("A19:B19").Font.Color = config.ColorWhite
       .Range("A19").Value = "■ 控制按鈕功能字典": .Range("A19").Font.Bold = True

        
        rowOffset = 20
        For i = LBound(dict) To UBound(dict)
            subArr = dict(i)
            .Cells(rowOffset, 1).Value = subArr(LBound(subArr))
            .Cells(rowOffset, 2).Value = subArr(LBound(subArr) + 1)
            .Cells(rowOffset, 1).Font.Bold = True
            .Cells(rowOffset, 1).Font.Color = config.ThemeIO
            rowOffset = rowOffset + 1
        Next i
        
        For Each r In Array("A3:B7", "A10:B11", "A14:B17", "A20:B" & (rowOffset - 1))
            With .Range(CStr(r))
                .Interior.Color = config.InputBg: .Borders.LineStyle = xlContinuous: .Borders.Color = config.Border
                .Font.Size = config.FontSizeBase: .Font.Color = config.ColorTextMain: .WrapText = True
            End With
        Next r
        .Range("A3:A7,A10,A11,A14:A17").Font.Bold = True: .Range("A3:A7,A10,A11,A14:A17").Font.Color = config.ThemeExp
        
        ' 確保所有列依據內容自然展開，不吃字
        .Rows.AutoFit
        
        ' 【穩健化】按鈕位置改用rowOffset動態計算，不再寫死行號——
        ' 前方內容增減時，按鈕永遠正確排在字典區塊結束後兩列，不會疊到內容
        Dim btnRow As Long: btnRow = rowOffset + 1
        Set btnToUI = .Shapes.AddShape(msoShapeRoundedRectangle, .Range("A" & btnRow).Left + 10, .Range("A" & btnRow).Top, 150, 26)
        Call FormatNavButton(btnToUI, config, strLib.BtnTextGoHome, config.UIColorNormal, config.ColorWhite, "Mod_UIReset.GoToUI")
    End With
    Set strLib = Nothing
End Sub

Private Sub ApplyValidationSafe(ByVal rng As Range, ByVal listVals As String)
    Dim sysSeparator As String
    sysSeparator = Application.International(xlListSeparator)
    If sysSeparator <> "," Then listVals = Replace(listVals, ",", sysSeparator)
    On Error Resume Next: rng.Validation.Delete: On Error GoTo 0
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listVals
    rng.Validation.IgnoreBlank = True: rng.Validation.InCellDropdown = True: rng.Validation.ShowInput = True: rng.Validation.ShowError = True
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：RenderControlPanel
' 本 Sub 現為協調者(Orchestrator)：僅負責流程順序，實際邏輯下沉至本函式後方的
' 私有輔助函式（SRP 拆分，ACDS 第八章／第二章）。對外簽章與行為皆未變動。
' ------------------------------------------------------------------------------
Private Sub RenderControlPanel(ws As Worksheet, wb As Workbook, config As cls_Config, ByVal isUILocked As Boolean, ByVal isGlobalLocked As Boolean, ByVal isLayoutLocked As Boolean)
    Dim strLib As New cls_StringLibrary
    Dim vals As Object: Set vals = LoadControlPanelCurrentValues(ws, wb, config)

    Call ClearControlPanelStaleShapes(ws)
    Call RenderControlPanelFields(ws, config, strLib, vals, isUILocked, isLayoutLocked)

    Dim startX As Double, currentY As Double, maxW As Double
    Call RenderControlPanelCard(ws, config, startX, currentY, maxW)

    Call RenderCoreActionButtons(ws, config, startX, currentY, maxW)
    Call RenderEngineToolButtons(ws, config, strLib, startX, currentY, maxW, CStr(vals("curAuto")))
    Call RenderSystemMaintenanceButtons(ws, config, strLib, startX, currentY, maxW, isGlobalLocked, isUILocked, CBool(vals("curLogVisible")))
    Call RenderFooterAndNavButton(ws, config, startX, currentY, maxW)

    Set strLib = Nothing
End Sub

' ------------------------------------------------------------------------------
' 以下為 RenderControlPanel 的 SRP 拆分區：每支函式只做一件事，
' 依第八章 100-150 行安全線與第二章 SRP 原則拆出，全部維持 Private，
' 不影響本模組 EXPORTS 對外清單（RenderControlPanel 本來就是 Private）。
' ------------------------------------------------------------------------------

' 職責：讀取控制面板九組設定值（含各自的預設回退），與 LOG 分頁目前顯示狀態，
'       彙整成單一字典回傳，供渲染函式取用
Private Function LoadControlPanelCurrentValues(ByVal ws As Worksheet, ByVal wb As Workbook, ByVal config As cls_Config) As Object
    Dim vals As Object: Set vals = CreateObject("Scripting.Dictionary")

    vals("curProj") = Trim(CStr(ws.Range(config.CellProjName).Value)): If vals("curProj") = "" Then vals("curProj") = config.UIProjectName
    vals("curW") = ws.Range(config.CellImgWidth).Value: If isEmpty(vals("curW")) Then vals("curW") = config.UIImgWidth
    vals("curS") = ws.Range(config.CellImgSpacing).Value: If isEmpty(vals("curS")) Then vals("curS") = config.UIImgSpacing
    vals("curMode") = Trim(CStr(ws.Range(config.CellDisplayModeFlagSet).Value)): If vals("curMode") = "" Then vals("curMode") = config.UIDisplayMode
    vals("curCat") = Trim(CStr(ws.Range(config.CellCatalogModeFlagSet).Value)): If vals("curCat") = "" Then vals("curCat") = config.CatalogMode
    vals("curNameMode") = Trim(CStr(ws.Range(config.CellNamingMode).Value)): If vals("curNameMode") = "" Then vals("curNameMode") = config.NamingMode
    vals("curPref") = Trim(CStr(ws.Range(config.CellSheetPrefix).Value)): If vals("curPref") = "" Then vals("curPref") = config.UISheetPrefix
    vals("curPdfImg") = Trim(CStr(ws.Range(config.CellPdfImgName).Value)): If vals("curPdfImg") = "" Then vals("curPdfImg") = config.PDFImageReportName
    vals("curPdfCat") = Trim(CStr(ws.Range(config.CellPdfCatName).Value)): If vals("curPdfCat") = "" Then vals("curPdfCat") = config.PDFCatalogReportName
    vals("curGrid") = ws.Range(config.CellPdfGrid).Value: If isEmpty(vals("curGrid")) Then vals("curGrid") = config.PDFGridLayout
    vals("curExpFilter") = Trim(CStr(ws.Range(config.CellExportFilter).Value)): If vals("curExpFilter") = "" Then vals("curExpFilter") = config.ExportFilterMode
    vals("curExpRange") = Trim(CStr(ws.Range(config.CellExportRange).Value)): If vals("curExpRange") = "" Then vals("curExpRange") = config.ExportCustomRange
    vals("curFmt") = Trim(CStr(ws.Range(config.CellDangerConsole).Value)): If vals("curFmt") = "" Then vals("curFmt") = config.FormatBehavior
    vals("curAuto") = Trim(CStr(ws.Range(config.CellAutoLayoutStateConsole).Value)): If vals("curAuto") = "" Then vals("curAuto") = "否"

    Dim curLogVisible As Boolean
    On Error Resume Next
    curLogVisible = (wb.Worksheets(config.SheetSysLog).Visible = xlSheetVisible)
    On Error GoTo 0
    vals("curLogVisible") = curLogVisible

    Set LoadControlPanelCurrentValues = vals
End Function

' 職責：清除面板上舊的 AutoShape／文字方塊殘留，避免每次重繪疊圖
Private Sub ClearControlPanelStaleShapes(ByVal ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoAutoShape Or shp.Type = msoTextBox Then On Error Resume Next: shp.Delete: On Error GoTo 0
    Next shp
End Sub

' 職責：渲染控制面板的資料欄位區（IO／UI／命名匯出／格式化四大區塊文字、樣式、資料驗證）
Private Sub RenderControlPanelFields(ByVal ws As Worksheet, ByVal config As cls_Config, ByVal strLib As cls_StringLibrary, ByVal vals As Object, ByVal isUILocked As Boolean, ByVal isLayoutLocked As Boolean)
    Dim r As Variant, dataRanges As Variant

    With ws
        .Cells.Clear
        .Cells.Font.Name = config.FontMain: .Cells.Font.Size = config.FontSizeBase: .Cells.Interior.Color = config.ColorWhite
        .Rows("1:50").RowHeight = 20: .Rows("1:1").RowHeight = 28: .Rows("2:2").RowHeight = 22

        ' 從解鎖清單中拔除 config.CellProjName (B3)，讓它永遠唯讀
        .Range("B9,B12:B18," & config.CellDangerConsole).Locked = False

        ' 【局部重排鎖】：B6:B8（寬度/間距/排版模式）不無條件解鎖，
        ' 尊重呼叫端傳入的既有狀態；forceDefaultLocks=True 時上層已將
        ' isLayoutLocked 設為 False，等同強制解鎖回預設
        .Range(config.CellImgWidth & ":" & config.CellDisplayModeFlagSet).Locked = isLayoutLocked

        On Error Resume Next: ActiveWindow.DisplayGridlines = False: On Error GoTo 0

        .Columns("A:A").ColumnWidth = 28: .Columns("B:B").ColumnWidth = 35
        .Columns("C:C").ColumnWidth = 6
        .Columns("D:I").ColumnWidth = 11
        .Columns("J:J").ColumnWidth = 6
        .Columns("K:M").Hidden = False
        .Columns("A:B").HorizontalAlignment = xlLeft

        .Range("A1").Value = config.AppTitleControlPanel
        .Range("A1").Font.Size = config.FontSizeSub: .Range("A1").Font.Bold = True: .Range("A1").Font.Color = config.ColorTitle
        .Range("A2:B2").Interior.Color = config.ThemeIO: .Range("A2:B2").Font.Color = config.ColorWhite
        .Range("A2").Value = "■ IO 設定 (檔案與路徑)": .Range("A2").Font.Bold = True
        .Range("A3").Value = "大母夾名稱 (唯讀)："
        .Range(config.CellProjName).Value = vals("curProj")

        .Range("A5:B5").Interior.Color = config.ThemeIO: .Range("A5:B5").Font.Color = config.ColorWhite
        .Range("A5").Value = "■ UI 設定 (排版與規格)": .Range("A5").Font.Bold = True
        .Range("A6").Value = strLib.LabelImgWidthConsole & "：": .Range(config.CellImgWidth).Value = vals("curW")
        .Range("A7").Value = strLib.LabelImgSpacingConsole & "：": .Range(config.CellImgSpacing).Value = vals("curS")
        .Range("A8").Value = strLib.LabelLayoutModeConsole & "：": .Range(config.CellDisplayModeFlagSet).Value = vals("curMode")
        .Range("A9").Value = "目錄模式 (覆蓋/新增)：": .Range(config.CellCatalogModeFlagSet).Value = vals("curCat")
        .Range("A11:B11").Interior.Color = config.ThemeExp: .Range("A11:B11").Font.Color = config.ColorWhite
        .Range("A11").Value = "■ 命名與匯出設定": .Range("A11").Font.Bold = True
        .Range("A12").Value = "命名開頭模式：": .Range(config.CellNamingMode).Value = vals("curNameMode")
        .Range("A13").Value = "自訂開頭文字：": .Range(config.CellSheetPrefix).Value = vals("curPref")
        .Range("A14").Value = "自訂圖片報告檔名：": .Range(config.CellPdfImgName).Value = vals("curPdfImg")
        .Range("A15").Value = "自訂目錄報告檔名：": .Range(config.CellPdfCatName).Value = vals("curPdfCat")
        .Range("A16").Value = "PDF 匯出宮格數 (2/4/6/8)：": .Range(config.CellPdfGrid).Value = vals("curGrid")
        .Range("A17").Value = "匯出篩選模式：": .Range(config.CellExportFilter).Value = vals("curExpFilter")
        .Range("A18").Value = "自訂分頁編號 (例: 1-5, 8)：": .Range(config.CellExportRange).Value = vals("curExpRange")
        .Range(config.CellExportRange).NumberFormat = "@"
        .Range("A20:B20").Interior.Color = config.ThemeExp: .Range("A20:B20").Font.Color = config.ColorWhite
        .Range("A20").Value = "■ 格式化行為設定": .Range("A20").Font.Bold = True
        .Range("A22").Value = "執行模式："
        .Range(config.CellDangerConsole).Value = vals("curFmt")

        dataRanges = Array("A3:B3", "A6:B9", "A12:B18", "A22:B22")
        For Each r In dataRanges
            With .Range(CStr(r))
                .Interior.Color = config.ReadonlyBg: .Borders.LineStyle = xlContinuous: .Borders.Color = config.Border
            End With
        Next r

        ' B3 不再被覆蓋為 InputBg，保持莫蘭迪灰色的唯讀外觀
        .Range("B9,B12:B18," & config.CellDangerConsole).Interior.Color = config.InputBg
        .Range(config.CellImgWidth & ":" & config.CellDisplayModeFlagSet).Interior.Color = IIf(isLayoutLocked, config.UIColorLockRed, config.UIColorPassGreen)
        .Range("A1:A3, A5:A9, A11:A18, A20:A22").IndentLevel = 1

        Call ApplyValidationSafe(.Range(config.CellDisplayModeFlagSet), config.OptDisplayMode)
        Call ApplyValidationSafe(.Range(config.CellCatalogModeFlagSet), config.OptCatalogMode)
        Call ApplyValidationSafe(.Range(config.CellNamingMode), config.OptNamingMode)
        Call ApplyValidationSafe(.Range(config.CellPdfGrid), config.OptGridSize)
        Call ApplyValidationSafe(.Range(config.CellExportFilter), config.OptExportFilter)
        Call ApplyValidationSafe(.Range(config.CellDangerConsole), config.OptFormatBehavior)

        If isUILocked Then
            .Rows("22:22").Hidden = True: .Range(config.CellDangerConsole).FormatConditions.Delete
            .Range("A21").Value = strLib.HintFormatLocked
            .Range("A21").Font.Color = config.ColorTextWarning: .Range("A21").Font.Bold = True: .Range("A21").IndentLevel = 1
        Else
            .Rows("22:22").Hidden = False
            .Range("A21").Value = strLib.HintFormatUnlocked
            .Range("A21").Font.Color = config.ColorTextSuccess: .Range("A21").Font.Bold = True: .Range("A21").IndentLevel = 1
            .Range(config.CellDangerConsole).Font.Bold = True: .Range(config.CellDangerConsole).FormatConditions.Delete
        End If

        .Range(config.CellAutoLayoutStateConsole).Value = vals("curAuto")

    End With
End Sub

' 職責：算出按鈕區的基準座標(startX/currentY/maxW)，並畫出面板卡片背景
Private Function RenderControlPanelCard(ByVal ws As Worksheet, ByVal config As cls_Config, ByRef startX As Double, ByRef currentY As Double, ByRef maxW As Double) As Shape
    Dim btnZone As Range
    Set btnZone = ws.Range("D2:I2")
    startX = btnZone.Left + 8
    currentY = ws.Range("A2").Top
    maxW = btnZone.Width - 16

    Dim panelCard As Shape
    Set panelCard = ws.Shapes.AddShape(msoShapeRoundedRectangle, startX - 10, currentY - 10, maxW + 20, 310)
    With panelCard
        .Fill.Solid: .Fill.ForeColor.RGB = config.ColorPanelBg: .Line.Visible = msoFalse: .Placement = xlFreeFloating
        On Error Resume Next: .Adjustments.Item(1) = 0.04: On Error GoTo 0
        .Shadow.Type = msoShadow21: .Shadow.Blur = 10: .Shadow.Transparency = 0.8: .Shadow.ForeColor.RGB = config.ColorShadow
    End With

    Set RenderControlPanelCard = panelCard
End Function

' 職責：渲染區塊1「核心管線與產出」標題與四顆按鈕，並前移 currentY
Private Sub RenderCoreActionButtons(ByVal ws As Worksheet, ByVal config As cls_Config, ByVal startX As Double, ByRef currentY As Double, ByVal maxW As Double)
    Dim lblHeader As Shape, btn1 As Shape, btn2 As Shape, btn3 As Shape, btn4 As Shape
    Dim rowH As Double: rowH = 32
    Dim headerH As Double: headerH = 16
    Dim qW As Double, x1 As Double, x2 As Double, x3 As Double, x4 As Double

    Set lblHeader = ws.Shapes.AddShape(msoShapeRectangle, startX, currentY, maxW, headerH)
    Call FormatSectionHeader(lblHeader, config, "■ 核心管線與產出", config.ThemeIO)
    currentY = currentY + headerH + config.GapY

    qW = (maxW - (config.GapX * 3)) / 4
    x1 = startX: x2 = startX + qW + config.GapX: x3 = startX + (qW * 2) + (config.GapX * 2): x4 = startX + (qW * 3) + (config.GapX * 3)

    Set btn1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x1, currentY, qW, rowH)
    Call FormatAppButton(btn1, config, "開始作業", config.UIColorNormal, config.ColorWhite, "Mod_Main.開始作業", 10)
    Set btn2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x2, currentY, qW, rowH)
    Call FormatAppButton(btn2, config, "匯出圖片", config.UIColorNormal, config.ColorWhite, "Mod_Main.匯出圖片PDF", 10)
    Set btn3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x3, currentY, qW, rowH)
    Call FormatAppButton(btn3, config, "匯出目錄", config.UIColorNormal, config.ColorWhite, "Mod_Main.匯出橫式目錄PDF", 10)
    Set btn4 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x4, currentY, qW, rowH)
    Call FormatAppButton(btn4, config, "匯出附件", config.UIColorNormal, config.ColorWhite, "Mod_Main.匯出附件目錄PDF", 10)
    currentY = currentY + rowH + config.GapY + 4
End Sub

' 職責：渲染區塊2「引擎操作與資料夾檢視」標題與兩排按鈕，並前移 currentY
Private Sub RenderEngineToolButtons(ByVal ws As Worksheet, ByVal config As cls_Config, ByVal strLib As cls_StringLibrary, ByVal startX As Double, ByRef currentY As Double, ByVal maxW As Double, ByVal curAuto As String)
    Dim lblHeader As Shape, btn1 As Shape, btn2 As Shape, btn3 As Shape, btn4 As Shape
    Dim rowH As Double: rowH = 32
    Dim headerH As Double: headerH = 16
    Dim qW As Double, x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim isLayoutLocked As Boolean

    Set lblHeader = ws.Shapes.AddShape(msoShapeRectangle, startX, currentY, maxW, headerH)
    Call FormatSectionHeader(lblHeader, config, "■ 引擎操作與資料夾檢視", config.ThemeExp)
    currentY = currentY + headerH + config.GapY

    ' 【局部重排】按鈕的鎖定外觀，讀取 RenderControlPanelFields 剛設定好的儲存格 Locked 狀態
    On Error Resume Next
    isLayoutLocked = CBool(ws.Range(config.CellImgWidth).Locked)
    On Error GoTo 0

    qW = (maxW - (config.GapX * 3)) / 4
    x1 = startX: x2 = startX + qW + config.GapX: x3 = startX + (qW * 2) + (config.GapX * 2): x4 = startX + (qW * 3) + (config.GapX * 3)

    Set btn1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x1, currentY, qW, rowH)
    Call FormatAppButton(btn1, config, "預覽排版", config.ThemeIO, config.ColorWhite, "Mod_Main.UI_切換正式排版", 10)
    Set btn2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x2, currentY, qW, rowH)
    Call FormatAppButton(btn2, config, "全面排版", config.ThemeIO, config.ColorWhite, "Mod_Main.UI_全面排版入口", 10)
    Set btn3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x3, currentY, qW, rowH)
    If curAuto = "是" Then
        Call FormatAppButton(btn3, config, strLib.BtnTextAutoOn, config.UIColorPassGreen, config.ColorWhite, "Mod_Main.自動排版開關設定", 9, config.BtnAutoLayout)
    Else
        Call FormatAppButton(btn3, config, strLib.BtnTextAutoOff, config.UIColorLockRed, config.ColorWhite, "Mod_Main.自動排版開關設定", 9, config.BtnAutoLayout)
    End If
    Set btn4 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x4, currentY, qW, rowH)
     If isLayoutLocked Then
        Call FormatAppButton(btn4, config, strLib.BtnTextLayoutOn, config.UIColorPassGreen, config.ColorWhite, "Mod_Main.局部重排鎖定切換", 9, config.BtnLayoutLock)
     Else
        Call FormatAppButton(btn4, config, strLib.BtnTextLayoutOff, config.UIColorLockRed, config.ColorWhite, "Mod_Main.局部重排鎖定切換", 9, config.BtnLayoutLock)
     End If
    currentY = currentY + rowH + config.GapY

    qW = (maxW - (config.GapX * 3)) / 4
    x1 = startX: x2 = startX + qW + config.GapX: x3 = startX + (qW * 2) + (config.GapX * 2): x4 = startX + (qW * 3) + (config.GapX * 3)
    Set btn1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x1, currentY, qW, rowH)
    Call FormatAppButton(btn1, config, "開圖片母夾", config.ThemeExp, config.ColorWhite, "Mod_SystemAdmin.開圖片母夾", 9)
    Set btn2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x2, currentY, qW, rowH)
    Call FormatAppButton(btn2, config, "開匯出成果", config.ThemeExp, config.ColorWhite, "Mod_SystemAdmin.開匯出成果", 9)
    Set btn3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x3, currentY, qW, rowH)
    Call FormatAppButton(btn3, config, "開歸檔附件", config.ThemeExp, config.ColorWhite, "Mod_SystemAdmin.開歸檔附件", 9)
    Set btn4 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x4, currentY, qW, rowH)
    Call FormatAppButton(btn4, config, "開雜物箱", config.ThemeExp, config.ColorWhite, "Mod_SystemAdmin.開雜物箱", 9)
    currentY = currentY + rowH + config.GapY + 4
End Sub

' 職責：渲染區塊3「系統安全與資料庫維護」標題與兩排按鈕，並前移 currentY
Private Sub RenderSystemMaintenanceButtons(ByVal ws As Worksheet, ByVal config As cls_Config, ByVal strLib As cls_StringLibrary, ByVal startX As Double, ByRef currentY As Double, ByVal maxW As Double, ByVal isGlobalLocked As Boolean, ByVal isUILocked As Boolean, ByVal curLogVisible As Boolean)
    Dim lblHeader As Shape, btn1 As Shape, btn2 As Shape, btn3 As Shape, btn4 As Shape
    Dim rowH As Double: rowH = 32
    Dim headerH As Double: headerH = 16
    Dim qW As Double, hW As Double, x1 As Double, x2 As Double, x3 As Double, x4 As Double

    Set lblHeader = ws.Shapes.AddShape(msoShapeRectangle, startX, currentY, maxW, headerH)
    Call FormatSectionHeader(lblHeader, config, "■ 系統安全與資料庫維護", config.ColorTextHint)
    currentY = currentY + headerH + config.GapY

    hW = (maxW - config.GapX) / 2
    x1 = startX: x2 = startX + hW + config.GapX

    Set btn1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x1, currentY, hW, rowH)
    If isGlobalLocked Then
        Call FormatAppButton(btn1, config, strLib.BtnTextGlobalLock, config.UIColorLockRed, config.ColorWhite, "Mod_Main.工作表鎖定切換", 9, config.btnGlobalLock)
    Else
        Call FormatAppButton(btn1, config, strLib.BtnTextGlobalUnlock, config.UIColorPassGreen, config.ColorWhite, "Mod_Main.工作表鎖定切換", 9, config.btnGlobalLock)
    End If
    Set btn2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x2, currentY, hW, rowH)
    If isUILocked Then
        Call FormatAppButton(btn2, config, strLib.BtnTextUILock, config.UIColorLockRed, config.ColorWhite, "Mod_Main.操作台鎖定設定", 9, config.btnLockUI)
    Else
        Call FormatAppButton(btn2, config, strLib.BtnTextUIUnlock, config.UIColorPassGreen, config.ColorWhite, "Mod_Main.操作台鎖定設定", 9, config.btnLockUI)
    End If
    currentY = currentY + rowH + config.GapY

    qW = (maxW - (config.GapX * 3)) / 4
    x1 = startX: x2 = startX + qW + config.GapX: x3 = startX + (qW * 2) + (config.GapX * 2): x4 = startX + (qW * 3) + (config.GapX * 3)
    Set btn1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x1, currentY, qW, rowH)

    If curLogVisible Then
        Call FormatAppButton(btn1, config, strLib.BtnTextLogHide, config.UIColorPassGreen, config.ColorTextMain, "Mod_Main.LOG顯示切換設定", 9, config.BtnToggleLog)
    Else
        Call FormatAppButton(btn1, config, strLib.BtnTextLogShow, config.UIColorLockRed, config.ColorWhite, "Mod_Main.LOG顯示切換設定", 9, config.BtnToggleLog)
    End If

    Set btn2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x2, currentY, qW, rowH)
    Call FormatAppButton(btn2, config, "重整清單索引", config.UIColorNormal, config.ColorWhite, "Mod_UIReset.UI_重整匯出選單", 9, config.BtnRebuildIndex)

    Set btn3 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x3, currentY, qW, rowH)
    Call FormatAppButton(btn3, config, "系統重置與還原", config.ThemeManual, config.ColorWhite, "Mod_SystemAdmin.執行系統緊急重啟", 9, config.BtnSystemRestart)

    Set btn4 = ws.Shapes.AddShape(msoShapeRoundedRectangle, x4 + 5, currentY + 3, qW - 10, rowH - 6)
    Call FormatAppButton(btn4, config, "格式化", config.UIColorLockRed, config.ColorWhite, "Mod_Main.操作台格式化", 9, config.BtnFormat)
    currentY = currentY + rowH + 10
End Sub

' 職責：渲染版權文字與「查看系統說明書」導覽按鈕
Private Sub RenderFooterAndNavButton(ByVal ws As Worksheet, ByVal config As cls_Config, ByVal startX As Double, ByVal currentY As Double, ByVal maxW As Double)
    Dim lblCopyright As Shape
    Set lblCopyright = ws.Shapes.AddShape(msoShapeRectangle, startX, currentY, maxW, 12)
    With lblCopyright
        .Fill.Visible = msoFalse: .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = config.CopyrightText
        .TextFrame2.TextRange.Font.Name = "Consolas": .TextFrame2.TextRange.Font.Size = 7.3: .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = config.ColorTextHint
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft: .TextFrame2.VerticalAnchor = msoAnchorMiddle: .Placement = xlFreeFloating
    End With

    Dim btnToMan As Shape
    Set btnToMan = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("A25").Left + 10, ws.Range("A25").Top, 150, 22)
    Call FormatNavButton(btnToMan, config, "查看系統說明書", config.UIColorNormal, config.ColorWhite, "Mod_UIReset.GoToManual")
End Sub


Public Sub UI_重整匯出選單()
    Dim ws As Worksheet, config As cls_Config
    Set config = New cls_Config
    Set ws = Mod_Utils.GetSheetSafe(ThisWorkbook, config.SheetNameUI)
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    Call Mod_Utils.SafeUnprotect(ws, config.SecurityPassword)
    ws.Columns("J:M").Hidden = False
    Call RefreshExportDrawer(ws, config)
    Call Mod_Utils.SafeProtect(ws, config.SecurityPassword)
    Application.ScreenUpdating = True
    Set config = Nothing: Set ws = Nothing
End Sub

' ------------------------------------------------------------------------------
' 程式名稱：RefreshExportDrawer
' 功用與目的：重整 M 欄的動態匯出抽屜，實裝視覺一致性與自動列高展開。
' ------------------------------------------------------------------------------
Private Sub RefreshExportDrawer(ws As Worksheet, config As cls_Config)
    Dim scanWs As Worksheet, rowIdx As Long, customIdx As Long, imgIdx As Long
    
    ws.Range("J:M").Clear
    ws.Columns("J:J").ColumnWidth = 2: ws.Columns("K:K").ColumnWidth = 35
    ws.Columns("L:L").ColumnWidth = 15: ws.Columns("M:M").ColumnWidth = 2
    
    rowIdx = 2: customIdx = 1
    
    For Each scanWs In config.TargetWB.Worksheets
        If Left(scanWs.Name, Len(config.PrefixCustom)) = config.PrefixCustom Then
            If customIdx = 1 Then
                ws.Range("K" & rowIdx & ":L" & rowIdx).Merge
                ws.Range("K" & rowIdx).Value = "■ 自訂分頁索引"
                ws.Range("K" & rowIdx).Font.Name = config.FontMain: ws.Range("K" & rowIdx).Font.Bold = True: ws.Range("K" & rowIdx).Font.Color = config.ColorWhite
                ws.Range("K" & rowIdx & ":L" & rowIdx).Interior.Color = config.ThemeManual ' 視覺對齊：套用控制台色系
                ws.Range("K" & rowIdx & ":L" & rowIdx).RowHeight = 18: ws.Range("K" & rowIdx).VerticalAlignment = xlVAlignCenter
                rowIdx = rowIdx + 1
            End If
            
            ws.Cells(rowIdx, 11).Value = "‧ " & scanWs.Name
            ws.Cells(rowIdx, 12).Formula = "=HYPERLINK(""#'" & scanWs.Name & "'!A1"", ""點擊預覽"")"
            
            With ws.Range(ws.Cells(rowIdx, 11), ws.Cells(rowIdx, 12))
                .Font.Name = config.FontMain: .Font.Size = config.FontSizeBase
                .Borders.LineStyle = xlContinuous: .Borders.Color = config.DrawerItemBorder
                .Interior.Color = config.DrawerItemBg
                .WrapText = True ' 允許自然換行不吃字
                .VerticalAlignment = xlVAlignCenter
            End With
            
            ws.Cells(rowIdx, 11).Font.Color = config.DrawerItemText
            ws.Cells(rowIdx, 11).Font.Bold = True
            ws.Cells(rowIdx, 12).Font.Color = config.DrawerLinkText
            ws.Cells(rowIdx, 12).Font.Underline = xlUnderlineStyleNone ' 拔除原生超連結底線
            rowIdx = rowIdx + 1: customIdx = customIdx + 1
        End If
    Next scanWs
    
    If customIdx > 1 Then rowIdx = rowIdx + 1
    
    ws.Range("K" & rowIdx & ":L" & rowIdx).Merge
    ws.Range("K" & rowIdx).Value = "■ 動態匯出清單索引"
    ws.Range("K" & rowIdx).Font.Name = config.FontMain: ws.Range("K" & rowIdx).Font.Bold = True: ws.Range("K" & rowIdx).Font.Color = config.ColorWhite
    ws.Range("K" & rowIdx & ":L" & rowIdx).Interior.Color = config.ThemeExp ' 視覺對齊：套用主面板輸出引擎色系
    ws.Range("K" & rowIdx & ":L" & rowIdx).RowHeight = 18: ws.Range("K" & rowIdx).VerticalAlignment = xlVAlignCenter
    rowIdx = rowIdx + 1
    
    imgIdx = 1
    For Each scanWs In config.TargetWB.Worksheets
        If Mod_Rules.IsStandardImgSheet(scanWs.Name, config) Or scanWs.Name = config.SheetNameAllPics Or scanWs.Name = config.SheetNameRescue Then
            ws.Cells(rowIdx, 11).Value = "[" & imgIdx & "] " & scanWs.Name
            ws.Cells(rowIdx, 12).Formula = "=HYPERLINK(""#'" & scanWs.Name & "'!A1"", ""點擊預覽"")"
            
            With ws.Range(ws.Cells(rowIdx, 11), ws.Cells(rowIdx, 12))
                .Font.Name = config.FontMain: .Font.Size = config.FontSizeBase
                .Borders.LineStyle = xlContinuous: .Borders.Color = config.DrawerItemBorder
                .Interior.Color = config.DrawerItemBg
                .WrapText = True ' 允許自然換行不吃字
                .VerticalAlignment = xlVAlignCenter
            End With
            
            ws.Cells(rowIdx, 11).Font.Color = config.ColorTextMain
            ws.Cells(rowIdx, 12).Font.Color = config.DrawerLinkText
            ws.Cells(rowIdx, 12).Font.Bold = True
            ws.Cells(rowIdx, 12).Font.Underline = xlUnderlineStyleNone ' 拔除原生超連結底線
            rowIdx = rowIdx + 1: imgIdx = imgIdx + 1
        End If
    Next scanWs
    
    If imgIdx = 1 Then
        ws.Range("K" & rowIdx & ":L" & rowIdx).Merge
        ws.Range("K" & rowIdx).Value = "(目前尚無圖片歸檔工作表)"
        ws.Range("K" & rowIdx).Font.Color = config.ColorTextHint
        rowIdx = rowIdx + 1
    End If
    
    ws.Range("K" & rowIdx & ":L" & rowIdx).Merge
    ws.Range("K" & rowIdx).Value = "※ 系統提示：建立名稱以 「" & config.PrefixCustom & "」 開頭的工作表，系統將自動為您保留並免於排版清理。"
    ws.Range("K" & rowIdx).Font.Name = config.FontMain: ws.Range("K" & rowIdx).Font.Size = 9: ws.Range("K" & rowIdx).Font.Color = config.ColorTextHint
    ws.Range("K" & rowIdx).WrapText = True
    
    ' 強制執行 AutoFit，確保所有抽屜項目依照文字長度自然展開列高
    ws.Range("K2:L" & rowIdx).Rows.AutoFit
End Sub

Public Sub GoToUI()
    Dim config As cls_Config
    Set config = New cls_Config
    On Error Resume Next: Worksheets(config.SheetNameUI).Activate: On Error GoTo 0
    Set config = Nothing
End Sub

Public Sub GoToManual()
    Dim config As cls_Config
    Set config = New cls_Config
    On Error Resume Next: Worksheets(config.SheetNameManual).Activate: On Error GoTo 0
    Set config = Nothing
End Sub

Private Sub FormatAppButton(shp As Shape, config As cls_Config, txt As String, bg As Long, fg As Long, macro As String, sz As Double, Optional id As String)
    With shp
        If id <> "" Then .Name = id
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginRight = 0: .TextFrame2.MarginTop = 0: .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = txt
        .TextFrame2.TextRange.Font.Name = config.FontMain: .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fg: .TextFrame2.TextRange.Font.Size = sz
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.Solid: .Fill.ForeColor.RGB = bg: .Line.Visible = msoFalse
        .Shadow.Type = msoShadow21: .Shadow.Transparency = 0.8: .Shadow.Blur = 5
        .Placement = xlFreeFloating
        .OnAction = "'" & ThisWorkbook.Name & "'!" & macro
    End With
End Sub

Private Sub FormatNavButton(shp As Shape, config As cls_Config, txt As String, bg As Long, fg As Long, macro As String)
    With shp
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginRight = 0: .TextFrame2.MarginTop = 0: .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = txt
        .TextFrame2.TextRange.Font.Name = config.FontMain: .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fg
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.Solid: .Fill.ForeColor.RGB = bg: .Line.Visible = msoFalse
        .OnAction = "'" & ThisWorkbook.Name & "'!" & macro
    End With
End Sub

Private Sub FormatSectionHeader(shp As Shape, config As cls_Config, txt As String, bgColor As Long)
    With shp
        .TextFrame2.MarginLeft = 5: .TextFrame2.MarginRight = 0: .TextFrame2.MarginTop = 0: .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = txt
        .TextFrame2.TextRange.Font.Name = config.FontMain: .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = config.ColorWhite
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.Solid: .Fill.ForeColor.RGB = bgColor: .Line.Visible = msoFalse
        .Placement = xlFreeFloating
    End With
End Sub

Public Sub RenderSheetNavButtons(ByVal ws As Worksheet, ByVal config As cls_Config)
    If ws.Name = config.SheetNameCatalog Or ws.Name = config.SheetNameAttachment Then Exit Sub
    
    Dim btnExport As Shape, btnHome As Shape
    Dim startX As Double, startY As Double
    Dim strLib As New cls_StringLibrary
    
    On Error Resume Next
    ws.Shapes("Btn_ExportThis").Delete
    ws.Shapes("Btn_GoHome").Delete
    On Error GoTo 0
    
    startX = ws.Range("A2").Left + 5
    startY = ws.Range("A2").Top + 5
    
    Set btnExport = ws.Shapes.AddShape(msoShapeRoundedRectangle, startX, startY, 110, 26)
    With btnExport
        .Name = "Btn_ExportThis"
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginRight = 0: .TextFrame2.MarginTop = 0: .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = strLib.BtnTextExportSheet
        .TextFrame2.TextRange.Font.Name = config.FontMain: .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = config.ColorWhite
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.Solid: .Fill.ForeColor.RGB = config.ThemeManual
        .Line.Visible = msoFalse: .Shadow.Type = msoShadow21: .Placement = xlFreeFloating
        .OnAction = "'" & config.TargetWB.Name & "'!" & "Mod_Main.UI_匯出當前畫布"
        
        On Error Resume Next
        .ControlFormat.PrintObject = False
        On Error GoTo 0
    End With
    
    Set btnHome = ws.Shapes.AddShape(msoShapeRoundedRectangle, startX + 120, startY, 110, 26)
    With btnHome
        .Name = "Btn_GoHome"
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginRight = 0: .TextFrame2.MarginTop = 0: .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = strLib.BtnTextGoHome
        .TextFrame2.TextRange.Font.Name = config.FontMain: .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = config.ColorWhite
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.Solid: .Fill.ForeColor.RGB = config.UIColorNormal
        .Line.Visible = msoFalse: .Shadow.Type = msoShadow21: .Placement = xlFreeFloating
        .OnAction = "'" & config.TargetWB.Name & "'!" & "Mod_UIReset.GoToUI"
        
        On Error Resume Next
        .ControlFormat.PrintObject = False
        On Error GoTo 0
    End With
    Set strLib = Nothing
End Sub
