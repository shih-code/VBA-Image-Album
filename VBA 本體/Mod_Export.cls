Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_Export (標準模組)
' PURPOSE: PDF 報告匯出引擎。
'          實裝「畫布即事實」算繪基準、全靜默黑白名單、宮格渲染引擎與極簡單頁匯出。
' EXPORTS: 執行匯出圖片PDF, 執行單一畫布匯出PDF, 執行匯出橫式目錄PDF, 執行匯出附件目錄PDF
' IMPORTS: cls_ExecutionContext, Mod_Utils, Mod_UIMessenger, Mod_Actions, Mod_Rules, cls_StringLibrary, Mod_UIReset, cls_Log
' FORBIDDEN: 1. 嚴禁將 TEMP 臨時分頁殘留於活頁簿中 (必須確保清理流程完整執行)。
'            2. 嚴禁算繪時逆向依賴硬碟實體檔案 (必須以 Excel 畫布 Shape 為唯一依據)。
' DEPENDENCIES: Microsoft Scripting Runtime (FileSystemObject)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_Export"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
' 區塊 1：宮格圖片匯出引擎 (Grid Export)
' ==============================================================================

' ------------------------------------------------------------------------------
' 程式名稱：執行匯出圖片PDF
' 目的：依據選擇模式，將畫布圖片或實體圖檔，以不扭曲、精準貼合 A4 宮格的方式匯出 PDF。
' 修正重點：
'   1. 【格線固定 + 等比例適應】：算出原圖長寬比，嚴格同比例縮放置中，絕不改變圖框大小。
'   2. 【A4 頁面精算】：重算 2/4/6/8 宮格高度，確保剛好一頁，不被 Excel 自動裁切分頁。
' 本 Sub 現為協調者(Orchestrator)：僅負責流程順序，實際邏輯下沉至本區塊末端的
' 私有輔助函式（SRP 拆分，ACDS 第八章／第二章）。對外簽章與行為皆未變動。
' ------------------------------------------------------------------------------
Public Sub 執行匯出圖片PDF(ByVal Context As cls_ExecutionContext)
    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True

   ' 【熔斷網提前】：旗標鎖上後立刻生效，避免後續初始化拋錯導致 isRunning 卡死 True
   On Error GoTo ExportErrorHandler

    Dim config As cls_Config: Set config = Context.config
    Dim fso As Object, wb As Workbook
    Dim targetSheets As Collection, ws As Worksheet
    Dim wsPdf As Worksheet
    Dim pdfPath As String, picExportDir As String
    Dim origScreen As Boolean, origEvents As Boolean, origCalc As XlCalculation
    Dim isStateSaved As Boolean, wasProtected As Boolean
    Dim strLib As New cls_StringLibrary

    ' 【垃圾車】：收集所有產生過的暫存檔路徑，最後集中銷毀
    Dim cleanupFiles As Collection: Set cleanupFiles = New Collection

    ' 【快取字典】：快取 Catalog 登記之「匯入時間」與「硬碟原始路徑」
    Dim dictTimes As Object: Set dictTimes = CreateObject("Scripting.Dictionary")
    Dim dictDiskPaths As Object: Set dictDiskPaths = CreateObject("Scripting.Dictionary")

    ' 模式選擇彈窗 (是 = 畫布優先 / 否 = 資料夾直匯)
    Dim isCanvasPriority As Boolean
    isCanvasPriority = AskExportMode(strLib)
    Context.LogStore.Record "TX_START", "EXPORT", "開始匯出圖片PDF | 模式：" & IIf(isCanvasPriority, "畫布優先", "資料夾直匯")

    isStateSaved = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wb = config.TargetWB

    ' 1. 取得過濾後的工作表集合
    Set targetSheets = GetTargetSheetsBySilentPipeline(Context)
    If targetSheets.count = 0 Then
        Call Mod_UIMessenger.ShowWarning(strLib.ErrNoValidPicToExport, strLib.TitleSysWarning)
        GoTo ExportCleanup
    End If

    Call Mod_Actions.InitializeSystemFolders(config)
    picExportDir = fso.BuildPath(config.IOExportFolder, "圖片")
    Call Mod_Utils.EnsureFolderExists(fso, picExportDir)

    ' 2. 檔名與宮格數解析
    Dim layout As Integer: layout = Context.DynamicPdfGrid
    If layout <> 2 And layout <> 4 And layout <> 6 And layout <> 8 Then layout = 4
    pdfPath = ResolveExportPdfPath(fso, picExportDir, Context, layout, isCanvasPriority)

    ' 3. 凍結畫布與環境防護罩
    Call FreezeExportEnvironment(wb, config, origScreen, origEvents, origCalc, wasProtected)
    isStateSaved = True

    ' 建立 Catalog 資訊快取
    Call BuildCatalogCache(wb, config, dictTimes, dictDiskPaths)

    ' 清除舊 TEMP 表
    Call PurgeOldTempSheets(wb)

    ' 4. 建立臨時工作表 TEMP_PDF_GRID
    Dim textH As Double, imgH As Double, gapH As Double, numCols As Integer
    Call GetGridLayoutMetrics(layout, textH, imgH, gapH, numCols)
    Set wsPdf = CreateGridTempSheet(wb, config, Context, numCols)

    ' 5. 掃描目標表，填入宮格
    Dim itemIdx As Long
    itemIdx = 0
    For Each ws In targetSheets
        Dim sortedShapes As Collection
        Set sortedShapes = CollectSortedPictureShapes(ws)

        If sortedShapes.count > 0 Then
            Call RenderPicturesForSheet(ws, wsPdf, sortedShapes, dictTimes, dictDiskPaths, isCanvasPriority, fso, layout, numCols, textH, imgH, gapH, config, Context, itemIdx, cleanupFiles)
        End If
    Next ws

    If itemIdx = 0 Then
        Call Mod_UIMessenger.ShowWarning(strLib.ErrNoValidPicToExport, strLib.TitleSysWarning)
        GoTo ExportCleanup
    End If

    ' 輸出實體 PDF 檔前開啟畫面更新
    Application.ScreenUpdating = True
    Application.StatusBar = "EXPORT: 正在輸出 PDF 實體檔案..."
    wsPdf.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Context.LogStore.Record "TX_COMMIT_FS", "EXPORT", "圖片PDF匯出完成 | 共 " & itemIdx & " 張 | 路徑：" & pdfPath


    Call Mod_UIMessenger.ShowInfo(strLib.InfoBatchExportSuccess(pdfPath), strLib.TitleComplete)

ExportCleanup:
    On Error Resume Next

    Call PurgeCleanupFiles(fso, cleanupFiles)
    Call RemoveTempGridSheet(wsPdf)
    Call RestoreExportEnvironment(wb, config, wasProtected, isStateSaved, origScreen, origEvents, origCalc)

    Application.StatusBar = False
    Call Mod_UIReset.GoToUI
    Set fso = Nothing: Set targetSheets = Nothing: Set ws = Nothing: Set wsPdf = Nothing
    Set dictTimes = Nothing: Set dictDiskPaths = Nothing: Set strLib = Nothing
    isRunning = False
    Exit Sub

ExportErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    If Err.Number = 1004 Or InStr(errDesc, "1004") > 0 Then errDesc = "PDF 檔案可能正被開啟中，請先關閉舊檔案後重試。"
    Context.LogStore.Record "TX_ERROR", "EXPORT", "圖片PDF匯出失敗：" & errDesc
    Call Mod_UIMessenger.ShowError(strLib.ErrExportPic(errDesc), strLib.TitleSysWarning)
    Resume ExportCleanup
End Sub

' ------------------------------------------------------------------------------
' 以下為 執行匯出圖片PDF 的 SRP 拆分區：每支函式只做一件事，
' 依第八章 100-150 行安全線與第二章 SRP 原則拆出，全部維持 Private，
' 不影響本模組 EXPORTS 對外清單。
' ------------------------------------------------------------------------------

' 職責：詢問使用者要「畫布優先」還是「資料夾直匯」
Private Function AskExportMode(ByVal strLib As cls_StringLibrary) As Boolean
    Dim userChoice As VbMsgBoxResult
    userChoice = Mod_UIMessenger.AskQuestion(strLib.PromptSelectExportMode, strLib.TitleSelectExportMode)
    AskExportMode = (userChoice = vbYes)
End Function

' 職責：算出本次匯出的 PDF 檔名，含撞名流水號防覆蓋
Private Function ResolveExportPdfPath(ByVal fso As Object, ByVal picExportDir As String, ByVal Context As cls_ExecutionContext, ByVal layout As Integer, ByVal isCanvasPriority As Boolean) As String
    Dim basePdfName As String, pdfPath As String, counter As Long
    Dim modeTag As String: modeTag = IIf(isCanvasPriority, "畫布版", "原始檔版")

    basePdfName = fso.BuildPath(picExportDir, Context.DynamicPdfImgName & "_" & layout & "宮格_" & modeTag)
    pdfPath = basePdfName & ".pdf"
    counter = 1
    Do While fso.FileExists(pdfPath)
        pdfPath = basePdfName & "_" & counter & ".pdf"
        counter = counter + 1
    Loop

    ResolveExportPdfPath = pdfPath
End Function

' 職責：凍結畫面更新/事件/計算並解鎖工作簿結構，回傳還原用的原始狀態
Private Sub FreezeExportEnvironment(ByVal wb As Workbook, ByVal config As cls_Config, ByRef origScreen As Boolean, ByRef origEvents As Boolean, ByRef origCalc As XlCalculation, ByRef wasProtected As Boolean)
    origScreen = Application.ScreenUpdating
    origEvents = Application.EnableEvents
    origCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "EXPORT: 正在準備圖片版面渲染..."

    wasProtected = wb.ProtectStructure
    If wasProtected Then wb.Unprotect Password:=config.SecurityPassword
End Sub

' 職責：還原工作簿保護狀態與畫面/事件/計算模式（僅在 FreezeExportEnvironment 已執行過才還原）
Private Sub RestoreExportEnvironment(ByVal wb As Workbook, ByVal config As cls_Config, ByVal wasProtected As Boolean, ByVal isStateSaved As Boolean, ByVal origScreen As Boolean, ByVal origEvents As Boolean, ByVal origCalc As XlCalculation)
    On Error Resume Next
    If wasProtected And Not wb Is Nothing Then wb.Protect Password:=config.SecurityPassword, Structure:=True, Windows:=False
    If isStateSaved Then
        Application.ScreenUpdating = origScreen
        Application.EnableEvents = origEvents
        If origCalc <> 0 Then Application.Calculation = origCalc
    End If
End Sub

' 職責：從 Catalog 分頁讀出「匯入時間」與「硬碟原始路徑」快取字典
Private Sub BuildCatalogCache(ByVal wb As Workbook, ByVal config As cls_Config, ByRef dictTimes As Object, ByRef dictDiskPaths As Object)
    Dim wsCat As Worksheet
    Dim lastRow As Long, r As Long, cName As String, cPath As String

    Set wsCat = Mod_Utils.GetSheetSafe(wb, config.SheetNameCatalog)
    If wsCat Is Nothing Then Exit Sub

    lastRow = wsCat.Cells(wsCat.Rows.count, "A").End(xlUp).row
    For r = 5 To lastRow
        cName = Trim(CStr(wsCat.Cells(r, 1).Value))
        cPath = Trim(CStr(wsCat.Cells(r, 3).Value))
        If cName <> "" Then
            dictTimes(cName) = Trim(CStr(wsCat.Cells(r, 5).Value))
            If cPath <> "" Then dictDiskPaths(cName) = cPath
        End If
    Next r
End Sub

' 職責：刪除上次執行殘留的 TEMP_PDF 系列工作表
Private Sub PurgeOldTempSheets(ByVal wb As Workbook)
    Dim wsTemp As Worksheet
    Application.DisplayAlerts = False
    For Each wsTemp In wb.Worksheets
        If InStr(1, wsTemp.Name, "TEMP_PDF", vbTextCompare) > 0 Then wsTemp.Delete
    Next wsTemp
    Application.DisplayAlerts = True
End Sub

' 職責：依宮格數決定文字列高/圖片列高/間距/欄數這四個版面參數
Private Sub GetGridLayoutMetrics(ByVal layout As Integer, ByRef textH As Double, ByRef imgH As Double, ByRef gapH As Double, ByRef numCols As Integer)
    If layout = 2 Then
        textH = 30: imgH = 320: gapH = 15: numCols = 1
    ElseIf layout = 4 Then
        textH = 25: imgH = 320: gapH = 15: numCols = 2
    ElseIf layout = 6 Then
        textH = 22: imgH = 210: gapH = 10: numCols = 2
    Else
        textH = 18: imgH = 155: gapH = 7: numCols = 2
    End If
End Sub

' 職責：建立 TEMP_PDF_GRID 臨時工作表，套用 A4 版面設定、欄寬與標題列
Private Function CreateGridTempSheet(ByVal wb As Workbook, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal numCols As Integer) As Worksheet
    Dim wsPdf As Worksheet
    Set wsPdf = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
    wsPdf.Name = "TEMP_PDF_GRID"

    With wsPdf.PageSetup
        .PaperSize = xlPaperA4: .Orientation = xlPortrait
        .LeftMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .RightMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .TopMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .BottomMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .CenterHeader = "": .CenterFooter = "&P": .PrintTitleRows = ""
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .CenterHorizontally = True: .CenterVertically = False
    End With

    wsPdf.Cells.Font.Name = config.FontMain: wsPdf.Cells.Interior.Color = config.ColorWhite
    ActiveWindow.DisplayGridlines = False

    If numCols = 1 Then
        wsPdf.Columns("A:A").ColumnWidth = 78: wsPdf.Range("A1:A1").Merge
    Else
        wsPdf.Columns("A:A").ColumnWidth = 38: wsPdf.Columns("B:B").ColumnWidth = 2
        wsPdf.Columns("C:C").ColumnWidth = 38: wsPdf.Range("A1:C1").Merge
    End If

    wsPdf.Range("A1").Value = "■ 影像歸檔報告 - " & Context.DynamicProjName
    wsPdf.Range("A1").Font.Size = config.FontSizeTitle: wsPdf.Range("A1").Font.Bold = True
    wsPdf.Range("A1").HorizontalAlignment = xlCenter: wsPdf.Range("A1").VerticalAlignment = xlCenter
    wsPdf.Rows("1:1").RowHeight = 30: wsPdf.Rows("2:2").RowHeight = 5

    Set CreateGridTempSheet = wsPdf
End Function

' 職責：收集某工作表內所有圖片 Shape，並依左上到右下的視覺順序排序後回傳
Private Function CollectSortedPictureShapes(ByVal ws As Worksheet) As Collection
    Dim shp As Shape
    Dim nPic As Long: nPic = 0
    Dim result As New Collection

    For Each shp In ws.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then nPic = nPic + 1
    Next shp

    If nPic = 0 Then
        Set CollectSortedPictureShapes = result
        Exit Function
    End If

    Dim shpArr() As Shape, shpMetric() As Double
    ReDim shpArr(1 To nPic): ReDim shpMetric(1 To nPic)
    Dim picCount As Long: picCount = 0
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            picCount = picCount + 1
            Set shpArr(picCount) = shp
            shpMetric(picCount) = (shp.Top * 10000) + shp.Left
        End If
    Next shp

    ' 視覺排序（左上到右下）
    Dim i As Long, j As Long, ts As Shape, tM As Double
    For i = 1 To nPic - 1
        For j = 1 To nPic - i
            If shpMetric(j) > shpMetric(j + 1) Then
                Set ts = shpArr(j): tM = shpMetric(j)
                Set shpArr(j) = shpArr(j + 1): shpMetric(j) = shpMetric(j + 1)
                Set shpArr(j + 1) = ts: shpMetric(j + 1) = tM
            End If
        Next j
    Next i

    For i = 1 To nPic
        result.Add shpArr(i)
    Next i
    Set CollectSortedPictureShapes = result
End Function

' 職責：依目前第幾張圖片與欄數，算出應落在的欄列索引
Private Sub ResolveGridCellPosition(ByVal itemIdx As Long, ByVal numCols As Integer, ByRef colIdx As Long, ByRef rowIdx As Long)
    If numCols = 1 Then
        colIdx = 1: rowIdx = 3 + (itemIdx - 1) * 3
    Else
        If itemIdx Mod 2 <> 0 Then colIdx = 1 Else colIdx = 3
        rowIdx = 3 + (Int((itemIdx - 1) / 2) * 3)
    End If
End Sub

' 職責：畫出單一宮格的文字標題列與圖框邊界（不含圖片本身）
Private Sub RenderGridCellFrame(ByVal wsPdf As Worksheet, ByVal rowIdx As Long, ByVal colIdx As Long, ByVal sName As String, ByVal sTime As String, ByVal layout As Integer, ByVal textH As Double, ByVal imgH As Double, ByVal gapH As Double, ByVal config As cls_Config)
    With wsPdf.Cells(rowIdx, colIdx)
        .Value = " ■ 名稱: " & sName & vbCrLf & " ■ 匯入: " & sTime
        .Interior.Color = RGB(242, 244, 247): .Font.Color = RGB(50, 60, 70)
        .Font.Size = IIf(layout = 8, 9, 10): .Font.Bold = True
        .VerticalAlignment = xlCenter: .HorizontalAlignment = xlLeft
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(200, 200, 200)
    End With
    wsPdf.Rows(rowIdx).RowHeight = textH

    With wsPdf.Cells(rowIdx + 1, colIdx)
        .Interior.Color = config.ColorWhite
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(200, 200, 200)
    End With
    wsPdf.Rows(rowIdx + 1).RowHeight = imgH
    wsPdf.Rows(rowIdx + 2).RowHeight = gapH
End Sub

' 職責：依 isCanvasPriority 決定的優先順序，嘗試「剪貼簿→暫存檔→硬碟路徑」
'       （資料夾直匯模式則為「硬碟路徑→剪貼簿」）三層降級貼圖，回傳成功貼上的 Shape
Private Function TryPastePictureWithFallback(ByVal wsPdf As Worksheet, ByVal ws As Worksheet, ByVal shp As Shape, ByVal sName As String, ByVal isCanvasPriority As Boolean, ByVal dictDiskPaths As Object, ByVal fso As Object, ByVal rowIdx As Long, ByVal colIdx As Long, ByVal cellLeft As Double, ByVal cellTop As Double, ByRef cleanupFiles As Collection) As Shape
    Dim pastedShp As Shape, isSuccess As Boolean
    Dim tmpImgPath As String, diskFilePath As String
    Dim shapesCountBefore As Long

    isSuccess = False

    If isCanvasPriority Then
        ' ==========================================================
        ' 【靜默渲染機制】：無畫面切換、無 Select，全程背景處理
        ' ==========================================================
        ' 1. 嘗試直接剪貼簿複製貼上 (無 Select/Activate，不閃爍)
        shapesCountBefore = wsPdf.Shapes.count

        On Error Resume Next
        shp.Copy
        wsPdf.Paste Destination:=wsPdf.Cells(rowIdx + 1, colIdx)

        If Err.Number = 0 And wsPdf.Shapes.count > shapesCountBefore Then
            Set pastedShp = wsPdf.Shapes(wsPdf.Shapes.count)
            isSuccess = True
        End If
        Err.Clear
        On Error GoTo 0

        ' 【每次剪貼簿操作後立即清空】：畫布優先模式反覆佔用系統剪貼簿，
        ' 是唯一會導致Excel本身不穩定甚至自行關閉的模式，不能等到60張才清一次
        Application.CutCopyMode = False

        ' 2. 若剪貼簿失敗，降流至暫存檔轉繪
        If Not isSuccess Then
            tmpImgPath = Mod_Utils.ExportShapeToTempFile(ws, shp)
            If tmpImgPath <> "" And fso.FileExists(tmpImgPath) Then
                Set pastedShp = wsPdf.Shapes.AddPicture(fileName:=tmpImgPath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                                Left:=cellLeft, Top:=cellTop, Width:=-1, Height:=-1)
                cleanupFiles.Add tmpImgPath
                isSuccess = True
            End If
        End If

        ' 3. 若依然失敗，降流至硬碟實體檔
        If Not isSuccess And dictDiskPaths.Exists(sName) Then
            diskFilePath = dictDiskPaths(sName)
            If fso.FileExists(diskFilePath) Then
                Set pastedShp = wsPdf.Shapes.AddPicture(fileName:=diskFilePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                                Left:=cellLeft, Top:=cellTop, Width:=-1, Height:=-1)
                isSuccess = True
            End If
        End If
    Else
        ' 【資料夾直匯模式】優先讀取實體檔
        If dictDiskPaths.Exists(sName) Then
            diskFilePath = dictDiskPaths(sName)
            If fso.FileExists(diskFilePath) Then
                Set pastedShp = wsPdf.Shapes.AddPicture(fileName:=diskFilePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                                Left:=cellLeft, Top:=cellTop, Width:=-1, Height:=-1)
                isSuccess = True
            End If
        End If

        ' 降流至剪貼簿
        If Not isSuccess Then
            shapesCountBefore = wsPdf.Shapes.count

            On Error Resume Next
            shp.Copy
            wsPdf.Paste Destination:=wsPdf.Cells(rowIdx + 1, colIdx)
            If Err.Number = 0 And wsPdf.Shapes.count > shapesCountBefore Then
                Set pastedShp = wsPdf.Shapes(wsPdf.Shapes.count)
                isSuccess = True
            End If
            Err.Clear
            On Error GoTo 0

            ' 【資源釋放】：這裡同樣經過剪貼簿，操作完立刻清空
            Application.CutCopyMode = False
        End If
    End If

    Set TryPastePictureWithFallback = pastedShp
End Function

' 職責：將貼上的 Shape 依原始長寬比等比例縮放，並在宮格內置中
Private Sub ScaleAndCenterPastedShape(ByVal pastedShp As Shape, ByVal wsPdf As Worksheet, ByVal rowIdx As Long, ByVal colIdx As Long, ByVal targetW As Double, ByVal targetH As Double)
    Dim origW As Double, origH As Double, scaleRatio As Double

    ' 重置真實尺寸
    pastedShp.ScaleWidth 1#, msoTrue, msoScaleFromTopLeft
    pastedShp.ScaleHeight 1#, msoTrue, msoScaleFromTopLeft
    origW = pastedShp.Width
    origH = pastedShp.Height

    ' 算同比例縮放因子
    If (targetW / origW) < (targetH / origH) Then
        scaleRatio = targetW / origW
    Else
        scaleRatio = targetH / origH
    End If

    ' 套用尺寸
    pastedShp.LockAspectRatio = msoFalse
    pastedShp.Width = origW * scaleRatio
    pastedShp.Height = origH * scaleRatio
    pastedShp.LockAspectRatio = msoTrue

    ' 格子內精準置中
    pastedShp.Left = wsPdf.Cells(rowIdx + 1, colIdx).Left + (wsPdf.Cells(rowIdx + 1, colIdx).Width - pastedShp.Width) / 2
    pastedShp.Top = wsPdf.Cells(rowIdx + 1, colIdx).Top + (wsPdf.Cells(rowIdx + 1, colIdx).Height - pastedShp.Height) / 2
End Sub

' 職責：把某一張來源工作表裡已排序好的圖片，逐張渲染進宮格頁面。
'       itemIdx 是跨多張來源工作表累加的全域計數器，以 ByRef 傳入延續累加；
'       本函式自帶單張失敗即跳過的錯誤處理（Fail-Fast 但不中斷整體匯出）。
Private Sub RenderPicturesForSheet(ByVal ws As Worksheet, ByVal wsPdf As Worksheet, ByVal sortedShapes As Collection, ByVal dictTimes As Object, ByVal dictDiskPaths As Object, ByVal isCanvasPriority As Boolean, ByVal fso As Object, ByVal layout As Integer, ByVal numCols As Integer, ByVal textH As Double, ByVal imgH As Double, ByVal gapH As Double, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByRef itemIdx As Long, ByRef cleanupFiles As Collection)
    Dim shp As Shape, sName As String, sTime As String
    Dim colIdx As Long, rowIdx As Long
    Dim targetW As Double, targetH As Double, cellLeft As Double, cellTop As Double
    Dim pastedShp As Shape

    For Each shp In sortedShapes
        On Error GoTo ItemErrorHandler
        sName = shp.Name: sTime = "未知"
        If dictTimes.Exists(sName) Then sTime = dictTimes(sName)

        itemIdx = itemIdx + 1
        Application.StatusBar = "EXPORT: 正在渲染圖片 (" & itemIdx & ")"

        ' 【批次防線】：對齊全系統既有節奏，每20張回報進度、每60張存檔+釋放資源
        If itemIdx Mod 20 = 0 Then
            Context.LogStore.Record "TX_PROGRESS", "EXPORT", "宮格匯出進度：" & itemIdx
            DoEvents
        End If
        If itemIdx Mod 60 = 0 Then config.TargetWB.Save

        ' 計算欄列索引
        Call ResolveGridCellPosition(itemIdx, numCols, colIdx, rowIdx)

        ' 繪製文字標題格與圖框邊界
        Call RenderGridCellFrame(wsPdf, rowIdx, colIdx, sName, sTime, layout, textH, imgH, gapH, config)

        ' 格子尺寸限制與目標座標
        targetW = wsPdf.Cells(rowIdx + 1, colIdx).Width - 8
        targetH = wsPdf.Cells(rowIdx + 1, colIdx).Height - 8
        cellLeft = wsPdf.Cells(rowIdx + 1, colIdx).Left + 4
        cellTop = wsPdf.Cells(rowIdx + 1, colIdx).Top + 4

        ' 【靜默渲染機制】：無畫面切換、無 Select，全程背景處理，三層降級貼圖
        Set pastedShp = TryPastePictureWithFallback(wsPdf, ws, shp, sName, isCanvasPriority, dictDiskPaths, fso, rowIdx, colIdx, cellLeft, cellTop, cleanupFiles)

        ' 【不變形同比例縮放】
        If Not pastedShp Is Nothing Then
            Call ScaleAndCenterPastedShape(pastedShp, wsPdf, rowIdx, colIdx, targetW, targetH)
        End If

        ' 宮格分頁控制
        If itemIdx Mod layout = 0 Then wsPdf.HPageBreaks.Add Before:=wsPdf.Rows(rowIdx + 3)

ItemNextPic:
    Next shp
    Exit Sub

ItemErrorHandler:
    Context.LogStore.Record "TX_ERROR", "EXPORT", "宮格匯出單張圖片失敗，已跳過：" & sName & " | 原因：" & Err.Description
    Err.Clear
    Resume ItemNextPic
End Sub

' 職責：銷毀本次匯出過程中產生的暫存圖檔（Cleanup Label 的其中一段）
Private Sub PurgeCleanupFiles(ByVal fso As Object, ByVal cleanupFiles As Collection)
    On Error Resume Next
    If cleanupFiles Is Nothing Then Exit Sub
    Dim fPath As Variant
    For Each fPath In cleanupFiles
        If fso.FileExists(fPath) Then fso.DeleteFile fPath, True
    Next fPath
End Sub

' 職責：刪除本次建立的 TEMP_PDF_GRID 臨時工作表（Cleanup Label 的其中一段）
Private Sub RemoveTempGridSheet(ByVal wsPdf As Worksheet)
    On Error Resume Next
    If Not wsPdf Is Nothing Then
        Application.DisplayAlerts = False: wsPdf.Delete: Application.DisplayAlerts = True
    End If
End Sub

' ==============================================================================
' 區塊 2：單頁畫布與目錄報表匯出引擎 (Single Sheet & Catalog Export)
' ==============================================================================

' ------------------------------------------------------------------------------
' 程式名稱：執行單一畫布匯出PDF (就地圖章專用)
' ------------------------------------------------------------------------------
Public Sub 執行單一畫布匯出PDF(ByVal ws As Worksheet, ByVal config As cls_Config)
    Dim fso As Object, pdfPath As String, origScreen As Boolean
    Dim strLib As New cls_StringLibrary
    Dim exportLogStore As New cls_Log
    Call exportLogStore.Initialize(config, config.TargetWB, False)
    
    On Error GoTo SingleExportErr
    Set fso = CreateObject("Scripting.FileSystemObject")
    exportLogStore.Record "TX_START", "EXPORT", "開始匯出單一畫布PDF：" & ws.Name
    
    Call Mod_Actions.InitializeSystemFolders(config)
    
    Dim picExportDir As String: picExportDir = fso.BuildPath(config.IOExportFolder, "圖片")
    Call Mod_Utils.EnsureFolderExists(fso, picExportDir)
    
    pdfPath = fso.BuildPath(picExportDir, Replace(ws.Name, config.PrefixCustom, "") & "_專屬報告.pdf")
    
    If Mod_Utils.IsFileLocked(fso, pdfPath) Then
        Call Mod_UIMessenger.ShowError(strLib.ErrFileLocked(fso.GetFileName(pdfPath)), strLib.TitleSysWarning)
        Set fso = Nothing: Set strLib = Nothing: Exit Sub
    End If
    
    origScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Call Mod_Utils.SafeUnprotect(ws, config.SecurityPassword)
    Call ApplyCustomSheetPagination(ws, config)
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
                            
    If Left(ws.Name, Len(config.PrefixCustom)) <> config.PrefixCustom Then Call Mod_Utils.SafeProtect(ws, config.SecurityPassword)
    
    Application.ScreenUpdating = origScreen
    exportLogStore.Record "TX_COMMIT_FS", "EXPORT", "單一畫布PDF匯出完成：" & pdfPath
    Call Mod_UIMessenger.ShowInfo(strLib.InfoSingleExportSuccess(Replace(ws.Name, config.PrefixCustom, "")), strLib.TitleComplete)

SingleExportCleanup:
    Set fso = Nothing: Set strLib = Nothing
    Exit Sub

SingleExportErr:
    Dim errDesc As String: errDesc = Err.Description
    If Err.Number = 1004 Or InStr(errDesc, "1004") > 0 Then errDesc = "PDF 檔案可能正被佔用，請關閉後重試。"
    Application.ScreenUpdating = origScreen
    exportLogStore.Record "TX_ERROR", "EXPORT", "單一畫布PDF匯出失敗：" & ws.Name & " | " & errDesc
    Call Mod_UIMessenger.ShowError(strLib.ErrSingleExport(errDesc), strLib.TitleSysWarning)
    Resume SingleExportCleanup
End Sub

' ------------------------------------------------------------------------------
' [輔助引擎]：ApplyCustomSheetPagination (動態偵測列印範圍)
' ------------------------------------------------------------------------------
Private Sub ApplyCustomSheetPagination(ByVal ws As Worksheet, ByVal config As cls_Config)
    On Error Resume Next
    ws.ResetAllPageBreaks
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Borders.LineStyle = xlNone
    
    ' 【動態範圍偵測】：動態找出畫布上真正的最右下角 (包含 Shape 所在位置)
    Dim maxRow As Long, maxCol As Long
    Dim rngLastRow As Range, rngLastCol As Range
    
    Set rngLastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set rngLastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    If Not rngLastRow Is Nothing Then maxRow = rngLastRow.row Else maxRow = 1
    If Not rngLastCol Is Nothing Then maxCol = rngLastCol.Column Else maxCol = 1
    
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Visible = msoTrue And shp.Name <> "Btn_ExportThis" And shp.Name <> "Btn_GoHome" Then
            If shp.BottomRightCell.row > maxRow Then maxRow = shp.BottomRightCell.row
            If shp.BottomRightCell.Column > maxCol Then maxCol = shp.BottomRightCell.Column
        End If
    Next shp
    
    maxRow = maxRow + 2: maxCol = maxCol + 1
    Dim colLetter As String: colLetter = Split(ws.Cells(1, maxCol).Address, "$")(1)
    
    With ws.PageSetup
        .PrintArea = "$A$1:$" & colLetter & "$" & maxRow
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .CenterHorizontally = True: .CenterVertically = False
        .TopMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .BottomMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .LeftMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .RightMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .CenterHeader = "": .CenterFooter = "&P": .PrintTitleRows = "": .PrintGridlines = False
    End With
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' [輔助引擎]：GetTargetSheetsBySilentPipeline (白名單/黑名單解析)
' ------------------------------------------------------------------------------
Private Function GetTargetSheetsBySilentPipeline(ByVal Context As cls_ExecutionContext) As Collection
    Dim config As cls_Config: Set config = Context.config
    Dim colAllImgSheets As New Collection
    Dim colFiltered As New Collection
    Dim ws As Worksheet, filterMode As String, customRange As String, menuID As Long
    
    For Each ws In config.TargetWB.Worksheets
        If Mod_Rules.IsStandardImgSheet(ws.Name, config) Or ws.Name = config.SheetNameAllPics Or ws.Name = config.SheetNameRescue Then
            colAllImgSheets.Add ws
        End If
    Next ws
    
    If colAllImgSheets.count = 0 Then
        Set GetTargetSheetsBySilentPipeline = colFiltered: Exit Function
    End If
    
    filterMode = Context.DynamicExportFilter
    customRange = Context.DynamicExportRange
    
    If filterMode = "全額匯出" Or Trim(customRange) = "" Then
        Set GetTargetSheetsBySilentPipeline = colAllImgSheets: Exit Function
    End If
    
    Dim parseArr() As String, k As Long, token As String
    Dim dictUserNumbers As Object: Set dictUserNumbers = CreateObject("Scripting.Dictionary")
    
    parseArr = Split(customRange, ",")
    For k = LBound(parseArr) To UBound(parseArr)
        token = Trim(parseArr(k))
        If InStr(token, "-") > 0 Then
            Dim rangeParts() As String: rangeParts = Split(token, "-")
            If UBound(rangeParts) = 1 Then
                If IsNumeric(Trim(rangeParts(0))) And IsNumeric(Trim(rangeParts(1))) Then
                    Dim startIdx As Long, endIdx As Long
                    startIdx = CLng(Trim(rangeParts(0))): endIdx = CLng(Trim(rangeParts(1)))
                    If startIdx > endIdx Then
                        Dim tmpVal As Long: tmpVal = startIdx: startIdx = endIdx: endIdx = tmpVal
                    End If
                    Dim rIdx As Long
                    For rIdx = startIdx To endIdx
                        If rIdx >= 1 And rIdx <= colAllImgSheets.count Then dictUserNumbers(rIdx) = True
                    Next rIdx
                End If
            End If
        ElseIf IsNumeric(token) Then
            Dim numVal As Long: numVal = CLng(token)
            If numVal >= 1 And numVal <= colAllImgSheets.count Then dictUserNumbers(numVal) = True
        End If
    Next k
    
    Dim shouldExport As Boolean
    For menuID = 1 To colAllImgSheets.count
        If dictUserNumbers.Exists(menuID) Then
            shouldExport = (filterMode = "白名單")
        Else
            shouldExport = (filterMode = "黑名單")
        End If
        If shouldExport Then colFiltered.Add colAllImgSheets(menuID)
    Next menuID
    
    Set dictUserNumbers = Nothing
    Set GetTargetSheetsBySilentPipeline = colFiltered
End Function

' ------------------------------------------------------------------------------
' 程式名稱：執行匯出橫式目錄PDF / 執行匯出附件目錄PDF
' 目的：生成 A4 橫式的歸檔報表。同樣採用臨時工作表隔離處理，匯出後即刪除。
' ------------------------------------------------------------------------------
Public Sub 執行匯出橫式目錄PDF(ByVal Context As cls_ExecutionContext)
    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True

    Dim config As cls_Config: Set config = Context.config
    Dim wb As Workbook, wsCat As Worksheet, wsPdf As Worksheet
    Dim pdfPath As String, catExportDir As String
    Dim fso As Object, lastRow As Long
    Dim origEvents As Boolean, origScreen As Boolean, wasProtected As Boolean
    Dim strLib As New cls_StringLibrary
    
    On Error GoTo ErrCatExport
    Set wb = config.TargetWB
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    origEvents = Application.EnableEvents: origScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False: Application.EnableEvents = False
    
    wasProtected = wb.ProtectStructure
    If wasProtected Then wb.Unprotect Password:=config.SecurityPassword
    Context.LogStore.Record "TX_START", "EXPORT", "開始匯出橫式目錄PDF"
    
    On Error Resume Next
    Set wsCat = Mod_Utils.GetSheetSafe(wb, config.SheetNameCatalog)
    On Error GoTo ErrCatExport
    
    If wsCat Is Nothing Then
        Call Mod_UIMessenger.ShowWarning(strLib.ErrNoCatalog, strLib.TitleSysWarning)
        GoTo ExpCatExit
    End If
    
    lastRow = wsCat.Cells(wsCat.Rows.count, "A").End(xlUp).row
    If lastRow < 5 Then
        Call Mod_UIMessenger.ShowWarning(strLib.ErrNoCatRecord, strLib.TitleSysWarning)
        GoTo ExpCatExit
    End If
    
    Application.StatusBar = "EXPORT: 正在生成目錄 PDF..."
    Application.DisplayAlerts = False
    Dim wsTemp As Worksheet
    For Each wsTemp In wb.Worksheets
        If InStr(1, wsTemp.Name, "TEMP_PDF_CAT", vbTextCompare) > 0 Then wsTemp.Delete
    Next wsTemp
    Application.DisplayAlerts = True
    
    wsCat.Copy After:=wb.Sheets(wb.Sheets.count)
    Set wsPdf = wb.Sheets(wb.Sheets.count)
    wsPdf.Name = "TEMP_PDF_CAT"
    Call Mod_Utils.SafeUnprotect(wsPdf, config.SecurityPassword)
    
    wsPdf.Columns("I:K").Delete: wsPdf.Rows("1:3").Delete
    wsPdf.Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    wsPdf.Range("A1").Value = "■ 專案圖片歸檔目錄 - " & Context.DynamicProjName
    wsPdf.Range("A1").Font.Name = config.FontMain: wsPdf.Range("A1").Font.Size = config.FontSizeSub: wsPdf.Range("A1").Font.Bold = True
    wsPdf.Range("A1:H1").Merge
    wsPdf.Rows("1:1").RowHeight = 35: wsPdf.Rows("2:2").RowHeight = 10
    
    With wsPdf.PageSetup
        .PaperSize = xlPaperA4: .Orientation = xlLandscape
        .LeftMargin = Application.CentimetersToPoints(config.PDFMarginCM): .RightMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .TopMargin = Application.CentimetersToPoints(config.PDFMarginCM): .BottomMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .CenterHeader = "": .CenterFooter = "&P": .PrintTitleRows = ""
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
    End With
    
    wsPdf.Cells.Font.Name = config.FontMain: wsPdf.Cells.Font.Size = 10
    wsPdf.Rows("3:3").Font.Size = 11: wsPdf.Rows("3:3").Font.Bold = True
    
    wsPdf.Columns("A:H").AutoFit
    If wsPdf.Columns("C:C").ColumnWidth > 35 Then wsPdf.Columns("C:C").ColumnWidth = 35
    If wsPdf.Columns("D:D").ColumnWidth > 35 Then wsPdf.Columns("D:D").ColumnWidth = 35
    wsPdf.Columns("C:D").WrapText = True: wsPdf.Rows.AutoFit
    
    Call Mod_Actions.InitializeSystemFolders(config)
    catExportDir = fso.BuildPath(config.IOExportFolder, "目錄")
    Call Mod_Utils.EnsureFolderExists(fso, catExportDir)
    
    Dim basePdfName As String, counter As Long
    basePdfName = fso.BuildPath(catExportDir, Context.DynamicPdfCatName & "_圖片清單報告")
    pdfPath = basePdfName & ".pdf"
    counter = 1
    Do While fso.FileExists(pdfPath)
        pdfPath = basePdfName & "_" & counter & ".pdf"
        counter = counter + 1
    Loop
    
    wsPdf.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Context.LogStore.Record "TX_COMMIT_FS", "EXPORT", "橫式目錄PDF匯出完成：" & pdfPath
    Call Mod_UIMessenger.ShowInfo(strLib.InfoExportCatSuccess, strLib.TitleComplete)

ExpCatExit:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wsPdf Is Nothing Then wsPdf.Delete
    Application.DisplayAlerts = True
    If wasProtected And Not wb Is Nothing Then wb.Protect Password:=config.SecurityPassword, Structure:=True, Windows:=False
    
    Call Mod_UIReset.GoToUI
    Application.ScreenUpdating = origScreen: Application.EnableEvents = origEvents: Application.StatusBar = False
    Set fso = Nothing: Set wsCat = Nothing: Set strLib = Nothing
    isRunning = False
    Exit Sub

ErrCatExport:
    isRunning = False
    Dim errDesc As String: errDesc = Err.Description
    If Err.Number = 1004 Or InStr(errDesc, "1004") > 0 Then errDesc = "PDF 檔案可能正被佔用，請關閉後重試。"
    Context.LogStore.Record "TX_ERROR", "EXPORT", "橫式目錄PDF匯出失敗：" & errDesc
    Call Mod_UIMessenger.ShowError(strLib.ErrExportCat(errDesc), strLib.TitleSysWarning)
    Resume ExpCatExit
End Sub

Public Sub 執行匯出附件目錄PDF(ByVal Context As cls_ExecutionContext)
    Dim config As cls_Config: Set config = Context.config
    Dim fso As Object, file As Object, varFile As Variant
    Dim wb As Workbook, wsPdf As Worksheet
    Dim pdfPath As String, catExportDir As String
    Dim origEvents As Boolean, origScreen As Boolean, wasProtected As Boolean, rowIdx As Long
    Dim colFiles As Collection, strLib As New cls_StringLibrary
    
    On Error GoTo ErrAttExport
    Set wb = config.TargetWB
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set colFiles = New Collection
    
    origEvents = Application.EnableEvents: origScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False: Application.EnableEvents = False
    
    wasProtected = wb.ProtectStructure
    If wasProtected Then wb.Unprotect Password:=config.SecurityPassword
    Context.LogStore.Record "TX_START", "EXPORT", "開始匯出附件目錄PDF"
    
    Call Mod_Actions.InitializeSystemFolders(config)
    If fso.FolderExists(config.IOAttachmentFolder) Then Call Mod_Actions.GatherFilesRecursive(fso, config.IOAttachmentFolder, colFiles, config)
    If fso.FolderExists(config.IOQuarantineFolder) Then Call Mod_Actions.GatherFilesRecursive(fso, config.IOQuarantineFolder, colFiles, config)
    
    ' 【空清單防呆】：預先算一次非圖片檔案數量，沒有就直接拒絕匯出，不再產生只有一句提示語的空報告
    Dim nonImageCount As Long: nonImageCount = 0
    Dim probeFile As Variant, probeObj As Object
    For Each probeFile In colFiles
        Set probeObj = fso.GetFile(CStr(probeFile))
        If Not Mod_Rules.IsImageExtension(fso.GetExtensionName(probeObj.Name)) Then nonImageCount = nonImageCount + 1
    Next probeFile
    If nonImageCount = 0 Then
        Call Mod_UIMessenger.ShowWarning("目前專案內無任何非圖片之附件檔案，無需匯出。", strLib.TitleSysWarning)
        GoTo ExpAttExit
    End If

    
    Application.StatusBar = "EXPORT: 正在整理附件資料並生成 PDF..."
    Application.DisplayAlerts = False
    Dim wsTemp As Worksheet
    For Each wsTemp In wb.Worksheets
        If InStr(1, wsTemp.Name, "TEMP_PDF_ATT", vbTextCompare) > 0 Then wsTemp.Delete
    Next wsTemp
    Application.DisplayAlerts = True
    
    Set wsPdf = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
    wsPdf.Name = "TEMP_PDF_ATT"
    
    wsPdf.Range("A1").Value = "■ 專案附件歸檔目錄 - " & Context.DynamicProjName
    wsPdf.Range("A1").Font.Name = config.FontMain: wsPdf.Range("A1").Font.Size = config.FontSizeSub: wsPdf.Range("A1").Font.Bold = True
    wsPdf.Range("A1:E1").Merge
    wsPdf.Rows("1:1").RowHeight = 35: wsPdf.Rows("2:2").RowHeight = 10
    
    wsPdf.Range("A3:E3").Value = Array("副檔名", "檔案名稱", "檔案大小 (KB)", "修改時間", "實體路徑")
    With wsPdf.Range("A3:E3")
        .Interior.Color = RGB(130, 150, 170): .Font.Color = config.ColorWhite: .Font.Name = config.FontMain: .Font.Bold = True: .RowHeight = 22: .HorizontalAlignment = xlCenter
    End With
    
    rowIdx = 4
    For Each varFile In colFiles
        Set file = fso.GetFile(CStr(varFile))
        If Not Mod_Rules.IsImageExtension(fso.GetExtensionName(file.Name)) Then
            wsPdf.Cells(rowIdx, 1).Value = UCase(fso.GetExtensionName(file.Name))
            wsPdf.Cells(rowIdx, 2).Value = file.Name
            wsPdf.Cells(rowIdx, 3).Value = Format(file.Size / 1024, "#,##0.0")
            wsPdf.Cells(rowIdx, 4).Value = Format(file.DateLastModified, "yyyy/mm/dd hh:mm")
            wsPdf.Cells(rowIdx, 5).Value = file.Path
            
            With wsPdf.Range(wsPdf.Cells(rowIdx, 1), wsPdf.Cells(rowIdx, 5))
                .Font.Name = config.FontMain: .Font.Size = 10: .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(220, 220, 220)
                .VerticalAlignment = xlVAlignCenter: .EntireRow.RowHeight = 20
            End With
            rowIdx = rowIdx + 1
        End If
    Next varFile
    
    If rowIdx = 4 Then
        wsPdf.Range("A4:E4").Merge
        wsPdf.Range("A4").Value = "（目前專案內無任何非圖片之附件檔案）"
        wsPdf.Range("A4").HorizontalAlignment = xlCenter: wsPdf.Range("A4").Font.Color = config.ColorTextHint
    End If
    
    With wsPdf.PageSetup
        .PaperSize = xlPaperA4: .Orientation = xlLandscape
        .LeftMargin = Application.CentimetersToPoints(config.PDFMarginCM): .RightMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .TopMargin = Application.CentimetersToPoints(config.PDFMarginCM): .BottomMargin = Application.CentimetersToPoints(config.PDFMarginCM)
        .CenterHeader = "": .CenterFooter = "&P": .PrintTitleRows = ""
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
    End With
    
    wsPdf.Columns("A:E").AutoFit
    If wsPdf.Columns("B:B").ColumnWidth > 35 Then wsPdf.Columns("B:B").ColumnWidth = 35
    If wsPdf.Columns("E:E").ColumnWidth > 45 Then wsPdf.Columns("E:E").ColumnWidth = 45
    wsPdf.Columns("B:B").WrapText = True: wsPdf.Columns("E:E").WrapText = True
    wsPdf.Rows.AutoFit: ActiveWindow.DisplayGridlines = False
    
    catExportDir = fso.BuildPath(config.IOExportFolder, "目錄")
    Call Mod_Utils.EnsureFolderExists(fso, catExportDir)
    
    Dim basePdfName As String, counter As Long
    basePdfName = fso.BuildPath(catExportDir, Context.DynamicPdfCatName & "_附件清單報告")
    pdfPath = basePdfName & ".pdf"
    counter = 1
    Do While fso.FileExists(pdfPath)
        pdfPath = basePdfName & "_" & counter & ".pdf"
        counter = counter + 1
    Loop
    
    wsPdf.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Context.LogStore.Record "TX_COMMIT_FS", "EXPORT", "附件目錄PDF匯出完成：" & pdfPath
    Call Mod_UIMessenger.ShowInfo(strLib.InfoBatchExportSuccess(pdfPath), strLib.TitleComplete)

ExpAttExit:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wsPdf Is Nothing Then wsPdf.Delete
    Application.DisplayAlerts = True
    If wasProtected And Not wb Is Nothing Then wb.Protect Password:=config.SecurityPassword, Structure:=True, Windows:=False
    
    Call Mod_UIReset.GoToUI
    Application.ScreenUpdating = origScreen: Application.EnableEvents = origEvents: Application.StatusBar = False
    Set fso = Nothing: Set colFiles = Nothing: Set strLib = Nothing
    Exit Sub

ErrAttExport:
    Dim errDesc As String: errDesc = Err.Description
    If Err.Number = 1004 Or InStr(errDesc, "1004") > 0 Then errDesc = "PDF 檔案可能正被佔用，請關閉後重試。"
    Context.LogStore.Record "TX_ERROR", "EXPORT", "橫式附件目錄PDF匯出失敗：" & errDesc
    Call Mod_UIMessenger.ShowError("附件目錄匯出失敗：" & errDesc, strLib.TitleSysWarning)
    Resume ExpAttExit
End Sub
