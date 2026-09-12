Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_LayoutEngine (標準模組)
' PURPOSE: 幾何排版引擎。
'          專職負責將工作表內的圖片，依據執行期狀態進行雙欄瀑布流幾何算繪。
' EXPORTS: ExecuteMasonryLayout
' IMPORTS: cls_ExecutionContext, Mod_Utils, Mod_Rules, cls_Config
' FORBIDDEN: 1. 嚴禁在此撰寫任何彈窗 (MsgBox)，所有異常交由上層管線熔斷。
'            2. 嚴禁在 Shape 迴圈內部使用 Columns.Find 進行慢速跨表探測。
' DEPENDENCIES: Microsoft Scripting Runtime (Dictionary)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_LayoutEngine"
Private Const MODULE_VERSION As String = "1.0.0"

Public Sub ExecuteMasonryLayout(ByVal ws As Worksheet, ByVal Context As cls_ExecutionContext)
    Dim config As cls_Config
    Dim shp As Shape, picCount As Long, targetW As Double, targetS As Double, isRescueSheet As Boolean
    Dim errNum As Long, errDesc As String

    ' 【錯誤處理閉環】：確保算繪崩潰時能安全釋放陣列與字典
    On Error GoTo LayoutErrorHandler

    Set config = Context.config
    targetW = Context.DynamicImgWidth
    targetS = Context.DynamicImgSpacing
    isRescueSheet = (ws.Name = config.SheetNameRescue)
    
    ' 【修正】救援表的圖片尺寸與間距本來就該固定（見 cmd_Rebuild.ProcessSingleFile
    ' V2-062 說明），不該跟著操作介面的寬度/間距設定變動，否則調整一般設定
    ' 會連帶讓救援表的圖片跟著縮放，跟掛載當下鎖定的固定尺寸自相矛盾
    If isRescueSheet Then
        targetW = config.RescueEmbedWidth
        targetS = config.RescueEmbedGap
    End If

    Call Mod_Utils.SafeUnprotect(ws, config.SecurityPassword)

    picCount = 0
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then picCount = picCount + 1
    Next shp

    If picCount = 0 Then GoTo LayoutCleanup

    ' 幾何陣列動態分配
    Dim shpArr() As Shape, shpTop() As Double
    Dim txtName() As String, txtImport() As String, txtSrc() As String, txtHash() As String
    ReDim shpArr(1 To picCount): ReDim shpTop(1 To picCount)
    ReDim txtName(1 To picCount): ReDim txtImport(1 To picCount): ReDim txtSrc(1 To picCount): ReDim txtHash(1 To picCount)

    Dim nPic As Long: nPic = 0
    Dim catWs As Worksheet
    Set catWs = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameCatalog)

    Dim dictCatCache As Object
    Set dictCatCache = BuildCatalogCache(catWs)

    Call CollectAndBindPictures(ws, dictCatCache, isRescueSheet, shpArr, shpTop, txtName, txtImport, txtSrc, txtHash, nPic)

    Call SortShapesByTop(shpArr, shpTop, txtName, txtImport, txtSrc, txtHash, nPic)

    Dim rCurrent As Long, shpL As Shape, shpR As Shape, formatRange As Range
    rCurrent = PrepareCanvasGrid(ws, isRescueSheet, targetW, targetS)

    Call RenderMasonryPairs(ws, config, isRescueSheet, targetW, targetS, shpArr, txtName, txtImport, txtSrc, txtHash, nPic, rCurrent, shpL, shpR, formatRange)

LayoutCleanup:
    On Error Resume Next
    Set dictCatCache = Nothing
    Erase shpArr: Erase shpTop: Erase txtName: Erase txtImport: Erase txtSrc: Erase txtHash
    Set config = Nothing: Set catWs = Nothing: Set shpL = Nothing: Set shpR = Nothing: Set formatRange = Nothing
    Exit Sub

LayoutErrorHandler:
    errNum = Err.Number: errDesc = Err.Description
    If Not Context Is Nothing Then
        If Not Context.LogStore Is Nothing Then
            Context.LogStore.Record "TX_ERROR", "LAYOUT_ENGINE", "畫布算繪失敗 (" & ws.Name & "): " & errDesc
        End If
    End If
    Resume LayoutCleanup
End Sub

' --------------------------------------------------------------------------
' 【職責 3】建立目錄快取字典：一次性讀取《圖片目錄》塞入 Dictionary，供後續 O(1) 配對
' --------------------------------------------------------------------------
Private Function BuildCatalogCache(ByVal catWs As Worksheet) As Object
    Dim dictCatCache As Object
    Set dictCatCache = CreateObject("Scripting.Dictionary")
    dictCatCache.CompareMode = 1

    If Not catWs Is Nothing Then
        Dim r As Long, lastCatRow As Long
        lastCatRow = catWs.Cells(catWs.Rows.count, "A").End(xlUp).row
        For r = 5 To lastCatRow
            Dim kName As String
            kName = Trim(CStr(catWs.Cells(r, 1).Value))
            If kName <> "" Then
                dictCatCache(kName) = Array( _
                    Trim(CStr(catWs.Cells(r, 2).Value)), _
                    Trim(CStr(catWs.Cells(r, 5).Value)), _
                    Trim(CStr(catWs.Cells(r, 7).Value)) _
                )
            End If
        Next r
    End If

    Set BuildCatalogCache = dictCatCache
End Function

' --------------------------------------------------------------------------
' 【職責 4】收集圖片 Shape，並與目錄快取進行 metadata 綁定
' --------------------------------------------------------------------------
Private Sub CollectAndBindPictures(ByVal ws As Worksheet, ByVal dictCatCache As Object, ByVal isRescueSheet As Boolean, _
                                    ByRef shpArr() As Shape, ByRef shpTop() As Double, _
                                    ByRef txtName() As String, ByRef txtImport() As String, ByRef txtSrc() As String, ByRef txtHash() As String, _
                                    ByRef nPic As Long)
    Dim shp As Shape
    nPic = 0
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            nPic = nPic + 1: Set shpArr(nPic) = shp: shpTop(nPic) = shp.Top
            txtName(nPic) = "■ 名稱: " & shp.Name

            If dictCatCache.Exists(shp.Name) Then
                Dim catData As Variant
                catData = dictCatCache(shp.Name)
                txtHash(nPic) = "■ 雜湊: " & catData(0)
                txtImport(nPic) = "■ 匯入: " & catData(1)
                txtSrc(nPic) = "■ 來源: " & catData(2)
            Else
                txtImport(nPic) = "■ 匯入: " & Format(Now, "yyyy/mm/dd hh:mm")
                txtSrc(nPic) = "■ 來源: " & IIf(isRescueSheet, "原始目錄紀錄遺失", "系統重新排版")
                txtHash(nPic) = "■ 雜湊: 等待管線重新驗證"
            End If
        End If
    Next shp
End Sub

' --------------------------------------------------------------------------
' 【職責 5】依 Top 座標氣泡排序（純陣列操作，無 I/O，符合 AVS 第二節 Rule 定義）
' --------------------------------------------------------------------------
Private Sub SortShapesByTop(ByRef shpArr() As Shape, ByRef shpTop() As Double, _
                             ByRef txtName() As String, ByRef txtImport() As String, ByRef txtSrc() As String, ByRef txtHash() As String, _
                             ByVal nPic As Long)
    Dim i As Long, j As Long, ts As Shape, tT As Double, tt1 As String, tt2 As String, tt3 As String, tt4 As String
    For i = 1 To nPic - 1
        For j = 1 To nPic - i
            If shpTop(j) > shpTop(j + 1) Then
                Set ts = shpArr(j): tT = shpTop(j): tt1 = txtName(j): tt2 = txtImport(j): tt3 = txtSrc(j): tt4 = txtHash(j)
                Set shpArr(j) = shpArr(j + 1): shpTop(j) = shpTop(j + 1): txtName(j) = txtName(j + 1): txtImport(j) = txtImport(j + 1): txtSrc(j) = txtSrc(j + 1): txtHash(j) = txtHash(j + 1)
                Set shpArr(j + 1) = ts: shpTop(j + 1) = tT: txtName(j + 1) = tt1: txtImport(j + 1) = tt2: txtSrc(j + 1) = tt3: txtHash(j + 1) = tt4
            End If
        Next j
    Next i
End Sub

' --------------------------------------------------------------------------
' 【職責 6】版面清除＋欄寬設定（含 Excel 列數/欄寬上限防禦），回傳算好的起始列 rCurrent
' --------------------------------------------------------------------------
Private Function PrepareCanvasGrid(ByVal ws As Worksheet, ByVal isRescueSheet As Boolean, ByVal targetW As Double, ByVal targetS As Double) As Long
    Dim lastRowCheck As Long, clearTarget As Long
    lastRowCheck = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    If ws.Cells(ws.Rows.count, 6).End(xlUp).row > lastRowCheck Then lastRowCheck = ws.Cells(ws.Rows.count, 6).End(xlUp).row

    clearTarget = lastRowCheck + 500
    If clearTarget > ws.Rows.count Then clearTarget = ws.Rows.count

    ws.Range("A4:H" & clearTarget).ClearContents
    ws.Range("A4:A" & clearTarget).EntireRow.RowHeight = 15

    Dim cwB As Double, cwE As Double
    cwB = targetW / 5.25: If cwB > 254 Then cwB = 254
    cwE = IIf(targetS < 10, 10, targetS) / 5.25: If cwE > 254 Then cwE = 254

    ws.Columns("B").ColumnWidth = cwB: ws.Columns("C:D").ColumnWidth = 2
    ws.Columns("E").ColumnWidth = cwE
    ws.Columns("F").ColumnWidth = cwB: ws.Columns("G:H").ColumnWidth = 2

    PrepareCanvasGrid = IIf(isRescueSheet, 5, 4)
End Function

' --------------------------------------------------------------------------
' 【職責 7】雙欄瀑布流幾何算繪與資料落地
' --------------------------------------------------------------------------
Private Sub RenderMasonryPairs(ByVal ws As Worksheet, ByVal config As cls_Config, ByVal isRescueSheet As Boolean, _
                                ByVal targetW As Double, ByVal targetS As Double, _
                                ByRef shpArr() As Shape, ByRef txtName() As String, ByRef txtImport() As String, ByRef txtSrc() As String, ByRef txtHash() As String, _
                                ByVal nPic As Long, ByVal rCurrent As Long, _
                                ByRef shpL As Shape, ByRef shpR As Shape, ByRef formatRange As Range)
    Dim i As Long, hL As Double, hR As Double, maxH As Double, ratioL As Double, ratioR As Double

    For i = 1 To nPic Step 2
        If rCurrent + 20 > ws.Rows.count Then Exit For

        Set shpL = shpArr(i): ratioL = IIf(shpL.Width > 0, shpL.Height / shpL.Width, 1): hL = targetW * ratioL
        shpL.LockAspectRatio = msoFalse: shpL.Width = targetW: shpL.Height = hL: shpL.LockAspectRatio = msoTrue

        hR = 0
        If i + 1 <= nPic Then
            Set shpR = shpArr(i + 1): ratioR = IIf(shpR.Width > 0, shpR.Height / shpR.Width, 1): hR = targetW * ratioR
            shpR.LockAspectRatio = msoFalse: shpR.Width = targetW: shpR.Height = hR: shpR.LockAspectRatio = msoTrue
        End If

        maxH = IIf(hL > hR, hL, hR)

        ws.Cells(rCurrent, 2).Value = txtName(i): ws.Cells(rCurrent + 1, 2).Value = txtImport(i): ws.Cells(rCurrent + 2, 2).Value = txtSrc(i): ws.Cells(rCurrent + 3, 2).Value = txtHash(i)
        If i + 1 <= nPic Then
            ws.Cells(rCurrent, 6).Value = txtName(i + 1): ws.Cells(rCurrent + 1, 6).Value = txtImport(i + 1): ws.Cells(rCurrent + 2, 6).Value = txtSrc(i + 1): ws.Cells(rCurrent + 3, 6).Value = txtHash(i + 1)
        End If

        If i + 1 <= nPic Then Set formatRange = ws.Range(ws.Cells(rCurrent, 2), ws.Cells(rCurrent + 3, 6)) Else Set formatRange = ws.Range(ws.Cells(rCurrent, 2), ws.Cells(rCurrent + 3, 2))
        With formatRange.Font
            .Name = config.FontMain
            .Size = 9: .Bold = True: .Color = IIf(isRescueSheet, RGB(180, 70, 70), RGB(40, 80, 150))
        End With

        shpL.Left = ws.Range("B2").Left: shpL.Top = ws.Cells(rCurrent + 5, 2).Top
        If i + 1 <= nPic Then shpR.Left = ws.Range("F2").Left: shpR.Top = ws.Cells(rCurrent + 5, 6).Top

        rCurrent = rCurrent + 5 + Int((maxH + targetS) / 15) + 2
    Next i
End Sub
