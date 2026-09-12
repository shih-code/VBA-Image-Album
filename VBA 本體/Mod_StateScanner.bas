Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_StateScanner (標準模組)
' PURPOSE: 狀態掃描器。開局極速掃描實體硬碟與畫布，
'          嚴格區分整理夾與救援夾兩套範圍，建立雙軌電子名冊。
' EXPORTS: ScanCurrentState
' IMPORTS: cls_ExecutionContext, Mod_Rules, Mod_Actions, Mod_Utils
' FORBIDDEN: 1. 嚴禁跳出 MsgBox 或干涉 UI。
'            2. 嚴禁修改、刪除任何實體檔案或 Excel 儲存格。
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_StateScanner"
Private Const MODULE_VERSION As String = "1.0.0"

Public Sub ScanCurrentState(ByVal Context As cls_ExecutionContext)
    Dim fso As Object
    Dim config As cls_Config
    Dim dictHashCache As Object
    Dim dictOrgAlive As Object    ' 標準整理夾存活名單
    Dim dictRescueAlive As Object ' 救援夾存活名單
    Dim dictCatHashByName As Object

    Set config = Context.config
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dictHashCache = CreateObject("Scripting.Dictionary"): dictHashCache.CompareMode = 1
    Set dictOrgAlive = CreateObject("Scripting.Dictionary"): dictOrgAlive.CompareMode = 1
    Set dictRescueAlive = CreateObject("Scripting.Dictionary"): dictRescueAlive.CompareMode = 1

    ' ==========================================================================
    ' 步驟 1: 讀取舊目錄 (建立雜湊效能快取與身分對照表)
    ' ==========================================================================
    Set dictCatHashByName = BuildHashCacheAndCatalogIndex(fso, config, Context, dictHashCache)

    ' ==========================================================================
    ' 步驟 2: 實體掃描 (嚴格分離整理夾與救援夾) -> 登錄全域指紋庫
    ' ==========================================================================
    Call ScanFolderIntoRegistry(fso, config, Context, dictHashCache, config.IOOrganizeFolder, False, Context.DictGlobalHash, dictOrgAlive)
    Call ScanFolderIntoRegistry(fso, config, Context, dictHashCache, config.IORescueFolder, True, Context.DictRescueHash, dictRescueAlive)

    ' ==========================================================================
    ' 步驟 3: 巡查畫布，找出對應不到實體檔案的圖片
    ' ==========================================================================
    Call DetectAllMissingShapes(config, Context, dictCatHashByName, dictOrgAlive, dictRescueAlive)

    ' 釋放記憶體
    Set fso = Nothing: Set dictHashCache = Nothing
    Set dictOrgAlive = Nothing: Set dictRescueAlive = Nothing
    Set dictCatHashByName = Nothing
End Sub

' --------------------------------------------------------------------------
' 【職責 2】讀取舊目錄，建立雜湊效能快取（寫入傳入的 dictHashCache）、
' 名稱→雜湊身分對照表（回傳），並登記救援血統候選清單（Context.RestoreCandidates）
' --------------------------------------------------------------------------
Private Function BuildHashCacheAndCatalogIndex(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal dictHashCache As Object) As Object
    Dim wsCat As Worksheet
    Dim r As Long, lastRow As Long
    Dim catName As String, catHash As String, catPath As String, catDate As String, catOrigHash As String
    Dim catSize As String

    Dim dictCatHashByName As Object
    Set dictCatHashByName = CreateObject("Scripting.Dictionary"): dictCatHashByName.CompareMode = 1

    Set wsCat = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameCatalog)
    If Not wsCat Is Nothing Then
        lastRow = wsCat.Cells(wsCat.Rows.count, "A").End(xlUp).row
        For r = 5 To lastRow
            catName = Trim(CStr(wsCat.Cells(r, 1).Value))
            catHash = Trim(CStr(wsCat.Cells(r, 2).Value))
            catPath = Trim(CStr(wsCat.Cells(r, 3).Value))
            catDate = ""
            If IsDate(wsCat.Cells(r, 6).Value) Then catDate = Format(CDate(wsCat.Cells(r, 6).Value), "yyyy/mm/dd hh:mm:ss")
            catSize = Trim(CStr(wsCat.Cells(r, 4).Value))
            catOrigHash = Trim(CStr(wsCat.Cells(r, 9).Value))

            If catPath <> "" And catDate <> "" And catSize <> "" And catHash <> "" Then
                dictHashCache(catPath & "|" & catDate & "|" & catSize) = catHash
                 ' 【備援key】：以實體檔名（非識別名稱欄位，避免格式不一致）建立備援查詢依據，
                 ' 沿用原本「首次出現者優先」語意，故用 Add + Exists 檢查，而非直接指派覆蓋
                 Dim catFileName As String
                 catFileName = fso.GetFileName(catPath)
                 Dim metaKey As String
                 metaKey = catFileName & "|" & catDate & "|" & catSize
                 If Not dictHashCache.Exists(metaKey) Then
                     dictHashCache.Add metaKey, catHash
                 End If
            End If

            ' 【身分對照表】：名稱→目前雜湊，供步驟3判斷畫布 Shape 身分時查表使用
            If catName <> "" And catHash <> "" Then dictCatHashByName(catName) = catHash

            If catOrigHash <> "" And catName Like "MANUAL_RESCUE_*" Then
                Context.RestoreCandidates(catOrigHash) = catName
            End If
        Next r
    End If

    Set BuildHashCacheAndCatalogIndex = dictCatHashByName
End Function

' --------------------------------------------------------------------------
' 【職責 3a+3b 合併】掃描指定資料夾，登錄雜湊指紋庫。原本整理夾／救援夾兩段
' 幾乎相同的程式碼合併為一支，用 useRescueGather 切換收檔規則（救援夾不套用
' IsProtectedFolder排除），用 targetHashDict/aliveDict 切換寫入的目標字典
' --------------------------------------------------------------------------
Private Sub ScanFolderIntoRegistry(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, _
                                    ByVal dictHashCache As Object, ByVal folderPath As String, ByVal useRescueGather As Boolean, _
                                    ByVal targetHashDict As Object, ByVal aliveDict As Object)
    Dim colFiles As Collection, varFile As Variant, fileObj As Object
    Dim fileHash As String, cacheKey As String

    Set colFiles = New Collection
    If fso.FolderExists(folderPath) Then
        If useRescueGather Then
            Call GatherRescueFilesRecursive(fso, folderPath, colFiles, config)
        Else
            Call Mod_Actions.GatherFilesRecursive(fso, folderPath, colFiles, config)
        End If

        For Each varFile In colFiles
            Set fileObj = fso.GetFile(CStr(varFile))
            If Mod_Rules.IsImageExtension(fso.GetExtensionName(fileObj.Name)) Then
                cacheKey = fileObj.Path & "|" & Format(fileObj.DateLastModified, "yyyy/mm/dd hh:mm:ss") & "|" & CStr(fileObj.Size)
                If dictHashCache.Exists(cacheKey) Then
                    fileHash = dictHashCache(cacheKey)
                Else
                    ' 路徑對不上，退回「檔名+時間+大小」備援比對，
                    ' 命中代表僅搬家未改名，沿用既有雜湊，不重算
                    Dim metaKeyScan As String
                    metaKeyScan = fileObj.Name & "|" & Format(fileObj.DateLastModified, "yyyy/mm/dd hh:mm:ss") & "|" & CStr(fileObj.Size)
                    If dictHashCache.Exists(metaKeyScan) Then
                        fileHash = dictHashCache(metaKeyScan)
                    Else
                        fileHash = Mod_Utils.GetMD5(fso, fileObj.Path)
                    End If
                    dictHashCache(cacheKey) = fileHash
                End If
                targetHashDict(fileHash) = fileObj.Path
                Context.DictPathToHash(fileObj.Path) = fileHash  ' 【V2-047】供cmd_Deduplicate查詢，避免重算
                aliveDict(fileHash) = fileObj.Path ' 改用雜湊登記存活，值順便存實體路徑供之後查用
            End If
        Next varFile
    End If
End Sub

' --------------------------------------------------------------------------
' 【職責 4 調度】依工作表類型分派到對應的存活名單進行孤兒偵測
' --------------------------------------------------------------------------
Private Sub DetectAllMissingShapes(ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal dictCatHashByName As Object, ByVal dictOrgAlive As Object, ByVal dictRescueAlive As Object)
    Context.MissingShapes.RemoveAll

    Dim ws As Worksheet
    For Each ws In config.TargetWB.Worksheets
        If Mod_Rules.IsStandardImgSheet(ws.Name, config) Or ws.Name = config.SheetNameAllPics Then
            Call DetectMissingShapesInSheet(ws, dictCatHashByName, dictOrgAlive, Context)
        ElseIf ws.Name = config.SheetNameRescue Then
            Call DetectMissingShapesInSheet(ws, dictCatHashByName, dictRescueAlive, Context)
        End If
    Next ws
End Sub

' --------------------------------------------------------------------------
' 【職責 4a+4b 合併】巡查單一工作表，找出對應不到實體檔案的圖片。
' 原本標準畫布／救援畫布兩軌幾乎相同的邏輯合併為一支，用 aliveDict 切換比對基準
' --------------------------------------------------------------------------
Private Sub DetectMissingShapesInSheet(ByVal ws As Worksheet, ByVal dictCatHashByName As Object, ByVal aliveDict As Object, ByVal Context As cls_ExecutionContext)
    Dim shp As Shape, shpHash As String, uniqueKey As String
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            If dictCatHashByName.Exists(shp.Name) Then
                shpHash = dictCatHashByName(shp.Name)
                ' 目錄表有登記：查雜湊存活清單，雜湊對得上就是活的，不管檔名怎麼被改過
                If Not aliveDict.Exists(shpHash) Then
                    uniqueKey = ws.Name & "||" & shp.Name
                    If Not Context.MissingShapes.Exists(uniqueKey) Then Context.MissingShapes.Add uniqueKey, shp.Name
                End If
            Else
                ' 目錄表從沒登記過這個名字（野生貼圖等情況）：一律視為遺失，
                ' 交由既有救援流程處理，順便正式登記進目錄表
                uniqueKey = ws.Name & "||" & shp.Name
                If Not Context.MissingShapes.Exists(uniqueKey) Then Context.MissingShapes.Add uniqueKey, shp.Name
            End If
        End If
    Next shp
End Sub

' ------------------------------------------------------------------------------
' [輔助引擎]：GatherRescueFilesRecursive
' 目的：專供掃描救援夾自身使用，不套用 IsProtectedFolder 排除規則
'      （該規則是為了防止匯入時誤抓保護區，救援夾掃描不適用這條）
' ------------------------------------------------------------------------------
Private Sub GatherRescueFilesRecursive(ByVal fso As Object, ByVal currentPath As String, ByRef sourceFiles As Collection, ByVal config As cls_Config)
    Dim folder As Object, subFolder As Object, file As Object
    ' 已歸位／已廢棄的子夾屬於塵埃落定的封存區，不參與救援比對，跳過遞迴
    If UCase(currentPath) = UCase(config.IORescueDowncastFolder) Or UCase(currentPath) = UCase(config.IORescueDiscardFolder) Then Exit Sub
    On Error Resume Next
    Set folder = fso.GetFolder(currentPath)
    If folder Is Nothing Then Exit Sub
    
    For Each file In folder.Files: sourceFiles.Add file.Path: Next file
    For Each subFolder In folder.SubFolders: Call GatherRescueFilesRecursive(fso, subFolder.Path, sourceFiles, config): Next subFolder
    On Error GoTo 0
End Sub
