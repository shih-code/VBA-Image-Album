Option Explicit
Implements ICommand

' ==========================================================
' MODULE: cmd_Rebuild (類別模組)
' PURPOSE: 專職目錄重建與排版掛載命令。遵守單一職責原則 (SRP)。
'          以硬碟現狀為準，將實體資產掛載至畫布並寫入目錄。
' EXPORTS: ICommand_Execute 等介面合約
' IMPORTS: cls_ExecutionContext, Mod_Rules, Mod_Actions, Mod_Utils, Mod_LayoutEngine, Mod_ThumbnailCache, Mod_EmptyFolderCleaner, Mod_UIReset, Mod_ImportCollisionGate
' FORBIDDEN: 1. 嚴禁在此實作任何指紋比對與替代邏輯 (已交由 cmd_Deduplicate)。
'            2. 嚴禁呼叫 MsgBox 中斷管線。
' DEPENDENCIES: Microsoft Scripting Runtime (FileSystemObject, Dictionary)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "cmd_Rebuild"
Private Const MODULE_VERSION As String = "1.0.0"

Private m_NewCount As Long
Private m_RescueCount As Long
Private m_DictValidShapes As Object     ' 合法通行證：記錄已在實體硬碟找到的檔案名稱
Private m_DictCatWritten As Object      ' 防止重複寫入目錄之快取
Private m_DictHashCache As Object       ' 效能快取：記錄 [路徑|修改時間|大小] -> HASH
Private m_HashCacheHit As Long          ' 診斷用：本輪快取命中次數
Private m_HashCacheMiss As Long         ' 診斷用：本輪快取未命中、現算次數
Private m_DictDNA As Object             ' 備份機制：記錄 [原始HASH -> Array(狀態, 是否救援, 原HASH)]
Private m_DictShapeHash As Object      ' 工作表||Shape名稱 → 雜湊
Private m_DictHashToShape As Object    ' 雜湊 → 工作表||Shape名稱（真正的O(1)查表用）

Private m_DictCatNameToHash As Object

' 【局部重排】本輪被 ProcessSingleFile／CleanInvalidShapes 實際改動過畫布的工作表名稱
Private m_DictTouchedSheets As Object

' ==============================================================================
' ICommand 介面實作區
' ==============================================================================
Private Property Get ICommand_IsCritical() As Boolean
    ICommand_IsCritical = True
End Property

Private Property Get ICommand_CommandName() As String
    ICommand_CommandName = "REBUILD"
End Property

Private Property Get ICommand_Dependencies() As Variant
    ICommand_Dependencies = Array("DEDUPLICATE")
End Property

Private Sub ICommand_Execute(ByVal Context As cls_ExecutionContext)
    Dim fso As Object, config As cls_Config, LogStore As Object
    Dim catWs As Worksheet, attWs As Worksheet, wsSingle As Worksheet, wsRescue As Worksheet
    Dim catRow As Long, attRow As Long
    Dim errNum As Long, errDesc As String
    
    On Error GoTo RebuildErrorHandler
    
    Set config = Context.config
    Set LogStore = Context.LogStore
    If config Is Nothing Or LogStore Is Nothing Then Err.Raise 9991, MODULE_NAME & ".Execute", "Context 遺失"
    
    m_NewCount = 0: m_RescueCount = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set m_DictValidShapes = CreateObject("Scripting.Dictionary"): m_DictValidShapes.CompareMode = 1
    Set m_DictCatWritten = CreateObject("Scripting.Dictionary"): m_DictCatWritten.CompareMode = 1
    Set m_DictHashCache = CreateObject("Scripting.Dictionary"): m_DictHashCache.CompareMode = 1
    m_HashCacheHit = 0: m_HashCacheMiss = 0
    Set m_DictDNA = CreateObject("Scripting.Dictionary"): m_DictDNA.CompareMode = 1
    Set m_DictTouchedSheets = CreateObject("Scripting.Dictionary"): m_DictTouchedSheets.CompareMode = 1
    
    LogStore.Record "TX_START", "REBUILD", "開始執行畫布掛載與雙目錄重建作業"
    Call Mod_Actions.InitializeSystemFolders(config)
    
   ' 【安全網】REBUILD會產生大量縮圖快取與實體同步的複製暫存，同樣需要確認磁碟空間
   If Not CheckDiskSpaceGate(config, LogStore) Then GoTo RebuildCleanup
    
    Call ConsolidateLooseFilesInOrganizeRoot(fso, config)
    
    Dim wsCatForCollisionCheck As Worksheet
    Set wsCatForCollisionCheck = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameCatalog)
    If Not wsCatForCollisionCheck Is Nothing Then
        Call Mod_ImportCollisionGate.GatherConflicts(fso, config, Context, wsCatForCollisionCheck, "REBUILD")
        Call Mod_ImportCollisionGate.AskAndApplyDecisions(fso, config, Context, wsCatForCollisionCheck)
    End If
    
    ' 【修正】拍照時機挪到撞名詢問之後——AskAndApplyDecisions 可能剛把某些列
    ' 標記「已取代」，BackupCatalogState 必須拍到這個最新狀態，CleanInvalidShapes
    ' 的「等待救援保護傘」才不會誤用取代前的舊 isBounty 狀態保護一張其實已經
    ' 結案的 Shape
    Call BackupCatalogState(Context)
    Call Mod_Actions.CleanUpOldSheets(Context.config, Context.isCatalogAppend)

    ' 【建立 Shape 雜湊索引】
    Call BuildShapeHashIndex(config)
    
    Set catWs = InitializeCatalogSheet(config, Context)
    Set attWs = InitializeAttachmentSheet(config)
    catRow = 5: attRow = 5
    
    If Context.IsSingleMode Then
        Set wsSingle = Mod_Utils.GetOrCreateSheet(config.TargetWB, config.SheetNameAllPics)
        wsSingle.Cells.Clear
    End If
    
    Call MountOrganizeAndAttachmentFiles(fso, config, Context, LogStore, catWs, catRow, wsSingle, attWs, attRow)

    Set wsRescue = MountRescueFolderFiles(fso, config, Context, LogStore, catWs, catRow, wsSingle)
    
    ' 【修正】孤兒清除必須排在排版寫文字之前——所需資料掛載階段就已備妥，
    ' 不需要等排版跑完。順序對調後，畫布上不會再殘留「文字寫了、圖片被刪」的鬼影，
    ' 也讓局部重排的「摸過」標記能正確涵蓋孤兒清除造成的異動
    Call CleanInvalidShapes(config, LogStore)
     
    Call RenderFinalLayoutsAndStats(config, Context, LogStore, catWs, wsRescue)
    
    Call FinalizeRebuildCleanupAndFormatting(fso, config, Context, LogStore, catWs, attWs, wsRescue)
    
    LogStore.Record "TX_INFO", "REBUILD", "雜湊快取：命中 " & m_HashCacheHit & " | 現算 " & m_HashCacheMiss
    LogStore.Record "TX_COMMIT_FS", "REBUILD", "目錄重建與畫布排版作業完成"

RebuildCleanup:
    Set fso = Nothing: Set m_DictValidShapes = Nothing: Set m_DictCatWritten = Nothing
    Set m_DictHashCache = Nothing: Set m_DictDNA = Nothing
    Set m_DictTouchedSheets = Nothing
    If errNum <> 0 Then Err.Raise errNum, MODULE_NAME & ".Execute", errDesc
    Exit Sub

RebuildErrorHandler:
    errNum = Err.Number: errDesc = Err.Description
    If Not LogStore Is Nothing Then LogStore.Record "TX_ERROR", "REBUILD", "執行失敗: " & errDesc
    Resume RebuildCleanup
End Sub

' --------------------------------------------------------------------------
' 【職責 2】磁碟空間 Fail-Fast 查核。True=可繼續，False=空間不足
' --------------------------------------------------------------------------
Private Function CheckDiskSpaceGate(ByVal config As cls_Config, ByVal LogStore As Object) As Boolean
    CheckDiskSpaceGate = True
    If Not Mod_Actions.IsDiskSpaceSafe(config, LogStore) Then
        LogStore.Record "TX_ERROR", "REBUILD", "磁碟空間不足，全面排版作業中止，請先騰出空間再重新執行"
        CheckDiskSpaceGate = False
    End If
End Function

' --------------------------------------------------------------------------
' 建立 Shape 雜湊索引：掃描既有標準圖片工作表，比對目錄快照的名稱→雜湊記錄，
' 建立雙向索引（工作表||Shape名稱 to 雜湊），供後續掛載與孤兒清除使用
' --------------------------------------------------------------------------
Private Sub BuildShapeHashIndex(ByVal config As cls_Config)
    Set m_DictShapeHash = CreateObject("Scripting.Dictionary"): m_DictShapeHash.CompareMode = 1
    Set m_DictHashToShape = CreateObject("Scripting.Dictionary"): m_DictHashToShape.CompareMode = 1
    Dim idxWs As Worksheet, idxShp As Shape
    For Each idxWs In config.TargetWB.Worksheets
        If Mod_Rules.IsStandardImgSheet(idxWs.Name, config) Or idxWs.Name = config.SheetNameAllPics Or idxWs.Name = config.SheetNameRescue Then
            For Each idxShp In idxWs.Shapes
                If idxShp.Type = msoPicture Or idxShp.Type = msoLinkedPicture Then
                   If m_DictCatNameToHash.Exists(idxShp.Name) Then
                       m_DictShapeHash(idxWs.Name & "||" & idxShp.Name) = m_DictCatNameToHash(idxShp.Name)
                       m_DictHashToShape(m_DictCatNameToHash(idxShp.Name)) = idxWs.Name & "||" & idxShp.Name

                    End If
                End If
            Next idxShp
        End If
    Next idxWs
End Sub

' --------------------------------------------------------------------------
' 收集整理夾與附件區所有檔案，逐一掛載畫布（圖片）或寫入附件表（非圖片），
' 迴圈結束後補上附件表無資料時的提示文字與欄寬調整
' --------------------------------------------------------------------------
Private Sub MountOrganizeAndAttachmentFiles(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal LogStore As Object, _
                                             ByVal catWs As Worksheet, ByRef catRow As Long, ByVal wsSingle As Worksheet, _
                                             ByVal attWs As Worksheet, ByRef attRow As Long)
    Dim colAllFiles As Collection
    Set colAllFiles = New Collection
    If fso.FolderExists(config.IOOrganizeFolder) Then Call Mod_Actions.GatherFilesRecursive(fso, config.IOOrganizeFolder, colAllFiles, config)
    If fso.FolderExists(config.IOAttachmentFolder) Then Call Mod_Actions.GatherFilesRecursive(fso, config.IOAttachmentFolder, colAllFiles, config)
    
    Dim totalScanFiles As Long, currentScanIdx As Long
    totalScanFiles = colAllFiles.count: currentScanIdx = 0
    
    Dim file As Object, varFile As Variant
    For Each varFile In colAllFiles
        Set file = fso.GetFile(CStr(varFile))
        currentScanIdx = currentScanIdx + 1
        
        If currentScanIdx Mod 20 = 0 Then
            Dim progressMsg As String
            progressMsg = "正在執行畫布掛載與報表寫入 (" & currentScanIdx & "/" & totalScanFiles & ")"
            Application.StatusBar = progressMsg
            LogStore.Record "TX_PROGRESS", "REBUILD", progressMsg
            DoEvents
        End If
        
        If currentScanIdx Mod 60 = 0 Then
            config.TargetWB.Save
        End If
        On Error GoTo ProcessFileErrorHandler
        If Mod_Rules.IsImageExtension(fso.GetExtensionName(file.Name)) Then
            Call ProcessSingleFile(file, fso, catWs, catRow, wsSingle, Nothing, Context)
        Else
            attWs.Cells(attRow, 1).Value = UCase(fso.GetExtensionName(file.Name))
            attWs.Cells(attRow, 2).Value = file.Name
            attWs.Cells(attRow, 3).Value = Format(file.Size / 1024, "#,##0.0")
            attWs.Cells(attRow, 4).Value = Format(file.DateLastModified, "yyyy/mm/dd hh:mm")
            attWs.Cells(attRow, 5).Value = file.Path
            With attWs.Range(attWs.Cells(attRow, 1), attWs.Cells(attRow, 5))
                .Font.Name = config.FontMain: .Font.Size = config.FontSizeBase: .Borders.LineStyle = xlContinuous: .Borders.Color = config.Border: .VerticalAlignment = xlVAlignCenter: .EntireRow.RowHeight = 20
            End With
            attRow = attRow + 1
        End If
    On Error GoTo 0

ProcessFileNextItem:
    Next varFile
    
    If attRow = 5 Then
        attWs.Range("A5:E5").Merge: attWs.Range("A5").Value = "（目前專案內無任何非圖片之附件檔案）"
        attWs.Range("A5").HorizontalAlignment = xlCenter: attWs.Range("A5").Font.Color = config.ColorTextHint
    End If
    attWs.Columns("A:E").AutoFit: attWs.Rows.AutoFit
    Exit Sub

ProcessFileErrorHandler:
    LogStore.Record "TX_ERROR", "REBUILD", "單一檔案處理失敗，已跳過繼續下一個：" & file.Path & " | 原因：" & Err.Description
    Err.Clear
    Resume ProcessFileNextItem
End Sub

' --------------------------------------------------------------------------
' 救援夾掃描：檔名逆寫修正（方案4）＋掛載畫布，回傳處理好的救援工作表
' --------------------------------------------------------------------------
Private Function MountRescueFolderFiles(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal LogStore As Object, _
                                         ByVal catWs As Worksheet, ByRef catRow As Long, ByVal wsSingle As Worksheet) As Worksheet
    Dim wsRescue As Worksheet, file As Object
    Set wsRescue = Mod_Utils.GetOrCreateSheet(config.TargetWB, config.SheetNameRescue)
    wsRescue.Cells.Clear
    If fso.FolderExists(config.IORescueFolder) Then
        Dim rescueFolder As Object: Set rescueFolder = fso.GetFolder(config.IORescueFolder)
        For Each file In rescueFolder.Files
            If Mod_Rules.IsImageExtension(fso.GetExtensionName(file.Name)) Then
            
                ' 【方案4：檔名逆寫】實體檔名若跟血統紀錄該有的名稱不一致，主動改回去，
                ' 不讓使用者手滑改到的檔名持續存在，維持救援夾的檔名秩序
                Dim rescueFileHash As String
                Dim rescueCacheKey As String: rescueCacheKey = file.Path & "|" & Format(file.DateLastModified, "yyyy/mm/dd hh:mm:ss") & "|" & CStr(file.Size)
                If m_DictHashCache.Exists(rescueCacheKey) Then
                    rescueFileHash = m_DictHashCache(rescueCacheKey)
                Else
                    rescueFileHash = Mod_Utils.GetMD5(fso, file.Path)
                    m_DictHashCache(rescueCacheKey) = rescueFileHash
                End If

                If m_DictDNA.Exists(rescueFileHash) Then
                    Dim rescueDna As Variant: rescueDna = m_DictDNA(rescueFileHash)
                    If UBound(rescueDna) >= 3 Then
                        Dim shouldBeName As String: shouldBeName = Trim(CStr(rescueDna(3)))
                        Dim currentBaseName As String: currentBaseName = fso.GetBaseName(file.Name)
                        If shouldBeName <> "" And UCase(currentBaseName) <> UCase(shouldBeName) Then
                            Dim correctedPath As String
                            correctedPath = Mod_Utils.BuildSafeWindowsPath(fso, config.IORescueFolder, shouldBeName & "." & fso.GetExtensionName(file.Name))
                            Dim oldWrongPath As String: oldWrongPath = file.Path
                            
                            On Error Resume Next
                            
                            If Not fso.FileExists(correctedPath) Then
                               fso.CopyFile oldWrongPath, correctedPath, False
                               On Error GoTo 0

                                ' 核對複製結果：新檔案確實存在，且大小跟原檔一致，才刪除舊的錯誤命名檔案
                                If fso.FileExists(correctedPath) Then
                                    Dim origSize As Double, newSize As Double
                                    origSize = fso.GetFile(oldWrongPath).Size
                                    newSize = fso.GetFile(correctedPath).Size
                                    If origSize = newSize Then
                                        On Error Resume Next
                                        fso.DeleteFile oldWrongPath, True
                                        On Error GoTo 0
                                        Set file = fso.GetFile(correctedPath)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            On Error GoTo ProcessRescueFileErrorHandler
            Call ProcessSingleFile(file, fso, catWs, catRow, wsSingle, wsRescue, Context)
            On Error GoTo 0
            
ProcessRescueFileNextItem:
            End If
        Next file:
    End If
     
    Set MountRescueFolderFiles = wsRescue
    Exit Function

ProcessRescueFileErrorHandler:
    LogStore.Record "TX_ERROR", "REBUILD", "救援夾單一檔案處理失敗，已跳過繼續下一個：" & file.Path & " | 原因：" & Err.Description
    Err.Clear
    Resume ProcessRescueFileNextItem
End Function

' --------------------------------------------------------------------------
' 統計文字寫入＋逐一標準圖片工作表呼叫排版引擎與導覽按鈕渲染
' --------------------------------------------------------------------------
Private Sub RenderFinalLayoutsAndStats(ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal LogStore As Object, ByVal catWs As Worksheet, ByVal wsRescue As Worksheet)
    catWs.Range("E2").Value = "正常圖片：" & m_NewCount & " 張"
    catWs.Range("F2").Value = "救援圖片：" & m_RescueCount & " 張"
    With catWs.Range("E2:F2")
        .Font.Name = config.FontMain: .Font.Size = 9.5: .Font.Bold = True: .Font.Color = config.ColorTextMain: .VerticalAlignment = xlVAlignCenter
    End With
    
    ' 【局部重排】讀取操作介面 B6（寬度格）的 Locked 狀態，決定本輪是全量重排
    ' 還是只重排被 ProcessSingleFile／CleanInvalidShapes 摸過的表
    Dim isLayoutLocked As Boolean
    Dim wsUIForLock As Worksheet
    Set wsUIForLock = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameUI)
    If Not wsUIForLock Is Nothing Then
        On Error Resume Next
        isLayoutLocked = CBool(wsUIForLock.Range(config.CellImgWidth).Locked)
        On Error GoTo 0
    End If

    Dim wsLayout As Worksheet
    Dim layoutIdx As Long: layoutIdx = 0
    Dim shouldRenderLayout As Boolean
    Dim skippedLayoutCount As Long: skippedLayoutCount = 0
    For Each wsLayout In config.TargetWB.Worksheets
        If Mod_Rules.IsStandardImgSheet(wsLayout.Name, config) Or wsLayout.Name = config.SheetNameAllPics Or wsLayout.Name = config.SheetNameRescue Then
            shouldRenderLayout = True
            If isLayoutLocked And Not m_DictTouchedSheets.Exists(wsLayout.Name) Then shouldRenderLayout = False

            If shouldRenderLayout Then
                layoutIdx = layoutIdx + 1

                If layoutIdx Mod 20 = 0 Then
                    Application.StatusBar = "正在渲染工作表排版與按鈕 (" & layoutIdx & ")"
                    LogStore.Record "TX_PROGRESS", "REBUILD", "排版進度：" & layoutIdx
                    DoEvents
                End If
                If layoutIdx Mod 60 = 0 Then config.TargetWB.Save

                Call Mod_LayoutEngine.ExecuteMasonryLayout(wsLayout, Context)
                ' 【修正】導覽按鈕不在這裡渲染——CleanInvalidShapes 稍後會無條件清掉
                ' 本輪新畫的 AutoShape，這裡畫了也是白畫，改到 FinalizeRebuildCleanupAndFormatting
                ' 收尾階段、CleanInvalidShapes 執行「之後」統一補畫，見該函式
            Else
                skippedLayoutCount = skippedLayoutCount + 1
            End If
        End If
    Next wsLayout
    
    LogStore.Record "TX_INFO", "REBUILD", "局部重排：" & IIf(isLayoutLocked, "鎖定中", "未鎖定") & " | 重排 " & layoutIdx & " 表 | 跳過 " & skippedLayoutCount & " 表"
End Sub

' --------------------------------------------------------------------------
' 收尾：清孤兒圖形／清縮圖快取／清空資料夾／救援表Tab顏色／目錄表格式化／重整匯出選單
' --------------------------------------------------------------------------
Private Sub FinalizeRebuildCleanupAndFormatting(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal LogStore As Object, _
                                                 ByVal catWs As Worksheet, ByVal attWs As Worksheet, ByVal wsRescue As Worksheet)
' 【修正】CleanInvalidShapes 已提前到 ICommand_Execute 主流程、排版之前執行，
' 必須在這之後才補畫，否則按鈕會在同一輪 REBUILD 內被建立又立刻消失

    Dim wsNav As Worksheet
    For Each wsNav In config.TargetWB.Worksheets
        If Mod_Rules.IsStandardImgSheet(wsNav.Name, config) Or wsNav.Name = config.SheetNameAllPics Or wsNav.Name = config.SheetNameRescue _
           Or Left(wsNav.Name, Len(config.PrefixCustom)) = config.PrefixCustom Then
            Call Mod_UIReset.RenderSheetNavButtons(wsNav, config)
        End If
    Next wsNav
    
    Call Mod_ThumbnailCache.CleanOrphanedCache(fso, config)
    
    Dim emptyFolderCount As Long, emptyFolderPreview As String
    Call Mod_EmptyFolderCleaner.ScanAndRemoveEmptyFolders(fso, config, config.IOOrganizeFolder, emptyFolderCount, emptyFolderPreview, False)
    
    If m_RescueCount > 0 Then wsRescue.Tab.Color = config.UIColorLockRed Else wsRescue.Tab.ColorIndex = xlColorIndexNone
    
    catWs.Columns("A:K").AutoFit
    Call Mod_UIReset.RenderSheetNavButtons(catWs, config)
    Call Mod_UIReset.RenderSheetNavButtons(attWs, config)
    Call Mod_UIReset.UI_重整匯出選單
End Sub

' ------------------------------------------------------------------------------
' 輔助函數：處理單一硬碟檔案 (掛載與寫目錄，並恢復備份資料)
' ------------------------------------------------------------------------------
Private Sub ProcessSingleFile(ByVal file As Object, ByVal fso As Object, ByVal catWs As Worksheet, ByRef catRow As Long, ByVal wsSingle As Worksheet, ByVal wsRescueTarget As Worksheet, ByVal Context As cls_ExecutionContext)
    Dim config As cls_Config, fileHash As String, finalName As String
    Dim wsImg As Worksheet, shp As Shape
    Dim isRescueFile As Boolean
    
    Set config = Context.config
    finalName = fso.GetBaseName(file.Name)
    
    isRescueFile = DetermineIsRescueFile(file, fso, config, wsRescueTarget)

    If isRescueFile Then m_RescueCount = m_RescueCount + 1 Else m_NewCount = m_NewCount + 1

    fileHash = GetOrComputeFileHash(file, fso)

    ' 【格式統一】：非救援檔案（正常歸檔）要正式進整理夾前，強制洗成標準旗幟格式
    ' （拔除任何意外殘留的烙印、確保結尾有這張圖自己正確的雜湊尾巴）
    ' 救援檔案不動，因為它的正確身分本來就該帶著烙印，直到 ApplyFileMatch 正名那一刻才能拔
    Dim cleanedNameBeforeHash As String
    If Not isRescueFile Then
        finalName = NormalizeAndSyncPhysicalFilename(file, fso, finalName, fileHash, cleanedNameBeforeHash, Context.LogStore, config)
    End If

    ' --- 畫布掛載邏輯 ---
    ' 【V2-064修正】同名合併分組邏輯改讀洗雜湊「前」的乾淨名稱，不再讀被HASH尾巴污染的finalName
    Dim preHashMergeSource As String
    preHashMergeSource = IIf(Not isRescueFile, cleanedNameBeforeHash, finalName)
    Set wsImg = ResolveTargetSheet(fso, config, Context, isRescueFile, wsSingle, preHashMergeSource, finalName, file)

    Dim alreadyHasShape As Boolean
    Set shp = LocateExistingShapeByHash(config, fileHash)
    alreadyHasShape = Not shp Is Nothing
    
    ' 救援檔案的識別名稱，優先沿用血統延續紀錄，不採用現在的實體檔名
    ' （避免使用者手動改救援夾檔名後，目錄表被洗成新檔名、跟畫布上Shape名稱永久對不上）
    If isRescueFile Then
        finalName = ResolveRescueDisplayName(fileHash, finalName)
    End If

    If Not alreadyHasShape Then
        Dim previewPath As String
        previewPath = Mod_ThumbnailCache.GetOrCreateCachedPreview(fso, Context.config, file.Path)
        Set shp = wsImg.Shapes.AddPicture(fileName:=previewPath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=10, Top:=10, Width:=100, Height:=100)
        shp.Name = finalName
        m_DictShapeHash(wsImg.Name & "||" & finalName) = fileHash
        m_DictHashToShape(fileHash) = wsImg.Name & "||" & finalName
        shp.ScaleHeight 1, msoTrue: shp.ScaleWidth 1, msoTrue: shp.LockAspectRatio = msoTrue
        ' 【V2-062修正】救援夾檔案統一改用config.RescueEmbedWidth固定寬度，
        ' 不再用Context.DynamicImgWidth——避免同一張救援圖片，因為是被cmd_Rescue直接處理
        ' 還是被這裡重新掃描到，而套用兩種不同的尺寸標準
        If isRescueFile Then
            shp.Width = config.RescueEmbedWidth
        Else
            shp.Width = Context.DynamicImgWidth
        End If
        m_DictTouchedSheets(wsImg.Name) = True
   ElseIf shp.Name <> finalName Or shp.Parent.Name <> wsImg.Name Then
      ' 【V2-086修正】V2-074只處理「同一張工作表內改名」，沒處理「改名連帶
      ' 讓歸屬工作表也變了」（例如同名合併模式下，改檔名讓分組結果變成另一張表）。
      ' 原邏輯只會原地把shp改名，字典卻登記成新算出的目標表wsImg——兩者不一致時，
      ' Shape實際位置跟字典記錄位置對不上，下一輪CleanInvalidShapes判定孤兒直接刪除。
      ' 這正是「改檔名+搬子資料夾」測試情境下圖片消失的具體根因
      If shp.Parent.Name = wsImg.Name Then
          ' 同一張工作表，單純改名即可，原邏輯不變
          Dim oldShapeKey As String: oldShapeKey = wsImg.Name & "||" & shp.Name
          On Error Resume Next
          shp.Name = finalName
          If Err.Number = 0 Then
              m_DictShapeHash.Remove oldShapeKey
              m_DictShapeHash(wsImg.Name & "||" & finalName) = fileHash
              m_DictHashToShape(fileHash) = wsImg.Name & "||" & finalName
              m_DictTouchedSheets(wsImg.Name) = True
          Else
              Context.LogStore.Record "TX_ERROR", "REBUILD", "Shape名稱同步失敗（可能撞名），目錄與畫布短暫不一致：" & shp.Name & " → " & finalName
          End If
          On Error GoTo 0
      Else
          ' 歸屬工作表不同，單純改名不夠，必須真正搬移——比照新建Shape的既有
          ' AddPicture邏輯，從縮圖快取重新嵌入到正確工作表，確認新的建立成功
          ' 後才刪除舊的（複製優先於刪除，P0第2條）
          Dim oldShapeKeyMove As String: oldShapeKeyMove = shp.Parent.Name & "||" & shp.Name
          Dim oldSheetNameForTouch As String: oldSheetNameForTouch = shp.Parent.Name
          Dim relocPreviewPath As String
          relocPreviewPath = Mod_ThumbnailCache.GetOrCreateCachedPreview(fso, Context.config, file.Path)
          Dim newShp As Shape
          Set newShp = Nothing
          On Error Resume Next
          Set newShp = wsImg.Shapes.AddPicture(fileName:=relocPreviewPath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=10, Top:=10, Width:=100, Height:=100)
          If Not newShp Is Nothing Then newShp.Name = finalName
          On Error GoTo 0

          If Not newShp Is Nothing And Err.Number = 0 Then
              newShp.ScaleHeight 1, msoTrue: newShp.ScaleWidth 1, msoTrue: newShp.LockAspectRatio = msoTrue
              If isRescueFile Then
                  newShp.Width = config.RescueEmbedWidth
              Else
                  newShp.Width = Context.DynamicImgWidth
              End If
              m_DictShapeHash.Remove oldShapeKeyMove
              m_DictShapeHash(wsImg.Name & "||" & finalName) = fileHash
              m_DictHashToShape(fileHash) = wsImg.Name & "||" & finalName
              m_DictTouchedSheets(wsImg.Name) = True
              m_DictTouchedSheets(oldSheetNameForTouch) = True
              ' 確認新的已成功嵌入到正確工作表後，才刪除舊工作表上的原Shape
              shp.Delete
          Else
              Context.LogStore.Record "TX_ERROR", "REBUILD", "Shape跨工作表搬移失敗，原Shape保留於" & shp.Parent.Name & "：" & shp.Name & " → " & wsImg.Name & "||" & finalName
          End If
      End If
    End If
        
    Call WriteCatalogRow(catWs, catRow, finalName, fileHash, file, config, isRescueFile, shp)
End Sub

' --------------------------------------------------------------------------
' 【職責 1】判斷檔案是否為救援檔案（路徑比對，或呼叫端明確指定 wsRescueTarget）
' --------------------------------------------------------------------------
Private Function DetermineIsRescueFile(ByVal file As Object, ByVal fso As Object, ByVal config As cls_Config, ByVal wsRescueTarget As Worksheet) As Boolean
    Dim parentPath As String, rescuePath As String
    parentPath = UCase(fso.GetParentFolderName(file.Path)): If Right(parentPath, 1) <> "\" Then parentPath = parentPath & "\"
    rescuePath = UCase(config.IORescueFolder): If Right(rescuePath, 1) <> "\" Then rescuePath = rescuePath & "\"
    DetermineIsRescueFile = (parentPath = rescuePath Or Not wsRescueTarget Is Nothing)
End Function

' --------------------------------------------------------------------------
' 【職責 2】雜湊取得：優先查快取，查無則現算並寫入快取，並登記合法通行證字典
' --------------------------------------------------------------------------
Private Function GetOrComputeFileHash(ByVal file As Object, ByVal fso As Object) As String
    Dim cacheKey As String, fileHash As String
    cacheKey = file.Path & "|" & Format(file.DateLastModified, "yyyy/mm/dd hh:mm:ss") & "|" & CStr(file.Size)
    
    'Debug.Print "[快取除錯] " & cacheKey & " | 命中=" & m_DictHashCache.Exists(cacheKey)                                   ' 先註銷
    
    If m_DictHashCache.Exists(cacheKey) Then
        fileHash = m_DictHashCache(cacheKey)
        m_HashCacheHit = m_HashCacheHit + 1
    Else
        fileHash = Mod_Utils.GetMD5(fso, file.Path)
        m_DictHashCache(cacheKey) = fileHash
        m_HashCacheMiss = m_HashCacheMiss + 1
    End If
    
    ' 【身分登記改用雜湊】：不管檔名是不是被手動改過，只要內容雜湊對得上就算同一張圖
    m_DictValidShapes(fileHash) = True
    
    GetOrComputeFileHash = fileHash
End Function

' --------------------------------------------------------------------------
' 【職責 3】非救援檔案的格式統一清洗＋實體檔案改名同步。file 以 ByRef 傳入，
' 若實體檔案被改名，file 會被更新指向新檔案，呼叫端必須沿用更新後的物件
' --------------------------------------------------------------------------
Private Function NormalizeAndSyncPhysicalFilename(ByRef file As Object, ByVal fso As Object, ByVal baseName As String, ByVal fileHash As String, ByRef cleanedNameBeforeHash As String, ByVal LogStore As Object, ByVal config As cls_Config) As String
   ' 【V2.3.0】複製貼上產生的「- 複製」文字與括號流水號一律直接洗掉，不轉換保留——
   ' 反正結尾一定會重新貼上依實際內容算出的HASH尾巴，數字本身不承擔防撞名責任。
   ' 兩支底層工具（StripTrailingParenNumber／StripKnownCopySuffix）各自單一職責、
   ' 只查一次，交錯清洗的順序與輪數由這裡的迴圈決定，不藏進底層函式裡
   Dim cleanedName As String: cleanedName = baseName
   Dim copyStripLoop As Integer: copyStripLoop = 0
   Do
        Dim beforeCopyStrip As String: beforeCopyStrip = cleanedName
        cleanedName = Mod_Utils.StripTrailingParenNumber(cleanedName)
        cleanedName = Mod_Utils.StripKnownCopySuffix(cleanedName, config.CopySuffixPatterns)
        copyStripLoop = copyStripLoop + 1
   Loop While cleanedName <> beforeCopyStrip And copyStripLoop < 30
   cleanedName = Mod_Utils.ExtractOriginalIdentifier(cleanedName)
   ' 不再用「結尾對不對」判斷要不要清洗，一律無條件拔到底、重新貼上正確雜湊
   ' 避免結尾剛好碰巧正確、卻漏掉中間卡著的舊雜湊殘留
   Dim safetyLoop As Integer: safetyLoop = 0
   Do While True
        Dim beforeStrip As String: beforeStrip = cleanedName
        cleanedName = Mod_Utils.StripHashSuffix(cleanedName)
        If cleanedName = beforeStrip Then Exit Do
        safetyLoop = safetyLoop + 1
        If safetyLoop >= 10 Then Exit Do
   Loop
       
   '【V2-064新增】_HASH尾巴加上去之前，先記下這個乾淨版本，供同名合併分組用
    cleanedNameBeforeHash = cleanedName
    cleanedName = cleanedName & "_#HASH" & Left(fileHash, 8)
    
    ' 【實體同步】：目錄/Shape已經算出乾淨名稱，硬碟上的實體檔案也要跟著改名，
    ' 否則目錄表跟畫布看起來乾淨，但Windows總管看到的檔名依然是骯髒的原樣
    If cleanedName <> baseName Then
        Dim cleanedPhysicalPath As String
        cleanedPhysicalPath = Mod_Utils.BuildSafeWindowsPath(fso, fso.GetParentFolderName(file.Path), cleanedName & "." & fso.GetExtensionName(file.Name))
        Dim oldDirtyPath As String: oldDirtyPath = file.Path

        If Not fso.FileExists(cleanedPhysicalPath) Then
            On Error Resume Next
            fso.CopyFile oldDirtyPath, cleanedPhysicalPath, False
            On Error GoTo 0
 
            If fso.FileExists(cleanedPhysicalPath) Then
                If fso.GetFile(oldDirtyPath).Size = fso.GetFile(cleanedPhysicalPath).Size Then
                    On Error Resume Next
                    fso.DeleteFile oldDirtyPath, True
                    On Error GoTo 0
                    Set file = fso.GetFile(cleanedPhysicalPath)
                Else
                    LogStore.Record "TX_ERROR", "REBUILD", "檔名正規化：複製後大小不符，實體未改名，仍為舊檔名：" & oldDirtyPath & " | 目錄/Shape將使用新名稱：" & cleanedName
                End If
             Else
                LogStore.Record "TX_ERROR", "REBUILD", "檔名正規化：複製失敗，實體未改名，仍為舊檔名：" & oldDirtyPath & " | 目錄/Shape將使用新名稱：" & cleanedName
             End If
         End If
     End If
    
    NormalizeAndSyncPhysicalFilename = cleanedName
End Function

' --------------------------------------------------------------------------
' 【職責 4】依模式（救援/單頁/資料夾合併/同名合併/預設）決定目標工作表
' --------------------------------------------------------------------------
Private Function ResolveTargetSheet(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal isRescueFile As Boolean, _
                                     ByVal wsSingle As Worksheet, ByVal preHashMergeSource As String, ByVal finalName As String, ByVal file As Object) As Worksheet
    Dim wsImg As Worksheet
    If isRescueFile Then
        Set wsImg = Mod_Utils.GetOrCreateSheet(config.TargetWB, config.SheetNameRescue)
    ElseIf Context.IsSingleMode Then
        Set wsImg = wsSingle
    ElseIf Context.IsFolderMergeMode Then
        Dim pName As String: pName = GetTopLevelFolderName(fso, file.Path, config)
        pName = Mod_Utils.SanitizeSheetName(pName, config)
        Set wsImg = Mod_Utils.GetOrCreateSheet(config.TargetWB, pName)
    ElseIf Context.IsMergeMode Then
        Dim mergeGroupName As String
        Dim lastUnderscorePos As Long: lastUnderscorePos = InStrRev(preHashMergeSource, "_")
        If lastUnderscorePos > 1 Then
            mergeGroupName = Left(preHashMergeSource, lastUnderscorePos - 1)
        Else
            mergeGroupName = preHashMergeSource
        End If
        mergeGroupName = Mod_Utils.SanitizeSheetName(mergeGroupName, config)
        Set wsImg = Mod_Utils.GetOrCreateSheet(config.TargetWB, mergeGroupName)
    Else
        Dim safeSingleName As String
        safeSingleName = Mod_Utils.SanitizeSheetName(finalName, config)
        Set wsImg = Mod_Utils.GetOrCreateSheet(config.TargetWB, safeSingleName)
    End If
    Set ResolveTargetSheet = wsImg
End Function

' --------------------------------------------------------------------------
' 【職責 5】透過雜湊查找既有 Shape，查無或索引失效時回傳 Nothing
' --------------------------------------------------------------------------
Private Function LocateExistingShapeByHash(ByVal config As cls_Config, ByVal fileHash As String) As Shape
    Dim shp As Shape
    If m_DictHashToShape.Exists(fileHash) Then
        Dim keyParts() As String: keyParts = Split(CStr(m_DictHashToShape(fileHash)), "||")
        On Error Resume Next
        Set shp = config.TargetWB.Worksheets(keyParts(0)).Shapes(keyParts(1))
        If Err.Number <> 0 Then Set shp = Nothing
        On Error GoTo 0
    End If
    Set LocateExistingShapeByHash = shp
End Function

' --------------------------------------------------------------------------
' 【職責 6】救援檔案識別名稱：血統延續紀錄優先，查無則沿用傳入的預設名稱
' --------------------------------------------------------------------------
Private Function ResolveRescueDisplayName(ByVal fileHash As String, ByVal fallbackName As String) As String
    ResolveRescueDisplayName = fallbackName
    If m_DictDNA.Exists(fileHash) Then
        Dim dnaPreserved As Variant: dnaPreserved = m_DictDNA(fileHash)
        If UBound(dnaPreserved) >= 3 Then
            If Trim(CStr(dnaPreserved(3))) <> "" Then ResolveRescueDisplayName = Trim(CStr(dnaPreserved(3)))
        End If
    End If
End Function

' --------------------------------------------------------------------------
' 【職責 8】寫入目錄表（11 欄資料＋Hover 預覽＋格式化），已寫入過則跳過
' --------------------------------------------------------------------------
Private Sub WriteCatalogRow(ByVal catWs As Worksheet, ByRef catRow As Long, ByVal finalName As String, ByVal fileHash As String, ByVal file As Object, ByVal config As cls_Config, ByVal isRescueFile As Boolean, ByVal shp As Shape)
    If Not m_DictCatWritten.Exists(finalName) Then
        catWs.Cells(catRow, 1).Value = finalName
        catWs.Cells(catRow, 2).Value = fileHash
        catWs.Cells(catRow, 3).Value = file.Path
        catWs.Cells(catRow, 4).Value = file.Size
        catWs.Cells(catRow, 5).Value = Format(Now, "yyyy/mm/dd hh:mm")
        catWs.Cells(catRow, 6).Value = Format(file.DateLastModified, "yyyy/mm/dd hh:mm:ss")
        
        If m_DictDNA.Exists(fileHash) Then
            Dim dnaData As Variant: dnaData = m_DictDNA(fileHash)

            catWs.Cells(catRow, 7).Value = dnaData(0)
            catWs.Cells(catRow, 8).Value = dnaData(1)
            catWs.Cells(catRow, 9).Value = dnaData(2)
        Else
            catWs.Cells(catRow, 7).Value = IIf(isRescueFile, "系統救援", "正常歸檔")
            catWs.Cells(catRow, 8).Value = IIf(isRescueFile, "是", "否")
            catWs.Cells(catRow, 9).Value = IIf(isRescueFile, fileHash, "")
        End If
        
        catWs.Cells(catRow, 10).Value = "檔案正常"
        catWs.Cells(catRow, 11).Value = Mod_Utils.StripHashSuffix(finalName)
        
        Call Mod_Utils.AddHoverPreview(catWs.Cells(catRow, 1), file.Path, IIf(Not shp Is Nothing, shp.Width, 150), IIf(Not shp Is Nothing, shp.Height, 150))
        
        With catWs.Range(catWs.Cells(catRow, 1), catWs.Cells(catRow, 11))
            .Font.Name = config.FontMain: .Font.Size = config.FontSizeBase: .Borders.LineStyle = xlContinuous: .Borders.Color = config.Border: .VerticalAlignment = xlVAlignCenter: .EntireRow.RowHeight = 20
            If isRescueFile Then
                .Interior.Color = config.ThemeManual: .Font.Color = config.ColorWhite
            Else
                .Interior.Color = config.ColorWhite: .Font.Color = config.ColorTextMain
            End If
        End With
        
        m_DictCatWritten(finalName) = True
        catRow = catRow + 1
    End If
End Sub

' ------------------------------------------------------------------------------
' 輔助函數：備份目錄狀態 (BackupCatalogState) - 防止清除重建時遺失既有資料
' ------------------------------------------------------------------------------
Private Sub BackupCatalogState(ByVal Context As cls_ExecutionContext)

    Dim config As cls_Config: Set config = Context.config
    Dim catWs As Worksheet: Set catWs = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameCatalog)
    
    Set m_DictCatNameToHash = CreateObject("Scripting.Dictionary"): m_DictCatNameToHash.CompareMode = 1
    
    If catWs Is Nothing Then Exit Sub
    
    Dim cRow As Long, lastC As Long
    Dim cName As String, cHash As String, cPath As String, cTime As String, cSize As String
    Dim cStatus As String, cIsRescued As String, cOrigHash As String
    Dim cacheKey As String
    
    lastC = catWs.Cells(catWs.Rows.count, "A").End(xlUp).row
    If lastC < 5 Then Exit Sub
    
    For cRow = 5 To lastC
        cName = Trim(CStr(catWs.Cells(cRow, 1).Value))
        cHash = Trim(CStr(catWs.Cells(cRow, 2).Value))
        cPath = Trim(CStr(catWs.Cells(cRow, 3).Value))
        cSize = Trim(CStr(catWs.Cells(cRow, 4).Value))
        cTime = ""
        If IsDate(catWs.Cells(cRow, 6).Value) Then cTime = Format(CDate(catWs.Cells(cRow, 6).Value), "yyyy/mm/dd hh:mm:ss")
        
        cStatus = Trim(CStr(catWs.Cells(cRow, 7).Value))
        cIsRescued = Trim(CStr(catWs.Cells(cRow, 8).Value))
        cOrigHash = Trim(CStr(catWs.Cells(cRow, 9).Value))
        
        ' BackupCatalogState
        If cHash <> "" Then
            m_DictDNA(cHash) = Array(cStatus, cIsRescued, cOrigHash, cName)
        End If
        If cName <> "" And cHash <> "" Then m_DictCatNameToHash(cName) = cHash
        
        If cHash <> "" And cPath <> "" And cTime <> "" And cSize <> "" Then
            cacheKey = cPath & "|" & cTime & "|" & cSize
            m_DictHashCache(cacheKey) = cHash
        End If
    Next cRow
End Sub

' ------------------------------------------------------------------------------
' 輔助函數：清除畫布上的無效圖形 (CleanInvalidShapes)
' ------------------------------------------------------------------------------
Private Sub CleanInvalidShapes(ByVal config As cls_Config, ByVal LogStore As Object)
    Dim ws As Worksheet, sIdx As Long, shpClean As Shape
    Dim cleanedCount As Long: cleanedCount = 0
    For Each ws In config.TargetWB.Worksheets
        If Mod_Rules.IsStandardImgSheet(ws.Name, config) Or ws.Name = config.SheetNameAllPics Or ws.Name = config.SheetNameRescue Then
            For sIdx = ws.Shapes.count To 1 Step -1
                Set shpClean = ws.Shapes(sIdx)
                If shpClean.Type = msoPicture Or shpClean.Type = msoLinkedPicture Then
                    Dim cleanKey As String: cleanKey = ws.Name & "||" & shpClean.Name
                    Dim isStillValid As Boolean: isStillValid = False
                    If m_DictShapeHash.Exists(cleanKey) Then
                           If m_DictValidShapes.Exists(m_DictShapeHash(cleanKey)) Then isStillValid = True
                    End If
                    
                     ' 【補漏】即使此輪掃描不到對應實體檔案，若目錄表記錄此Shape目前正處於
                     ' 「等待救援中」狀態（isBounty=是），不能視為孤兒直接刪除——這是合法的
                     ' 暫時性狀態，否則cmd_Rescue辛苦保留下來的圖形，會在REBUILD這一步被悄悄清掉
                     If Not isStillValid Then
                         Dim isPendingRescue As Boolean: isPendingRescue = False
                         If m_DictCatNameToHash.Exists(shpClean.Name) Then
                             Dim checkHash As String: checkHash = m_DictCatNameToHash(shpClean.Name)
                             If m_DictDNA.Exists(checkHash) Then
                                 If m_DictDNA(checkHash)(1) = "是" Then isPendingRescue = True
                             End If
                         End If
                         
                          If Not isPendingRescue Then
                              ' 【V2-087新增】此前這裡完全沒有任何記錄，一旦誤刪，
                              ' LOG查不到任何線索。補上完整診斷資訊：Shape名稱、
              ' 所在工作表、目錄表有沒有查到這個名字、有沒有找到對應雜湊、
              ' 這個雜湊這輪掃描算不算「有效」——足夠事後回溯判斷誤刪原因
                              Dim diagInCatalog As String: diagInCatalog = IIf(m_DictCatNameToHash.Exists(shpClean.Name), "是(雜湊=" & m_DictCatNameToHash(shpClean.Name) & ")", "否")
                              Dim diagInShapeHash As String: diagInShapeHash = IIf(m_DictShapeHash.Exists(cleanKey), "是", "否")
                              LogStore.Record "TX_ERROR", "REBUILD", "CleanInvalidShapes判定孤兒並刪除：工作表=" & ws.Name & _
                                  " | 名稱=" & shpClean.Name & " | 目錄表查得到嗎=" & diagInCatalog & _
                                  " | Shape雜湊索引查得到嗎=" & diagInShapeHash
                              cleanedCount = cleanedCount + 1
                              m_DictTouchedSheets(ws.Name) = True
                              shpClean.Delete
                          End If

                     End If
                ElseIf shpClean.Type = msoAutoShape Or shpClean.Type = msoTextBox Then
                    shpClean.Delete
                End If
            Next sIdx
        End If
    Next ws
    If cleanedCount > 0 Then
        LogStore.Record "TX_INFO", "REBUILD", "CleanInvalidShapes本輪共清除 " & cleanedCount & " 個孤兒圖形，詳見上方個別記錄"
    End If
End Sub

' ------------------------------------------------------------------------------
' 內部重構：修正後的初始化目錄與附件工作表
' ------------------------------------------------------------------------------
Private Function InitializeCatalogSheet(ByVal config As cls_Config, ByVal Context As cls_ExecutionContext) As Worksheet
    Dim catWs As Worksheet, backupName As String, shp As Shape, oldWs As Worksheet
    Set catWs = Mod_Utils.GetOrCreateSheet(config.TargetWB, config.SheetNameCatalog)
    backupName = ""
    
    If Context.isCatalogAppend And catWs.Range("A4").Value <> "" Then
        backupName = config.PrefixOldCatalog & Format(Now, "mmdd_hhmmss")
        catWs.Copy After:=config.TargetWB.Sheets(config.TargetWB.Sheets.count)
        Set oldWs = config.TargetWB.Sheets(config.TargetWB.Sheets.count)
        oldWs.Name = backupName
        oldWs.Cells(2, 3).Formula = "=HYPERLINK(""#'" & config.SheetNameCatalog & "'!A1"", "" ▋前往最新圖片目錄 "")"
    End If
    
    catWs.Cells.Clear
    For Each shp In catWs.Shapes
        If shp.Type = msoAutoShape Or shp.Type = msoTextBox Then On Error Resume Next: shp.Delete: On Error GoTo 0
    Next shp
    
    catWs.Range("A1:K1").Merge: catWs.Range("A1").Value = "IMAGE ALBUM | 專案自動化圖片歸檔目錄"
    catWs.Range("A1").Font.Name = config.FontMain: catWs.Range("A1").Font.Size = config.FontSizeSub: catWs.Range("A1").Font.Bold = True: catWs.Range("A1").Font.Color = config.ColorWhite
    catWs.Range("A1:K1").Interior.Color = config.ThemeIO: catWs.Range("A1:K1").RowHeight = 32: catWs.Range("A1").VerticalAlignment = xlVAlignCenter
    
    catWs.Rows("2:2").RowHeight = 22: catWs.Rows("3:3").RowHeight = 0
    catWs.Cells(2, 1).Formula = "=HYPERLINK(""#'操作介面'!A1"", "" ▋返回控制台 "")"
    If backupName <> "" Then catWs.Cells(2, 3).Formula = "=HYPERLINK(""#'" & backupName & "'!A1"", "" ▋檢視歷史備份快照 "")"
    
    With catWs.Range("A2:K2")
        .Font.Name = config.FontMain: .Font.Size = 10: .Font.Bold = True: .VerticalAlignment = xlVAlignCenter: .Interior.Color = RGB(240, 244, 248)
    End With
    catWs.Cells(2, 1).Font.Color = config.ColorLinkText: catWs.Cells(2, 3).Font.Color = config.ColorLinkText
    
    catWs.Range("A4:K4").Value = Array("圖片編號(滑鼠移入預覽)", "HASH 特徵碼", "目前物理路徑", "檔案大小", "登錄時間", "檔案修改時間", "管線執行狀態", "是否觸發救援", "指紋動態對齊", "實體狀態", "顯示名稱")
    With catWs.Range("A4:K4")
        .Interior.Color = config.ThemeExp: .Font.Color = config.ColorWhite: .Font.Name = config.FontMain: .Font.Bold = True: .RowHeight = 22: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlVAlignCenter
    End With
    Set InitializeCatalogSheet = catWs
End Function

Private Function InitializeAttachmentSheet(ByVal config As cls_Config) As Worksheet
    Dim attWs As Worksheet, shp As Shape
    Set attWs = Mod_Utils.GetOrCreateSheet(config.TargetWB, config.SheetNameAttachment)
    attWs.Cells.Clear
    
    For Each shp In attWs.Shapes
        If shp.Type = msoAutoShape Or shp.Type = msoTextBox Then On Error Resume Next: shp.Delete: On Error GoTo 0
    Next shp
    
    attWs.Range("A1:E1").Merge: attWs.Range("A1").Value = "IMAGE ALBUM | 專案附件歸檔目錄"
    With attWs.Range("A1"): .Font.Name = config.FontMain: .Font.Size = config.FontSizeSub: .Font.Bold = True: .Font.Color = config.ColorWhite: End With
    With attWs.Range("A1:E1"): .Interior.Color = config.ThemeIO: .RowHeight = 32: .VerticalAlignment = xlVAlignCenter: End With
    
    attWs.Rows("2:2").RowHeight = 22: attWs.Rows("3:3").RowHeight = 0
    attWs.Cells(2, 1).Formula = "=HYPERLINK(""#'操作介面'!A1"", "" ▋返回控制台 "")"
    With attWs.Range("A2:E2")
        .Font.Name = config.FontMain: .Font.Size = 10: .Font.Bold = True: .VerticalAlignment = xlVAlignCenter: .Interior.Color = RGB(240, 244, 248)
    End With
    attWs.Cells(2, 1).Font.Color = config.ColorLinkText
    
    attWs.Range("A4:E4").Value = Array("副檔名", "檔案名稱", "檔案大小 (KB)", "修改時間", "實體路徑")
    With attWs.Range("A4:E4")
        .Interior.Color = config.ThemeExp: .Font.Color = config.ColorWhite: .Font.Name = config.FontMain: .Font.Bold = True: .RowHeight = 22: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlVAlignCenter
    End With
    Set InitializeAttachmentSheet = attWs
End Function

Private Function GetTopLevelFolderName(ByVal fso As Object, ByVal filePath As String, ByVal config As cls_Config) As String
    Dim orgFolder As String: orgFolder = config.IOOrganizeFolder
    If Right(orgFolder, 1) <> "\" Then orgFolder = orgFolder & "\"
    
    Dim relativePath As String
    If Left(UCase(filePath), Len(orgFolder)) = UCase(orgFolder) Then
        relativePath = Mid(filePath, Len(orgFolder) + 1)
    Else
        GetTopLevelFolderName = "未分類"
        Exit Function
    End If
    
    Dim slashPos As Long: slashPos = InStr(relativePath, "\")
    If slashPos > 0 Then
        GetTopLevelFolderName = Left(relativePath, slashPos - 1)
    Else
        ' 【野生檔案兜底】：使用者繞過匯入流程、直接把檔案塞進整理夾根目錄，沒有任何資料夾包著
        ' 這種情況不分別開表，統一歸進固定分類，避免每個野生檔案各自變成獨立一張表
        GetTopLevelFolderName = "未分類"
    End If
End Function


' cmd_Rebuild.cls，ICommand_Execute 開頭，InitializeCatalogSheet 之前呼叫
Private Sub ConsolidateLooseFilesInOrganizeRoot(ByVal fso As Object, ByVal config As cls_Config)
    On Error Resume Next   ' 這段本身失敗也不影響後續排版，退回既有的「未分類」機制即可
    
    If Not fso.FolderExists(config.IOOrganizeFolder) Then Exit Sub
    
    Dim rootFld As Object, looseFile As Object
    Set rootFld = fso.GetFolder(config.IOOrganizeFolder)
    Call Mod_Utils.EnsureFolderExists(fso, config.IOLooseImportFolder)
    
    Dim colLoose As New Collection
    For Each looseFile In rootFld.Files
        colLoose.Add looseFile.Path
    Next looseFile
    
    Dim vFile As Variant
    For Each vFile In colLoose
        Dim destPath As String
        destPath = Mod_Utils.BuildSafeWindowsPath(fso, config.IOLooseImportFolder, fso.GetFileName(CStr(vFile)))
        Err.Clear
        Call Mod_Utils.CopyPhysicalFile(fso, CStr(vFile), destPath)
        If Err.Number = 0 Then fso.DeleteFile CStr(vFile), True
    Next vFile
    
    Set rootFld = Nothing: Set colLoose = Nothing
    On Error GoTo 0
End Sub
