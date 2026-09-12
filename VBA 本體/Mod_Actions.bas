Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_Actions
' PURPOSE: 專職負責與作業系統、硬碟實體 IO 或 Excel 畫布進行互動。
'          實裝開機真值表防禦、延遲建置輕量化，並解決 OS 目錄快取延遲。
' EXPORTS: InitializeSystemFolders, IsEnvironmentHealthy, GetBootQuadrant, ExecuteProactiveRescueAndSnapshot, 等
' IMPORTS: cls_Config, Mod_Utils, Mod_Rules, Mod_GDIPlusExport, Mod_UIReset, cls_Log
' FORBIDDEN: 1. 嚴禁在此撰寫任何 UI 彈窗 (MsgBox/AskQuestion)！決策必須交還 Pipeline。
'            2. 嚴禁預先生成無內容物的空白子夾。
' DEPENDENCIES: Windows Scripting Host (FSO, Shell)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_Actions"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
' 區塊 1：基礎實體目錄結構建置與檔案蒐集 (Physical Environment Init)
' ==============================================================================

' ------------------------------------------------------------------------------
' [實體建造]：InitializeSystemFolders
' 目的：開機時或環境修復時，僅建立「絕對必要」的平級三柱核心結構。
' 防禦：加入 DoEvents 強制驅動 Windows 釋放佇列，讓作業系統有時間刷新 MFT。
' ------------------------------------------------------------------------------
Public Sub InitializeSystemFolders(ByVal config As cls_Config)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 【最終防線】：不管呼叫端是誰、有沒有自己先檢查過，這裡是實際動手寫入硬碟的唯一入口，
   ' 危險區一律靜默拒絕，不主動彈窗（本模組 FORBIDDEN 條款禁止在此彈窗，交由呼叫端自行決定要不要告知使用者）。
   If Mod_Rules.IsDangerousPath(config.BaseFolder) Then
       Set fso = Nothing
       Exit Sub
   End If
 
    ' 【V2-051】統一改讀 GetAllManagedFolderPaths，此清單同時是 IsEnvironmentHealthy 的唯一事實來源，
    ' 兩處不再各自維護一份清單，避免「這裡建了、那裡沒檢查」或反過來的清單漂移
    Dim folderPath As Variant
    For Each folderPath In GetAllManagedFolderPaths(config)
        Call Mod_Utils.EnsureFolderExists(fso, CStr(folderPath))
    Next folderPath

    ' 【隱藏屬性】：僅此二者需要額外設定 Windows 隱藏屬性，非全體資料夾適用，維持獨立於共用清單之外
    Call Mod_Utils.SetFolderHidden(fso, config.IOSandboxFolder)
    Call Mod_Utils.SetFolderHidden(fso, config.IOThumbnailCacheFolder)
    
    ' 【延遲防線】：短暫延遲 0.2 秒。防止 FSO 非同步建立資料夾後，後續檢驗瞬間讀取不到引發假警報。
    Dim tSleep As Single: tSleep = Timer
    Do While Abs(Timer - tSleep) < 0.2: DoEvents: Loop
    
    Set fso = Nothing
End Sub

Public Sub InitializeExportFolders(ByVal config As cls_Config)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call Mod_Utils.EnsureFolderExists(fso, config.IOExportFolder)
    Call Mod_Utils.EnsureFolderExists(fso, config.IOExportCatalogFolder)
    Call Mod_Utils.EnsureFolderExists(fso, config.IOExportImageFolder)
    Set fso = Nothing
End Sub

Public Sub GatherFilesRecursive(ByVal fso As Object, ByVal currentPath As String, ByRef sourceFiles As Collection, ByVal config As cls_Config)
    Dim folder As Object, subFolder As Object, file As Object
    If Mod_Rules.IsProtectedFolder(currentPath, config) Then Exit Sub
    
    On Error Resume Next
    Set folder = fso.GetFolder(currentPath)
    If folder Is Nothing Then Exit Sub
    
    For Each file In folder.Files: sourceFiles.Add file.Path: Next file
    For Each subFolder In folder.SubFolders: Call GatherFilesRecursive(fso, subFolder.Path, sourceFiles, config): Next subFolder
    On Error GoTo 0
End Sub

' ==============================================================================
' 區塊 2：核心環境智慧探測 (State Probes - 回傳狀態，由 Pipeline 決策)
' ==============================================================================
 ' ------------------------------------------------------------------------------
 ' [SSOT清單]：GetAllManagedFolderPaths
 ' 目的：全系統唯一維護的「系統自行管理資料夾」清單。InitializeSystemFolders（建立）
 '      與 IsEnvironmentHealthy（檢查）皆改讀此清單，新增/移除資料夾只需改這一處。
 ' 排除項：IOExportCatalogFolder／IOExportImageFolder 刻意不列入——此二者為
 '      Lazy Loading（延遲生成），僅在使用者實際執行匯出時建立，不屬於開機骨架範疇。
 ' ------------------------------------------------------------------------------
 Public Function GetAllManagedFolderPaths(ByVal config As cls_Config) As Collection
     Dim col As New Collection

     col.Add config.BaseFolder
     col.Add config.IOExportFolder

     col.Add config.IOTargetFolder
     col.Add config.IOOrganizeFolder
     col.Add config.IORescueFolder
     col.Add config.IORescueDowncastFolder
     col.Add config.IORescueDiscardFolder
     col.Add config.IODuplicateDiscardFolder
     col.Add config.IOLooseImportFolder
     col.Add config.IOThumbnailCacheFolder
     col.Add config.IOSandboxFolder

     col.Add config.IOAttachmentFolder
     col.Add config.IOAttachmentDuplicateDiscardFolder  ' 【V2-051補漏】此前未被建立過

     col.Add config.IOQuarantineFolder

     col.Add config.IOArchiveFolder
     col.Add config.IOBackupSnapshotFolder
     col.Add config.IOLogArchiveFolder

     Set GetAllManagedFolderPaths = col
 End Function
 
' ------------------------------------------------------------------------------
' [狀態探測]：IsEnvironmentHealthy
' 目的：專供管線啟動前呼叫的狀態確認。純粹回傳 True/False，不彈窗！
' ------------------------------------------------------------------------------

Public Function IsEnvironmentHealthy(ByRef config As cls_Config) As Boolean
      Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
      Dim folderPath As Variant
      Dim allExist As Boolean: allExist = True

     For Each folderPath In GetAllManagedFolderPaths(config)
          If Not fso.FolderExists(CStr(folderPath)) Then
              allExist = False
              Exit For
          End If
     Next folderPath

      IsEnvironmentHealthy = allExist
      Set fso = Nothing
 End Function

' ------------------------------------------------------------------------------
' [狀態探測]：GetBootQuadrant
' 目的：對齊 4 象限開機判斷表。精準區分「改名」、「誤刪」與「新專案」。
' 回傳值：1 (健康吻合), 2 (健康被改名), 3 (殘缺吻合-誤刪), 4 (殘缺改名-新專案), -1 (危險區熔斷)
' ------------------------------------------------------------------------------
Public Function GetBootQuadrant(ByRef config As cls_Config, ByRef outActualName As String, ByRef outMemoryName As String) As Integer
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim currentDir As String: currentDir = ThisWorkbook.Path
    If currentDir = "" Then currentDir = Application.DefaultFilePath
    
    ' [防禦前哨] 危險區探測
    If Mod_Rules.IsDangerousPath(currentDir) Then
        GetBootQuadrant = -1
        Set fso = Nothing: Exit Function
    End If
    
    ' [讀取真理]
    Dim wsMeta As Worksheet
    Set wsMeta = Mod_Utils.GetOrCreateSheet(ThisWorkbook, config.SheetSysMeta)
    outMemoryName = Trim(CStr(wsMeta.Range("A1").Value))
    outActualName = fso.GetFileName(currentDir)
    
    Dim isHealthy As Boolean: isHealthy = IsEnvironmentHealthy(config)
    
    ' [四象限判定]
    If isHealthy And (outMemoryName = outActualName Or outMemoryName = "") Then
        GetBootQuadrant = 1 ' 象限 1：正常啟動
    ElseIf isHealthy And outMemoryName <> outActualName Then
        GetBootQuadrant = 2 ' 象限 2：外部改名
    ElseIf Not isHealthy And outMemoryName = outActualName Then
        GetBootQuadrant = 3 ' 象限 3：環境誤刪 (需觸發自癒)
    Else
        GetBootQuadrant = 4 ' 象限 4：新專案初始化
    End If
    
    Set fso = Nothing: Set wsMeta = Nothing
End Function

' ==============================================================================
' 區塊 3：Excel 畫布與工作表維護操作 (Canvas & Workspace Operations)
' ==============================================================================

Public Sub UpdateMemoryPoint(ByVal wb As Workbook, ByVal newName As String, ByVal config As cls_Config)
    Dim wsMeta As Worksheet, wasProtected As Boolean
    Set wsMeta = Mod_Utils.GetOrCreateSheet(wb, config.SheetSysMeta)
    wasProtected = wb.ProtectStructure
    If wasProtected Then wb.Unprotect Password:=config.SecurityPassword
    
    Call Mod_Utils.SafeUnprotect(wsMeta, config.SecurityPassword)
    wsMeta.Range("A1").Value = newName
    wsMeta.Visible = xlSheetVeryHidden
    Call Mod_Utils.SafeProtect(wsMeta, config.SecurityPassword)
    
    If wasProtected Then wb.Protect Password:=config.SecurityPassword, Structure:=True, Windows:=False
    Set wsMeta = Nothing
End Sub

Public Sub ExecuteStandardBoot(ByVal config As cls_Config, ByVal actualName As String)
    Dim wsUI As Worksheet: Set wsUI = Mod_Utils.GetOrCreateSheet(ThisWorkbook, config.SheetNameUI)
    Call Mod_Utils.SafeUnprotect(wsUI, config.SecurityPassword)
    Application.EnableEvents = False
    wsUI.Range(config.CellProjName).Value = actualName
    Application.EnableEvents = True
    Call Mod_Utils.SafeProtect(wsUI, config.SecurityPassword)

    wsUI.Visible = xlSheetVisible
    wsUI.Activate
    Call Mod_UIReset.重建操作介面(True, False)
    wsUI.Range("A1").Select
End Sub

Public Sub CleanUpOldSheets(ByVal config As cls_Config, ByVal isCatalogAppend As Boolean)
     Dim wsCheck As Worksheet
     Dim iLoop As Long, oldCatCount As Long, wasProtected As Boolean
     If config Is Nothing Then Exit Sub
     
     Application.DisplayAlerts = False
     wasProtected = config.TargetWB.ProtectStructure
     If wasProtected Then config.TargetWB.Unprotect Password:=config.SecurityPassword
     
    If isCatalogAppend = False Then
         Set wsCheck = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameCatalog)
         If Not wsCheck Is Nothing Then On Error Resume Next: wsCheck.Delete: On Error GoTo 0
         Set wsCheck = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameAttachment)
         If Not wsCheck Is Nothing Then On Error Resume Next: wsCheck.Delete: On Error GoTo 0
     End If
    
    oldCatCount = 0
    For iLoop = config.TargetWB.Worksheets.count To 1 Step -1
        Set wsCheck = config.TargetWB.Worksheets(iLoop)
        If Left(wsCheck.Name, Len(config.PrefixCustom)) = config.PrefixCustom Then
            ' 保留自訂白名單
        ElseIf InStr(wsCheck.Name, config.PrefixOldCatalog) = 1 Then
            oldCatCount = oldCatCount + 1
            If oldCatCount > 2 Then On Error Resume Next: wsCheck.Delete: On Error GoTo 0
        ElseIf wsCheck.Name = config.SheetNameUI Or wsCheck.Name = config.SheetNameManual Or wsCheck.Name = config.SheetNameQuarantine Or wsCheck.Name = config.SheetNameRescue Or wsCheck.Name = config.SheetNameCatalog Or wsCheck.Name = config.SheetNameAttachment Or wsCheck.Name = config.SheetSysMeta Or wsCheck.Name = config.SheetSysLog Or wsCheck.Name = config.SheetSysTempSettings Then
            ' 保留系統底板表
        Else
            On Error Resume Next: wsCheck.Delete: On Error GoTo 0
        End If
    Next iLoop
    
    If wasProtected Then config.TargetWB.Protect Password:=config.SecurityPassword, Structure:=True, Windows:=False
    Application.DisplayAlerts = True
    Set wsCheck = Nothing
End Sub

Public Sub ResetWorkspaceToZero(ByVal wb As Workbook, ByVal config As cls_Config)
    Dim ws As Worksheet, shName As Variant, sheetsToDelete As Object, safeWs As Worksheet
    Application.DisplayAlerts = False: Application.EnableEvents = False: Application.ScreenUpdating = False
    Call Mod_Utils.GlobalUnprotectAll(wb, config.SecurityPassword)
    
    On Error Resume Next
    Set safeWs = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
    safeWs.Name = config.SheetNameTempSafeRoom
    safeWs.Visible = xlSheetVisible
    safeWs.Activate
    On Error GoTo 0
    
    Set sheetsToDelete = CreateObject("Scripting.Dictionary")
    For Each ws In wb.Worksheets
        If ws.Name <> config.SheetNameManual And ws.Name <> config.SheetNameUI And ws.Name <> config.SheetSysTempSettings And ws.Name <> config.SheetSysMeta And ws.Name <> config.SheetSysLog And ws.Name <> config.SheetNameTempSafeRoom And Left(ws.Name, Len(config.PrefixCustom)) <> config.PrefixCustom Then
            sheetsToDelete.Add ws.Name, 1
        End If
    Next ws
    
    For Each shName In sheetsToDelete.Keys
        Set ws = wb.Worksheets(shName)
        On Error Resume Next: ws.Visible = xlSheetVisible: ws.Delete: On Error GoTo 0
    Next shName
    
    On Error Resume Next
    wb.Worksheets(config.SheetNameUI).Visible = xlSheetVisible
    wb.Worksheets(config.SheetNameUI).Activate
    safeWs.Delete
    On Error GoTo 0
    
    Application.DisplayAlerts = True: Application.EnableEvents = True
    Set sheetsToDelete = Nothing: Set safeWs = Nothing
End Sub

' ==============================================================================
' 區塊 4：日誌、隔離與異常復原 (Logs & Emergency Actions)
' ==============================================================================

Public Function CountTotalMissing(ByVal config As cls_Config) As Long
    Dim catWs As Worksheet, lastRow As Long, r As Long, fso As Object, count As Long, curPath As String
    CountTotalMissing = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set catWs = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameCatalog)
    If catWs Is Nothing Then Exit Function
    lastRow = catWs.Cells(catWs.Rows.count, "A").End(xlUp).row
    If lastRow < 5 Then Exit Function
    
    count = 0
    For r = 5 To lastRow
        curPath = Trim(CStr(catWs.Cells(r, 3).Value))
        If curPath <> "" Then If Not fso.FileExists(curPath) Then count = count + 1
    Next r
    CountTotalMissing = count
    Set fso = Nothing: Set catWs = Nothing
End Function

Public Sub WriteQuarantineLog(ByVal config As cls_Config, ByVal origName As String, ByVal newName As String, ByVal reason As String, ByVal actualSavedPath As String)
    Dim wsQ As Worksheet, lastRow As Long, sheetName As String: sheetName = config.SheetNameQuarantine
    Dim wasWbProtected As Boolean
    Set wsQ = Mod_Utils.GetSheetSafe(config.TargetWB, sheetName)
    
    If wsQ Is Nothing Then
        wasWbProtected = config.TargetWB.ProtectStructure
        If wasWbProtected Then On Error Resume Next: config.TargetWB.Unprotect Password:=config.SecurityPassword: On Error GoTo 0
        Set wsQ = config.TargetWB.Worksheets.Add(After:=config.TargetWB.Sheets(config.TargetWB.Sheets.count))
        wsQ.Name = sheetName
        wsQ.Range("A1:D1").Merge: wsQ.Range("A1").Value = "ROUTING & QUARANTINE LOG | 智慧分流與隔離清單"
        wsQ.Range("A1").Font.Name = config.FontMain: wsQ.Range("A1").Font.Size = 14: wsQ.Range("A1").Font.Bold = True: wsQ.Range("A1").Font.Color = RGB(255, 255, 255)
        wsQ.Range("A1:D1").Interior.Color = RGB(70, 110, 120): wsQ.Range("A1:D1").RowHeight = 32: wsQ.Range("A1").VerticalAlignment = xlVAlignCenter
        wsQ.Range("A2:D2").Value = Array("處理時間", "原始檔案名稱", "分流後狀態 (動態偵測)", "分流原因與類別")
        With wsQ.Range("A2:D2")
            .Font.Name = config.FontMain: .Font.Size = 10: .Font.Bold = True: .Interior.Color = RGB(120, 140, 150): .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlHAlignCenter: .Borders.LineStyle = xlContinuous
        End With
        wsQ.Rows("2:2").RowHeight = 22: wsQ.Cells(2, 1).Formula = "=HYPERLINK(""#'操作介面'!A1"", "" ▋返回控制台 "")"
        wsQ.Cells(2, 1).Font.Color = config.ColorLinkText
        If wasWbProtected Then On Error Resume Next: config.TargetWB.Protect Password:=config.SecurityPassword, Structure:=True, Windows:=False: On Error GoTo 0
    End If
    
    Call Mod_Utils.SafeUnprotect(wsQ, config.SecurityPassword)
    lastRow = wsQ.Cells(wsQ.Rows.count, 2).End(xlUp).row + 1
    If lastRow < 3 Then lastRow = 3
    
    wsQ.Cells(lastRow, 1).Value = Format(Now, "yyyy/mm/dd hh:mm:ss")
    wsQ.Cells(lastRow, 2).Value = origName
    wsQ.Cells(lastRow, 3).Formula = "=IF(IsFileAlive(""" & actualSavedPath & """), HYPERLINK(""" & actualSavedPath & """, ""[開啟] " & newName & """), """ & newName & " (實體已刪除)"")"
    wsQ.Cells(lastRow, 4).Value = reason
    With wsQ.Range(wsQ.Cells(lastRow, 1), wsQ.Cells(lastRow, 4))
        .Font.Name = config.FontMain: .Font.Size = 9: .Font.Color = RGB(50, 50, 50): .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(220, 220, 220): .EntireRow.RowHeight = 20: .VerticalAlignment = xlVAlignCenter
    End With
    wsQ.Columns("A:D").AutoFit
    Call Mod_Utils.SafeProtect(wsQ, config.SecurityPassword)
    Set wsQ = Nothing
End Sub

Public Sub ArchiveRestoredFile(ByVal fso As Object, ByVal filePath As String, ByVal config As cls_Config, ByVal LogStore As Object)
    Dim restoredDir As String, fileName As String, finalName As String, destPath As String
    
    On Error GoTo ArchiveErrorHandler
    
    restoredDir = config.IORescueDowncastFolder
    Call Mod_Utils.EnsureFolderExists(fso, restoredDir)
    fileName = fso.GetFileName(filePath)
    finalName = Mod_Utils.GenerateUniqueFileName(fso, restoredDir, fileName)
    destPath = Mod_Utils.BuildSafeWindowsPath(fso, restoredDir, finalName)
    fso.MoveFile filePath, destPath
    Exit Sub

ArchiveErrorHandler:
    If Not LogStore Is Nothing Then
        LogStore.Record "TX_ERROR", "ARCHIVE", "歸檔失敗，檔案未被搬移，原地保留：" & filePath & " | 原因：" & Err.Description
    End If
End Sub

' ------------------------------------------------------------------------------
' [環境毀損復原]：ExecuteProactiveRescueAndSnapshot
' 目的：當環境遭實體抹除時，將當下畫布所有殘影強制回存為實體，並將舊目錄封存快照。
' ------------------------------------------------------------------------------
Public Sub ExecuteProactiveRescueAndSnapshot(ByVal config As cls_Config)
    Dim fso As Object, wsLoop As Worksheet, shp As Shape
    Dim targetWs As Worksheet, sIdx As Long, waitIdx As Integer
    Dim timeStamp As String, snapshotName As String
    Dim safeFileName As String, exportFilename As String, apiResult As Boolean
    Dim origW As Double, origH As Double, nativeW As Double, nativeH As Double
    Dim targetFactor As Double, chtObj As ChartObject
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    
    Dim rescueLogStore As New cls_Log
    Call rescueLogStore.Initialize(config, config.TargetWB, False)
    rescueLogStore.Record "TX_START", "PROACTIVE_RESCUE", "偵測到環境異常，開始執行緊急搶救與快照作業"
    
    Dim targetLedgers As Variant, ledgerName As Variant
    targetLedgers = Array(config.SheetNameCatalog, config.SheetNameAttachment)
    
' 1. 宣告變數並預先記錄活頁簿是否處於保護狀態
    Dim wasWbProtected As Boolean

    wasWbProtected = config.TargetWB.ProtectStructure

    Application.DisplayAlerts = False
    
    ' 2. 全域解鎖 TargetWB 結構與所有工作表
    Call Mod_Utils.GlobalUnprotectAll(config.TargetWB, config.SecurityPassword)
    
    For Each ledgerName In targetLedgers
        Set targetWs = Mod_Utils.GetSheetSafe(config.TargetWB, CStr(ledgerName))
        
        If Not targetWs Is Nothing Then
            snapshotName = config.PrefixCustom & CStr(ledgerName) & "_快照_" & timeStamp
            If Len(snapshotName) > 31 Then snapshotName = Left(snapshotName, 31)
            
            targetWs.Name = snapshotName
            Call Mod_Utils.SafeProtect(targetWs, config.SecurityPassword)
        End If
    Next ledgerName
    
    ' 3. 若原本有鎖，處理完畢後鎖回 TargetWB 活頁簿結構
    If wasWbProtected Then
        config.TargetWB.Protect Password:=config.SecurityPassword, Structure:=True, Windows:=False
    End If
    
    Application.DisplayAlerts = True
    
    
     Dim gdipToken As LongPtr
     gdipToken = Mod_GDIPlusExport.StartGdiplusEngine()
     If gdipToken = 0 Then
         Call Mod_Actions.WriteQuarantineLog(config, "系統", "系統", "緊急搶救中止：GDI+引擎啟動失敗，作業系統圖形元件異常", "")
         rescueLogStore.Record "TX_ERROR", "PROACTIVE_RESCUE", "搶救中止：GDI+引擎啟動失敗，本次未搶救任何圖形"
         Set fso = Nothing: Set targetWs = Nothing: Set shp = Nothing
         Exit Sub
     End If

    
    Dim rescueSnapshotProgressCounter As Long: rescueSnapshotProgressCounter = 0
    Dim rescueSuccessCount As Long: rescueSuccessCount = 0
    Dim rescueQuarantineCount As Long: rescueQuarantineCount = 0
    For Each wsLoop In ThisWorkbook.Worksheets
        If Mod_Rules.IsStandardImgSheet(wsLoop.Name, config) Or wsLoop.Name = config.SheetNameRescue Then
            Call Mod_Utils.SafeUnprotect(wsLoop, config.SecurityPassword)
            For sIdx = wsLoop.Shapes.count To 1 Step -1
                Set shp = wsLoop.Shapes(sIdx)
                If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
                
                
                    Err.Clear
                    On Error Resume Next
                    origW = shp.Width: origH = shp.Height
                    shp.ScaleWidth 1, msoTrue: shp.ScaleHeight 1, msoTrue
                    nativeW = shp.Width: nativeH = shp.Height
                    If nativeW > 300 Or nativeH > 300 Then
                        If nativeW >= nativeH Then targetFactor = 300 / nativeW Else targetFactor = 300 / nativeH
                        shp.Width = nativeW * targetFactor: shp.Height = nativeH * targetFactor
                    End If

                    ' 【關鍵防呆】：只信任「CopyPicture 這一刻」的執行結果
                    Err.Clear
                    shp.CopyPicture Appearance:=2, Format:=2
                    Dim copySucceeded As Boolean: copySucceeded = (Err.Number = 0)

                    For waitIdx = 1 To 10: DoEvents: Next waitIdx
                    Dim extractedName As String: extractedName = shp.Name
                   ' 【V2-083修正，比照V2-057】不再無條件先刪除Shape——先嘗試存檔，
                   ' 確認硬碟上真的有檔案落地，才刪除原Shape。複製優先於刪除，P0第2條。
                   ' 此前不管CopyPicture有沒有成功，Delete都會無條件先執行，若後續存檔
                   ' （含Chart備援）也失敗，圖形就徹底消失、且完全沒有任何記錄可查
                   Dim rescueFileSaved As Boolean: rescueFileSaved = False

                   If copySucceeded Then
                        safeFileName = Mod_Utils.GenerateUniqueFileName(fso, config.IORescueFolder, "AutoRescue_" & extractedName & ".png")
                        exportFilename = Mod_Utils.BuildSafeWindowsPath(fso, config.IORescueFolder, safeFileName)
                        apiResult = Mod_GDIPlusExport.SaveClipboardToPNG(exportFilename, gdipToken)
                        If Not apiResult Then
                            Err.Clear
                            Set chtObj = wsLoop.ChartObjects.Add(Left:=0, Top:=0, Width:=nativeW, Height:=nativeH)
                            If Not chtObj Is Nothing Then
                                chtObj.Activate: chtObj.Chart.Paste: chtObj.Chart.ChartArea.Format.Line.Visible = msoFalse
                                chtObj.Chart.Export fileName:=exportFilename, FilterName:="PNG"
                                chtObj.Delete: Set chtObj = Nothing
                            End If
                        End If
                       ' 不管走GDI+還是Chart備援，都要實際確認檔案真的落地，不能只信任函式回傳值
                       rescueFileSaved = fso.FileExists(exportFilename)
                   End If

                   If rescueFileSaved Then
                       shp.Delete
                       rescueSuccessCount = rescueSuccessCount + 1
                   ElseIf copySucceeded Then
                       ' 複製到剪貼簿成功，但存檔（含Chart備援）全部失敗——原圖形保留不刪除
                       Call Mod_Actions.WriteQuarantineLog(config, extractedName, extractedName, "緊急搶救時PNG與備援存檔皆失敗，原圖形已保留於畫布，未刪除", "")
                       rescueQuarantineCount = rescueQuarantineCount + 1
                   Else
                       ' 剪貼簿沒拿到有效圖像，原圖形保留不刪除
                       Call Mod_Actions.WriteQuarantineLog(config, extractedName, extractedName, "緊急搶救時剪貼簿擷取失敗，原圖形已保留於畫布，未刪除", "")
                       rescueQuarantineCount = rescueQuarantineCount + 1
                   End If
                    Err.Clear: On Error GoTo 0
                    

                    ' 【補上】此函式原本完全沒有任何存檔機制，跟cmd_Rescue過去閃退的根因是同一類風險：
                    ' 若中途應用程式層級崩潰，前面已經處理過的搶救全部作廢。比照cmd_Rescue既有的
                    ' 60張一次頻率，給Excel定期一次強制重置內部狀態的機會
                    rescueSnapshotProgressCounter = rescueSnapshotProgressCounter + 1
                    If rescueSnapshotProgressCounter Mod 60 = 0 Then
                        config.TargetWB.Save
                    End If

                    
                End If
            Next sIdx
            Call Mod_Utils.SafeProtect(wsLoop, config.SecurityPassword)
        End If
    Next wsLoop
     ' 【補上】迴圈全部跑完後，若有處理過任何圖形，最後再存一次檔，
    ' 確保最後一批不滿60張的搶救結果也確實落地，不留在只存在記憶體裡的狀態
    If rescueSnapshotProgressCounter > 0 Then config.TargetWB.Save
    Call Mod_GDIPlusExport.StopGdiplusEngine(gdipToken)
    rescueLogStore.Record "TX_COMMIT_FS", "PROACTIVE_RESCUE", "搶救作業完成 | 共處理 " & rescueSnapshotProgressCounter & " 張 | 成功搶救 " & rescueSuccessCount & " 張 | 隔離保留 " & rescueQuarantineCount & " 張（詳見隔離清單）"
    Set fso = Nothing: Set targetWs = Nothing: Set shp = Nothing
End Sub

' ------------------------------------------------------------------------------
' [P0安全檢查]：IsDiskSpaceSafe
' 目的：批次作業開始前與過程中皆可呼叫。同時檢查專案所在磁碟區與C槽——
'      C槽即使不是專案所在位置，仍因Windows分頁檔預設放在此處，
'      空間見底一樣會拖累整台電腦穩定性，不能只查專案磁碟區。
' ------------------------------------------------------------------------------
Public Function IsDiskSpaceSafe(ByVal config As cls_Config, ByVal LogStore As Object) As Boolean
    Dim projectFreeGB As Double: projectFreeGB = Mod_Utils.GetFreeDiskSpaceGB(config.BaseFolder)
    Dim systemFreeGB As Double: systemFreeGB = Mod_Utils.GetFreeDiskSpaceGB("C:\")

    IsDiskSpaceSafe = True

    If projectFreeGB >= 0 And projectFreeGB < config.MinFreeDiskSpaceGB Then
        IsDiskSpaceSafe = False
        LogStore.Record "TX_ERROR", "DISKSPACE", "專案磁碟區剩餘空間過低（" & Format(projectFreeGB, "0.0") & "GB），低於安全下限" & config.MinFreeDiskSpaceGB & "GB"
    End If
    
     Dim availMemMB As Double: availMemMB = Mod_Utils.GetAvailableMemoryMB()
     If availMemMB >= 0 And availMemMB < config.MinAvailableMemoryMB Then
         IsDiskSpaceSafe = False
         LogStore.Record "TX_ERROR", "DISKSPACE", "可用記憶體過低（" & Format(availMemMB, "0") & "MB），低於安全下限" & config.MinAvailableMemoryMB & "MB"
     End If

    If systemFreeGB >= 0 And systemFreeGB < config.MinFreeDiskSpaceGB Then
        IsDiskSpaceSafe = False
        LogStore.Record "TX_ERROR", "DISKSPACE", "C槽（系統分頁檔所在磁碟區）剩餘空間過低（" & Format(systemFreeGB, "0.0") & "GB），低於安全下限" & config.MinFreeDiskSpaceGB & "GB"
    End If
End Function
