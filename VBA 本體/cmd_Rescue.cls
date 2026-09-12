Option Explicit
Implements ICommand

' ==========================================================
' MODULE: cmd_Rescue (類別模組)
' PURPOSE: 遺失圖片救援命令。負責將畫布上找不到實體檔案的圖形匯出至救援夾，
'          統一加上救援標記 (MANUAL_RESCUE)，並保留《圖片目錄》中的原始 HASH 記錄。
' EXPORTS: ICommand_Execute 等介面合約
' IMPORTS: cls_ExecutionContext, Mod_Utils, Mod_GDIPlusExport, Mod_Actions
' FORBIDDEN: 1. 嚴禁在此撰寫任何 UI 彈窗 (已交由 Pipeline 提前決策)。
'            2. 嚴禁在此掃描硬碟與比對畫布 (已交由 Scanner 完成)。
'            3. 嚴禁操作 Application.ScreenUpdating (防止畫面閃爍異常)。
'            [例外授權]: 允許操作《圖片目錄》儲存格，以寫入原始 HASH 與救援狀態資訊。
' DEPENDENCIES: Microsoft Scripting Runtime (FileSystemObject)
' VERSION: 1.0.0 [Stability: Stable]
' ADR: 詳見 ACDS_ADR/ADR-（待補編號）.md — SRP 拆分達成有限度：迴圈本體因時序敏感
'      （DoEvents 排錯序列）與 P0 事故修復邏輯（V2-057/V2-084 先確認新圖嵌入成功
'      才刪除原圖形）高度耦合，刻意不拆，僅抽出迴圈外三塊無副作用的準備工作
' ==========================================================

Private Const MODULE_NAME As String = "cmd_Rescue"
Private Const MODULE_VERSION As String = "1.0.0"


' ==============================================================================
' ICommand 介面實作區
' ==============================================================================
Private Property Get ICommand_IsCritical() As Boolean
    ICommand_IsCritical = True
End Property

Private Property Get ICommand_CommandName() As String
    ICommand_CommandName = "RESCUE"
End Property

Private Property Get ICommand_Dependencies() As Variant
    ICommand_Dependencies = Array("IMPORT")
End Property

Private Sub ICommand_Execute(ByVal Context As cls_ExecutionContext)
    Dim fso As Object, config As cls_Config, LogStore As Object
    Dim uniqueKey As Variant, parts() As String
    Dim wsTarget As Worksheet, sh As Shape
    Dim rescueDir As String, finalRescueName As String, exportFilename As String, oldName As String
    Dim apiResult As Boolean, newHash As String
    Dim origW As Double, origH As Double, nativeW As Double, nativeH As Double, targetFactor As Double
    Dim rescueCounter As Long, timeStamp As String
    Dim rescueProgressCounter As Long
    Dim errNum As Long, errDesc As String

    On Error GoTo RescueErrorHandler

    Set config = Context.config
    Set LogStore = Context.LogStore

    If config Is Nothing Or LogStore Is Nothing Then
        Err.Raise 9991, MODULE_NAME, "Context 遺失"
    End If

    ' 【依使用者決策執行】：若使用者選擇「否」，或無遺失圖片，則不執行任何動作
    If Not Context.UserWantsRescue Or Context.MissingShapes.count = 0 Then
        LogStore.Record "INFO", "RESCUE", "已確認無需執行救援，結束。"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    rescueDir = config.IORescueFolder
    Call Mod_Utils.EnsureFolderExists(fso, rescueDir)

    LogStore.Record "TX_START", "RESCUE", "開始執行遺失圖片匯出與 HASH 記錄..."

    ' --------------------------------------------------------------------------
    ' 準備目錄快取 (定位待處理項目與原始 HASH)
    ' --------------------------------------------------------------------------
    Dim wsCat As Worksheet
    Dim dictCatRow As Object
    Set wsCat = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameCatalog)
    Set dictCatRow = BuildCatalogRowIndex(wsCat)

    rescueCounter = 0
    rescueProgressCounter = 0
    
  
    ' 【效能優化】：目標工作表全程只有一個，迴圈外查詢一次即可，不需要每張圖片都重新查詢
    Dim wsRescueTarget2 As Worksheet
    Set wsRescueTarget2 = Mod_Utils.GetOrCreateSheet(config.TargetWB, config.SheetNameRescue)

    Dim gdipToken As LongPtr
    If Not InitializeRescueEngine(config, LogStore, gdipToken) Then GoTo RescueCleanup
     
    ' 【V2-060】固定嵌入尺寸與垂直堆疊位置，不再全部疊在同一個(0,0)看不見彼此。
    ' 起始位置從救援夾既有Shape最底部算起，避免跟上一輪還沒被REBUILD收編的殘留圖形重疊
    Dim rescueStackTop As Double
    rescueStackTop = CalculateInitialStackTop(config, wsRescueTarget2)


    ' ==========================================================================
    ' 核心流程：截圖匯出、改名、寫入原始 HASH 記錄
    ' 【單張錯誤隔離】：任何一張圖片處理失敗，只記錄跳過，不影響其餘圖片繼續救援
    ' ==========================================================================
    For Each uniqueKey In Context.MissingShapes.Keys

        On Error GoTo RescueItemErrorHandler

        parts = Split(CStr(uniqueKey), "||")
        Set wsTarget = Mod_Utils.GetSheetSafe(config.TargetWB, parts(0))

        If Not wsTarget Is Nothing Then
            On Error Resume Next
            Set sh = wsTarget.Shapes(parts(1))
            If Err.Number = 0 And Not sh Is Nothing Then
                On Error GoTo RescueItemErrorHandler

                ' 【批次防線】：每 20 張回報進度，每 60 張中途存檔，避免大量救援時記憶體堆積
                rescueProgressCounter = rescueProgressCounter + 1
                If rescueProgressCounter Mod 20 = 0 Then
                    Application.StatusBar = "正在救援遺失圖片 (" & rescueProgressCounter & " / " & Context.MissingShapes.count & ")"
                    LogStore.Record "TX_PROGRESS", "RESCUE", "進度：" & rescueProgressCounter & " / " & Context.MissingShapes.count
                    DoEvents
                End If
                If rescueProgressCounter Mod 300 = 0 Then
                    config.TargetWB.Save
                    
                   ' 【安全網】跑到一半空間可能被吃光，每次存檔順便複查，
                    ' 發現危險立刻中止，不要死撐到系統自己崩潰
                    If Not Mod_Actions.IsDiskSpaceSafe(config, LogStore) Then
                        LogStore.Record "TX_ERROR", "RESCUE", "磁碟空間於執行過程中降至安全門檻以下，救援作業提前中止，已處理項目的存檔進度已保留"
                        GoTo RescueCleanup
                    End If
                    
                     ' 【存檔後緩衝】：存檔是I/O吃重操作，Excel可能需要一點時間才能真正恢復穩定，
                     ' 緊接著做剪貼簿操作容易撞上這段還沒恢復的空窗期，造成連續失敗
                     Dim saveSettleIdx As Integer
                     For saveSettleIdx = 1 To 15
                         DoEvents
                     Next saveSettleIdx
                End If

                oldName = sh.Name

                ' 【規則一：救援標記】：區分第一次遺失與重複遺失
                If oldName Like "MANUAL_RESCUE_*" Then
                    finalRescueName = oldName
                Else
                    finalRescueName = "MANUAL_RESCUE_" & oldName
                End If

                ' 尺寸防禦 (防範超大圖撐爆 GDI 記憶體)
                ' 【V2-060修正】ScaleWidth(1,msoTrue)取得的「原始尺寸」實際上是Shape首次
                ' 嵌入時的排版設定值，不是圖片真正原生解析度，導致擷取尺寸曾經跟排版PT數字
                ' 成正比、沒有下限保護。比例仍可信，只有絕對值不可信，改為只取比例、
                ' 一律強制設為固定寬度，不再判斷是否超過上限
                origW = sh.Width
                origH = sh.Height
                sh.ScaleWidth 1, msoTrue
                sh.ScaleHeight 1, msoTrue
                nativeW = sh.Width
                nativeH = sh.Height

                Dim captureAspectRatio As Double
                If nativeW > 0 Then
                    captureAspectRatio = nativeH / nativeW
                Else
                    captureAspectRatio = 1
                End If
                sh.Width = config.MaxSafetyPoint
                sh.Height = config.MaxSafetyPoint * captureAspectRatio

                ' 【V2-058】改尺寸後、擷取前也要給Excel重繪的時間，否則CopyPicture
                ' 可能擷取到還沒畫完的殘影，導致存檔異常小（症狀：2MB原圖救援後只剩10K）
                Dim resizeSettleIdx As Integer
                For resizeSettleIdx = 1 To 10
                    DoEvents
                Next resizeSettleIdx



                ' 【截圖匯出】
                ' 【V2-056】CopyPicture比照既有重試模式，COM偶發不穩定不代表真的失敗
                Dim copyPicAttempt As Integer
                Dim copyPicOk As Boolean: copyPicOk = False
                For copyPicAttempt = 1 To 3
                    Err.Clear
                    On Error Resume Next
                    sh.CopyPicture Appearance:=1, Format:=2
                    If Err.Number = 0 Then
                        copyPicOk = True
                        Err.Clear
                        On Error GoTo RescueItemErrorHandler
                        Exit For
                    End If
                    Err.Clear
                    On Error GoTo RescueItemErrorHandler
                    If copyPicAttempt < 3 Then
                        Dim copyPicSettleIdx As Integer
                        For copyPicSettleIdx = 1 To 15
                            DoEvents
                        Next copyPicSettleIdx
                    End If
                Next copyPicAttempt

                If Not copyPicOk Then
                    ' 【V2-085補漏】直接跳出前，尺寸還原這兩行完全沒機會執行，
                    ' Shape會永久卡在為了擷取而縮小過的尺寸，不算真正的「原封不動留在原地」
                    sh.Width = origW
                    sh.Height = origH
                    LogStore.Record "TX_ERROR", "RESCUE", "截圖擷取失敗（已重試3次），已跳過：" & oldName
                    GoTo RescueItemNextKey
                End If

                Dim waitIdx As Integer
                For waitIdx = 1 To 10
                    DoEvents
                Next waitIdx

                exportFilename = Mod_Utils.BuildSafeWindowsPath(fso, rescueDir, finalRescueName & ".png")
                apiResult = Mod_GDIPlusExport.SaveClipboardToPNG(exportFilename, gdipToken)

                ' 【資源釋放】：每次剪貼簿操作完立刻清空，不等到60張才清，
                ' 這是唯一逐張都會佔用系統剪貼簿的操作，累積下去會讓Excel本身不穩定
                Application.CutCopyMode = False

                ' 恢復原始尺寸
                sh.Width = origW
                sh.Height = origH

                ' ==============================================================
                ' 匯出成功後：執行改名與原始 HASH 記錄寫入
                ' ==============================================================
                If apiResult And fso.FileExists(exportFilename) Then
                    newHash = Mod_Utils.GetMD5(fso, exportFilename)
                    Context.DictRescueHash(newHash) = finalRescueName


                     ' 【同步修正】：Shape 實際物理位置要跟著檔案搬到救援夾，畫布跟硬碟才能對得起來
                     ' 【複製優先於刪除】：先複製、確認貼上成功，才刪除原圖形；
                     ' 貼上失敗則保留原圖形不動，目錄狀態不推進，交由下一輪全面排版重新判定為遺失、重新嘗試
                     
                     Dim shOriginal As Shape
                     Dim relocateSucceeded As Boolean
                     relocateSucceeded = True  ' 同工作表情境不需搬移，視為成功

                     If Not wsTarget Is wsRescueTarget2 Then
                        ' 【V2-084修正】不再於此提前改sh.Name——此前不管搬移最後成不成功，
                        ' 名字都會先被改成救援完成格式，導致搬移失敗時，原地保留的sh頂著
                        ' 「看起來已救援完成」的名字，跟它「其實還沒救援成功」的真實狀態不一致。
                        ' sh的命運只有兩種：搬移成功→被Delete(這裡改的名字用不到)；
                        ' 搬移失敗→原地保留(不該頂著錯誤的名字)，兩種情況都不該提前改名
                     
                     
                         ' 【V2-057修正】原V2-054先Delete原圖形、才嘗試AddPicture，
                         ' 失敗時原圖形已經沒了。改回：先確認新圖嵌入成功，才刪除原圖形
                         Dim newPic As Shape
                         Set newPic = Nothing
                         On Error Resume Next

                         Set newPic = wsRescueTarget2.Shapes.AddPicture( _
                             fileName:=exportFilename, LinkToFile:=False, SaveWithDocument:=True, _
                             Left:=0, Top:=0, Width:=-1, Height:=-1)
                         On Error GoTo RescueItemErrorHandler

                         If Not newPic Is Nothing Then
                             newPic.Name = finalRescueName
                             Dim embedRatio As Double
                          If newPic.Width > 0 Then
                              embedRatio = newPic.Height / newPic.Width
                          Else
                              embedRatio = 1
                          End If
                          newPic.Width = config.RescueEmbedWidth
                          newPic.Height = config.RescueEmbedWidth * embedRatio
                          newPic.Left = 0
                          newPic.Top = rescueStackTop
                          rescueStackTop = rescueStackTop + newPic.Height + config.RescueEmbedGap
                             
                            ' 【確認新圖嵌入成功後，才刪除原工作表上的圖形】
                            sh.Delete
                         Else
                             LogStore.Record "TX_ERROR", "RESCUE", "PNG重新嵌入救援夾失敗，原圖形已保留於原工作表，可重新觸發救援：" & oldName
                             relocateSucceeded = False
                             
                            ' 【V2-072補漏】V2-054重新設計搬移機制時遺漏了這段——舊版Cut/Paste
                            ' 機制原本就有孤兒PNG清理，改用AddPicture後沒有搬過來。PNG已存在
                            ' 硬碟但沒有被任何Shape引用，若不清掉，下次全面排版scan救援夾資料夾
                            ' 時可能被cmd_Rebuild.ProcessSingleFile誤認為獨立的新救援檔案，
                            ' 建出跟原本sh重複的另一個Shape，造成同一張圖出現兩份的狀態不一致
                            Call Mod_Actions.ArchiveRestoredFile(fso, exportFilename, config, LogStore)
 
                         End If
                    Else
                        ' 同工作表情境：不需要搬移，這裡才是真正該改名的地方，不會有失敗風險
                        sh.Name = finalRescueName
                     End If

                     ' 【計數時機修正】：搬移確認成功後才計入，避免「PNG匯出成功但搬移失敗」被誤計為救援成功
                     If relocateSucceeded Then
                         Context.SessionRescued = Context.SessionRescued + 1
                     End If
 
                     ' 2. 更新目錄，保留原始 HASH
                     ' 【原子性保護】：搬移失敗時不寫入任何「已救援」狀態，維持原紀錄不變
                     If relocateSucceeded Then
                     If Not wsCat Is Nothing Then
                         If dictCatRow.Exists(oldName) Then
                       
                        
                        
                            Dim catR As Long
                            catR = dictCatRow(oldName)
                            Dim currHash As String
                            currHash = Trim(CStr(wsCat.Cells(catR, 2).Value))
                            Dim origHash As String
                            origHash = Trim(CStr(wsCat.Cells(catR, 9).Value))

                            ' 【規則二：保留原始 HASH】若為第一次遺失，OrigHash 欄位為空，將目前 HASH 寫入第 9 欄保存
                            If origHash = "" Then
                                origHash = currHash
                            End If

                            ' 原子更新目錄
                            wsCat.Cells(catR, 1).Value = finalRescueName
                            wsCat.Cells(catR, 2).Value = newHash
                            wsCat.Cells(catR, 3).Value = exportFilename
                            wsCat.Cells(catR, 7).Value = "系統救援"
                            wsCat.Cells(catR, 8).Value = "是"
                            wsCat.Cells(catR, 9).Value = origHash

                            ' 更新快取，避免同名衝突
                            dictCatRow.Remove oldName
                            dictCatRow(finalRescueName) = catR
                        Else
                            ' 【新面孔】：目錄表從沒登記過這個名字（例如手動貼上工作表的圖片），開一筆全新紀錄
                            Dim newCatRow As Long
                            newCatRow = wsCat.Cells(wsCat.Rows.count, "A").End(xlUp).row + 1
                            wsCat.Cells(newCatRow, 1).Value = finalRescueName
                            wsCat.Cells(newCatRow, 2).Value = newHash
                            wsCat.Cells(newCatRow, 3).Value = exportFilename
                            wsCat.Cells(newCatRow, 7).Value = "系統救援"
                            wsCat.Cells(newCatRow, 8).Value = "是"
                            wsCat.Cells(newCatRow, 9).Value = newHash
                            dictCatRow(finalRescueName) = newCatRow
                        End If
                    End If
                End If
            End If
            Err.Clear
            On Error GoTo RescueItemErrorHandler
        End If
    End If

RescueItemNextKey:
    Next uniqueKey

    On Error GoTo 0


    Call Mod_GDIPlusExport.StopGdiplusEngine(gdipToken)
    LogStore.Record "TX_COMMIT_FS", "RESCUE", "救援作業完成，共處理 " & Context.SessionRescued & " 張圖片，並記錄原始 HASH。"

    Exit Sub

RescueItemErrorHandler:
    Context.SessionFailed = Context.SessionFailed + 1
    LogStore.Record "TX_ERROR", "RESCUE", "單張圖片救援失敗，已跳過繼續處理下一張：" & CStr(uniqueKey) & " | 原因：" & Err.Description
    Err.Clear
    Resume RescueItemNextKey

RescueCleanup:
    Call Mod_GDIPlusExport.StopGdiplusEngine(gdipToken)
    Set fso = Nothing
    Set config = Nothing
    Set LogStore = Nothing
    Set wsTarget = Nothing
    Set sh = Nothing
    Set dictCatRow = Nothing
    Exit Sub

RescueErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    If Not LogStore Is Nothing Then
        LogStore.Record "TX_ERROR", "RESCUE", "執行錯誤: " & errNum & " | 描述: " & errDesc
    End If
    Resume RescueCleanup
End Sub

' --------------------------------------------------------------------------
' 準備目錄快取：讀取《圖片目錄》，建立「顯示名稱 → 列號」的索引，供迴圈快速定位
' --------------------------------------------------------------------------
Private Function BuildCatalogRowIndex(ByVal wsCat As Worksheet) As Object
    Dim dictCatRow As Object
    Set dictCatRow = CreateObject("Scripting.Dictionary")
    dictCatRow.CompareMode = 1

    If Not wsCat Is Nothing Then
        Dim lr As Long, i As Long
        lr = wsCat.Cells(wsCat.Rows.count, "A").End(xlUp).row
        For i = 5 To lr
            dictCatRow(Trim(CStr(wsCat.Cells(i, 1).Value))) = i
        Next i
    End If

    Set BuildCatalogRowIndex = dictCatRow
End Function

' --------------------------------------------------------------------------
' 啟動 GDI+ 引擎並確認開始前磁碟空間充足，任一項失敗即回傳 False（呼叫端應直接
' GoTo RescueCleanup）。gdipToken 以 ByRef 傳回，即使失敗也可能已取得有效 token
' （例如引擎啟動成功但磁碟空間不足），確保 Cleanup 階段能正確停止引擎
' --------------------------------------------------------------------------
Private Function InitializeRescueEngine(ByVal config As cls_Config, ByVal LogStore As Object, ByRef gdipToken As LongPtr) As Boolean
    InitializeRescueEngine = True

    ' 【V2-052】GDI+引擎迴圈外啟動一次，這是220張左右觸發應用程式層級閃退的根因
    gdipToken = Mod_GDIPlusExport.StartGdiplusEngine()
    If gdipToken = 0 Then
        LogStore.Record "TX_ERROR", "RESCUE", "GDI+引擎啟動失敗，救援作業無法執行任何截圖匯出"
        InitializeRescueEngine = False
        Exit Function
    End If

    ' 【安全網】開始前先確認磁碟空間充足，避免明知會失敗還硬跑，跑到一半才崩潰
    If Not Mod_Actions.IsDiskSpaceSafe(config, LogStore) Then
        LogStore.Record "TX_ERROR", "RESCUE", "磁碟空間不足，救援作業中止，請先騰出空間再重新執行"
        InitializeRescueEngine = False
    End If
End Function

' --------------------------------------------------------------------------
' 【V2-060】計算救援夾垂直堆疊的起始位置：從既有 Shape 最底部算起，
' 避免跟上一輪還沒被 REBUILD 收編的殘留圖形重疊
' --------------------------------------------------------------------------
Private Function CalculateInitialStackTop(ByVal config As cls_Config, ByVal wsRescueTarget2 As Worksheet) As Double
    Dim rescueStackTop As Double: rescueStackTop = 0
    Dim existingShp As Shape
    For Each existingShp In wsRescueTarget2.Shapes
        If existingShp.Top + existingShp.Height > rescueStackTop Then
            rescueStackTop = existingShp.Top + existingShp.Height
        End If
    Next existingShp
    If rescueStackTop > 0 Then rescueStackTop = rescueStackTop + config.RescueEmbedGap
    CalculateInitialStackTop = rescueStackTop
End Function

