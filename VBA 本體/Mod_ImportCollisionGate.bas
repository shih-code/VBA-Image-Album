Option Explicit

' ==========================================================
' MODULE: Mod_ImportCollisionGate (標準模組)
' PURPOSE: 統一處理「即將進來的檔案，是否撞名到等待救援中的紀錄」，
'          不論來源是使用者選取的檔案，還是REBUILD重新掃描整理夾時遇到的檔案，
'          一次彙整、一次詢問、一次記錄決策。純粹負責「問不問」，
'          不碰任何實體複製或目錄表以外的寫入。
' EXPORTS: GatherConflicts, AskAndApplyDecisions
' IMPORTS: Mod_Utils, Mod_UIMessenger, Mod_Actions
' FORBIDDEN: 嚴禁在此複製、搬移一般業務檔案；僅允許歸檔被取代的舊救援殘影
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_ImportCollisionGate"
Private Const MODULE_VERSION As String = "1.0.0"


' ------------------------------------------------------------------------------
' 主流程：合併使用者選取檔案(IMPORT) + 整理夾遞迴掃描(REBUILD)，統一比對、統一詢問
' ------------------------------------------------------------------------------
Public Sub GatherConflicts(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal wsCat As Worksheet, ByVal opMode As String)
    Dim dictPendingByOrigName As Object
    Set dictPendingByOrigName = CreateObject("Scripting.Dictionary"): dictPendingByOrigName.CompareMode = 1
    Dim dictPendingRowByOrigName As Object
    Set dictPendingRowByOrigName = CreateObject("Scripting.Dictionary"): dictPendingRowByOrigName.CompareMode = 1

    Dim preLastRow As Long: preLastRow = wsCat.Cells(wsCat.Rows.count, "A").End(xlUp).row
    Dim preR As Long
    For preR = 5 To preLastRow
        Dim preName As String: preName = Trim(CStr(wsCat.Cells(preR, 1).Value))
        Dim preIsBounty As Boolean: preIsBounty = (Trim(CStr(wsCat.Cells(preR, 8).Value)) = "是")
        If preName <> "" And preIsBounty Then
            dictPendingByOrigName(UCase(preName)) = preName
            dictPendingRowByOrigName(UCase(preName)) = preR
        End If
    Next preR

    If dictPendingByOrigName.count = 0 Then Exit Sub

    Dim colAllCandidates As New Collection
    Dim selItem As Variant

    If opMode = "IMPORT" Then
        For Each selItem In Context.SelectedFiles
            colAllCandidates.Add CStr(selItem)
        Next selItem
    ElseIf opMode = "REBUILD" Then
        Dim colOrganizeFiles As New Collection
        Call Mod_Actions.GatherFilesRecursive(fso, config.IOOrganizeFolder, colOrganizeFiles, config)
        Dim orgItem As Variant
        For Each orgItem In colOrganizeFiles
            colAllCandidates.Add CStr(orgItem)
        Next orgItem
    End If

    Dim colNameMatched As New Collection
    Dim itemPath As Variant
    For Each itemPath In colAllCandidates
        Dim baseName As String: baseName = fso.GetBaseName(CStr(itemPath))
        If dictPendingByOrigName.Exists(UCase(baseName)) Then colNameMatched.Add CStr(itemPath)
    Next itemPath

    If colNameMatched.count = 0 Then Exit Sub

    Dim hashDict As Object: Set hashDict = Mod_Utils.BuildHashDictionary(fso, colNameMatched, config)
    Dim colTrueConflicts As New Collection

    For Each itemPath In colNameMatched
        Dim mBaseName As String: mBaseName = fso.GetBaseName(CStr(itemPath))
        Dim pendingRow As Long: pendingRow = dictPendingRowByOrigName(UCase(mBaseName))
        Dim pendingCurrHash As String: pendingCurrHash = Trim(CStr(wsCat.Cells(pendingRow, 2).Value))
        Dim pendingOrigHash As String: pendingOrigHash = Trim(CStr(wsCat.Cells(pendingRow, 9).Value))
        Dim selHash As String: selHash = ""
        If hashDict.Exists(CStr(itemPath)) Then selHash = hashDict(CStr(itemPath))

        If selHash <> "" And selHash <> pendingCurrHash And selHash <> pendingOrigHash Then
            colTrueConflicts.Add itemPath
        End If
    Next itemPath

    If colTrueConflicts.count = 0 Then Exit Sub

    ' 【這段放在這裡】：找到真衝突後，不彈窗，只負責記下來，交給AskAndApplyDecisions事後統一問
    Dim conflictItem As Variant
    For Each conflictItem In colTrueConflicts
        Dim conflictBaseName As String: conflictBaseName = fso.GetBaseName(CStr(conflictItem))
        Context.PendingConflicts(CStr(conflictItem)) = dictPendingByOrigName(UCase(conflictBaseName))
    Next conflictItem
End Sub


Public Sub AskAndApplyDecisions(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal wsCat As Worksheet)
    If Context.PendingConflicts.count = 0 Then Exit Sub
    
    Dim collisionList As String: collisionList = ""
    Dim collisionCount As Long: collisionCount = 0
    Dim conflictPath As Variant
    
    For Each conflictPath In Context.PendingConflicts.Keys
        collisionCount = collisionCount + 1
        If collisionCount <= 10 Then collisionList = collisionList & "- " & fso.GetBaseName(CStr(conflictPath)) & vbCrLf
    Next conflictPath
    
    If collisionCount > 10 Then collisionList = collisionList & "...(其餘 " & (collisionCount - 10) & " 筆省略)" & vbCrLf
    
    Application.ScreenUpdating = True
    Dim collisionMsg As String
    collisionMsg = "偵測到 " & collisionCount & " 張圖片，與正在等待救援的原始檔名相同：" & vbCrLf & vbCrLf & _
                   collisionList & vbCrLf & _
                   "是否以當前圖片取代原始圖片？" & vbCrLf & _
                   "警告：圖片歷史脈絡將一併消失。" & vbCrLf & _
                   "備註：被取代之救援圖片，可至【已歸位】資料夾尋找，取代錯誤可以「新圖片」身分再次匯入。"
    Dim collisionAnswer As Boolean
    collisionAnswer = (Mod_UIMessenger.AskQuestion(collisionMsg, "偵測到同名圖片") = vbYes)
    Application.ScreenUpdating = False
    Context.LogStore.Record "TX_INFO", "IMPORT_COLLISION", "同名撞名詢問，使用者選擇：" & IIf(collisionAnswer, "是(取代)", "否(不取代)") & "，共 " & collisionCount & " 筆"
    
    
    Set Context.NameCollisionDecisions = CreateObject("Scripting.Dictionary")
    Context.NameCollisionDecisions.CompareMode = 1
    
    For Each conflictPath In Context.PendingConflicts.Keys
        Dim originalIdentifier As String: originalIdentifier = Context.PendingConflicts(conflictPath)
        Dim cBaseName As String: cBaseName = fso.GetBaseName(CStr(conflictPath))
        Context.NameCollisionDecisions(UCase(cBaseName)) = collisionAnswer
        
        If collisionAnswer = True Then
            Dim freshRow As Long: freshRow = 0
            Dim freshLastRow As Long: freshLastRow = wsCat.Cells(wsCat.Rows.count, "A").End(xlUp).row
            Dim searchR As Long
            For searchR = 5 To freshLastRow
                If Trim(CStr(wsCat.Cells(searchR, 1).Value)) = originalIdentifier Then
                    freshRow = searchR
                    Exit For
                End If
            Next searchR
            
            If freshRow = 0 Then
                Context.LogStore.Record "TX_ERROR", "IMPORT_COLLISION", "取代失敗：目錄表查無此識別名稱=" & originalIdentifier
            End If
            
            If freshRow > 0 Then
                ' 【修正】不再憑「識別名稱.副檔名」猜測舊檔路徑，改讀目錄表第3欄
                ' 自己記錄的實際路徑——這是唯一該信任的來源，不管副檔名、命名規則
                ' 未來怎麼變，這裡永遠對得上目前真正的實體位置
                Dim priorRescuePath As String: priorRescuePath = Trim(CStr(wsCat.Cells(freshRow, 3).Value))
                Context.LogStore.Record "TX_INFO", "IMPORT_COLLISION", "取代成功：識別名稱=" & originalIdentifier & " | 第" & freshRow & "列 | 舊路徑=" & priorRescuePath & " | 舊路徑存在嗎=" & fso.FileExists(priorRescuePath)
                If priorRescuePath <> "" And fso.FileExists(priorRescuePath) Then
                    Call Mod_Actions.ArchiveRestoredFile(fso, priorRescuePath, config, Context.LogStore)
                End If
                wsCat.Cells(freshRow, 7).Value = "已取代"
                wsCat.Cells(freshRow, 8).Value = "否"
                ' 【補漏，Grok審閱發現】明確取代時，血統錨點(OrigHash)不該留著舊值——
                ' 雖然isBounty=否已經讓cmd_Deduplicate的候選清單不會誤配對到這筆，
                ' 功能上安全，但資料本身具誤導性，看起來像還在追蹤某個雜湊，實際上已作廢
                wsCat.Cells(freshRow, 9).Value = ""
            End If
        End If
    Next conflictPath
    
    Context.PendingConflicts.RemoveAll
End Sub
