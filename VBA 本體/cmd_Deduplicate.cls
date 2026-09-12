Option Explicit
Implements ICommand

' ==========================================================
' MODULE: cmd_Deduplicate (類別模組)
' PURPOSE: 全域資產去重命令。
'          依優先權（原始 HASH > 新 HASH）比對是否命中既有目錄列；
'          未命中之檔案，名稱用於判斷該雜湊是否為首次出現、應否登記為新項目，
'          並非比對優先權的一環。
'          執行「下架、取代、正名、清除」的原子操作序列。
' EXPORTS: ICommand_Execute 等介面合約
' IMPORTS: cls_ExecutionContext, Mod_Rules, Mod_Utils, Mod_Actions, cls_Config, cls_StringLibrary
' FORBIDDEN: 嚴禁直接更改 UI，所有狀態與錯誤皆須交由 LogStore 與 Pipeline 統一處置。
' DEPENDENCIES: Windows FileSystemObject (FSO), Dictionary
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "cmd_Deduplicate"
Private Const MODULE_VERSION As String = "1.0.0"

Private m_DeletedCount As Long
Private m_FailedCount As Long
Private m_MatchedCount As Long

' 核心比對狀態 (紀錄 比對條件 -> 對應的目錄 Row ID)
Private m_DictCandidateOrigHash As Object
Private m_DictCandidateRescueHash As Object


Private m_DictFileHash As Object
Private m_DictFileBaseName As Object
Private m_DictFileExt As Object
Private m_DictFileCat As Object
Private m_DictFileProcessed As Object
Private m_DictFileFinalPath As Object
Private m_DictAttachmentGlobalHash As Object  ' 【V2-046】附件線獨立重複偵測，不與圖片共用 Context.DictGlobalHash

' ==============================================================================
' ICommand 介面實作區
' ==============================================================================
Private Property Get ICommand_IsCritical() As Boolean
    ICommand_IsCritical = True
End Property

Private Property Get ICommand_CommandName() As String
    ICommand_CommandName = "DEDUPLICATE"
End Property

Private Property Get ICommand_Dependencies() As Variant
    ICommand_Dependencies = Array("RESCUE", "IMPORT")
End Property

Private Sub ICommand_Execute(ByVal Context As cls_ExecutionContext)
    Dim fso As Object, config As cls_Config, LogStore As Object
    Dim errNum As Long, errDesc As String
    Dim wsCat As Worksheet

    On Error GoTo DeduplicateErrorHandler

    Set config = Context.config
    Set LogStore = Context.LogStore
    If config Is Nothing Or LogStore Is Nothing Then
        Err.Raise 9991, MODULE_NAME & ".Execute", "Context 遺失"
    End If

    m_DeletedCount = 0
    m_FailedCount = 0
    m_MatchedCount = 0
    Set fso = CreateObject("Scripting.FileSystemObject")

    LogStore.Record "TX_START", "DEDUPLICATE", "開始執行去重比對作業"

    ' --------------------------------------------------------------------------
    ' 步驟 1：初始化比對用字典
    ' --------------------------------------------------------------------------
    Call InitializeDeduplicationState

    ' --------------------------------------------------------------------------
    ' 步驟 2：讀取目錄，登記待救援清單 (將條件綁定到具體的 Row ID)
    ' --------------------------------------------------------------------------
    Set wsCat = Mod_Utils.GetSheetSafe(config.TargetWB, config.SheetNameCatalog)
    Call RegisterRescueCandidates(config, wsCat)

    ' --------------------------------------------------------------------------
    ' 步驟 3：收集整理夾與附件區所有檔案，逐一計算雜湊
    ' --------------------------------------------------------------------------
    Call CollectAndHashFiles(fso, config, Context, LogStore)

    ' ==========================================================================
    ' 步驟 4：分優先權比對 (順序：原HASH > 新HASH)
    ' ==========================================================================
    Call MatchByOriginalHash(fso, config, Context, LogStore, wsCat)
    Call MatchByRescueHash(fso, config, Context, LogStore, wsCat)

    ' ==========================================================================
    ' 步驟 5：未比對命中檔案的降級與清理
    ' ==========================================================================
    Call DowngradeUnmatchedFiles(fso, config, Context, LogStore)

    Dim finalReport As String
    finalReport = "去重作業結束。正名歸位: " & m_MatchedCount & " | 刪除重複: " & m_DeletedCount
    LogStore.Record "TX_COMMIT_FS", "DEDUPLICATE", finalReport
    
        config.TargetWB.Save

    Exit Sub

DeduplicateCleanup:
    Set m_DictCandidateOrigHash = Nothing
    Set m_DictCandidateRescueHash = Nothing
    Set m_DictFileHash = Nothing
    Set m_DictFileBaseName = Nothing
    Set m_DictFileExt = Nothing
    Set m_DictFileCat = Nothing
    Set m_DictFileProcessed = Nothing
    Set m_DictFileFinalPath = Nothing
    Set m_DictAttachmentGlobalHash = Nothing
    Set fso = Nothing
    Set config = Nothing
    Set LogStore = Nothing
    Exit Sub

DeduplicateErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    If Not LogStore Is Nothing Then
        LogStore.Record "TX_ERROR", "DEDUPLICATE", "代碼: " & errNum & " | 描述: " & errDesc
    End If
    Resume DeduplicateCleanup
End Sub

' --------------------------------------------------------------------------
' 【步驟 1】初始化本次執行所需的 9 個比對用字典（Class 層級狀態，供本模組其餘方法共用）
' --------------------------------------------------------------------------
Private Sub InitializeDeduplicationState()
    Set m_DictCandidateOrigHash = CreateObject("Scripting.Dictionary")
    m_DictCandidateOrigHash.CompareMode = 1
    Set m_DictCandidateRescueHash = CreateObject("Scripting.Dictionary")
    m_DictCandidateRescueHash.CompareMode = 1

    Set m_DictFileHash = CreateObject("Scripting.Dictionary")
    m_DictFileHash.CompareMode = 1
    Set m_DictFileBaseName = CreateObject("Scripting.Dictionary")
    m_DictFileBaseName.CompareMode = 1
    Set m_DictFileExt = CreateObject("Scripting.Dictionary")
    m_DictFileExt.CompareMode = 1
    Set m_DictFileCat = CreateObject("Scripting.Dictionary")
    m_DictFileCat.CompareMode = 1
    Set m_DictFileProcessed = CreateObject("Scripting.Dictionary")
    m_DictFileProcessed.CompareMode = 1
    Set m_DictFileFinalPath = CreateObject("Scripting.Dictionary")
    m_DictFileFinalPath.CompareMode = 1
     Set m_DictAttachmentGlobalHash = CreateObject("Scripting.Dictionary")
     m_DictAttachmentGlobalHash.CompareMode = 1
End Sub

' --------------------------------------------------------------------------
' 【步驟 2】讀取目錄，登記待救援清單（將比對條件綁定到具體的 Row ID）
' --------------------------------------------------------------------------
Private Sub RegisterRescueCandidates(ByVal config As cls_Config, ByVal wsCat As Worksheet)
    Dim r As Long, lastR As Long
    If Not wsCat Is Nothing Then
        lastR = wsCat.Cells(wsCat.Rows.count, "A").End(xlUp).row
        For r = 5 To lastR
            Dim cName As String
            cName = Trim(CStr(wsCat.Cells(r, 1).Value))
            Dim cHash As String
            cHash = Trim(CStr(wsCat.Cells(r, 2).Value))
            Dim cOrigHash As String
            cOrigHash = Trim(CStr(wsCat.Cells(r, 9).Value))
            Dim isBounty As Boolean
            isBounty = (Trim(CStr(wsCat.Cells(r, 8).Value)) = "是") Or (InStr(1, CStr(wsCat.Cells(r, 7).Value), "救援") > 0)

            ' 【V2-082修正】OrigHash比對資格從「限定isBounty=是」擴大為「只要曾經被救援過就永遠有效」，
            ' 不管目前是不是已經結案。OrigHash存在的目的本來就是永久記住「這一列真正的原始身分」，
            ' 供未來真原圖重新出現時比對——即使救援已用低畫質快照結案，真原圖回歸時仍應被識別、
            ' 自動取代（沿用ApplyFileMatch既有機制），不該被當成全新獨立項目。
            ' RescueHash（比對目前救援夾內容）維持isBounty限定，因為已結案列的現行雜湊已無意義追蹤。
            If cName <> "" Then
                If cOrigHash <> "" Then
                    m_DictCandidateOrigHash(cOrigHash) = r
                End If
                If isBounty And cHash <> "" Then
                    m_DictCandidateRescueHash(cHash) = r
                End If
            End If

        Next r
    End If
End Sub

' --------------------------------------------------------------------------
' 【步驟 3】收集整理夾與附件區/救援夾所有檔案，逐一計算雜湊
' 【單檔案錯誤隔離】：任何一個檔案算雜湊失敗，只記錄跳過，不影響其餘檔案
' --------------------------------------------------------------------------
Private Sub CollectAndHashFiles(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal LogStore As Object)
    Dim colFiles As New Collection
    If fso.FolderExists(config.IOOrganizeFolder) Then
        Call Mod_Actions.GatherFilesRecursive(fso, config.IOOrganizeFolder, colFiles, config)
    End If
    If fso.FolderExists(config.IOAttachmentFolder) Then
        Call Mod_Actions.GatherFilesRecursive(fso, config.IOAttachmentFolder, colFiles, config)
    End If
    If fso.FolderExists(config.IORescueFolder) Then
        Call GatherRescueFilesRecursive(fso, config.IORescueFolder, colFiles, config)
    End If

    Dim varFile As Variant, filePath As String, fHash As String, baseName As String, ext As String
    Dim hashProgressCounter As Long
    hashProgressCounter = 0

    For Each varFile In colFiles

        On Error GoTo HashLoopErrorHandler

        filePath = CStr(varFile)
        ext = LCase(fso.GetExtensionName(filePath))
        baseName = fso.GetBaseName(filePath)

         Dim fCat As String
         fCat = Mod_Rules.GetFileCategory(ext)

         ' 【V2-047】圖片類別優先查詢 ScanCurrentState 已算好的雜湊（Context.DictPathToHash），
         ' 避免同一次管線執行內重複計算；查無快取（附件、或掃描後才新出現的檔案）才現算
         If fCat = "IMAGE" And Context.DictPathToHash.Exists(filePath) Then
             fHash = Context.DictPathToHash(filePath)
         Else
             fHash = Mod_Utils.GetMD5(fso, filePath)
         End If

        m_DictFileHash.Add filePath, fHash
        m_DictFileBaseName.Add filePath, baseName
        m_DictFileExt.Add filePath, ext
        m_DictFileCat.Add filePath, fCat
        m_DictFileProcessed.Add filePath, False
        m_DictFileFinalPath.Add filePath, filePath

        ' 【批次防線】：對齊全系統既有節奏，此步驟不涉及剪貼簿，僅需進度回報與存檔
        hashProgressCounter = hashProgressCounter + 1
        If hashProgressCounter Mod 20 = 0 Then
            Application.StatusBar = "正在計算雜湊 (" & hashProgressCounter & " / " & colFiles.count & ")"
            LogStore.Record "TX_PROGRESS", "DEDUPLICATE", "雜湊計算進度：" & hashProgressCounter
            DoEvents
        End If
        
        ' 【V2-045】不碰剪貼簿，無COM瞬斷風險；目錄表為硬碟衍生紀錄（見ADR V2-044），
        ' 中途崩潰造成的落差由下次全面排版自動修正，改為整支命令結束後統一存檔一次

HashLoopNextFile:
    Next varFile

    On Error GoTo 0
    Exit Sub

HashLoopErrorHandler:
    Context.SessionFailed = Context.SessionFailed + 1
    LogStore.Record "TX_ERROR", "DEDUPLICATE", "雜湊計算失敗，已跳過：" & filePath & " | 原因：" & Err.Description
    Err.Clear
    Resume HashLoopNextFile
End Sub

' --------------------------------------------------------------------------
' 【步驟 4 上半：第二輪】原 HASH 命中 → 正名取代
' --------------------------------------------------------------------------
Private Sub MatchByOriginalHash(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal LogStore As Object, ByVal wsCat As Worksheet)
    Dim kPath As Variant, fHash As String
    Dim round2ProgressCounter As Long
    round2ProgressCounter = 0

    For Each kPath In m_DictFileHash.Keys
        On Error GoTo Round2ErrorHandler
        If Not m_DictFileProcessed(kPath) And m_DictFileCat(kPath) = "IMAGE" And InStr(1, UCase(CStr(kPath)), UCase(config.IORescueFolder)) = 0 Then
            fHash = m_DictFileHash(kPath)
            If m_DictCandidateOrigHash.Exists(fHash) Then
                Call ApplyFileMatch(kPath, m_DictCandidateOrigHash(fHash), fso, config, Context, wsCat)
            End If
        End If

        ' 【批次防線】：對齊既有雜湊迴圈節奏，每 20 筆回報進度、每 60 筆存檔
        round2ProgressCounter = round2ProgressCounter + 1
        If round2ProgressCounter Mod 20 = 0 Then
            Application.StatusBar = "正在比對原HASH (" & round2ProgressCounter & " / " & m_DictFileHash.count & ")"
            LogStore.Record "TX_PROGRESS", "DEDUPLICATE", "第二輪比對進度：" & round2ProgressCounter
            DoEvents
        End If
        
        ' 【V2-045】同上，理由見雜湊計算迴圈註解
       
Round2NextFile:
    Next kPath
    On Error GoTo 0
    Exit Sub

Round2ErrorHandler:
    LogStore.Record "TX_ERROR", "DEDUPLICATE", "第二輪原HASH比對失敗，已跳過：" & CStr(kPath) & " | 原因：" & Err.Description
    Err.Clear
    Resume Round2NextFile
End Sub

' --------------------------------------------------------------------------
' 【步驟 4 下半：第三輪】新 HASH 命中 → 正名取代
' --------------------------------------------------------------------------
Private Sub MatchByRescueHash(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal LogStore As Object, ByVal wsCat As Worksheet)
    Dim kPath As Variant, fHash As String
    Dim round3ProgressCounter As Long
    round3ProgressCounter = 0

    For Each kPath In m_DictFileHash.Keys
        On Error GoTo Round3ErrorHandler
        If Not m_DictFileProcessed(kPath) And m_DictFileCat(kPath) = "IMAGE" And InStr(1, UCase(CStr(kPath)), UCase(config.IORescueFolder)) = 0 Then
            fHash = m_DictFileHash(kPath)
            If m_DictCandidateRescueHash.Exists(fHash) Then
                Call ApplyFileMatch(kPath, m_DictCandidateRescueHash(fHash), fso, config, Context, wsCat)
            End If
        End If
  

        ' 【批次防線】：對齊既有雜湊迴圈節奏，每 20 筆回報進度、每 60 筆存檔
        round3ProgressCounter = round3ProgressCounter + 1
        If round3ProgressCounter Mod 20 = 0 Then
            Application.StatusBar = "正在比對救援HASH (" & round3ProgressCounter & " / " & m_DictFileHash.count & ")"
            LogStore.Record "TX_PROGRESS", "DEDUPLICATE", "第三輪比對進度：" & round3ProgressCounter
            DoEvents
        End If
        
Round3NextFile:
    Next kPath
    On Error GoTo 0
    Exit Sub

Round3ErrorHandler:
    LogStore.Record "TX_ERROR", "DEDUPLICATE", "第三輪救援HASH比對失敗，已跳過：" & CStr(kPath) & " | 原因：" & Err.Description
    Err.Clear
    Resume Round3NextFile
End Sub

' --------------------------------------------------------------------------
' 【步驟 5】未比對命中檔案的降級與清理（實際執行刪除/搬移重複檔案）
' --------------------------------------------------------------------------
Private Sub DowngradeUnmatchedFiles(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal LogStore As Object)
    Dim kPath As Variant, fHash As String, baseName As String

    Context.DictGlobalHash.RemoveAll
    Context.DictRescueHash.RemoveAll

    For Each kPath In m_DictFileHash.Keys

        On Error GoTo Step5FileErrorHandler

        If Not m_DictFileProcessed(kPath) Then
            fHash = m_DictFileHash(kPath)
            Dim fFinalPath As String
            fFinalPath = m_DictFileFinalPath(kPath)
            baseName = m_DictFileBaseName(kPath)

            Dim isRescueLeftover As Boolean
            isRescueLeftover = (InStr(1, UCase(kPath), UCase(config.IORescueFolder)) > 0)

            If isRescueLeftover Then
                If Context.DictRescueHash.Exists(fHash) Then
                    If UCase(Context.DictRescueHash(fHash)) <> UCase(baseName) Then
                        Dim discardName As String
                        discardName = "[" & Mod_Utils.GetFolderDisplayName(config.IORescueDiscardFolder) & "]_" & baseName & "." & m_DictFileExt(kPath)
                        Dim discardPath As String
                        discardPath = Mod_Utils.BuildSafeWindowsPath(fso, config.IORescueDiscardFolder, discardName)

                        On Error Resume Next
                        If fso.FileExists(discardPath) Then
                            fso.DeleteFile discardPath, True
                        End If
                        fso.MoveFile fFinalPath, discardPath
                        If Err.Number = 0 Then
                            m_DeletedCount = m_DeletedCount + 1
                            LogStore.Record "TX_INFO", "DEDUPLICATE", "救援殘影重複，已丟棄：" & baseName & " | 與既有救援紀錄撞HASH | 目的地：" & discardPath
                        Else
                            m_FailedCount = m_FailedCount + 1
                            LogStore.Record "TX_ERROR", "DEDUPLICATE", "救援殘影丟棄失敗：" & baseName & " | 原因：" & Err.Description
                        End If
                        On Error GoTo Step5FileErrorHandler
                    End If
                Else
                    Context.DictRescueHash.Add fHash, baseName
                End If
            Else
                 If m_DictFileCat(kPath) = "IMAGE" Then
                     ' 圖片線：邏輯不動，維持原本查 Context.DictGlobalHash、丟進 IODuplicateDiscardFolder
                     If Context.DictGlobalHash.Exists(fHash) Then
                         If UCase(Context.DictGlobalHash(fHash)) <> UCase(baseName) Then
                             Dim dupDiscardName As String
                             dupDiscardName = "[" & Mod_Utils.GetFolderDisplayName(config.IODuplicateDiscardFolder) & "]_" & baseName & "." & m_DictFileExt(kPath)
                             Dim dupDiscardPath As String
                             dupDiscardPath = Mod_Utils.BuildSafeWindowsPath(fso, config.IODuplicateDiscardFolder, dupDiscardName)

                             On Error Resume Next
                             If fso.FileExists(dupDiscardPath) Then
                                 fso.DeleteFile dupDiscardPath, True
                             End If
                             fso.MoveFile fFinalPath, dupDiscardPath
                             If Err.Number = 0 Then
                                 m_DeletedCount = m_DeletedCount + 1
                                 LogStore.Record "TX_INFO", "DEDUPLICATE", "圖片重複，已丟棄：" & baseName & " | 與既有登記「" & Context.DictGlobalHash(fHash) & "」撞HASH | 目的地：" & dupDiscardPath
                             Else
                                 m_FailedCount = m_FailedCount + 1
                                 LogStore.Record "TX_ERROR", "DEDUPLICATE", "圖片重複丟棄失敗：" & baseName & " | 原因：" & Err.Description
                             End If
                             On Error GoTo Step5FileErrorHandler
                         End If
                     Else
                         Context.DictGlobalHash.Add fHash, baseName
                     End If
                 Else
                     ' 【V2-046新增】附件線：獨立重複偵測，查 m_DictAttachmentGlobalHash（不查圖片專用字典），
                     ' 命中丟進 IOAttachmentDuplicateDiscardFolder（歸檔附錄底下），不再誤入圖片丟棄夾
                     If m_DictAttachmentGlobalHash.Exists(fHash) Then
                         If UCase(m_DictAttachmentGlobalHash(fHash)) <> UCase(baseName) Then
                             Dim attDiscardName As String
                             attDiscardName = "[" & Mod_Utils.GetFolderDisplayName(config.IOAttachmentDuplicateDiscardFolder) & "]_" & baseName & "." & m_DictFileExt(kPath)
                             Dim attDiscardPath As String
                             attDiscardPath = Mod_Utils.BuildSafeWindowsPath(fso, config.IOAttachmentDuplicateDiscardFolder, attDiscardName)

                             On Error Resume Next
                             If fso.FileExists(attDiscardPath) Then
                                 fso.DeleteFile attDiscardPath, True
                             End If
                             fso.MoveFile fFinalPath, attDiscardPath
                             If Err.Number = 0 Then
                                 m_DeletedCount = m_DeletedCount + 1
                                 LogStore.Record "TX_INFO", "DEDUPLICATE", "附件重複，已丟棄：" & baseName & " | 與既有登記「" & m_DictAttachmentGlobalHash(fHash) & "」撞HASH | 目的地：" & attDiscardPath
                             Else
                                 m_FailedCount = m_FailedCount + 1
                                 LogStore.Record "TX_ERROR", "DEDUPLICATE", "附件重複丟棄失敗：" & baseName & " | 原因：" & Err.Description
                                 
                             End If
                             On Error GoTo Step5FileErrorHandler
                         End If
                     Else
                         m_DictAttachmentGlobalHash.Add fHash, baseName
                     End If
                 End If
            End If
        End If

Step5NextFile:
    Next kPath

    On Error GoTo 0
    Exit Sub

Step5FileErrorHandler:
    LogStore.Record "TX_ERROR", "DEDUPLICATE", "步驟5單一檔案去重失敗，已跳過：" & CStr(kPath) & " | 原因：" & Err.Description
    Err.Clear
    Resume Step5NextFile
End Sub



' ------------------------------------------------------------------------------
' 輔助引擎：命中檔案的正名與取代序列
' ------------------------------------------------------------------------------
Private Sub ApplyFileMatch(ByVal origPath As String, ByVal catRow As Long, ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, ByVal wsCat As Worksheet)
    Dim oldName As String, newName As String
    oldName = Trim(CStr(wsCat.Cells(catRow, 1).Value))
    
    ' [1. 清除救援夾內的舊檔]
    Dim oldRescuePath As String
    oldRescuePath = Mod_Utils.BuildSafeWindowsPath(fso, config.IORescueFolder, oldName & ".png")
    If fso.FileExists(oldRescuePath) Then Call Mod_Actions.ArchiveRestoredFile(fso, oldRescuePath, config, Context.LogStore)
    oldRescuePath = Mod_Utils.BuildSafeWindowsPath(fso, config.IORescueFolder, oldName & ".jpg")
    If fso.FileExists(oldRescuePath) Then Call Mod_Actions.ArchiveRestoredFile(fso, oldRescuePath, config, Context.LogStore)
    
    ' [2. 正名]：若是名稱命中且帶有 MANUAL_RESCUE_ 標記，拔除烙印還原原始旗幟識別名稱
    If oldName Like "MANUAL_RESCUE_*" Then
       ' 拔除烙印：MANUAL_RESCUE_日期_ 這段固定前綴拔掉，剩下的就是原本的旗幟識別名稱
       Dim restoredName As String: restoredName = Mod_Utils.ExtractOriginalIdentifier(oldName)

       If restoredName <> "" And Not fso.FileExists(Mod_Utils.BuildSafeWindowsPath(fso, config.IOOrganizeFolder, restoredName & "." & m_DictFileExt(origPath))) Then
           newName = restoredName
       Else
           ' 拔不出來（舊格式殘留）或撞名，退回生成序號名字保底
           Dim basePattern As String: basePattern = Trim(Context.DynamicSheetPrefix)
           If basePattern = "" Then basePattern = "CUSTOM"
           Dim fileSerial As Long: fileSerial = 1
           Dim candidateName As String

           Do
               candidateName = basePattern & "_" & Format(fileSerial, "000")
               fileSerial = fileSerial + 1
           Loop While fso.FileExists(Mod_Utils.BuildSafeWindowsPath(fso, config.IOOrganizeFolder, candidateName & "_#HASH" & Left(m_DictFileHash(origPath), 8) & "." & m_DictFileExt(origPath)))
           newName = candidateName & "_#HASH" & Left(m_DictFileHash(origPath), 8)
       End If
    Else
        ' 原 HASH 命中：沿用原本名稱
        newName = oldName
    End If
    
    ' [3. 實體取代]：將硬碟中的檔案改名就位
    Dim ext As String: ext = m_DictFileExt(origPath)
    Dim newPath As String
    newPath = Mod_Utils.BuildSafeWindowsPath(fso, config.IOOrganizeFolder, newName & IIf(ext <> "", "." & ext, ""))
    
    ' 【修正】這筆目錄列若先前已經被正名取代過一次（例如先用救援複本歸位，
    ' 這次才換真原圖進來取代），整理夾裡可能還留著上一輪正名時的舊實體檔案
    ' （副檔名往往不同，從未被清理過），會在下一輪排版被誤認成全新獨立圖片。
    ' 取代前先讀目錄表目前記錄的路徑，如果跟這次的來源、目的地都不同，先歸檔掉
    Dim priorPath As String: priorPath = Trim(CStr(wsCat.Cells(catRow, 3).Value))
    If priorPath <> "" And UCase(priorPath) <> UCase(newPath) And UCase(priorPath) <> UCase(origPath) Then
        If fso.FileExists(priorPath) Then Call Mod_Actions.ArchiveRestoredFile(fso, priorPath, config, Context.LogStore)
    End If
    
    If UCase(origPath) <> UCase(newPath) Then
        If fso.FileExists(newPath) Then On Error Resume Next: fso.DeleteFile newPath, True: On Error GoTo 0
        fso.MoveFile origPath, newPath
    End If
     
    ' [4. 更新目錄資訊與全域雜湊記錄]
    Dim fileObj As Object: Set fileObj = fso.GetFile(newPath)
    wsCat.Cells(catRow, 1).Value = newName
    wsCat.Cells(catRow, 2).Value = m_DictFileHash(origPath) ' 寫入新 Current HASH
    wsCat.Cells(catRow, 3).Value = newPath
    wsCat.Cells(catRow, 4).Value = fileObj.Size
    wsCat.Cells(catRow, 7).Value = "已歸位" ' 清除救援字眼
    wsCat.Cells(catRow, 8).Value = "否"      ' 解除救援標記
    ' 註：第 9 欄 (OrigHash) 不動，保留原始比對依據！
    
    ' 更新內部字典，標記此檔案處理完畢
    m_DictFileFinalPath(origPath) = newPath
    m_DictFileProcessed(origPath) = True
    Context.DictGlobalHash(m_DictFileHash(origPath)) = newName
    
    ' [5. 移除已處理的比對條件]：將這個 catRow 從待比對清單中拔除，避免其他檔案重複命中
    Dim k As Variant
    For Each k In m_DictCandidateOrigHash.Keys
        If m_DictCandidateOrigHash(k) = catRow Then m_DictCandidateOrigHash.Remove k
    Next k
    For Each k In m_DictCandidateRescueHash.Keys
        If m_DictCandidateRescueHash(k) = catRow Then m_DictCandidateRescueHash.Remove k
    Next k

    m_MatchedCount = m_MatchedCount + 1
End Sub



' ------------------------------------------------------------------------------
' 輔助引擎：GatherRescueFilesRecursive
' 目的：專供本模組收集救援夾檔案使用，不套用 Mod_Rules.IsProtectedFolder 排除規則
'      （該規則是為了防止 IMPORT 匯入時誤抓保護區，去重掃描救援夾不適用這條）
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
