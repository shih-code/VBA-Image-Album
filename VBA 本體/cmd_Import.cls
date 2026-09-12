Option Explicit
Implements ICommand

' ==========================================================
' MODULE: cmd_Import (類別模組)
' PURPOSE: 智慧匯入命令。實裝分段進度回報 (Mod 20 更新狀態列 / Mod 60 儲存活頁簿) 與巨型檔案處理機制。
' EXPORTS: ICommand_Execute, ICommand_CommandName, ICommand_IsCritical, ICommand_Dependencies
' IMPORTS: cls_ExecutionContext, Mod_Actions, Mod_Rules, Mod_Utils
' FORBIDDEN: 1. 嚴禁直接呼叫 MsgBox。
'            2. 嚴禁在此執行「去重比對」(已移交 cmd_Deduplicate)。
' DEPENDENCIES: Microsoft Scripting Runtime (FileSystemObject, Dictionary)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "cmd_Import"
Private Const MODULE_VERSION As String = "1.0.0"



' ==============================================================================
' ICommand 介面規範實作區
' ==============================================================================
Private Property Get ICommand_IsCritical() As Boolean
    ICommand_IsCritical = True
End Property

Private Property Get ICommand_CommandName() As String
    ICommand_CommandName = "IMPORT"
End Property

Private Property Get ICommand_Dependencies() As Variant
    ICommand_Dependencies = Array()
End Property

Private Sub ICommand_Execute(ByVal Context As cls_ExecutionContext)
    Call ProcessImport(Context)
End Sub



' ==============================================================================
' 核心主流程：智慧匯入、散落檔案收編與命名分流
' ==============================================================================
Private Sub ProcessImport(ByVal Context As cls_ExecutionContext)
    Dim fso As Object, config As cls_Config, LogStore As Object
    Dim ctxItem As Variant, checkPath As String
    Dim currentIdx As Long
    Dim dictBatchHash As Object
    Set dictBatchHash = CreateObject("Scripting.Dictionary"): dictBatchHash.CompareMode = 1
    
    Set config = Context.config
    Set LogStore = Context.LogStore
    
    If config Is Nothing Or LogStore Is Nothing Then
        Err.Raise 9991, MODULE_NAME & ".Execute", "[系統報警] 執行期上下文 (Context) 遺失核心參照，管線拒絕啟動"
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    LogStore.Record "TX_START", "IMPORT", "開始執行實體檔案匯入與平級三柱分流"
    
     ' 【安全網】開始前先確認磁碟空間充足，避免明知會失敗還硬跑
    If Not Mod_Actions.IsDiskSpaceSafe(config, LogStore) Then
        LogStore.Record "TX_ERROR", "IMPORT", "磁碟空間不足，匯入作業中止，請先騰出空間再重新執行"
        Exit Sub
    End If
    
    ' [路徑防護] 禁止匯入系統保護目錄，避免遞迴誤刪
   ' 【V2-073修正】原本用InStr鬆散比對寫死的中文字串，改用config屬性動態組出
   ' 保護清單，並改為嚴謹的「路徑開頭比對」（Left比對，非InStr任意位置比對）——
   ' 避免使用者自己的資料夾名稱剛好包含這幾個字被誤攔，也避免資料夾改名後
   ' 這段防護悄悄失效卻無任何錯誤提示（Grok審閱發現，經查證屬實）
   Dim protectedFolders As Variant
   protectedFolders = Array(config.IOOrganizeFolder, config.IOTargetFolder, config.IORescueFolder, _
                             config.IOAttachmentFolder, config.IOExportFolder, config.IOQuarantineFolder)
       Dim protectedItem As Variant
       For Each protectedItem In protectedFolders
           Dim protectedPath As String: protectedPath = UCase(CStr(protectedItem))
           If Right(protectedPath, 1) <> "\" Then protectedPath = protectedPath & "\"
           If Left(checkPath, Len(protectedPath)) = protectedPath Then
               Err.Raise 9992, MODULE_NAME, "使用者試圖匯入系統保護目錄，已由防禦系統攔截。"
           End If
       Next protectedItem

    Call Mod_Actions.InitializeSystemFolders(config)

    ' ==========================================================================
    ' 【第二階段：使用者選取檔案匯入】命名與歸檔各自獨立判斷，統一走沙盒
    ' ==========================================================================
    currentIdx = 0
    For Each ctxItem In Context.SelectedFiles
        If fso.FolderExists(CStr(ctxItem)) Then
            Call ImportFolderRecursive(fso, config, Context, CStr(ctxItem), CStr(ctxItem), dictBatchHash, currentIdx)
        ElseIf fso.FileExists(CStr(ctxItem)) Then
            On Error GoTo SingleErrorHandler
            currentIdx = currentIdx + 1
            
            If currentIdx Mod 20 = 0 Then
                Application.StatusBar = "正在匯入 (" & currentIdx & ")"
                LogStore.Record "TX_PROGRESS", "IMPORT", "正在匯入 (" & currentIdx & ")"
                DoEvents
            End If
            If currentIdx Mod 60 = 0 Then config.TargetWB.Save
            
            If ImportSingleFileViaSandbox(fso, config, Context, CStr(ctxItem), "", False, dictBatchHash) Then
                Context.SessionImported = Context.SessionImported + 1
            End If
        End If
SingleNextItem:
    Next ctxItem
    On Error GoTo 0
    
    LogStore.Record "TX_COMMIT_FS", "IMPORT", "成功完成匯入作業，全域雜湊記錄已更新"

ImportCleanup:
    LogStore.Record "TX_PROGRESS", "IMPORT", "正在執行後序空資料夾清理作業..."
    DoEvents
    Call Mod_Utils.CleanEmptyFoldersBottomUp(fso, config.IOOrganizeFolder, config.IOOrganizeFolder, config)

    Set fso = Nothing: Set config = Nothing: Set LogStore = Nothing
    Exit Sub

SingleErrorHandler:
    Context.SessionFailed = Context.SessionFailed + 1
    LogStore.Record "TX_ERROR", "IMPORT", "個別檔案匯入失敗：" & CStr(ctxItem) & " | 原因：" & Err.Description
    Context.SessionFailedReasons = Context.SessionFailedReasons & fso.GetFileName(CStr(ctxItem)) & "：" & Err.Description & vbCrLf
    Err.Clear
    Resume SingleNextItem
End Sub

' ------------------------------------------------------------------------------
' 內部輔助引擎：ImportFolderRecursive
' ------------------------------------------------------------------------------
Private Sub ImportFolderRecursive(ByVal fso As Object, ByVal config As cls_Config, ByVal Context As cls_ExecutionContext, _
                                    ByVal currentPath As String, ByVal rootFolder As String, _
                                    ByVal dictBatchHash As Object, ByRef currentIdx As Long)
    Dim fld As Object, subFld As Object, fileObj As Object
    If Not fso.FolderExists(currentPath) Then Exit Sub
    Set fld = fso.GetFolder(currentPath)
    
    For Each fileObj In fld.Files
        On Error GoTo FolderFileErrorHandler
        currentIdx = currentIdx + 1
        If currentIdx Mod 20 = 0 Then
            Application.StatusBar = "正在匯入 (" & currentIdx & ")"
            Context.LogStore.Record "TX_PROGRESS", "IMPORT", "正在匯入 (" & currentIdx & ")"
            DoEvents
        End If
        If currentIdx Mod 60 = 0 Then config.TargetWB.Save
        
        If ImportSingleFileViaSandbox(fso, config, Context, fileObj.Path, rootFolder, True, dictBatchHash) Then
            Context.SessionImported = Context.SessionImported + 1
        End If
FolderFileNextItem:
    Next fileObj
    On Error GoTo 0
    
    For Each subFld In fld.SubFolders
        Call ImportFolderRecursive(fso, config, Context, subFld.Path, rootFolder, dictBatchHash, currentIdx)
    Next subFld
    
    Set fld = Nothing: Set fileObj = Nothing: Set subFld = Nothing
    Exit Sub

FolderFileErrorHandler:
    Context.SessionFailed = Context.SessionFailed + 1
    Context.LogStore.Record "TX_ERROR", "IMPORT", "資料夾匯入失敗：" & fileObj.Path & " | 原因：" & Err.Description
    Context.SessionFailedReasons = Context.SessionFailedReasons & fileObj.Name & "：" & Err.Description & vbCrLf
    Err.Clear
    Resume FolderFileNextItem
End Sub


' ==============================================================================
' 沙盒複製核心：所有匯入情境（含母夾根目錄散落檔案收編）統一呼叫這一支
' ==============================================================================
Private Function ImportSingleFileViaSandbox(ByVal fso As Object, ByVal config As cls_Config, _
                                              ByVal Context As cls_ExecutionContext, _
                                              ByVal sourcePath As String, _
                                              ByVal sourceRootFolder As String, _
                                              ByVal isFolderImport As Boolean, _
                                              ByVal dictBatchHash As Object) As Boolean
    ImportSingleFileViaSandbox = False
    
    Dim ext As String: ext = LCase(fso.GetExtensionName(sourcePath))
    Dim cat As String: cat = Mod_Rules.GetFileCategory(ext)
    
    Dim fHash As String: fHash = GetSmartHash(fso, sourcePath, config)
    
    ' 批次內去重：這批已經處理過同雜湊的檔案，不重複匯入
    If dictBatchHash.Exists(fHash) Then Exit Function
    
    ' 1. 沙盒：複製進暫存區，固定檔名主體，副檔名照原始帶過去
    Dim sandboxPath As String
    sandboxPath = Mod_Utils.BuildSafeWindowsPath(fso, config.IOSandboxFolder, "temp_working." & ext)
    Call Mod_Utils.CopyPhysicalFile(fso, sourcePath, sandboxPath)
    
    ' 2. 清洗：決定正式名稱（僅圖片類別套用雜湊尾巴旗幟命名，非圖片走原本的GenerateUniqueFileName）
    Dim sourceFileObj As Object: Set sourceFileObj = fso.GetFile(sourcePath)
    Dim finalName As String
    
    If cat = "IMAGE" Then
       ' 【修正】圖片的雜湊會被永久寫進檔名跟目錄表，錯誤代價遠高於附件的
       ' 批次查重，這裡改用真正的MD5，不沿用上面GetSmartHash可能給的替代值
       Dim trueImageHash As String: trueImageHash = Mod_Utils.GetMD5(fso, sourcePath)
       finalName = DetermineFinalName(Context, fso.GetBaseName(sourcePath), trueImageHash, sourceFileObj) & "." & ext
    Else
        ' 非圖片附件，維持原有流水號防撞名邏輯，不套用雜湊旗幟
        Dim targetCatFolder As String: targetCatFolder = Mod_Rules.GetDestinationCategory(ext, config)
        Call Mod_Utils.EnsureFolderExists(fso, targetCatFolder)
        finalName = Mod_Utils.GenerateUniqueFileName(fso, targetCatFolder, fso.GetFileName(sourcePath))
    End If
    
    ' 3. 歸檔：決定正式資料夾
    Dim finalFolder As String
    If cat = "IMAGE" Then
        finalFolder = DetermineFinalFolder(fso, config, sourcePath, sourceRootFolder, isFolderImport)
    Else
        finalFolder = Mod_Rules.GetDestinationCategory(ext, config)
    End If
    Call Mod_Utils.EnsureFolderExists(fso, finalFolder)
    
    Dim finalPath As String
    finalPath = Mod_Utils.BuildSafeWindowsPath(fso, finalFolder, finalName)
    
    ' 4. 從沙盒複製到正式位置，核對，核對通過才刪除沙盒殘留
    Call Mod_Utils.CopyPhysicalFile(fso, sandboxPath, finalPath)
    
    On Error Resume Next
    fso.DeleteFile sandboxPath, True
    On Error GoTo 0
    
    dictBatchHash(fHash) = finalPath
    
    If cat = "UNKNOWN" Then
        On Error Resume Next
        Call Mod_Actions.WriteQuarantineLog(config, fso.GetFileName(sourcePath), finalName, "未知格式，安全隔離", finalPath)
        On Error GoTo 0
    End If
    
    Set sourceFileObj = Nothing
    ImportSingleFileViaSandbox = True
End Function

' ------------------------------------------------------------------------------
' 職責單一：只回答「這個檔案該取什麼名字」
' ------------------------------------------------------------------------------
Private Function DetermineFinalName(ByVal Context As cls_ExecutionContext, _
                                      ByVal originalBaseName As String, _
                                      ByVal fileHash As String, _
                                      ByVal sourceFile As Object) As String
    Dim basePattern As String
    
    If Context.DynamicNamingMode = "原始名稱" Then
        If Context.NameCollisionDecisions.Exists(UCase(originalBaseName)) Then
            ' 不管選是或否，這都是雜湊不同的全新獨立身分，一律拔除不屬於這張圖的烙印與雜湊尾巴
            basePattern = Mod_Utils.StripHashSuffix(Mod_Utils.ExtractOriginalIdentifier(originalBaseName))
        Else
            basePattern = originalBaseName
        End If
    ElseIf Context.DynamicNamingMode = "今日日期" Then
        basePattern = Format(Now, "yyyymmdd")
    ElseIf Context.DynamicNamingMode = "修改日期" Then
        basePattern = Format(sourceFile.DateLastModified, "yyyymmdd")
    ElseIf Context.DynamicNamingMode = "自訂~今日日期" Then
        Dim customPrefixToday As String: customPrefixToday = Trim(Context.DynamicSheetPrefix)
        If customPrefixToday = "" Then customPrefixToday = "CUSTOM"
        basePattern = customPrefixToday & "~" & Format(Now, "yyyymmdd")
    ElseIf Context.DynamicNamingMode = "自訂~修改日期" Then
        Dim customPrefixModified As String: customPrefixModified = Trim(Context.DynamicSheetPrefix)
        If customPrefixModified = "" Then customPrefixModified = "CUSTOM"
        basePattern = customPrefixModified & "~" & Format(sourceFile.DateLastModified, "yyyymmdd")
    Else
        basePattern = Trim(Context.DynamicSheetPrefix)
        If basePattern = "" Then basePattern = "CUSTOM"
    End If
    
   ' 【V2.3.0】統一成「丟棄」，不再「轉換保留」——跟REBUILD那邊的邏輯一致，
   ' 不管圖片是從匯入進來還是REBUILD直接掃到，括號流水號跟複製文字一律
   ' 直接洗掉，不留痕跡。若使用者需要保留原本的流水號數字，須自行先用
   ' Mod_FilenameNormalizer（轉換保留）處理過，再執行匯入
   Dim copyStripLoopImport As Integer: copyStripLoopImport = 0
   Do
        Dim beforeCopyStripImport As String: beforeCopyStripImport = basePattern
        basePattern = Mod_Utils.StripTrailingParenNumber(basePattern)
        basePattern = Mod_Utils.StripKnownCopySuffix(basePattern, Context.config.CopySuffixPatterns)
        copyStripLoopImport = copyStripLoopImport + 1
   Loop While basePattern <> beforeCopyStripImport And copyStripLoopImport < 30
    
    DetermineFinalName = basePattern & "_#HASH" & Left(fileHash, 8)
End Function

' ------------------------------------------------------------------------------
' 職責單一：只回答「這個檔案該放進哪個資料夾」
' ------------------------------------------------------------------------------
Private Function DetermineFinalFolder(ByVal fso As Object, ByVal config As cls_Config, _
                                        ByVal sourcePath As String, _
                                        ByVal sourceRootFolder As String, _
                                        ByVal isFolderImport As Boolean) As String
    If Not isFolderImport Then
        DetermineFinalFolder = config.IOLooseImportFolder
        Exit Function
    End If
    
    ' 資料夾匯入：保留相對於「使用者選擇的來源根目錄」的巢狀結構
    Dim outerFolderName As String: outerFolderName = fso.GetFileName(sourceRootFolder)
    Dim srcParent As String: srcParent = fso.GetParentFolderName(sourcePath)
    
    Dim relativePath As String
    If Len(srcParent) >= Len(sourceRootFolder) Then
        relativePath = Mid(srcParent, Len(sourceRootFolder) + 1)
    Else
        relativePath = ""
    End If
    
    If Left(relativePath, 1) = "\" Then relativePath = Mid(relativePath, 2)
    
    If relativePath = "" Then
        DetermineFinalFolder = Mod_Utils.BuildSafeWindowsPath(fso, config.IOOrganizeFolder, outerFolderName)
    Else
        DetermineFinalFolder = Mod_Utils.BuildSafeWindowsPath(fso, config.IOOrganizeFolder, outerFolderName & "\" & relativePath)
    End If
End Function
