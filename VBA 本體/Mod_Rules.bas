Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_Rules (標準模組)
' PURPOSE: 全系統的規則判斷與防禦中樞。
'          專職執行純邏輯運算、邊界判定，並回傳 True/False 或類別字串。
' EXPORTS: IsProtectedFolder, IsProcessableImage, GetFileCategory, 等
' IMPORTS: cls_Config
' FORBIDDEN: 1. 嚴禁改變任何實體檔案、Excel 儲存格或狀態。
'            2. 嚴禁呼叫 MsgBox 或干涉 UI，異常一律用 Err.Raise 往上拋。
' DEPENDENCIES: Windows FileSystemObject (FSO)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_Rules"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
' 區塊 1：實體硬碟疆界防禦 (Physical Boundary Defense)
' 目的：防止系統遞迴吞噬自身，或將系統檔案誤認為使用者資產。
' ==============================================================================

' ------------------------------------------------------------------------------
' [防禦法規]：判斷是否為系統禁區
' 為什麼需要：防止 cmd_Import 遞迴掃描時，把「救援夾」或「匯出成果」裡的圖又重新匯入一次，造成無限套娃。
' ------------------------------------------------------------------------------
Public Function IsProtectedFolder(ByVal folderPath As String, ByVal config As cls_Config) As Boolean
    Dim uPath As String
    
    If folderPath = "" Then
        IsProtectedFolder = True ' 空路徑視為無效輸入，直接攔截
        Exit Function
    End If
    
    If config Is Nothing Then
        Err.Raise 9995, MODULE_NAME & ".IsProtectedFolder", "[系統報警] Context Config 物件遺失，無法驗證路徑安全。路徑=" & folderPath
    End If
    
    uPath = UCase(Trim(folderPath))
    If Right(uPath, 1) <> "\" Then uPath = uPath & "\"
    
    IsProtectedFolder = False
    
    ' 【保護目錄清單】：救援夾、匯出夾、雜物夾、備份快照
    If InStr(uPath, UCase(config.IORescueFolder)) > 0 Or _
        InStr(uPath, UCase(config.IOExportFolder)) > 0 Or _
        InStr(uPath, UCase(config.IOQuarantineFolder)) > 0 Or _
        InStr(uPath, UCase(config.IOThumbnailCacheFolder)) > 0 Or _
        InStr(uPath, UCase(config.IOSandboxFolder)) > 0 Or _
        InStr(uPath, UCase(config.IODuplicateDiscardFolder)) > 0 Or _
        InStr(uPath, UCase(config.IOBackupSnapshotFolder)) > 0 Then
        IsProtectedFolder = True
    End If
End Function

' ------------------------------------------------------------------------------
' [資產法規]：是否為合法可處理之圖片
' 為什麼需要：Windows 和 Mac 會產生隱藏的系統縮圖快取檔 (如 ~$ 或 ._ 開頭)，必須物理排除。
' ------------------------------------------------------------------------------
Public Function IsProcessableImage(ByVal fso As Object, ByVal fileName As String) As Boolean
    ' 物理排除：Office 鎖定臨時檔 (~$) 與 macOS 隱藏索引檔 (._)
    If Left(fileName, 2) = "~$" Or Left(fileName, 2) = "._" Then
        IsProcessableImage = False
        Exit Function
    End If
    
    ' 對齊 SSOT：將副檔名交由總分類器裁決
    IsProcessableImage = IsImageExtension(fso.GetExtensionName(fileName))
End Function

' ------------------------------------------------------------------------------
' [資產法規]：是否為圖片副檔名 (捷徑函數)
' ------------------------------------------------------------------------------
Public Function IsImageExtension(ByVal ext As String) As Boolean
    IsImageExtension = (GetFileCategory(ext) = "IMAGE")
End Function

' ==============================================================================
' 區塊 2：SSOT 全域分類器
' 目的：決定檔案的分類，是進入畫布的圖片 (IMAGE)，還是歸檔附錄的其他格式 (WORD/PDF...)。
' ==============================================================================

' ------------------------------------------------------------------------------
' [分類法規]：傳回檔案類別特徵碼
' 為什麼需要：將雜亂無章的副檔名降維成 8 大類別，供後續管線進行 O(1) 路由分發。
' ------------------------------------------------------------------------------
Public Function GetFileCategory(ByVal ext As String) As String
    Dim e As String
    e = "|" & LCase(Trim(ext)) & "|"
    
    ' 圖片類：進入圖片整理夾與畫布
    If InStr("|jpg|jpeg|png|bmp|jfif|webp|tiff|tif|", e) > 0 Then
        GetFileCategory = "IMAGE"
        Exit Function
    End If
    
    ' 文件類：進入歸檔附錄
    If e = "|pdf|" Then: GetFileCategory = "PDF": Exit Function
    If InStr("|doc|docx|docm|odt|rtf|", e) > 0 Then: GetFileCategory = "WORD": Exit Function
    If InStr("|xls|xlsx|xlsm|xlsb|csv|ods|", e) > 0 Then: GetFileCategory = "EXCEL": Exit Function
    If InStr("|ppt|pptx|pptm|odp|", e) > 0 Then: GetFileCategory = "PPT": Exit Function
    If InStr("|txt|ini|md|", e) > 0 Then: GetFileCategory = "TXT": Exit Function
    If InStr("|mp4|mov|avi|mkv|webm|flv|wmv|mp3|wav|m4a|flac|aac|ogg|", e) > 0 Then: GetFileCategory = "MEDIA": Exit Function
    If InStr("|zip|rar|7z|tar|gz|", e) > 0 Then: GetFileCategory = "ZIP": Exit Function
    
    ' 未知格式：送往 Quarantine 隔離區
    GetFileCategory = "UNKNOWN"
End Function

' ------------------------------------------------------------------------------
' [路由法規]：傳回檔案最終應當落地的實體資料夾
' ------------------------------------------------------------------------------
Public Function GetDestinationCategory(ByVal ext As String, ByVal config As cls_Config) As String
    Dim category As String
    category = GetFileCategory(ext)
    
    Select Case category
        Case "IMAGE":    GetDestinationCategory = config.IOOrganizeFolder
        Case "PDF":      GetDestinationCategory = config.IOAttachPDFFolder
        Case "WORD":     GetDestinationCategory = config.IOAttachWordFolder
        Case "EXCEL":    GetDestinationCategory = config.IOAttachExcelFolder
        Case "PPT":      GetDestinationCategory = config.IOAttachPPTFolder
        Case "TXT":      GetDestinationCategory = config.IOAttachTXTFolder
        Case "MEDIA":    GetDestinationCategory = config.IOAttachMediaFolder
        Case "ZIP":      GetDestinationCategory = config.IOAttachZipFolder
        Case Else:       GetDestinationCategory = config.IOQuarantineFolder & "未分類\"
    End Select
End Function

' ==============================================================================
' 區塊 3：系統級高風險路徑防禦 (Catastrophe Prevention)
' ==============================================================================

' ------------------------------------------------------------------------------
' [高風險防線]：判斷是否為高危險系統路徑
' 為什麼需要：如果使用者手殘把 "C:\" 或 "桌面" 設為圖片母夾，cmd_Import 的遞迴掃描會直接把整台電腦卡死甚至搞壞檔案。
' ------------------------------------------------------------------------------
Public Function IsDangerousPath(ByVal targetPath As String) As Boolean
    Dim wsh As Object, uPath As String
    Dim sysPaths As Variant, i As Integer
    
    IsDangerousPath = False
    If targetPath = "" Then Exit Function
    
    uPath = UCase(Trim(targetPath))
    If Right(uPath, 1) = "\" Then uPath = Left(uPath, Len(uPath) - 1)
    
    ' [防線 A]：禁止直接選取磁碟機根目錄 (如 C: 或 D:)
    If Len(uPath) <= 3 Then
        IsDangerousPath = True
        Exit Function
    End If
    
    ' [防線 B]：呼叫 WScript.Shell 探測，絕對禁止在桌面、文件、下載區直接發育
    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    sysPaths = Array( _
        UCase(wsh.SpecialFolders("Desktop")), _
        UCase(wsh.SpecialFolders("MyDocuments")), _
        UCase(wsh.SpecialFolders("Downloads")), _
        UCase(Environ("USERPROFILE") & "\Downloads"), _
        UCase(Environ("USERPROFILE") & "\Desktop"), _
        UCase(Environ("USERPROFILE") & "\Documents") _
    )
    Set wsh = Nothing
    On Error GoTo 0
    
    For i = LBound(sysPaths) To UBound(sysPaths)
        If sysPaths(i) <> "" Then
            If uPath = CStr(sysPaths(i)) Then
                IsDangerousPath = True
                Exit Function
            End If
        End If
    Next i
End Function

' ------------------------------------------------------------------------------
' [高風險防線]：探測磁碟剩餘空間是否足以執行搬運或壓縮
' ------------------------------------------------------------------------------
Public Function IsDiskSpaceAvailable(ByVal fso As Object, ByVal folderPath As String, ByVal minSpaceGB As Double) As Boolean
    Dim driveLetter As String, drv As Object, freeGB As Double
    
    On Error GoTo DiskError
    driveLetter = fso.GetDriveName(folderPath)
    
    If driveLetter <> "" Then
        If fso.DriveExists(driveLetter) Then
            Set drv = fso.GetDrive(driveLetter)
            freeGB = drv.FreeSpace / 1024 / 1024 / 1024
            IsDiskSpaceAvailable = (freeGB >= minSpaceGB)
            Set drv = Nothing
            Exit Function
        End If
    End If

    Err.Raise 9996, MODULE_NAME & ".IsDiskSpaceAvailable", "[硬體報警] 無法讀取目標磁碟，磁碟機可能不存在或已斷線。路徑=" & folderPath

DiskError:
    Dim errDesc As String: errDesc = Err.Description
    Err.Clear
    Err.Raise 9996, MODULE_NAME & ".IsDiskSpaceAvailable", "[硬體報警] 磁碟空間探測失敗 (" & errDesc & ")。路徑=" & folderPath
End Function

' ==============================================================================
' 區塊 4：Excel 畫布虛擬疆界防禦 (Virtual Boundary Defense)
' ==============================================================================

' ------------------------------------------------------------------------------
' [畫布法規]：判斷是否為「標準圖片畫布」
' 為什麼需要：排版引擎 (cmd_Rebuild / LayoutEngine) 在重整時會無差別刪除上面的 Shape，
'             必須用白名單嚴格保護系統表與使用者自訂表免遭誤殺。
' ------------------------------------------------------------------------------
Public Function IsStandardImgSheet(ByVal sheetName As String, ByVal config As cls_Config) As Boolean
    If config Is Nothing Then
        Err.Raise 9995, MODULE_NAME & ".IsStandardImgSheet", "[系統報警] Config 物件為空，無法執行畫布白名單驗證。"
    End If

    IsStandardImgSheet = False
    
    ' [防線 A：系統黑名單] 排除所有系統級、日誌與臨時防護工作表
    If InStr(sheetName, config.PrefixOldCatalog) > 0 Or _
       sheetName = config.SheetNameUI Or _
       sheetName = config.SheetNameManual Or _
       sheetName = config.SheetNameQuarantine Or _
       sheetName = config.SheetNameCatalog Or _
       sheetName = config.SheetNameAttachment Or _
       sheetName = config.SheetNameRescue Or _
       sheetName = config.SheetNameAllPics Or _
       sheetName = config.SheetSysMeta Or _
       sheetName = config.SheetSysLog Or _
       sheetName = config.SheetSysTempSettings Or _
       sheetName = config.SheetNameTempSafeRoom Then
        Exit Function
    End If
    
    ' [防線 B：自訂保護區] 使用者自訂分頁絕對不可被視為標準圖片表
    If Left(sheetName, Len(config.PrefixCustom)) = config.PrefixCustom Then
        Exit Function
    End If
    
    ' 通過考驗，認定為標準工作表 (有資格接受圖片掛載與排版)
    IsStandardImgSheet = True
End Function

' ------------------------------------------------------------------------------
' [畫布法規]：判斷是否為「合法圖片表」(包含所有圖表+合併大表)
' ------------------------------------------------------------------------------
Public Function IsValidImageSheet(ByVal sheetName As String, ByVal config As cls_Config) As Boolean
    If config Is Nothing Then Exit Function
    
    ' 合併總表 (所有圖片) 是合法圖表，但它不是標準分流表
    If sheetName = config.SheetNameAllPics Then
        IsValidImageSheet = True
        Exit Function
    End If
    
    IsValidImageSheet = IsStandardImgSheet(sheetName, config)
End Function

' ------------------------------------------------------------------------------
' [畫布法規]：探測工作表是否存在
' ------------------------------------------------------------------------------
Public Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    If wb Is Nothing Then
        Err.Raise 9995, MODULE_NAME & ".SheetExists", "[系統報警] Workbook 物件為空，無法探測工作表是否存在。"
    End If
    
    SheetExists = False
    For Each ws In wb.Worksheets
        If UCase(ws.Name) = UCase(sheetName) Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function
