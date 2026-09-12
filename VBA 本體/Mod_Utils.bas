Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_Utils (標準模組)
' PURPOSE: 基礎建設通用工具箱。
'          提供與特定業務無關的底層原子操作（FSO 檔案處理、字串過濾、空夾清理、WIA/MD5 引擎）。
' EXPORTS: EnsureFolderExists, CleanEmptyFoldersBottomUp, CopyPhysicalFile, GetMD5, 等
' IMPORTS: 無 (純原生 API 與外部依賴注入，實現 100% 解耦)
' FORBIDDEN: 1. 嚴禁在此模組寫入任何具體的業務邏輯 (如判斷是否為救援圖)。
'            2. 嚴禁呼叫 MsgBox 或干涉 UI，異常必須封裝為 Err.Raise 向上拋出。
' DEPENDENCIES: 需由外部注入 Microsoft Scripting Runtime (FileSystemObject)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_Utils"
Private Const MODULE_VERSION As String = "1.0.0"

' Mod_Utils.bas 新增
Private Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As LongLong
    ullAvailPhys As LongLong
    ullTotalPageFile As LongLong
    ullAvailPageFile As LongLong
    ullTotalVirtual As LongLong
    ullAvailVirtual As LongLong
    ullAvailExtendedVirtual As LongLong
End Type

Private Declare PtrSafe Function GlobalMemoryStatusEx Lib "kernel32" (ByRef lpBuffer As MEMORYSTATUSEX) As Long

' ==============================================================================
' 區塊 1：實體檔案與路徑系統安全操作 (Physical File System Operations)
' ==============================================================================

' ------------------------------------------------------------------------------
' [路徑工程]：EnsureFolderExists
' 目的：確保指定資料夾存在。若上層母資料夾不存在，則遞迴往上建立。
' ------------------------------------------------------------------------------
Public Sub EnsureFolderExists(ByVal fso As Object, ByVal folderPath As String)
    Dim parentFolderPath As String
    If folderPath = "" Then Exit Sub
    
    If Right(folderPath, 1) = "\" Then folderPath = Left(folderPath, Len(folderPath) - 1)
    If fso.FolderExists(folderPath) Then Exit Sub
    
    parentFolderPath = fso.GetParentFolderName(folderPath)
    
    If parentFolderPath <> "" Then
        If Not fso.FolderExists(parentFolderPath) Then Call EnsureFolderExists(fso, parentFolderPath)
    End If
    
    fso.CreateFolder folderPath
End Sub

' ------------------------------------------------------------------------------
' [空間自癒]：CleanEmptyFoldersBottomUp
' 目的：由磁碟最深處往上遞迴（後序遍歷），自動拆除搬移檔案後殘留的空殼資料夾。
' ------------------------------------------------------------------------------
Public Sub CleanEmptyFoldersBottomUp(ByVal fso As Object, ByVal currentFolderPath As String, ByVal rootSandboxPath As String, ByVal config As cls_Config)
     On Error Resume Next
     If Not fso.FolderExists(currentFolderPath) Then Exit Sub
     
 
     ' 系統管轄資料夾（快取、沙盒、各種軟刪除區、已歸位圖片等），即使暫時是空的也不能被當成垃圾清掉
    If Mod_Rules.IsProtectedFolder(currentFolderPath, config) Then Exit Sub
 
 
     Dim currentFolder As Object: Set currentFolder = fso.GetFolder(currentFolderPath)
     Dim subFolder As Object
     
     For Each subFolder In currentFolder.SubFolders
        Call CleanEmptyFoldersBottomUp(fso, subFolder.Path, rootSandboxPath, config)
     Next subFolder
 
     If currentFolder.Files.count = 0 And currentFolder.SubFolders.count = 0 Then
         If UCase(Trim(currentFolderPath)) <> UCase(Trim(rootSandboxPath)) Then
             fso.DeleteFolder currentFolderPath, True
         End If
     End If
     
     Set currentFolder = Nothing
     On Error GoTo 0
 End Sub

' ------------------------------------------------------------------------------
' [空間自癒]：EmptyFolderContentsSafe
' 目的：強制清空資料夾底下的所有內容，但不刪除資料夾母體。支援略過系統鎖定檔。
' ------------------------------------------------------------------------------
Public Sub EmptyFolderContentsSafe(ByVal fso As Object, ByVal folderPath As String)
    On Error Resume Next
    If fso.FolderExists(folderPath) Then
        fso.DeleteFile fso.BuildPath(folderPath, "*.*"), True
        fso.DeleteFolder fso.BuildPath(folderPath, "*"), True
    End If
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' [原子操作]：DeletePhysicalFile / CopyPhysicalFile / MovePhysicalFile
' 目的：封裝 FSO 基礎操作，將失敗轉換為結構化報警單 (Err.Raise)，供 Pipeline 熔斷。
' ------------------------------------------------------------------------------
Public Sub DeletePhysicalFile(ByVal fso As Object, ByVal filePath As String)
    Dim errNum As Long, errDesc As String
    On Error Resume Next
    fso.DeleteFile filePath, True
    errNum = Err.Number: errDesc = Err.Description
    On Error GoTo 0
    
    If errNum <> 0 Then
        Err.Raise errNum, MODULE_NAME & ".DeletePhysicalFile", "[硬體報警] 無法刪除實體檔案 (" & errDesc & ") | 路徑=" & filePath
    End If
End Sub

Public Sub CopyPhysicalFile(ByVal fso As Object, ByVal src As String, ByVal dest As String)
    Dim errNum As Long, errDesc As String
    On Error Resume Next
    fso.CopyFile src, dest, True
    errNum = Err.Number: errDesc = Err.Description
    On Error GoTo 0
    
    If errNum <> 0 Then
        Err.Raise errNum, MODULE_NAME & ".CopyPhysicalFile", "[硬體報警] 無法複製實體檔案 (" & errDesc & ") | 來源=" & src
    End If
    
    ' 【核對完整性】：複製指令沒噴錯,不代表內容真的完整一致,比對檔案大小
    Dim srcSize As Double, destSize As Double
    On Error Resume Next
    srcSize = fso.GetFile(src).Size
    destSize = fso.GetFile(dest).Size
    On Error GoTo 0
    
    If srcSize <> destSize Then
        On Error Resume Next
        fso.DeleteFile dest, True
        On Error GoTo 0
        Err.Raise 9997, MODULE_NAME & ".CopyPhysicalFile", "[完整性報警] 複製後檔案大小不一致，來源可能已損毀或複製中斷 | 來源=" & src & " (原始 " & srcSize & " bytes) | 複製結果 (" & destSize & " bytes)"
    End If
End Sub

Public Sub MovePhysicalFile(ByVal fso As Object, ByVal src As String, ByVal dest As String)
    Dim errNum As Long, errDesc As String
    On Error Resume Next
    If fso.FileExists(dest) Then fso.DeleteFile dest, True
    fso.MoveFile src, dest
    errNum = Err.Number: errDesc = Err.Description
    On Error GoTo 0
    
    If errNum <> 0 Then
        Err.Raise errNum, MODULE_NAME & ".MovePhysicalFile", "[硬體報警] 無法移動實體檔案 (" & errDesc & ") | 來源=" & src
    End If
End Sub

' ------------------------------------------------------------------------------
' [命名安全]：GenerateUniqueFileName / GetUniqueFolderName
' 目的：檔案或資料夾撞名時，自動生成時間戳記或流水號防撞。
' ------------------------------------------------------------------------------
Public Function GenerateUniqueFileName(ByVal fso As Object, ByVal folderPath As String, ByVal originalName As String) As String
    Dim baseName As String, ext As String, newName As String, counter As Long, msCode As String
    baseName = fso.GetBaseName(originalName)
    ext = fso.GetExtensionName(originalName)
    newName = originalName
    counter = 1
    msCode = Replace(CStr(Format(Timer, "0.00")), ".", "")

    Do While fso.FileExists(BuildSafeWindowsPath(fso, folderPath, newName))
        newName = baseName & "_保留_" & Format(Now, "hhmmss") & "_" & msCode & "_" & counter
        If ext <> "" Then newName = newName & "." & ext
        counter = counter + 1
    Loop
    GenerateUniqueFileName = newName
End Function

Public Function GetUniqueFolderName(ByVal fso As Object, ByVal parentDir As String, ByVal proposedName As String) As String
    Dim counter As Long, testName As String
    counter = 2: testName = proposedName
    Do While fso.FolderExists(fso.BuildPath(parentDir, testName))
        testName = proposedName & " (" & counter & ")"
        counter = counter + 1
    Loop
    GetUniqueFolderName = testName
End Function

' ------------------------------------------------------------------------------
' [路徑安全]：SanitizeFileName / BuildSafeWindowsPath
' 目的：防禦作業系統層級的特殊符號崩潰與 255 字元長度上限。
' ------------------------------------------------------------------------------
Public Function SanitizeFileName(ByVal rawName As String) As String
    Dim cleanName As String, invalidChars As Variant, i As Long
    cleanName = rawName
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(invalidChars) To UBound(invalidChars)
        cleanName = Replace(cleanName, invalidChars(i), "_")
    Next i
    SanitizeFileName = cleanName
End Function

' ------------------------------------------------------------------------------
' [命名安全]：SanitizeSheetName
' 目的：清洗工作表名稱不合法字元，並防止與系統保留名稱撞名。
'      跟 SanitizeFileName 分開，因為工作表名稱的合法字元集合比檔名更嚴格
'      （額外禁用中括號、開頭結尾單引號），且需要額外的保留名單防撞邏輯。
' ------------------------------------------------------------------------------
Public Function SanitizeSheetName(ByVal rawName As String, ByVal config As cls_Config) As String
    Dim cleanName As String, invalidChars As Variant, i As Long
    cleanName = rawName
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|", "[", "]", "'")

    For i = LBound(invalidChars) To UBound(invalidChars)
        cleanName = Replace(cleanName, invalidChars(i), "_")
    Next i
    
    cleanName = Trim(cleanName)
    If cleanName = "" Then cleanName = "未命名"
    
    cleanName = Left(cleanName, 31)
    
    Dim reservedNames As Variant: reservedNames = Array( _
        config.SheetNameUI, config.SheetNameManual, config.SheetNameQuarantine, _
        config.SheetNameCatalog, config.SheetNameAttachment, config.SheetNameRescue, _
        config.SheetNameAllPics, config.SheetSysMeta, config.SheetSysLog, config.SheetSysTempSettings)
    
    Dim r As Variant
    For Each r In reservedNames
        If UCase(cleanName) = UCase(CStr(r)) Then
            cleanName = Left(cleanName, 27) & "_grp"
            Exit For
        End If
    Next r
    
    SanitizeSheetName = cleanName
End Function

Public Function BuildSafeWindowsPath(ByVal fso As Object, ByVal folderPath As String, ByVal fileName As String) As String
    Dim fullPath As String, ext As String, base As String, overLen As Long
    Static pathAccumulator As Long
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    fullPath = folderPath & fileName
    
    If Len(fullPath) > 250 Then
        ext = fso.GetExtensionName(fileName)
        base = fso.GetBaseName(fileName)
        overLen = Len(fullPath) - 250
        
        If Len(base) > overLen Then
            base = Left(base, Len(base) - overLen)
        Else
            pathAccumulator = pathAccumulator + 1
            base = "TRUNC_" & Replace(CStr(Format(Timer, "0.00")), ".", "") & "_" & pathAccumulator
        End If
        
        If ext <> "" Then fileName = base & "." & ext Else fileName = base
        fullPath = folderPath & fileName
    End If
    
    BuildSafeWindowsPath = fullPath
End Function

' ------------------------------------------------------------------------------
' [資源回收]：MoveToTrashSafe
' 目的：透過 Shell API 呼叫系統資源回收筒 (Namespace 10)，提供使用者反悔機制。
' ------------------------------------------------------------------------------
Public Sub MoveToTrashSafe(ByVal fso As Object, ByVal targetPath As String)
    Dim shellApp As Object
    If Not fso.FolderExists(targetPath) And Not fso.FileExists(targetPath) Then Exit Sub
    
    On Error Resume Next
    Set shellApp = CreateObject("Shell.Application")
    shellApp.Namespace(10).MoveHere targetPath
    Set shellApp = Nothing
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' [死鎖探測]：IsFileLocked
' 目的：嘗試以獨佔寫入模式開啟檔案，探測是否遭防毒軟體或其他程式佔用。
' ------------------------------------------------------------------------------
Public Function IsFileLocked(ByVal fso As Object, ByVal filePath As String) As Boolean
    Dim fileNum As Integer
    On Error Resume Next
    If Not fso.FileExists(filePath) Then
        IsFileLocked = False: On Error GoTo 0: Exit Function
    End If
    
    fileNum = FreeFile
    Open filePath For Input Lock Read Write As #fileNum
    If Err.Number <> 0 Then
        IsFileLocked = True
    Else
        IsFileLocked = False
        Close #fileNum
    End If
    On Error GoTo 0
End Function

' ==============================================================================
' 區塊 2：圖片與媒體處理引擎 (Image & Media Processing)
' ==============================================================================

' ------------------------------------------------------------------------------
' [圖片壓縮]：OptimizeAndCopyImage
' 目的：調用 WIA 引擎等比例縮放圖片，並實裝非同步 I/O 探測與 COM 註銷防線。
' ------------------------------------------------------------------------------
Public Sub OptimizeAndCopyImage(ByVal sourcePath As String, ByVal destPath As String, ByVal fso As Object, ByVal maxDim As Double)
    Dim img As Object, IP As Object, origW As Double, origH As Double, ratio As Double
    Dim t As Single, fileNum As Integer, isProbeSuccess As Boolean
    
    On Error GoTo WIAFallback
    Set img = CreateObject("WIA.ImageFile")
    Set IP = CreateObject("WIA.ImageProcess")
    img.LoadFile sourcePath
    origW = img.Width: origH = img.Height

    If origW > maxDim Or origH > maxDim Then
        If origW >= origH Then ratio = maxDim / origW Else ratio = maxDim / origH
        IP.Filters.Add IP.FilterInfos("Scale").FilterID
        IP.Filters(1).Properties("MaximumWidth") = Application.WorksheetFunction.Max(1, Int(origW * ratio))
        IP.Filters(1).Properties("MaximumHeight") = Application.WorksheetFunction.Max(1, Int(origH * ratio))
        IP.Filters(1).Properties("PreserveAspectRatio") = True

        Set img = IP.Apply(img)
        If fso.FileExists(destPath) Then fso.DeleteFile destPath, True
        img.SaveFile destPath
    Else
        fso.CopyFile sourcePath, destPath, True
    End If

    t = Timer
    Do While IsFileLocked(fso, destPath)
        DoEvents
        If Abs(Timer - t) > 3 Then Exit Do
    Loop
    
    On Error Resume Next
    fileNum = FreeFile
    Open destPath For Input Lock Read As #fileNum
    If Err.Number = 0 Then
        isProbeSuccess = True: Close #fileNum
    Else
        isProbeSuccess = False
    End If
    Err.Clear: On Error GoTo WIAFallback
    
    If Not isProbeSuccess Then Err.Raise 9998, MODULE_NAME & ".OptimizeAndCopyImage", "WIA 算繪輸出遭 Windows 核心鎖定。"

WIAFinalize:
    On Error Resume Next
    Set img = Nothing: Set IP = Nothing
    On Error GoTo 0
    Exit Sub

WIAFallback:
    On Error Resume Next
    Set img = Nothing: Set IP = Nothing
    Err.Clear
    fso.CopyFile sourcePath, destPath, True
    On Error GoTo 0
    Resume WIAFinalize
End Sub

' ------------------------------------------------------------------------------
' [檔案雜湊]：GetMD5
' 目的：產生檔案唯一 HASH。內建 OOM 降級防線，防止巨型檔案撐爆 ADODB 記憶體。
' ------------------------------------------------------------------------------
Public Function GetMD5(ByVal fso As Object, ByVal filePath As String) As String
    Dim adoStream As Object, xmlDoc As Object, xmlNode As Object, md5Provider As Object, bytes() As Byte
    Dim fSize As Double
    On Error GoTo MD5ErrorHandler
    
    If fso.FileExists(filePath) Then fSize = fso.GetFile(filePath).Size
    If fSize > 52428800 Then GoTo MD5ErrorHandler
    
    Set adoStream = CreateObject("ADODB.Stream"): adoStream.Type = 1: adoStream.Open: adoStream.LoadFromFile filePath: bytes = adoStream.Read: adoStream.Close
    Set md5Provider = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider"): bytes = md5Provider.ComputeHash_2((bytes))
    Set xmlDoc = CreateObject("MSXML2.DOMDocument"): Set xmlNode = xmlDoc.createElement("b64")
    xmlNode.DataType = "bin.hex": xmlNode.nodeTypedValue = bytes: GetMD5 = Replace(xmlNode.Text, vbLf, "")

MD5Cleanup:
    Set adoStream = Nothing: Set xmlNode = Nothing: Set xmlDoc = Nothing: Set md5Provider = Nothing
    Exit Function

MD5ErrorHandler:
    If fso.FileExists(filePath) Then
        Dim f As Object: Set f = fso.GetFile(filePath)
        GetMD5 = "LARGE_" & Hex(f.Size) & "_" & Hex(DateDiff("s", "1970/1/1", f.DateLastModified))
    Else
        GetMD5 = ""
    End If
    Resume MD5Cleanup
End Function

' ------------------------------------------------------------------------------
' [輕量探測]：GetImageDimensions
' 目的：無痛讀取實體圖片寬高，不鎖定檔案、不佔用 GDI 記憶體。
' ------------------------------------------------------------------------------
Public Sub GetImageDimensions(ByVal imagePath As String, ByRef outWidth As Double, ByRef outHeight As Double)
    Dim img As Object
    On Error GoTo DimensionErr
    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile imagePath
    outWidth = img.Width: outHeight = img.Height
    Set img = Nothing
    Exit Sub

DimensionErr:
    outWidth = 0: outHeight = 0
    Set img = Nothing: Err.Clear
End Sub

' ==============================================================================
' 區塊 3：Excel 物件與介面控制 (Excel Application API Wrapper)
' ==============================================================================

Public Function GetSheetSafe(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set GetSheetSafe = Nothing
    For Each ws In wb.Worksheets
        If UCase(ws.Name) = UCase(sheetName) Then Set GetSheetSafe = ws: Exit Function
    Next ws
End Function

Public Function GetShapeSafe(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    Dim shp As Shape
    Set GetShapeSafe = Nothing
    For Each shp In ws.Shapes
        If UCase(shp.Name) = UCase(shapeName) Then Set GetShapeSafe = shp: Exit Function
    Next shp
End Function

Public Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = sName
    End If
    Set GetOrCreateSheet = ws
End Function

' ------------------------------------------------------------------------------
' [結構防護]：SafeUnprotect / SafeProtect / GlobalUnprotectAll
' 目的：集中管理密碼鎖，防止硬解碼造成的 1004 崩潰。
' ------------------------------------------------------------------------------
Public Sub SafeUnprotect(ByVal ws As Worksheet, ByVal pwd As String)
    Dim cleanPwd As String: cleanPwd = Trim(pwd)
    On Error Resume Next
    If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
        If cleanPwd <> "" Then ws.Unprotect Password:=cleanPwd
        If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then ws.Unprotect Password:=""
    End If
    On Error GoTo 0
End Sub

Public Sub SafeProtect(ByVal ws As Worksheet, ByVal pwd As String)
    Dim cleanPwd As String: cleanPwd = Trim(pwd)
    On Error Resume Next
    If Not ws.ProtectContents Then
        If cleanPwd = "" Then
            ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, AllowSorting:=True
        Else
            ws.Protect Password:=cleanPwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True, AllowSorting:=True
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub GlobalUnprotectAll(ByVal wb As Workbook, ByVal pwd As String)
    Dim ws As Worksheet
    On Error Resume Next
    wb.Unprotect Password:=pwd
    wb.Unprotect
    For Each ws In wb.Worksheets
        Call SafeUnprotect(ws, pwd)
    Next ws
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' [實體匯出]：ExportShapeToTempFile
' 目的：借用 ChartObject 容器，將 Excel 虛擬 Shape 截圖轉存為實體 JPG。
' ------------------------------------------------------------------------------
Public Function ExportShapeToTempFile(ByVal ws As Worksheet, ByVal shp As Shape) As String
    Dim tmpPath As String, chtObj As Object
    Static loopCounter As Long

    On Error GoTo ShapeExportCleanup
    loopCounter = loopCounter + 1
    tmpPath = Environ("TEMP") & "\app_preview_" & Format(Now, "yyyymmddhhmmss") & "_" & loopCounter & ".jpg"
    
    shp.CopyPicture xlScreen, xlPicture
    DoEvents
    
    Set chtObj = ws.ChartObjects.Add(0, 0, shp.Width, shp.Height)
    chtObj.Activate
    chtObj.Chart.ChartArea.Format.Line.Visible = msoFalse
    chtObj.Chart.Paste
    chtObj.Chart.Export tmpPath, "JPG"
    
    chtObj.Delete
    Set chtObj = Nothing
    ExportShapeToTempFile = tmpPath
    Exit Function

ShapeExportCleanup:
    On Error Resume Next
    If Not chtObj Is Nothing Then
        chtObj.Delete: Set chtObj = Nothing
    End If
    On Error GoTo 0
    ExportShapeToTempFile = ""
End Function

' ------------------------------------------------------------------------------
' [介面美化]：AddHoverPreview
' 目的：在目錄儲存格內插入懸停預覽 (Comment 背景圖)。
' ------------------------------------------------------------------------------
Public Sub AddHoverPreview(ByVal targetCell As Range, ByVal imagePath As String, ByVal origWidth As Double, ByVal origHeight As Double)
    On Error Resume Next
    targetCell.ClearComments
    targetCell.AddComment Text:=" "
    With targetCell.Comment
        .Visible = False
        .Shape.Fill.UserPicture imagePath
        .Shape.Width = 150
        If origWidth > 0 Then .Shape.Height = 150 * (origHeight / origWidth) Else .Shape.Height = 150
    End With
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' [系統環保]：CleanSystemTempFiles
' 目的：定時清洗 Windows %TEMP% 目錄下的過期匯出殘骸 (app_preview_)。
' ------------------------------------------------------------------------------
Public Sub CleanSystemTempFiles(ByVal fso As Object)
    Dim tempFolderObj As Object, fileObj As Object
    Dim tempPath As String
    
    On Error GoTo TempCleanupErr
    tempPath = Environ("TEMP")
    If Not fso.FolderExists(tempPath) Then Exit Sub
    
    Set tempFolderObj = fso.GetFolder(tempPath)
    
    For Each fileObj In tempFolderObj.Files
        If Left(fileObj.Name, 12) = "app_preview_" Then
            If DateDiff("h", fileObj.DateLastModified, Now) >= 1 Then
                On Error Resume Next
                fso.DeleteFile fileObj.Path, True
                On Error GoTo TempCleanupErr
            End If
        End If
    Next fileObj

TempCleanupExit:
    Set tempFolderObj = Nothing: Set fileObj = Nothing
    Exit Sub
TempCleanupErr:
    Resume TempCleanupExit
End Sub

' ------------------------------------------------------------------------------
' [輔助工具]：StripHashSuffix
' 目的：從旗幟識別名稱（格式：原始檔名_雜湊前8碼）拔除雜湊尾巴，還原顯示用名稱。
'       若輸入字串不符合此格式（例如舊格式殘留檔案），原樣傳回，不強制截斷。
' ------------------------------------------------------------------------------
Public Function StripHashSuffix(ByVal identifierName As String) As String
     ' 【V2-066】改找「_#」這兩個字元組成的固定標記，不再只找單一底線，
     ' 避免使用者自己檔名裡的底線（如AA_001的AA/001分隔線）被誤判成雜湊尾巴起點
     Dim markerPos As Long: markerPos = InStrRev(identifierName, "_#")
     If markerPos = 0 Then
        StripHashSuffix = identifierName
        Exit Function
    End If
    
    Dim suffix As String: suffix = Mid(identifierName, markerPos + 2)
    If Len(suffix) = 12 And Left(suffix, 4) = "HASH" And IsHexString(Mid(suffix, 5)) Then
        StripHashSuffix = Left(identifierName, markerPos - 1)
    Else
        StripHashSuffix = identifierName
    End If
End Function

Private Function IsHexString(ByVal s As String) As Boolean
    Dim i As Integer
    IsHexString = True
    For i = 1 To Len(s)
        Dim c As String: c = Mid(s, i, 1)
        If Not ((c >= "0" And c <= "9") Or (LCase(c) >= "a" And LCase(c) <= "f")) Then
            IsHexString = False
            Exit Function
        End If
    Next i
End Function

 ' ------------------------------------------------------------------------------
 ' [輔助工具]：ConvertParenSuffixToUnderscore
 ' 目的：偵測結尾是否為Windows自動編號格式「 (數字)」，是則轉換成「_數字」，
 '      讓後續同名合併分組邏輯能正確辨識。不符合此格式（如"最終版(草稿)"
 '      這種純文字括號）一律原樣傳回，只咬純數字，不誤判文字內容。
 ' ------------------------------------------------------------------------------
  Public Function ConvertParenSuffixToUnderscore(ByVal baseName As String) As String
      ConvertParenSuffixToUnderscore = baseName

      Dim isFullWidth As Boolean: isFullWidth = False
      If Right(baseName, 1) = "）" Then
          isFullWidth = True
      ElseIf Right(baseName, 1) <> ")" Then
          Exit Function
      End If
 
      Dim openParenPos As Long
      openParenPos = InStrRev(baseName, IIf(isFullWidth, " （", " ("))
      If openParenPos < 1 Then Exit Function

      Dim innerContent As String
      innerContent = Mid(baseName, openParenPos + 2, Len(baseName) - openParenPos - 2)

      If innerContent = "" Or Not IsNumeric(innerContent) Then Exit Function

      ConvertParenSuffixToUnderscore = Left(baseName, openParenPos - 1) & "_" & innerContent
  End Function
  
  ' 純字串工具：偵測結尾是否為「(數字)」或「（數字）」，是則直接丟棄整段（不轉換保留）。
' 跟既有的 ConvertParenSuffixToUnderscore 是兩支獨立函式，那支是「轉換」給
' Mod_FilenameNormalizer 這類通用工具使用，這支是「丟棄」給圖片管線內部使用，
' 兩者服務對象不同，不合併，避免改一支波及到另一個呼叫端
Public Function StripTrailingParenNumber(ByVal baseName As String) As String
    StripTrailingParenNumber = baseName
    Dim isFullWidth As Boolean: isFullWidth = False
    If Right(baseName, 1) = "）" Then
        isFullWidth = True
    ElseIf Right(baseName, 1) <> ")" Then
        Exit Function
    End If
    Dim openParenPos As Long
    openParenPos = InStrRev(baseName, IIf(isFullWidth, " （", " ("))
    If openParenPos < 1 Then Exit Function
    Dim innerContent As String
    innerContent = Mid(baseName, openParenPos + 2, Len(baseName) - openParenPos - 2)
    If innerContent = "" Or Not IsNumeric(innerContent) Then Exit Function
    StripTrailingParenNumber = Left(baseName, openParenPos - 1)
End Function

' 純字串工具：偵測結尾是否命中 patterns 清單裡任何一個複製樣式，是則丟棄該樣式本身。
' 單次檢查、單次動作，不迴圈——要交錯多輪清洗由呼叫端自行決定
Public Function StripKnownCopySuffix(ByVal baseName As String, ByVal patterns As Variant) As String
    StripKnownCopySuffix = baseName
    Dim p As Variant
    For Each p In patterns
        If Right(baseName, Len(CStr(p))) = CStr(p) Then
            StripKnownCopySuffix = Left(baseName, Len(baseName) - Len(CStr(p)))
            Exit Function
        End If
    Next p
End Function


' ------------------------------------------------------------------------------
' [輔助工具]：ExtractOriginalIdentifier
' 目的：從救援烙印名稱（格式：MANUAL_RESCUE_日期_旗幟識別名稱）拔除前綴，
'       還原成救援前的旗幟識別名稱。非烙印格式的輸入原樣傳回。
' ------------------------------------------------------------------------------
Public Function ExtractOriginalIdentifier(ByVal brandedName As String) As String
    Const BRAND_PREFIX As String = "MANUAL_RESCUE_"
    Const BRAND_LEN As Long = 14
    
    Dim current As String: current = brandedName
    Dim safetyCounter As Integer: safetyCounter = 0
    
    Do While Left(current, BRAND_LEN) = BRAND_PREFIX And safetyCounter < 20
        Dim result As String: result = Mid(current, BRAND_LEN + 1)
        If result = current Or result = "" Then Exit Do
        current = result
        safetyCounter = safetyCounter + 1
    Loop
    
    ExtractOriginalIdentifier = current
End Function

' ==============================================================================
' 內部輔助：智慧 HASH 引擎 (攔截巨型檔案)
' ==============================================================================
Public Function GetSmartHash(ByVal fso As Object, ByVal filePath As String, ByVal config As Object) As String
    Dim fileObj As Object: Set fileObj = fso.GetFile(filePath)
    If fileObj.Size > config.HugeFileLimit Then
        ' 巨型檔案閃避 MD5 耗時，改用替代雜湊值
        GetSmartHash = "HUGE_" & fileObj.Size & "_" & Format(fileObj.DateLastModified, "yyyymmddhhmmss")
    Else
        GetSmartHash = Mod_Utils.GetMD5(fso, filePath)
    End If
    Set fileObj = Nothing
End Function

Public Sub SetFolderHidden(ByVal fso As Object, ByVal folderPath As String)
    On Error Resume Next
    Dim fld As Object: Set fld = fso.GetFolder(folderPath)
    If Not fld Is Nothing Then
        If (fld.Attributes And 2) = 0 Then fld.Attributes = fld.Attributes Or 2 ' 2 = vbHidden
    End If
    On Error GoTo 0
End Sub

Public Function GetFolderDisplayName(ByVal folderPath As String) As String
    Dim trimmedPath As String: trimmedPath = folderPath
    If Right(trimmedPath, 1) = "\" Then trimmedPath = Left(trimmedPath, Len(trimmedPath) - 1)
    Dim lastSlash As Long: lastSlash = InStrRev(trimmedPath, "\")
    If lastSlash > 0 Then
        GetFolderDisplayName = Mid(trimmedPath, lastSlash + 1)
    Else
        GetFolderDisplayName = trimmedPath
    End If
End Function


Public Function BuildHashDictionary(ByVal fso As Object, ByVal colFiles As Collection, ByVal config As cls_Config) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Dim vFile As Variant
    For Each vFile In colFiles
        On Error Resume Next
        Dim h As String: h = Mod_Utils.GetSmartHash(fso, CStr(vFile), config)
        If Err.Number = 0 And h <> "" Then dict(CStr(vFile)) = h
        Err.Clear
        On Error GoTo 0
    Next vFile

    Set BuildHashDictionary = dict
End Function

' ------------------------------------------------------------------------------
' [輔助工具]：GetFreeDiskSpaceGB
' 目的：查詢指定路徑所在磁碟區的剩餘空間，換算成GB回傳。
' ------------------------------------------------------------------------------
Public Function GetFreeDiskSpaceGB(ByVal pathOrDrive As String) As Double
    On Error GoTo FailSafe
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim drv As Object: Set drv = fso.GetDrive(fso.GetDriveName(pathOrDrive))
    GetFreeDiskSpaceGB = drv.FreeSpace / (1024# ^ 3)
    Exit Function
FailSafe:
    ' 查詢失敗（例如網路磁碟機、權限問題），保守回傳-1，呼叫端應視為「無法確認」而非「安全」
    GetFreeDiskSpaceGB = -1
End Function


' ------------------------------------------------------------------------------
' [輔助工具]：GetAvailableMemoryMB
' 目的：查詢目前系統可用實體記憶體，換算成MB回傳。查詢失敗回傳-1。
' ------------------------------------------------------------------------------
Public Function GetAvailableMemoryMB() As Double
    On Error GoTo FailSafe
    Dim memStatus As MEMORYSTATUSEX
    memStatus.dwLength = LenB(memStatus)
    If GlobalMemoryStatusEx(memStatus) <> 0 Then
        GetAvailableMemoryMB = memStatus.ullAvailPhys / (1024# * 1024#)
    Else
        GetAvailableMemoryMB = -1
    End If
    Exit Function
FailSafe:
    GetAvailableMemoryMB = -1
End Function
