Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_GDIPlusExport (標準模組)
' PURPOSE: Windows API 剪貼簿與 GDI+ 匯出引擎。
'          負責將剪貼簿的點陣圖無損轉存為 PNG 實體檔案，為救援管線提供底層硬體支援。
' 模組層級: Utils (純工具，底層 OS 呼叫，全面支援現代 Office 32/64 位元)
' EXPORTS: StartGdiplusEngine, SaveClipboardToPNG, StopGdiplusEngine
' IMPORTS: 無 (絕對孤立，純原生 Windows API 呼叫)
' FORBIDDEN: 嚴禁在此模組調用任何 Excel 儲存格或 UI 視圖，保持 100% 靜默與純粹。
 ' VERSION: 1.1.0 [Stability: Stable]
 ' ADR: 詳見 ACDS_ADR/ADR-V2-052.md — GDI+引擎啟停由「每張圖片各自啟停」改為「呼叫端迴圈外啟停一次」
' ==========================================================

Private Const MODULE_NAME As String = "Mod_GDIPlusExport"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
' 1. Windows API 與 GDI+ 核心宣告 (現代化 PtrSafe 宣告)
' ==============================================================================
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal uFormat As Long) As Long

' 64位元 C++ 結構對齊 (Struct Alignment) 記憶體崩潰修正鎖
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    #If Win64 Then
    Padding As Long ' 64位元環境下指標前必須補齊 4 Bytes 偏移量，否則 Excel 會直接閃退
    #End If
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, inputbuf As GdiplusStartupInput, ByVal outputbuf As LongPtr) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As LongPtr, ByVal hpal As LongPtr, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal bitmap As LongPtr, ByVal fileName As LongPtr, clsid As Any, ByVal encoderParams As LongPtr) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal bitmap As LongPtr) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpsz As LongPtr, pclsid As Any) As Long

' 剪貼簿點陣圖常數與 PNG 編碼器 CLSID
Private Const CF_BITMAP As Long = 2
Private Const PIX_ENC_PNG As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"


 ' ==============================================================================
 ' 2. GDI+ 引擎生命週期管理 (迴圈外呼叫，全程只啟停一次)
 ' ==============================================================================
 Public Function StartGdiplusEngine() As LongPtr
     Dim gdipInput As GdiplusStartupInput
     Dim gdipToken As LongPtr
     Dim gdiStatus As Long
     gdipInput.GdiplusVersion = 1
     gdiStatus = GdiplusStartup(gdipToken, gdipInput, ByVal 0&)
     If gdiStatus = 0 Then
         StartGdiplusEngine = gdipToken
     Else
         StartGdiplusEngine = 0
     End If
 End Function

 Public Sub StopGdiplusEngine(ByVal gdipToken As LongPtr)
     If gdipToken <> 0 Then Call GdiplusShutdown(gdipToken)
 End Sub

 ' ==============================================================================
 ' 3. 剪貼簿提取與檔案匯出主程序
 ' ==============================================================================
 Public Function SaveClipboardToPNG(ByVal targetPath As String, ByVal gdipToken As LongPtr) As Boolean
    Dim tGuid(0 To 15) As Byte
    Dim gdiStatus As Long
    Dim retryCount As Integer
    Dim isClipboardOpen As Boolean
    
    Dim hBitmap As LongPtr
    Dim gdipBitmap As LongPtr
    
    SaveClipboardToPNG = False
    If Trim(targetPath) = "" Then Exit Function
    
    On Error GoTo ErrorHandler
    
      If gdipToken = 0 Then
         Err.Raise 9961, MODULE_NAME & ".SaveClipboardToPNG", _
             "[系統報警] GDI+ 引擎 token 無效，呼叫端須先呼叫 StartGdiplusEngine 並確認回傳非零"
     End If

    
    ' --------------------------------------------------------------------------
    ' [第一關]：檢查剪貼簿狀態與佔用重試機制
    ' --------------------------------------------------------------------------
    If IsClipboardFormatAvailable(CF_BITMAP) = 0 Then Exit Function
    
    isClipboardOpen = False
    For retryCount = 1 To 3
        If OpenClipboard(ByVal 0&) <> 0 Then
            isClipboardOpen = True
            Exit For
        End If
        DoEvents ' 允許 OS 釋放鎖，在極短迴圈內安全自旋
    Next retryCount
    
    If Not isClipboardOpen Then Exit Function
    
    hBitmap = GetClipboardData(CF_BITMAP)
    If hBitmap = 0 Then GoTo Cleanup
    
    ' --------------------------------------------------------------------------
    ' [第二關]：建立點陣圖與寫入實體檔案（引擎已由呼叫端啟動，此處直接使用傳入的token）
    ' --------------------------------------------------------------------------
    gdiStatus = GdipCreateBitmapFromHBITMAP(hBitmap, ByVal 0&, gdipBitmap)
    If gdiStatus <> 0 Then
        Err.Raise 9962, MODULE_NAME & ".SaveClipboardToPNG", _
            "[系統報警] 點陣圖實例化失敗，無法從剪貼簿控制代碼建立圖像物件，狀態碼: " & gdiStatus
    End If
    
    ' 【覆寫防禦】：放棄危險的 Dir 探測，直接執行原子級 Kill 抹除阻礙
    On Error Resume Next
    Kill targetPath
    On Error GoTo ErrorHandler
    
    ' 解析 PNG 編碼器 GUID
    Dim clsidStatus As Long
    clsidStatus = CLSIDFromString(StrPtr(PIX_ENC_PNG), tGuid(0))
    
    ' 【漏洞修復】：嚴禁無聲忽略 CLSID 錯誤，失敗必須觸發硬熔斷
    If clsidStatus <> 0 Then
        Err.Raise 9965, MODULE_NAME & ".SaveClipboardToPNG", _
            "[系統報警] OLE 系統元件錯誤：無法解析 PNG 編碼器 CLSID，狀態碼: " & clsidStatus
    End If
    
    ' 正式將圖片數據寫入硬碟
    gdiStatus = GdipSaveImageToFile(gdipBitmap, StrPtr(targetPath), tGuid(0), ByVal 0&)
    If gdiStatus = 0 Then
        SaveClipboardToPNG = True
    Else
        Err.Raise 9963, MODULE_NAME & ".SaveClipboardToPNG", _
            "[系統報警] GDI+ 物理寫入硬碟失敗 (可能是磁碟唯讀或無寫入權限)，狀態碼: " & gdiStatus & " | 靶機路徑: " & targetPath
    End If

' ------------------------------------------------------------------------------
' 3. 記憶體釋放與異常捕捉安全閉環 (Deterministic Reclamation)
' ------------------------------------------------------------------------------
Cleanup:
    On Error Resume Next
    ' 強制釋放：無論如何，控制代碼與 OLE 記憶體令牌必須銷毀
    If gdipBitmap <> 0 Then Call GdipDisposeImage(gdipBitmap)
    If isClipboardOpen Then Call CloseClipboard
    On Error GoTo 0
    Exit Function

ErrorHandler:
    Dim savedErrNum As Long: savedErrNum = Err.Number
    Dim savedErrDesc As String: savedErrDesc = Err.Description
    Err.Clear
    
    ' 崩潰時強制就地釋放 Windows 核心資源
    On Error Resume Next
    If gdipBitmap <> 0 Then Call GdipDisposeImage(gdipBitmap)
    If isClipboardOpen Then Call CloseClipboard
    On Error GoTo 0
    
    ' 向上拋出結構化異常，交給上層模組與管線捕獲
    If savedErrNum >= 9961 And savedErrNum <= 9965 Then
        Err.Raise savedErrNum, MODULE_NAME & ".SaveClipboardToPNG", savedErrDesc
    Else
        Err.Raise 9964, MODULE_NAME & ".SaveClipboardToPNG", _
            "[系統報警] 剪貼簿與 GDI+ 核心發生未預期之記憶體崩潰: " & savedErrDesc
    End If
End Function
