Option Explicit
' ==========================================================
' MODULE: Mod_UDF (標準模組)
' PURPOSE: 供 Excel 儲存格使用的自訂函數。即時安全偵測實體硬碟檔案是否存在。
' EXPORTS: IsFileAlive
' IMPORTS: 無
' FORBIDDEN: 嚴禁在此撰寫修改試算表或硬碟的邏輯。
' DEPENDENCIES: Windows API (Dir)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_UDF"
Private Const MODULE_VERSION As String = "1.0.0"

Public Function IsFileAlive(ByVal filePath As String) As Boolean
    ' 宣告為易失性函數，確保當試算表手動重新計算或工作表變更時，自動即時重新審查硬碟狀態
    Application.Volatile
    
    ' --------------------------------------------------------------------------
    ' [第一步：高頻呼叫防呆阻斷]
    ' --------------------------------------------------------------------------
    ' 預防性優化：Excel 批次刷新時可能傳入未填寫的空白儲存格路徑。
    ' 若不在此處攔截，原生 Dir("") 會頻繁引發底層 Run-time Error 52，嚴重拖慢畫面渲染速度。
    If Trim(filePath) = "" Then
        IsFileAlive = False
        Exit Function
    End If
    
    ' --------------------------------------------------------------------------
    ' [第二步：實體硬碟高速探測區]
    ' --------------------------------------------------------------------------
    On Error Resume Next
    
    ' 使用 Windows 原生 Dir 函數，指定 vbNormal 屬性精準排除資料夾，
    ' 以極低的記憶體成本與開銷完成探測，效能優於建立 FileSystemObject 物件數倍。
    Dim fileCheck As String
    fileCheck = Dir(filePath, vbNormal)
    
    If Err.Number <> 0 Then
        ' 若路徑內含不合法字元、網路磁碟機斷線或權限不足，底層會拋出異常，此處安全熔斷回傳 False
        IsFileAlive = False
        Err.Clear
    Else
        ' 字串長度大於 0 代表實體硬碟確實存在該檔案
        IsFileAlive = (Len(fileCheck) > 0)
    End If
    
    On Error GoTo 0
End Function
