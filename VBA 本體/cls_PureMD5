' ==========================================================
' MODULE: cls_PureMD5 (類別模組)
' PURPOSE: 純 VBA 實作的 MD5 密碼學雜湊運算引擎，取代原本透過 CreateObject(
'          "System.Security.Cryptography.MD5CryptoServiceProvider") 呼叫
'          .NET COM Interop 元件的做法。緣起 V18 Automation 錯誤
'          -2146232576（0x80131700，.NET CLR／mscorlib 保留 HRESULT 區段，
'          代表該台機器 .NET Runtime 初始化 MD5CryptoServiceProvider 失敗，
'          常見成因為 Windows 以 CNG 取代舊版 CAPI 後端造成的相容性問題）。
'          本類別刻意提供與 .NET MD5CryptoServiceProvider 相同的方法簽章
'          （ComputeHash_2），故 Mod_Crypto.bas 呼叫端程式碼無須修改，僅
'          Mod_WorkflowManager.bas 組裝 sysTokens 時改注入 New cls_PureMD5
'          取代 CreateObject(PROGID_MD5)（RunPhase1／RunPhase2 各一處）。
' EXPORTS: ComputeHash_2
' IMPORTS: 無
' FORBIDDEN: 嚴禁在本類別內以任何形式呼叫 CreateObject 取得外部 COM／.NET
'            元件；本類別存在的唯一目的就是消除這類環境依賴，任何修改都
'            不得違背此設計初衷。
' DEPENDENCIES: ACDS_ContractRegistry.md
' AUTHOR: 鄭詩樺 (Cheng, Shih-Hua)
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================
' 【正確性驗證說明】演算法邏輯已於 Python 環境以相同的 32 位元無號運算規則
' （AddUnsigned／RotateLeft／RShift）逐步模擬，並對空字串、"a"／"abc"／標準
' 測試字串、跨越 56/64 bytes 填補邊界的長度（0, 1, 55, 56, 57, 63, 64, 65,
' 119, 120, 121）、以及 100,000 bytes 與 1,000,000 bytes 隨機二進位資料，
' 逐一比對 Python 內建 hashlib.md5，結果完全一致，才轉譯為本檔案的 VBA 版本。
' 此驗證僅涵蓋演算法邏輯本身，VBA 語法層面的實機編譯測試仍需你自行執行。
' ==========================================================
Option Explicit

' 說明：VBA 不支援「Const 陣列」語法（Const K(0 To 63) As Long = [...] 無法編譯），
' 因此改用模組層級 Variant 陣列，於 Class_Initialize 以 Array() 一次性建立。
' K/S 元素雖存成 Variant，但內含值仍是正確的 Long／Byte 數值，後續在運算式中
' 使用時 VBA 會自動轉型，不影響任何計算結果。
Private K As Variant   ' MD5 標準規範的 64 個常數項（RFC 1321）
Private S As Variant   ' MD5 標準規範的 64 個循環位移量（RFC 1321）

Private Sub Class_Initialize()
    K = Array( _
        &HD76AA478, &HE8C7B756, &H242070DB, &HC1BDCEEE, _
        &HF57C0FAF, &H4787C62A, &HA8304613, &HFD469501, _
        &H698098D8, &H8B44F7AF, &HFFFF5BB1, &H895CD7BE, _
        &H6B901122, &HFD987193, &HA679438E, &H49B40821, _
        &HF61E2562, &HC040B340, &H265E5A51, &HE9B6C7AA, _
        &HD62F105D, &H2441453, &HD8A1E681, &HE7D3FBC8, _
        &H21E1CDE6, &HC33707D6, &HF4D50D87, &H455A14ED, _
        &HA9E3E905, &HFCEFA3F8, &H676F02D9, &H8D2A4C8A, _
        &HFFFA3942, &H8771F681, &H6D9D6122, &HFDE5380C, _
        &HA4BEEA44, &H4BDECFA9, &HF6BB4B60, &HBEBFBC70, _
        &H289B7EC6, &HEAA127FA, &HD4EF3085, &H4881D05, _
        &HD9D4D039, &HE6DB99E5, &H1FA27CF8, &HC4AC5665, _
        &HF4292244, &H432AFF97, &HAB9423A7, &HFC93A039, _
        &H655B59C3, &H8F0CCC92, &HFFEFF47D, &H85845DD1, _
        &H6FA87E4F, &HFE2CE6E0, &HA3014314, &H4E0811A1, _
        &HF7537E82, &HBD3AF235, &H2AD7D2BB, &HEB86D391)

    S = Array( _
        7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, _
        5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, _
        4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, _
        6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21)
End Sub

' =========================================================================
' 函數名稱：ComputeHash_2
' 功用說明：與 .NET MD5CryptoServiceProvider.ComputeHash_2 介面相容，輸入任意
'           長度的位元組陣列，回傳固定 16 bytes 的 MD5 雜湊結果陣列。
' =========================================================================
Public Function ComputeHash_2(ByRef bytArr() As Byte) As Byte()
    Dim lMsgLen     As Long
    Dim bytPadded() As Byte
    Dim lPaddedLen  As Long
    Dim lBlockCount As Long
    Dim lBlockIdx   As Long

    Dim lA0 As Long, lB0 As Long, lC0 As Long, lD0 As Long
    Dim lA As Long, lB As Long, lC As Long, lD As Long
    Dim lF As Long
    Dim lG As Long
    Dim M(0 To 15) As Long
    Dim i As Long, j As Long

    ' --- 1. 計算原始長度（若陣列未初始化，視為 0 bytes）---
    On Error Resume Next
    lMsgLen = UBound(bytArr) - LBound(bytArr) + 1
    If Err.Number <> 0 Then lMsgLen = 0
    On Error GoTo 0

    ' --- 2. 填補（Padding）---
    lPaddedLen = lMsgLen + 1
    Do While (lPaddedLen Mod 64) <> 56
        lPaddedLen = lPaddedLen + 1
    Loop
    lPaddedLen = lPaddedLen + 8

    ReDim bytPadded(0 To lPaddedLen - 1)
    If lMsgLen > 0 Then
        Dim lSrcBase As Long
        lSrcBase = LBound(bytArr)
        For i = 0 To lMsgLen - 1
            bytPadded(i) = bytArr(lSrcBase + i)
        Next i
    End If
    bytPadded(lMsgLen) = &H80
    ' 其餘填補位元組保持 ReDim 預設值 0，不需另行清零

    ' --- 3. 寫入原始訊息長度（bit 數），以 little-endian 64 位元寫入最後 8 bytes ---
    Dim dblBitLen As Double, lLenHigh As Long, lLenLow As Long, dblLow As Double
    dblBitLen = CDbl(lMsgLen) * 8
    lLenHigh = CLng(Int(dblBitLen / 4294967296#))
    dblLow = dblBitLen - (CDbl(lLenHigh) * 4294967296#)
    If dblLow >= 2147483648# Then
        lLenLow = CLng(dblLow - 4294967296#)
    Else
        lLenLow = CLng(dblLow)
    End If

    Call WriteLongLE(bytPadded, lPaddedLen - 8, lLenLow)
    Call WriteLongLE(bytPadded, lPaddedLen - 4, lLenHigh)

    ' --- 4. 初始向量 ---
    lA0 = &H67452301
    lB0 = CLng(&HEFCDAB89)
    lC0 = CLng(&H98BADCFE)
    lD0 = CLng(&H10325476)

    ' --- 5. 逐 64 bytes 區塊運算 ---
    lBlockCount = lPaddedLen \ 64
    For lBlockIdx = 0 To lBlockCount - 1
        Dim lBase As Long
        lBase = lBlockIdx * 64

        For j = 0 To 15
            M(j) = ReadLongLE(bytPadded, lBase + j * 4)
        Next j

        lA = lA0: lB = lB0: lC = lC0: lD = lD0

        For i = 0 To 63
            If i < 16 Then
                lF = (lB And lC) Or ((Not lB) And lD)
                lG = i
            ElseIf i < 32 Then
                lF = (lD And lB) Or ((Not lD) And lC)
                lG = (5 * i + 1) Mod 16
            ElseIf i < 48 Then
                lF = lB Xor lC Xor lD
                lG = (3 * i + 5) Mod 16
            Else
                lF = lC Xor (lB Or (Not lD))
                lG = (7 * i) Mod 16
            End If

            lF = AddUnsigned(lF, lA)
            lF = AddUnsigned(lF, K(i))
            lF = AddUnsigned(lF, M(lG))

            lA = lD
            lD = lC
            lC = lB
            lB = AddUnsigned(lB, RotateLeft(lF, S(i)))
        Next i

        lA0 = AddUnsigned(lA0, lA)
        lB0 = AddUnsigned(lB0, lB)
        lC0 = AddUnsigned(lC0, lC)
        lD0 = AddUnsigned(lD0, lD)
    Next lBlockIdx

    ' --- 6. 組裝輸出（16 bytes，A/B/C/D 各 4 bytes，little-endian）---
    Dim bytOut(0 To 15) As Byte
    Call WriteLongLE(bytOut, 0, lA0)
    Call WriteLongLE(bytOut, 4, lB0)
    Call WriteLongLE(bytOut, 8, lC0)
    Call WriteLongLE(bytOut, 12, lD0)

    ComputeHash_2 = bytOut
End Function

' =========================================================================
' 內部工具：從 Byte 陣列指定位移處讀取 4 bytes，組成 little-endian 32 位元 Long
' （以 &H80000000 符號位安全組裝，避免 VBA Overflow）
' =========================================================================
Private Function ReadLongLE(ByRef bytArr() As Byte, ByVal lOffset As Long) As Long
    Dim b0 As Long, b1 As Long, b2 As Long, b3 As Long
    b0 = bytArr(lOffset)
    b1 = bytArr(lOffset + 1)
    b2 = bytArr(lOffset + 2)
    b3 = bytArr(lOffset + 3)

    If b3 < 128 Then
        ReadLongLE = b0 + (b1 * 256) + (b2 * 65536) + (b3 * 16777216)
    Else
        ReadLongLE = (b0 + (b1 * 256) + (b2 * 65536) + ((b3 - 128) * 16777216)) Or &H80000000
    End If
End Function

' =========================================================================
' 內部工具：把一個 32 位元 Long（可能為負數的兩補數表示）拆成 4 bytes，
' 以 little-endian 順序寫入 Byte 陣列指定位移處
' =========================================================================
Private Sub WriteLongLE(ByRef bytArr() As Byte, ByVal lOffset As Long, ByVal lValue As Long)
    bytArr(lOffset) = CByte(lValue And 255)
    bytArr(lOffset + 1) = CByte(RShift(lValue, 8) And 255)
    bytArr(lOffset + 2) = CByte(RShift(lValue, 16) And 255)
    bytArr(lOffset + 3) = CByte(RShift(lValue, 24) And 255)
End Sub

' =========================================================================
' 內部工具：32 位元無號加法（避免 VBA Long 兩補數溢位時拋出 Overflow 例外）
' 經典進位安全技巧：拆成最高位、次高位與剩餘 30 位元分開運算後再組合。
' =========================================================================
Private Function AddUnsigned(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lX8 As Long, lY8 As Long, lX4 As Long, lY4 As Long, lResult As Long

    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000

    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

    If (lX4 And lY4) <> 0 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf (lX4 Or lY4) <> 0 Then
        If (lResult And &H40000000) <> 0 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If

    AddUnsigned = lResult
End Function

' =========================================================================
' 內部工具：32 位元邏輯右移（Logical Right Shift，非算術右移），n 範圍 0~31。
' 將 lValue 拆成符號位與剩餘 31 位元分開位移後再組合，避免正負號造成誤差。
' =========================================================================
Private Function RShift(ByVal lValue As Long, ByVal n As Byte) As Long
    Dim lSign As Long, lRest As Long

    If n = 0 Then
        RShift = lValue
        Exit Function
    End If

    If (lValue And &H80000000) <> 0 Then
        lSign = 1
    Else
        lSign = 0
    End If
    lRest = lValue And &H7FFFFFFF

    If n = 31 Then
        RShift = lSign
    Else
        RShift = (lSign * CLng(2 ^ (31 - n))) + (lRest \ CLng(2 ^ n))
    End If
End Function

' =========================================================================
' 內部工具：32 位元循環左移（Rotate Left），n 範圍 1~31。
' 原理：高 n 位元邏輯右移 (32-n) 搬到最低位；低 (32-n) 位元左移 n 搬到最高位，
' 左移過程以「先清除會溢位的那一個位元、乘完再補回符號位」的方式避免 Overflow。
' =========================================================================
Private Function RotateLeft(ByVal lValue As Long, ByVal n As Byte) As Long
    If n = 0 Then
        RotateLeft = lValue
        Exit Function
    End If

    Dim lShiftedHigh As Long
    Dim lowerMask As Long, lowerBits As Long, topBitMask As Long
    Dim lShiftedLow As Long

    lShiftedHigh = RShift(lValue, 32 - n)

    lowerMask = CLng(2 ^ (32 - n) - 1)
    lowerBits = lValue And lowerMask

    If (32 - n - 1) >= 0 Then
        topBitMask = CLng(2 ^ (32 - n - 1))
    Else
        topBitMask = 0
    End If

    If topBitMask <> 0 And (lowerBits And topBitMask) <> 0 Then
        lShiftedLow = (lowerBits Xor topBitMask) * CLng(2 ^ n)
        lShiftedLow = lShiftedLow Or &H80000000
    Else
        lShiftedLow = lowerBits * CLng(2 ^ n)
    End If

    RotateLeft = lShiftedLow Or lShiftedHigh
End Function

