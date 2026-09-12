Option Explicit
' ==========================================================
' MODULE: ThisWorkbook (Excel 物件: ThisWorkbook)
' PURPOSE: 活頁簿生命週期事件監聽器 (Trigger 層)。
' EXPORTS: Workbook_Open, Workbook_BeforeClose
' IMPORTS: Mod_Main
' FORBIDDEN: 絕對禁止在此撰寫任何業務邏輯，僅限觸發路由。
' DEPENDENCIES: 綁定於活頁簿本身
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "ThisWorkbook"
Private Const MODULE_VERSION As String = "1.0.0"

Private Sub Workbook_Open()
    ' 檔案開啟的第一秒，立刻將指揮權移交給 Controller (Mod_Main)
    Call Mod_Main.SYS_系統開機引導
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 關閉檔案前，強制還原 Excel 應用程式環境變數，避免因非預期崩潰導致的介面鎖死
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    On Error GoTo 0
End Sub



' ==========================================================
' AUTHOR: 鄭詩樺 (Cheng, Shih-Hua)
' LICENSE - CODE (本活頁簿內全部 VBA 原始碼適用)：MIT License
' Copyright (c) 2026 鄭詩樺 (Cheng, Shih-Hua)
' ==========================================================
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.
' ==========================================================
' LICENSE - DOCUMENTATION（隨附之使用說明書、規格書等文件適用）：
' CC BY-NC-SA 4.0（姓名標示-非商業性-相同方式分享）
'   - 姓名標示（BY）：轉載或改作須標明原作者
'   - 非商業性（NC）：不得作為商業用途
'   - 相同方式分享（SA）：改作後之衍生文件須採用相同授權條款釋出
'   完整條款：https://creativecommons.org/licenses/by-nc-sa/4.0/deed.zh-hant
' ==========================================================
