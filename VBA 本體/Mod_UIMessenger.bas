Option Explicit
Option Private Module

' ==========================================================
' MODULE: Mod_UIMessenger (標準模組)
' PURPOSE: 全系統唯一具備彈出視窗 (MsgBox) 權限的模組，供全系統呼叫。
' 模組層級: View (前端介面)
' EXPORTS: AskQuestion, AskQuestionCancel, AskCritical, ShowInfo, ShowWarning, ShowError
' IMPORTS: 無
' FORBIDDEN: 嚴禁在此模組撰寫任何業務邏輯判斷或直接操作檔案/儲存格，僅負責包裝標準 MsgBox 呼叫。
' DEPENDENCIES: 無
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "Mod_UIMessenger"
Private Const MODULE_VERSION As String = "1.0.0"

' 【防禦修正】：回歸標準二元邏輯，防止在 cmd_Rescue 或 Rename 管線中，
' 取消 (Cancel) 按鈕錯誤落入 vbYes 或 vbNo 的盲區而導致錯誤刪除。
Public Function AskQuestion(ByVal promptText As String, ByVal titleText As String) As VbMsgBoxResult
    AskQuestion = MsgBox(promptText, vbQuestion + vbYesNo + vbDefaultButton2, titleText)
End Function

' 【新增】：專供需要三元判斷 (包含取消) 的特殊路由使用 (例如環境防禦斷層彈窗)
Public Function AskQuestionCancel(ByVal promptText As String, ByVal titleText As String) As VbMsgBoxResult
    AskQuestionCancel = MsgBox(promptText, vbQuestion + vbYesNoCancel + vbDefaultButton1, titleText)
End Function

Public Function AskCritical(ByVal promptText As String, ByVal titleText As String) As VbMsgBoxResult
    AskCritical = MsgBox(promptText, vbCritical + vbYesNo + vbDefaultButton2, titleText)
End Function

Public Sub ShowInfo(ByVal promptText As String, ByVal titleText As String)
    MsgBox promptText, vbInformation, titleText
End Sub

Public Sub ShowWarning(ByVal promptText As String, ByVal titleText As String)
    MsgBox promptText, vbExclamation, titleText
End Sub

Public Sub ShowError(ByVal promptText As String, ByVal titleText As String)
    MsgBox promptText, vbCritical, titleText
End Sub
