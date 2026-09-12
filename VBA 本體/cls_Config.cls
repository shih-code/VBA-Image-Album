Option Explicit

' ==========================================================
' MODULE: cls_Config (類別模組)
' PURPOSE: 集中所有靜態設定，作為全系統唯一的單一事實來源 (SSOT)，全唯讀保護。
' EXPORTS: 各種路徑常數、版面設定、顏色常數的唯讀屬性 (Getters)
' IMPORTS: 無 (底層基礎物件)
' FORBIDDEN: 嚴禁提供 Let/Set 屬性，確保運行期間靜態參數絕對不可被篡改。
' DEPENDENCIES: 無
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Private Const MODULE_NAME As String = "cls_Config"
Private Const MODULE_VERSION As String = "1.0.0"

' ==============================================================================
' 【底層私有變數 (全面封裝靜態常數)】
' ==============================================================================
' 1. 系統層級名稱與特徵字串
Private m_SheetNameManual As String
Private m_SheetNameAllPics As String
Private m_PrefixCustom As String
Private m_PrefixOldCatalog As String
Private m_SheetSysMeta As String
Private m_SheetSysLog As String
Private m_SheetSysTempSettings As String
Private m_SheetNameTempSafeRoom As String
Private m_CopySuffixPatterns As Variant

' 2. 系統門面標題與版權
Private m_AppTitleControlPanel As String
Private m_AppTitleManual As String
Private m_CopyrightText As String

' 3. 下拉選單選項常數
Private m_OptDisplayMode As String
Private m_OptCatalogMode As String
Private m_OptNamingMode As String
Private m_OptGridSize As String
Private m_OptExportFilter As String
Private m_OptFormatBehavior As String

' 4. UI 字體與版面級距
Private m_FontMain As String
Private m_FontSizeTitle As Double
Private m_FontSizeSub As Double
Private m_FontSizeBase As Double
Private m_FontSizeSmall As Double

' 5. 系統全域色彩與幾何 (莫蘭迪色系)
Private m_ThemeIO As Long
Private m_ThemeExp As Long
Private m_ThemeManual As Long
Private m_DrawerCustomBg As Long
Private m_DrawerImageBg As Long
Private m_DrawerItemBg As Long
Private m_DrawerItemBorder As Long
Private m_DrawerItemText As Long
Private m_DrawerLinkText As Long

Private m_ColorWhite As Long
Private m_ColorTitle As Long
Private m_ColorTextMain As Long
Private m_ColorTextHint As Long
Private m_ColorTextWarning As Long
Private m_ColorTextSuccess As Long
Private m_ColorPanelBg As Long
Private m_ColorShadow As Long
Private m_InputBg As Long
Private m_ReadonlyBg As Long
Private m_Border As Long
Private m_ColorLinkText As Long

Private m_GapX As Double
Private m_GapY As Double
Private m_QuarterW As Double

' 6. 作業表實體網格座標 (SSOT 物理錨點)
Private m_RangeSettingsConsole As String
Private m_CellProjName As String
Private m_CellDangerConsole As String
Private m_CellAutoLayoutStateConsole As String

Private m_CellImgWidth As String
Private m_CellImgSpacing As String
Private m_CellNamingMode As String
Private m_CellSheetPrefix As String
Private m_CellPdfImgName As String
Private m_CellPdfCatName As String
Private m_CellPdfGrid As String
Private m_CellExportFilter As String
Private m_CellExportRange As String

' 7. UI 控件 ID
Private m_btnLockUI As String
Private m_btnLockDir As String
Private m_BtnAutoLayout As String
Private m_btnGlobalLock As String
Private m_BtnRebuildIndex As String
Private m_BtnSystemRestart As String
Private m_BtnFormat As String
Private m_BtnToggleLog As String
Private m_BtnLayoutLock As String

' 8. 引擎與旗幟座標
Private m_MaxDagLoopCountDAGEngine As Long
Private m_CellDebugStateFlagSet As String
Private m_CellDisplayModeFlagSet As String
Private m_CellCatalogModeFlagSet As String

Private m_MinImgWidth As Double
Private m_MaxImgWidth As Double
Private m_MinImgSpacing As Double
Private m_MaxImgSpacing As Double

' 9. 核心物件與環境路徑 (平級三柱拓撲)
Private m_TargetWB As Workbook
Private m_BaseFolder As String

' --- 獨立匯出體系 ---
Private m_IOExportFolder As String
Private m_IOExportCatalogFolder As String
Private m_IOExportImageFolder As String

' --- 第一柱：圖片體系 ---
Private m_IOTargetFolder As String
Private m_IOOrganizeFolder As String
Private m_IORescueFolder As String
Private m_IORescueDowncastFolder As String
Private m_IORescueDiscardFolder As String
Private m_IODuplicateDiscardFolder As String

Private m_IOThumbnailCacheFolder As String
Private m_ThumbnailCacheDimension As Double
Private m_MaxSafetyPoint As Double     '應該是匯出最大?
Private m_RescueEmbedWidth As Double  ' 【V2-062新增】救援夾圖片統一嵌入寬度，cmd_Rescue與cmd_Rebuild共用同一份SSOT
Private m_RescueEmbedGap As Double    ' 【V2-063新增】救援夾圖片垂直堆疊間距，修正V2-062遺漏（寬度收了、間距漏了）
Private m_MinFreeDiskSpaceGB As Double  ' 【新增】磁碟剩餘空間安全下限(GB)，低於此值中止批次作業
Private m_MinAvailableMemoryMB As Double
Private m_HugeFileLimit As Double

Private m_IOSandboxFolder As String


' --- 第二柱：歸檔附錄體系 ---
Private m_IOAttachmentFolder As String
Private m_IOAttachPDFFolder As String
Private m_IOAttachTXTFolder As String   ' [V4.0修正] TXT與MD專區
Private m_IOAttachWordFolder As String  ' Word 專區
Private m_IOAttachExcelFolder As String ' Excel (含 CSV) 專區
Private m_IOAttachPPTFolder As String   ' PPT 專區
Private m_IOAttachZipFolder As String   ' 壓縮封存專區
Private m_IOAttachMediaFolder As String
Private m_IOLooseImportFolder As String
Private m_IOAttachmentDuplicateDiscardFolder As String  ' 【V2-046新增】附件線雜湊衝突軟刪除區

' --- 第三柱：雜物垃圾體系 ---
Private m_IOQuarantineFolder As String


' --- 第四柱：封存備份體系 ---
Private m_IOArchiveFolder As String
Private m_IOBackupSnapshotFolder As String
Private m_IOLogArchiveFolder As String

' 10. 系統原廠預設參數
Private m_UIProjectName As String
Private m_ExpectedRootFolderName As String
Private m_SheetNameUI As String
Private m_SheetNameCatalog As String
Private m_SheetNameAttachment As String
Private m_SheetNameRescue As String
Private m_SheetNameQuarantine As String
Private m_SecurityPassword As String

Private m_UIImgWidth As Double
Private m_UIImgSpacing As Double
Private m_NamingMode As String
Private m_UISheetPrefix As String
Private m_UIDisplayMode As String
Private m_CatalogMode As String
Private m_FormatBehavior As String
Private m_ExportFilterMode As String
Private m_ExportCustomRange As String

' 11. 動態狀態控制色
Private m_UIColorPassGreen As Long
Private m_UIColorLockRed As Long
Private m_UIColorNormal As Long

' 12. PDF 報告輸出基礎設定
Private m_PDFImageReportName As String
Private m_PDFCatalogReportName As String
Private m_PDFGridLayout As Integer
Private m_PDFMarginCM As Double
Private m_PDFTitleFontSize As Integer
Private m_PDFHeaderHeight1 As Double
Private m_PDFHeaderHeight2 As Double
Private m_PDFTextHeight As Double
Private m_PDFImgHeight As Double
Private m_PDFGapHeight As Double
Private m_PDFNumCols As Integer
Private m_PDFColWidth As Double
Private m_PDFColGap As Double
Private m_PDFBaseFontSize As Integer

' ==============================================================================
' --- 1. 初始化流程 ---
' ==============================================================================

Private Sub Class_Initialize()
    LoadHardcodedDefaults
    BuildPathTopology
End Sub

' ==============================================================================
' --- 2. 內部靜態賦值程序 ---
' ==============================================================================

Private Sub LoadHardcodedDefaults()
    ' [系統層級名稱與特徵字串]
    m_SheetNameManual = "系統說明書"
    m_SheetNameAllPics = "歸檔_全部圖片"
    m_PrefixCustom = "###"
    m_PrefixOldCatalog = "舊目錄_"
    m_SheetSysMeta = "SYS_METADATA"
    m_SheetSysLog = "LOG_SYSTEM"
    m_SheetSysTempSettings = "TEMP_SETTINGS_PROTECTED"
    m_SheetNameTempSafeRoom = "TEMP_SAFE_DELETE_ROOM"
    m_CopySuffixPatterns = Array(" - 複製")
    
    ' [系統門面標題與版權]
    m_AppTitleControlPanel = "IMAGE ALBUM | Control Panel（控制台）"
    m_AppTitleManual = "IMAGE ALBUM | User Guide（使用手冊）"
    m_CopyrightText = "Code: MIT License | Documentation: CC BY-NC-SA 4.0｜Copyright (C) 2026 鄭詩樺｜完整條款見 LICENSE"
    
    ' [下拉選單選項常數]
    m_OptDisplayMode = "單一工作表,分開工作表,同名合併工作表,依資料夾合併"
    m_OptCatalogMode = "覆蓋,新增"
    m_OptNamingMode = "今日日期,修改日期,自訂,原始名稱,自訂~今日日期,自訂~修改日期"
    m_OptGridSize = "2,4,6,8"
    m_OptExportFilter = "全額匯出,白名單,黑名單"
    m_OptFormatBehavior = "僅重置工作表,保留備份"
    
    ' [UI 字體與版面級距]
    m_FontMain = "微軟正黑體"
    m_FontSizeTitle = 16
    m_FontSizeSub = 14
    m_FontSizeBase = 10
    m_FontSizeSmall = 8.5
    
    ' [UI 控件 ID]
    m_btnLockUI = "Btn_LockUI"
    m_btnLockDir = "Btn_LockDir"
    m_BtnAutoLayout = "Btn_AutoLayout"
    m_btnGlobalLock = "Btn_GlobalLock"
    m_BtnRebuildIndex = "Btn_RebuildIndex"
    m_BtnSystemRestart = "Btn_SystemRestart"
    m_BtnFormat = "Btn_Format"
    m_BtnToggleLog = "Btn_ToggleLog"
    m_BtnLayoutLock = "Btn_LayoutLock"
    
    ' [全域色彩與幾何 (莫蘭迪色系)]
    m_ThemeIO = RGB(65, 95, 115)
    m_ThemeExp = RGB(90, 120, 130)
    m_ThemeManual = RGB(155, 95, 95)
    m_DrawerCustomBg = RGB(110, 135, 155)
    m_DrawerImageBg = RGB(100, 140, 125)
    m_DrawerItemBg = RGB(250, 252, 253)
    m_DrawerItemBorder = RGB(225, 235, 240)
    m_DrawerItemText = RGB(50, 60, 70)
    m_DrawerLinkText = RGB(50, 100, 150)
    
    m_ColorWhite = RGB(255, 255, 255)
    m_ColorTitle = RGB(45, 65, 85)
    m_ColorTextMain = RGB(50, 60, 70)
    m_ColorTextHint = RGB(150, 160, 170)
    m_ColorTextWarning = RGB(180, 80, 80)
    m_ColorTextSuccess = RGB(60, 130, 90)
    m_ColorLinkText = RGB(50, 100, 150)
    
    m_ColorPanelBg = RGB(246, 248, 250)
    m_ColorShadow = RGB(170, 180, 190)
    m_InputBg = RGB(255, 255, 248)
    m_ReadonlyBg = RGB(248, 250, 252)
    m_Border = RGB(215, 225, 235)
    
    m_UIColorPassGreen = RGB(155, 180, 160)
    m_UIColorLockRed = RGB(215, 145, 145)
    m_UIColorNormal = RGB(80, 120, 170)
    
    m_GapX = 9
    m_GapY = 8
    m_QuarterW = 100
    
    ' [作業表網格與旗幟座標]
    m_RangeSettingsConsole = "B3:B22"
    m_CellProjName = "B3"
    m_CellDangerConsole = "B22"
    m_CellAutoLayoutStateConsole = "Z1"
    
    
    m_CellDebugStateFlagSet = "Z2"
    m_CellDisplayModeFlagSet = "B8"
    m_CellCatalogModeFlagSet = "B9"
    
    
    m_CellImgWidth = "B6"
    m_CellImgSpacing = "B7"
    m_CellNamingMode = "B12"
    m_CellSheetPrefix = "B13"
    m_CellPdfImgName = "B14"
    m_CellPdfCatName = "B15"
    m_CellPdfGrid = "B16"
    m_CellExportFilter = "B17"
    m_CellExportRange = "B18"
    
    m_MaxDagLoopCountDAGEngine = 50
    
    m_MinImgWidth = 50
    m_MaxImgWidth = 1333
    m_MinImgSpacing = 0
    m_MaxImgSpacing = 1333
    
    ' [底層原生物件預設值]
    Set m_TargetWB = ThisWorkbook
    m_SheetNameUI = "操作介面"
    m_SheetNameCatalog = "圖片目錄"
    m_SheetNameAttachment = "附件目錄"
    m_SheetNameRescue = "尚未救援"
    m_SheetNameQuarantine = "隔離與雜物清單"
    m_SecurityPassword = ""
    m_UISheetPrefix = "IMG"

    m_UIProjectName = "我的整理資料"
    m_ExpectedRootFolderName = "我的整理資料"
    m_UIImgWidth = 300
    m_UIImgSpacing = 15
    m_NamingMode = "自訂"
    m_UIDisplayMode = "同名合併工作表"
    m_CatalogMode = "覆蓋"
    m_PDFImageReportName = "影像歸檔報告"
    m_PDFCatalogReportName = "目錄歸檔報告"
    m_PDFGridLayout = 4
    m_ExportFilterMode = "全額匯出"
    m_ExportCustomRange = ""
    m_FormatBehavior = "僅重置工作表"
    
    m_PDFMarginCM = 1.5
    m_PDFTitleFontSize = 16
    m_PDFHeaderHeight1 = 35
    m_PDFHeaderHeight2 = 10
    m_PDFTextHeight = 30
    m_PDFImgHeight = 310
    m_PDFGapHeight = 20
    m_PDFNumCols = 2
    m_PDFColWidth = 38
    m_PDFColGap = 2
    m_PDFBaseFontSize = 10
End Sub

Private Sub BuildPathTopology()
    Dim currentDir As String
    currentDir = m_TargetWB.Path
    
    If currentDir = "" Then currentDir = Application.DefaultFilePath
    If Right(currentDir, 1) <> "\" Then currentDir = currentDir & "\"
    
    m_BaseFolder = currentDir
    
    ' ==========================================================================
    ' 【平級四柱拓撲 SSOT 實體建構】
    ' ==========================================================================
    
    ' --- 獨立產出目錄 ---
    m_IOExportFolder = m_BaseFolder & "匯出成果\"
    m_IOExportCatalogFolder = m_IOExportFolder & "目錄\"
    m_IOExportImageFolder = m_IOExportFolder & "圖片\"
    
    ' --- 第一柱：圖片體系 ---
    m_IOTargetFolder = m_BaseFolder & "圖片母夾\"
    m_IOOrganizeFolder = m_IOTargetFolder & "圖片整理夾\"
    m_IORescueFolder = m_IOTargetFolder & "救援夾\"
    m_IORescueDowncastFolder = m_IORescueFolder & "已歸位圖片\"
    m_IORescueDiscardFolder = m_IORescueFolder & "雜湊衝突淘汰救援圖片\"
    m_IODuplicateDiscardFolder = m_IOOrganizeFolder & "雜湊衝突軟刪除匯入圖片\"
    m_IOLooseImportFolder = m_IOOrganizeFolder & "散裝個別匯入\"
    m_IOThumbnailCacheFolder = m_IOOrganizeFolder & ".縮圖快取\"
    m_IOSandboxFolder = m_IOOrganizeFolder & ".暫存區\"
    m_ThumbnailCacheDimension = 1920       'A4 勉強清晰可見
    m_MaxSafetyPoint = 900     ' 【V2-061】約3508px；約2339px；約1754px-->A4印刷的話
    m_RescueEmbedWidth = 300    ' 【V2-062新增】救援夾圖片統一嵌入寬度
    m_RescueEmbedGap = 50       ' 【V2-063新增】救援夾圖片垂直堆疊間距
    m_MinFreeDiskSpaceGB = 5    ' 預設5GB，低於此值視為危險區
    m_MinAvailableMemoryMB = 500   ' 低於500MB可用記憶體視為危險區
    m_HugeFileLimit = 52428800 ' 50 MB
    
    ' --- 第二柱：歸檔附錄體系 (細化分類) ---
    m_IOAttachmentFolder = m_BaseFolder & "歸檔附錄\"
    m_IOAttachPDFFolder = m_IOAttachmentFolder & "PDF\"
    m_IOAttachTXTFolder = m_IOAttachmentFolder & "TXT\"   ' [V4.0修正] TXT與MD專區
    m_IOAttachWordFolder = m_IOAttachmentFolder & "Word\"
    m_IOAttachExcelFolder = m_IOAttachmentFolder & "Excel\"
    m_IOAttachPPTFolder = m_IOAttachmentFolder & "PPT\"
    m_IOAttachZipFolder = m_IOAttachmentFolder & "壓縮封存\"
    m_IOAttachMediaFolder = m_IOAttachmentFolder & "影音\"
    m_IOAttachmentDuplicateDiscardFolder = m_IOAttachmentFolder & "雜湊衝突軟刪除附件\"
    
    ' --- 第三柱：雜物垃圾體系 ---
    m_IOQuarantineFolder = m_BaseFolder & "雜物垃圾\"

    ' --- 第四柱：封存備份體系 ---
    m_IOArchiveFolder = m_BaseFolder & "封存備份\"
    m_IOBackupSnapshotFolder = m_IOArchiveFolder & "快照備份\"
    m_IOLogArchiveFolder = m_IOArchiveFolder & "LOG封存\"
End Sub

' ==============================================================================
' --- 3. 屬性開放區 (全唯讀 Read-Only Getters，捍衛 SSOT 中心) ---
' ==============================================================================

Public Property Get SheetNameManual() As String: SheetNameManual = m_SheetNameManual: End Property
Public Property Get SheetNameAllPics() As String: SheetNameAllPics = m_SheetNameAllPics: End Property
Public Property Get PrefixCustom() As String: PrefixCustom = m_PrefixCustom: End Property
Public Property Get PrefixOldCatalog() As String: PrefixOldCatalog = m_PrefixOldCatalog: End Property
Public Property Get SheetSysMeta() As String: SheetSysMeta = m_SheetSysMeta: End Property
Public Property Get SheetSysLog() As String: SheetSysLog = m_SheetSysLog: End Property
Public Property Get SheetSysTempSettings() As String: SheetSysTempSettings = m_SheetSysTempSettings: End Property
Public Property Get SheetNameTempSafeRoom() As String: SheetNameTempSafeRoom = m_SheetNameTempSafeRoom: End Property
Public Property Get CopySuffixPatterns() As Variant: CopySuffixPatterns = m_CopySuffixPatterns: End Property


Public Property Get AppTitleControlPanel() As String: AppTitleControlPanel = m_AppTitleControlPanel: End Property
Public Property Get AppTitleManual() As String: AppTitleManual = m_AppTitleManual: End Property
Public Property Get CopyrightText() As String: CopyrightText = m_CopyrightText: End Property

Public Property Get OptDisplayMode() As String: OptDisplayMode = m_OptDisplayMode: End Property
Public Property Get OptCatalogMode() As String: OptCatalogMode = m_OptCatalogMode: End Property
Public Property Get OptNamingMode() As String: OptNamingMode = m_OptNamingMode: End Property
Public Property Get OptGridSize() As String: OptGridSize = m_OptGridSize: End Property
Public Property Get OptExportFilter() As String: OptExportFilter = m_OptExportFilter: End Property
Public Property Get OptFormatBehavior() As String: OptFormatBehavior = m_OptFormatBehavior: End Property

Public Property Get FontMain() As String: FontMain = m_FontMain: End Property
Public Property Get FontSizeTitle() As Double: FontSizeTitle = m_FontSizeTitle: End Property
Public Property Get FontSizeSub() As Double: FontSizeSub = m_FontSizeSub: End Property
Public Property Get FontSizeBase() As Double: FontSizeBase = m_FontSizeBase: End Property
Public Property Get FontSizeSmall() As Double: FontSizeSmall = m_FontSizeSmall: End Property

Public Property Get ThemeIO() As Long: ThemeIO = m_ThemeIO: End Property
Public Property Get ThemeExp() As Long: ThemeExp = m_ThemeExp: End Property
Public Property Get ThemeManual() As Long: ThemeManual = m_ThemeManual: End Property
Public Property Get DrawerCustomBg() As Long: DrawerCustomBg = m_DrawerCustomBg: End Property
Public Property Get DrawerImageBg() As Long: DrawerImageBg = m_DrawerImageBg: End Property
Public Property Get DrawerItemBg() As Long: DrawerItemBg = m_DrawerItemBg: End Property
Public Property Get DrawerItemBorder() As Long: DrawerItemBorder = m_DrawerItemBorder: End Property
Public Property Get DrawerItemText() As Long: DrawerItemText = m_DrawerItemText: End Property
Public Property Get DrawerLinkText() As Long: DrawerLinkText = m_DrawerLinkText: End Property

Public Property Get ColorWhite() As Long: ColorWhite = m_ColorWhite: End Property
Public Property Get ColorTitle() As Long: ColorTitle = m_ColorTitle: End Property
Public Property Get ColorTextMain() As Long: ColorTextMain = m_ColorTextMain: End Property
Public Property Get ColorTextHint() As Long: ColorTextHint = m_ColorTextHint: End Property
Public Property Get ColorTextWarning() As Long: ColorTextWarning = m_ColorTextWarning: End Property
Public Property Get ColorTextSuccess() As Long: ColorTextSuccess = m_ColorTextSuccess: End Property
Public Property Get ColorPanelBg() As Long: ColorPanelBg = m_ColorPanelBg: End Property
Public Property Get ColorShadow() As Long: ColorShadow = m_ColorShadow: End Property
Public Property Get InputBg() As Long: InputBg = m_InputBg: End Property
Public Property Get ReadonlyBg() As Long: ReadonlyBg = m_ReadonlyBg: End Property
Public Property Get Border() As Long: Border = m_Border: End Property
Public Property Get ColorLinkText() As Long: ColorLinkText = m_ColorLinkText: End Property

Public Property Get GapX() As Double: GapX = m_GapX: End Property
Public Property Get GapY() As Double: GapY = m_GapY: End Property
Public Property Get QuarterW() As Double: QuarterW = m_QuarterW: End Property

Public Property Get RangeSettingsConsole() As String: RangeSettingsConsole = m_RangeSettingsConsole: End Property
Public Property Get CellProjName() As String: CellProjName = m_CellProjName: End Property
Public Property Get CellDangerConsole() As String: CellDangerConsole = m_CellDangerConsole: End Property
Public Property Get CellAutoLayoutStateConsole() As String: CellAutoLayoutStateConsole = m_CellAutoLayoutStateConsole: End Property

Public Property Get btnLockUI() As String: btnLockUI = m_btnLockUI: End Property
Public Property Get btnLockDir() As String: btnLockDir = m_btnLockDir: End Property
Public Property Get BtnAutoLayout() As String: BtnAutoLayout = m_BtnAutoLayout: End Property
Public Property Get btnGlobalLock() As String: btnGlobalLock = m_btnGlobalLock: End Property
Public Property Get BtnRebuildIndex() As String: BtnRebuildIndex = m_BtnRebuildIndex: End Property
Public Property Get BtnSystemRestart() As String: BtnSystemRestart = m_BtnSystemRestart: End Property
Public Property Get BtnFormat() As String: BtnFormat = m_BtnFormat: End Property
Public Property Get BtnToggleLog() As String: BtnToggleLog = m_BtnToggleLog: End Property
Public Property Get BtnLayoutLock() As String: BtnLayoutLock = m_BtnLayoutLock: End Property

Public Property Get MaxDagLoopCountDAGEngine() As Long: MaxDagLoopCountDAGEngine = m_MaxDagLoopCountDAGEngine: End Property
Public Property Get CellDebugStateFlagSet() As String: CellDebugStateFlagSet = m_CellDebugStateFlagSet: End Property
Public Property Get CellDisplayModeFlagSet() As String: CellDisplayModeFlagSet = m_CellDisplayModeFlagSet: End Property
Public Property Get CellCatalogModeFlagSet() As String: CellCatalogModeFlagSet = m_CellCatalogModeFlagSet: End Property

Public Property Get CellImgWidth() As String: CellImgWidth = m_CellImgWidth: End Property
Public Property Get CellImgSpacing() As String: CellImgSpacing = m_CellImgSpacing: End Property
Public Property Get CellNamingMode() As String: CellNamingMode = m_CellNamingMode: End Property
Public Property Get CellSheetPrefix() As String: CellSheetPrefix = m_CellSheetPrefix: End Property
Public Property Get CellPdfImgName() As String: CellPdfImgName = m_CellPdfImgName: End Property
Public Property Get CellPdfCatName() As String: CellPdfCatName = m_CellPdfCatName: End Property
Public Property Get CellPdfGrid() As String: CellPdfGrid = m_CellPdfGrid: End Property
Public Property Get CellExportFilter() As String: CellExportFilter = m_CellExportFilter: End Property
Public Property Get CellExportRange() As String: CellExportRange = m_CellExportRange: End Property


Public Property Get MinImgWidth() As Double: MinImgWidth = m_MinImgWidth: End Property
Public Property Get MaxImgWidth() As Double: MaxImgWidth = m_MaxImgWidth: End Property
Public Property Get MinImgSpacing() As Double: MinImgSpacing = m_MinImgSpacing: End Property
Public Property Get MaxImgSpacing() As Double: MaxImgSpacing = m_MaxImgSpacing: End Property

Public Property Get TargetWB() As Workbook: Set TargetWB = m_TargetWB: End Property
Public Property Get BaseFolder() As String: BaseFolder = m_BaseFolder: End Property
Public Property Get ExpectedRootFolderName() As String: ExpectedRootFolderName = m_ExpectedRootFolderName: End Property

Public Property Get IOExportFolder() As String: IOExportFolder = m_IOExportFolder: End Property
Public Property Get IOExportCatalogFolder() As String: IOExportCatalogFolder = m_IOExportCatalogFolder: End Property
Public Property Get IOExportImageFolder() As String: IOExportImageFolder = m_IOExportImageFolder: End Property

Public Property Get IOTargetFolder() As String: IOTargetFolder = m_IOTargetFolder: End Property
Public Property Get IOOrganizeFolder() As String: IOOrganizeFolder = m_IOOrganizeFolder: End Property
Public Property Get IORescueFolder() As String: IORescueFolder = m_IORescueFolder: End Property
Public Property Get IORescueDowncastFolder() As String: IORescueDowncastFolder = m_IORescueDowncastFolder: End Property
Public Property Get IORescueDiscardFolder() As String: IORescueDiscardFolder = m_IORescueDiscardFolder: End Property
Public Property Get IODuplicateDiscardFolder() As String: IODuplicateDiscardFolder = m_IODuplicateDiscardFolder: End Property
Public Property Get IOThumbnailCacheFolder() As String: IOThumbnailCacheFolder = m_IOThumbnailCacheFolder: End Property
Public Property Get IOLooseImportFolder() As String: IOLooseImportFolder = m_IOLooseImportFolder: End Property
Public Property Get ThumbnailCacheDimension() As Double: ThumbnailCacheDimension = m_ThumbnailCacheDimension: End Property
Public Property Get MaxSafetyPoint() As Double: MaxSafetyPoint = m_MaxSafetyPoint: End Property
Public Property Get RescueEmbedWidth() As Double: RescueEmbedWidth = m_RescueEmbedWidth: End Property
Public Property Get RescueEmbedGap() As Double: RescueEmbedGap = m_RescueEmbedGap: End Property
Public Property Get MinFreeDiskSpaceGB() As Double: MinFreeDiskSpaceGB = m_MinFreeDiskSpaceGB: End Property
Public Property Get MinAvailableMemoryMB() As Double: MinAvailableMemoryMB = m_MinAvailableMemoryMB: End Property
Public Property Get HugeFileLimit() As Double: HugeFileLimit = m_HugeFileLimit: End Property


Public Property Get IOAttachmentFolder() As String: IOAttachmentFolder = m_IOAttachmentFolder: End Property
Public Property Get IOAttachPDFFolder() As String: IOAttachPDFFolder = m_IOAttachPDFFolder: End Property
Public Property Get IOAttachTXTFolder() As String: IOAttachTXTFolder = m_IOAttachTXTFolder: End Property ' [V4.0修正]
Public Property Get IOAttachWordFolder() As String: IOAttachWordFolder = m_IOAttachWordFolder: End Property
Public Property Get IOAttachExcelFolder() As String: IOAttachExcelFolder = m_IOAttachExcelFolder: End Property
Public Property Get IOAttachPPTFolder() As String: IOAttachPPTFolder = m_IOAttachPPTFolder: End Property
Public Property Get IOAttachZipFolder() As String: IOAttachZipFolder = m_IOAttachZipFolder: End Property
Public Property Get IOAttachMediaFolder() As String: IOAttachMediaFolder = m_IOAttachMediaFolder: End Property
Public Property Get IOAttachmentDuplicateDiscardFolder() As String: IOAttachmentDuplicateDiscardFolder = m_IOAttachmentDuplicateDiscardFolder: End Property

Public Property Get IOQuarantineFolder() As String: IOQuarantineFolder = m_IOQuarantineFolder: End Property
Public Property Get IOBackupSnapshotFolder() As String: IOBackupSnapshotFolder = m_IOBackupSnapshotFolder: End Property
Public Property Get IOArchiveFolder() As String: IOArchiveFolder = m_IOArchiveFolder: End Property
Public Property Get IOLogArchiveFolder() As String: IOLogArchiveFolder = m_IOLogArchiveFolder: End Property

Public Property Get UIProjectName() As String: UIProjectName = m_UIProjectName: End Property
Public Property Get SheetNameUI() As String: SheetNameUI = m_SheetNameUI: End Property
Public Property Get SheetNameCatalog() As String: SheetNameCatalog = m_SheetNameCatalog: End Property
Public Property Get SheetNameAttachment() As String: SheetNameAttachment = m_SheetNameAttachment: End Property
Public Property Get SheetNameRescue() As String: SheetNameRescue = m_SheetNameRescue: End Property
Public Property Get SheetNameQuarantine() As String: SheetNameQuarantine = m_SheetNameQuarantine: End Property
Public Property Get SecurityPassword() As String: SecurityPassword = m_SecurityPassword: End Property

Public Property Get UIImgWidth() As Double: UIImgWidth = m_UIImgWidth: End Property
Public Property Get UIImgSpacing() As Double: UIImgSpacing = m_UIImgSpacing: End Property
Public Property Get NamingMode() As String: NamingMode = m_NamingMode: End Property
Public Property Get UISheetPrefix() As String: UISheetPrefix = m_UISheetPrefix: End Property
Public Property Get UIColorPassGreen() As Long: UIColorPassGreen = m_UIColorPassGreen: End Property
Public Property Get UIColorLockRed() As Long: UIColorLockRed = m_UIColorLockRed: End Property
Public Property Get UIColorNormal() As Long: UIColorNormal = m_UIColorNormal: End Property

Public Property Get PDFImageReportName() As String: PDFImageReportName = m_PDFImageReportName: End Property
Public Property Get PDFCatalogReportName() As String: PDFCatalogReportName = m_PDFCatalogReportName: End Property
Public Property Get UIDisplayMode() As String: UIDisplayMode = m_UIDisplayMode: End Property
Public Property Get CatalogMode() As String: CatalogMode = m_CatalogMode: End Property
Public Property Get ExportFilterMode() As String: ExportFilterMode = m_ExportFilterMode: End Property
Public Property Get ExportCustomRange() As String: ExportCustomRange = m_ExportCustomRange: End Property
Public Property Get FormatBehavior() As String: FormatBehavior = m_FormatBehavior: End Property

Public Property Get PDFGridLayout() As Integer: PDFGridLayout = m_PDFGridLayout: End Property
Public Property Get PDFMarginCM() As Double: PDFMarginCM = m_PDFMarginCM: End Property
Public Property Get PDFTitleFontSize() As Integer: PDFTitleFontSize = m_PDFTitleFontSize: End Property
Public Property Get PDFHeaderHeight1() As Double: PDFHeaderHeight1 = m_PDFHeaderHeight1: End Property
Public Property Get PDFHeaderHeight2() As Double: PDFHeaderHeight2 = m_PDFHeaderHeight2: End Property
Public Property Get PDFTextHeight() As Double: PDFTextHeight = m_PDFTextHeight: End Property
Public Property Get PDFImgHeight() As Double: PDFImgHeight = m_PDFImgHeight: End Property
Public Property Get PDFGapHeight() As Double: PDFGapHeight = m_PDFGapHeight: End Property
Public Property Get PDFNumCols() As Integer: PDFNumCols = m_PDFNumCols: End Property
Public Property Get PDFColWidth() As Double: PDFColWidth = m_PDFColWidth: End Property
Public Property Get PDFColGap() As Double: PDFColGap = m_PDFColGap: End Property
Public Property Get PDFBaseFontSize() As Integer: PDFBaseFontSize = m_PDFBaseFontSize: End Property

Public Property Get IOSandboxFolder() As String: IOSandboxFolder = m_IOSandboxFolder: End Property
