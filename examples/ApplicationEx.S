'COM._Application.GetProperty'  = list('Application' = function() {
	 ans = .COM(x, 'Application', .dispatch = as.integer(2))
	 as(ans, 'Application')
},
'Creator' = function() {
	 ans = .COM(x, 'Creator', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'Parent' = function() {
	 ans = .COM(x, 'Parent', .dispatch = as.integer(2))
	 as(ans, 'Application')
},
'ActiveCell' = function() {
	 ans = .COM(x, 'ActiveCell', .dispatch = as.integer(2))
	 as(ans, 'Range')
},
'ActiveChart' = function() {
	 ans = .COM(x, 'ActiveChart', .dispatch = as.integer(2))
	 as(ans, 'Chart')
},
'ActiveDialog' = function() {
	 ans = .COM(x, 'ActiveDialog', .dispatch = as.integer(2))
	 as(ans, 'DialogSheet')
},
'ActiveMenuBar' = function() {
	 ans = .COM(x, 'ActiveMenuBar', .dispatch = as.integer(2))
	 as(ans, 'MenuBar')
},
'ActivePrinter' = function() {
	 ans = .COM(x, 'ActivePrinter', .dispatch = as.integer(2))
ans
},
'ActiveSheet' = function() {
	 ans = .COM(x, 'ActiveSheet', .dispatch = as.integer(2))
	 ans
},
'ActiveWindow' = function() {
	 ans = .COM(x, 'ActiveWindow', .dispatch = as.integer(2))
	 as(ans, 'Window')
},
'ActiveWorkbook' = function() {
	 ans = .COM(x, 'ActiveWorkbook', .dispatch = as.integer(2))
	 as(ans, 'Workbook')
},
'AddIns' = function() {
	 ans = .COM(x, 'AddIns', .dispatch = as.integer(2))
	 as(ans, 'AddIns')
},
'Assistant' = function() {
	 ans = .COM(x, 'Assistant', .dispatch = as.integer(2))
	 as(ans, 'NA')
},
'Cells' = function() {
	 ans = .COM(x, 'Cells', .dispatch = as.integer(2))
	 as(ans, 'Range')
},
'Charts' = function() {
	 ans = .COM(x, 'Charts', .dispatch = as.integer(2))
	 as(ans, 'Sheets')
},
'Columns' = function() {
	 ans = .COM(x, 'Columns', .dispatch = as.integer(2))
	 as(ans, 'Range')
},
'CommandBars' = function() {
	 ans = .COM(x, 'CommandBars', .dispatch = as.integer(2))
	 as(ans, 'NA')
},
'DDEAppReturnCode' = function() {
	 ans = .COM(x, 'DDEAppReturnCode', .dispatch = as.integer(2))
ans
},
'DialogSheets' = function() {
	 ans = .COM(x, 'DialogSheets', .dispatch = as.integer(2))
	 as(ans, 'Sheets')
},
'MenuBars' = function() {
	 ans = .COM(x, 'MenuBars', .dispatch = as.integer(2))
	 as(ans, 'MenuBars')
},
'Modules' = function() {
	 ans = .COM(x, 'Modules', .dispatch = as.integer(2))
	 as(ans, 'Modules')
},
'Names' = function() {
	 ans = .COM(x, 'Names', .dispatch = as.integer(2))
	 as(ans, 'Names')
},
'Rows' = function() {
	 ans = .COM(x, 'Rows', .dispatch = as.integer(2))
	 as(ans, 'Range')
},
'Selection' = function() {
	 ans = .COM(x, 'Selection', .dispatch = as.integer(2))
	 ans
},
'Sheets' = function() {
	 ans = .COM(x, 'Sheets', .dispatch = as.integer(2))
	 as(ans, 'Sheets')
},
'ThisWorkbook' = function() {
	 ans = .COM(x, 'ThisWorkbook', .dispatch = as.integer(2))
	 as(ans, 'Workbook')
},
'Toolbars' = function() {
	 ans = .COM(x, 'Toolbars', .dispatch = as.integer(2))
	 as(ans, 'Toolbars')
},
'Windows' = function() {
	 ans = .COM(x, 'Windows', .dispatch = as.integer(2))
	 as(ans, 'Windows')
},
'Workbooks' = function() {
	 ans = .COM(x, 'Workbooks', .dispatch = as.integer(2))
	 as(ans, 'Workbooks')
},
'WorksheetFunction' = function() {
	 ans = .COM(x, 'WorksheetFunction', .dispatch = as.integer(2))
	 as(ans, 'WorksheetFunction')
},
'Worksheets' = function() {
	 ans = .COM(x, 'Worksheets', .dispatch = as.integer(2))
	 as(ans, 'Sheets')
},
'Excel4IntlMacroSheets' = function() {
	 ans = .COM(x, 'Excel4IntlMacroSheets', .dispatch = as.integer(2))
	 as(ans, 'Sheets')
},
'Excel4MacroSheets' = function() {
	 ans = .COM(x, 'Excel4MacroSheets', .dispatch = as.integer(2))
	 as(ans, 'Sheets')
},
'AlertBeforeOverwriting' = function() {
	 ans = .COM(x, 'AlertBeforeOverwriting', .dispatch = as.integer(2))
ans
},
'AltStartupPath' = function() {
	 ans = .COM(x, 'AltStartupPath', .dispatch = as.integer(2))
ans
},
'AskToUpdateLinks' = function() {
	 ans = .COM(x, 'AskToUpdateLinks', .dispatch = as.integer(2))
ans
},
'EnableAnimations' = function() {
	 ans = .COM(x, 'EnableAnimations', .dispatch = as.integer(2))
ans
},
'AutoCorrect' = function() {
	 ans = .COM(x, 'AutoCorrect', .dispatch = as.integer(2))
	 as(ans, 'AutoCorrect')
},
'Build' = function() {
	 ans = .COM(x, 'Build', .dispatch = as.integer(2))
ans
},
'CalculateBeforeSave' = function() {
	 ans = .COM(x, 'CalculateBeforeSave', .dispatch = as.integer(2))
ans
},
'Calculation' = function() {
	 ans = .COM(x, 'Calculation', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'CanPlaySounds' = function() {
	 ans = .COM(x, 'CanPlaySounds', .dispatch = as.integer(2))
ans
},
'CanRecordSounds' = function() {
	 ans = .COM(x, 'CanRecordSounds', .dispatch = as.integer(2))
ans
},
'Caption' = function() {
	 ans = .COM(x, 'Caption', .dispatch = as.integer(2))
ans
},
'CellDragAndDrop' = function() {
	 ans = .COM(x, 'CellDragAndDrop', .dispatch = as.integer(2))
ans
},
'DisplayClipboardWindow' = function() {
	 ans = .COM(x, 'DisplayClipboardWindow', .dispatch = as.integer(2))
ans
},
'ColorButtons' = function() {
	 ans = .COM(x, 'ColorButtons', .dispatch = as.integer(2))
ans
},
'CommandUnderlines' = function() {
	 ans = .COM(x, 'CommandUnderlines', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'ConstrainNumeric' = function() {
	 ans = .COM(x, 'ConstrainNumeric', .dispatch = as.integer(2))
ans
},
'CopyObjectsWithCells' = function() {
	 ans = .COM(x, 'CopyObjectsWithCells', .dispatch = as.integer(2))
ans
},
'Cursor' = function() {
	 ans = .COM(x, 'Cursor', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'CustomListCount' = function() {
	 ans = .COM(x, 'CustomListCount', .dispatch = as.integer(2))
ans
},
'CutCopyMode' = function() {
	 ans = .COM(x, 'CutCopyMode', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'DataEntryMode' = function() {
	 ans = .COM(x, 'DataEntryMode', .dispatch = as.integer(2))
ans
},
'_Default' = function() {
	 ans = .COM(x, '_Default', .dispatch = as.integer(2))
ans
},
'DefaultFilePath' = function() {
	 ans = .COM(x, 'DefaultFilePath', .dispatch = as.integer(2))
ans
},
'Dialogs' = function() {
	 ans = .COM(x, 'Dialogs', .dispatch = as.integer(2))
	 as(ans, 'Dialogs')
},
'DisplayAlerts' = function() {
	 ans = .COM(x, 'DisplayAlerts', .dispatch = as.integer(2))
ans
},
'DisplayFormulaBar' = function() {
	 ans = .COM(x, 'DisplayFormulaBar', .dispatch = as.integer(2))
ans
},
'DisplayFullScreen' = function() {
	 ans = .COM(x, 'DisplayFullScreen', .dispatch = as.integer(2))
ans
},
'DisplayNoteIndicator' = function() {
	 ans = .COM(x, 'DisplayNoteIndicator', .dispatch = as.integer(2))
ans
},
'DisplayCommentIndicator' = function() {
	 ans = .COM(x, 'DisplayCommentIndicator', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'DisplayExcel4Menus' = function() {
	 ans = .COM(x, 'DisplayExcel4Menus', .dispatch = as.integer(2))
ans
},
'DisplayRecentFiles' = function() {
	 ans = .COM(x, 'DisplayRecentFiles', .dispatch = as.integer(2))
ans
},
'DisplayScrollBars' = function() {
	 ans = .COM(x, 'DisplayScrollBars', .dispatch = as.integer(2))
ans
},
'DisplayStatusBar' = function() {
	 ans = .COM(x, 'DisplayStatusBar', .dispatch = as.integer(2))
ans
},
'EditDirectlyInCell' = function() {
	 ans = .COM(x, 'EditDirectlyInCell', .dispatch = as.integer(2))
ans
},
'EnableAutoComplete' = function() {
	 ans = .COM(x, 'EnableAutoComplete', .dispatch = as.integer(2))
ans
},
'EnableCancelKey' = function() {
	 ans = .COM(x, 'EnableCancelKey', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'EnableSound' = function() {
	 ans = .COM(x, 'EnableSound', .dispatch = as.integer(2))
ans
},
'EnableTipWizard' = function() {
	 ans = .COM(x, 'EnableTipWizard', .dispatch = as.integer(2))
ans
},
'FileSearch' = function() {
	 ans = .COM(x, 'FileSearch', .dispatch = as.integer(2))
	 as(ans, 'NA')
},
'FileFind' = function() {
	 ans = .COM(x, 'FileFind', .dispatch = as.integer(2))
	 as(ans, 'NA')
},
'FixedDecimal' = function() {
	 ans = .COM(x, 'FixedDecimal', .dispatch = as.integer(2))
ans
},
'FixedDecimalPlaces' = function() {
	 ans = .COM(x, 'FixedDecimalPlaces', .dispatch = as.integer(2))
ans
},
'Height' = function() {
	 ans = .COM(x, 'Height', .dispatch = as.integer(2))
ans
},
'IgnoreRemoteRequests' = function() {
	 ans = .COM(x, 'IgnoreRemoteRequests', .dispatch = as.integer(2))
ans
},
'Interactive' = function() {
	 ans = .COM(x, 'Interactive', .dispatch = as.integer(2))
ans
},
'Iteration' = function() {
	 ans = .COM(x, 'Iteration', .dispatch = as.integer(2))
ans
},
'LargeButtons' = function() {
	 ans = .COM(x, 'LargeButtons', .dispatch = as.integer(2))
ans
},
'Left' = function() {
	 ans = .COM(x, 'Left', .dispatch = as.integer(2))
ans
},
'LibraryPath' = function() {
	 ans = .COM(x, 'LibraryPath', .dispatch = as.integer(2))
ans
},
'MailSession' = function() {
	 ans = .COM(x, 'MailSession', .dispatch = as.integer(2))
	 ans
},
'MailSystem' = function() {
	 ans = .COM(x, 'MailSystem', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'MathCoprocessorAvailable' = function() {
	 ans = .COM(x, 'MathCoprocessorAvailable', .dispatch = as.integer(2))
ans
},
'MaxChange' = function() {
	 ans = .COM(x, 'MaxChange', .dispatch = as.integer(2))
ans
},
'MaxIterations' = function() {
	 ans = .COM(x, 'MaxIterations', .dispatch = as.integer(2))
ans
},
'MemoryFree' = function() {
	 ans = .COM(x, 'MemoryFree', .dispatch = as.integer(2))
ans
},
'MemoryTotal' = function() {
	 ans = .COM(x, 'MemoryTotal', .dispatch = as.integer(2))
ans
},
'MemoryUsed' = function() {
	 ans = .COM(x, 'MemoryUsed', .dispatch = as.integer(2))
ans
},
'MouseAvailable' = function() {
	 ans = .COM(x, 'MouseAvailable', .dispatch = as.integer(2))
ans
},
'MoveAfterReturn' = function() {
	 ans = .COM(x, 'MoveAfterReturn', .dispatch = as.integer(2))
ans
},
'MoveAfterReturnDirection' = function() {
	 ans = .COM(x, 'MoveAfterReturnDirection', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'RecentFiles' = function() {
	 ans = .COM(x, 'RecentFiles', .dispatch = as.integer(2))
	 as(ans, 'RecentFiles')
},
'Name' = function() {
	 ans = .COM(x, 'Name', .dispatch = as.integer(2))
ans
},
'NetworkTemplatesPath' = function() {
	 ans = .COM(x, 'NetworkTemplatesPath', .dispatch = as.integer(2))
ans
},
'ODBCErrors' = function() {
	 ans = .COM(x, 'ODBCErrors', .dispatch = as.integer(2))
	 as(ans, 'ODBCErrors')
},
'ODBCTimeout' = function() {
	 ans = .COM(x, 'ODBCTimeout', .dispatch = as.integer(2))
ans
},
'OnCalculate' = function() {
	 ans = .COM(x, 'OnCalculate', .dispatch = as.integer(2))
ans
},
'OnData' = function() {
	 ans = .COM(x, 'OnData', .dispatch = as.integer(2))
ans
},
'OnDoubleClick' = function() {
	 ans = .COM(x, 'OnDoubleClick', .dispatch = as.integer(2))
ans
},
'OnEntry' = function() {
	 ans = .COM(x, 'OnEntry', .dispatch = as.integer(2))
ans
},
'OnSheetActivate' = function() {
	 ans = .COM(x, 'OnSheetActivate', .dispatch = as.integer(2))
ans
},
'OnSheetDeactivate' = function() {
	 ans = .COM(x, 'OnSheetDeactivate', .dispatch = as.integer(2))
ans
},
'OnWindow' = function() {
	 ans = .COM(x, 'OnWindow', .dispatch = as.integer(2))
ans
},
'OperatingSystem' = function() {
	 ans = .COM(x, 'OperatingSystem', .dispatch = as.integer(2))
ans
},
'OrganizationName' = function() {
	 ans = .COM(x, 'OrganizationName', .dispatch = as.integer(2))
ans
},
'Path' = function() {
	 ans = .COM(x, 'Path', .dispatch = as.integer(2))
ans
},
'PathSeparator' = function() {
	 ans = .COM(x, 'PathSeparator', .dispatch = as.integer(2))
ans
},
'PivotTableSelection' = function() {
	 ans = .COM(x, 'PivotTableSelection', .dispatch = as.integer(2))
ans
},
'PromptForSummaryInfo' = function() {
	 ans = .COM(x, 'PromptForSummaryInfo', .dispatch = as.integer(2))
ans
},
'RecordRelative' = function() {
	 ans = .COM(x, 'RecordRelative', .dispatch = as.integer(2))
ans
},
'ReferenceStyle' = function() {
	 ans = .COM(x, 'ReferenceStyle', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'RollZoom' = function() {
	 ans = .COM(x, 'RollZoom', .dispatch = as.integer(2))
ans
},
'ScreenUpdating' = function() {
	 ans = .COM(x, 'ScreenUpdating', .dispatch = as.integer(2))
ans
},
'SheetsInNewWorkbook' = function() {
	 ans = .COM(x, 'SheetsInNewWorkbook', .dispatch = as.integer(2))
ans
},
'ShowChartTipNames' = function() {
	 ans = .COM(x, 'ShowChartTipNames', .dispatch = as.integer(2))
ans
},
'ShowChartTipValues' = function() {
	 ans = .COM(x, 'ShowChartTipValues', .dispatch = as.integer(2))
ans
},
'StandardFont' = function() {
	 ans = .COM(x, 'StandardFont', .dispatch = as.integer(2))
ans
},
'StandardFontSize' = function() {
	 ans = .COM(x, 'StandardFontSize', .dispatch = as.integer(2))
ans
},
'StartupPath' = function() {
	 ans = .COM(x, 'StartupPath', .dispatch = as.integer(2))
ans
},
'StatusBar' = function() {
	 ans = .COM(x, 'StatusBar', .dispatch = as.integer(2))
	 ans
},
'TemplatesPath' = function() {
	 ans = .COM(x, 'TemplatesPath', .dispatch = as.integer(2))
ans
},
'ShowToolTips' = function() {
	 ans = .COM(x, 'ShowToolTips', .dispatch = as.integer(2))
ans
},
'Top' = function() {
	 ans = .COM(x, 'Top', .dispatch = as.integer(2))
ans
},
'DefaultSaveFormat' = function() {
	 ans = .COM(x, 'DefaultSaveFormat', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'TransitionMenuKey' = function() {
	 ans = .COM(x, 'TransitionMenuKey', .dispatch = as.integer(2))
ans
},
'TransitionMenuKeyAction' = function() {
	 ans = .COM(x, 'TransitionMenuKeyAction', .dispatch = as.integer(2))
ans
},
'TransitionNavigKeys' = function() {
	 ans = .COM(x, 'TransitionNavigKeys', .dispatch = as.integer(2))
ans
},
'UsableHeight' = function() {
	 ans = .COM(x, 'UsableHeight', .dispatch = as.integer(2))
ans
},
'UsableWidth' = function() {
	 ans = .COM(x, 'UsableWidth', .dispatch = as.integer(2))
ans
},
'UserControl' = function() {
	 ans = .COM(x, 'UserControl', .dispatch = as.integer(2))
ans
},
'UserName' = function() {
	 ans = .COM(x, 'UserName', .dispatch = as.integer(2))
ans
},
'Value' = function() {
	 ans = .COM(x, 'Value', .dispatch = as.integer(2))
ans
},
'VBE' = function() {
	 ans = .COM(x, 'VBE', .dispatch = as.integer(2))
	 as(ans, 'NA')
},
'Version' = function() {
	 ans = .COM(x, 'Version', .dispatch = as.integer(2))
ans
},
'Visible' = function() {
	 ans = .COM(x, 'Visible', .dispatch = as.integer(2))
ans
},
'Width' = function() {
	 ans = .COM(x, 'Width', .dispatch = as.integer(2))
ans
},
'WindowsForPens' = function() {
	 ans = .COM(x, 'WindowsForPens', .dispatch = as.integer(2))
ans
},
'WindowState' = function() {
	 ans = .COM(x, 'WindowState', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'UILanguage' = function() {
	 ans = .COM(x, 'UILanguage', .dispatch = as.integer(2))
ans
},
'DefaultSheetDirection' = function() {
	 ans = .COM(x, 'DefaultSheetDirection', .dispatch = as.integer(2))
ans
},
'CursorMovement' = function() {
	 ans = .COM(x, 'CursorMovement', .dispatch = as.integer(2))
ans
},
'ControlCharacters' = function() {
	 ans = .COM(x, 'ControlCharacters', .dispatch = as.integer(2))
ans
},
'EnableEvents' = function() {
	 ans = .COM(x, 'EnableEvents', .dispatch = as.integer(2))
ans
},
'DisplayInfoWindow' = function() {
	 ans = .COM(x, 'DisplayInfoWindow', .dispatch = as.integer(2))
ans
},
'ExtendList' = function() {
	 ans = .COM(x, 'ExtendList', .dispatch = as.integer(2))
ans
},
'OLEDBErrors' = function() {
	 ans = .COM(x, 'OLEDBErrors', .dispatch = as.integer(2))
	 as(ans, 'OLEDBErrors')
},
'COMAddIns' = function() {
	 ans = .COM(x, 'COMAddIns', .dispatch = as.integer(2))
	 as(ans, 'NA')
},
'DefaultWebOptions' = function() {
	 ans = .COM(x, 'DefaultWebOptions', .dispatch = as.integer(2))
	 as(ans, 'DefaultWebOptions')
},
'ProductCode' = function() {
	 ans = .COM(x, 'ProductCode', .dispatch = as.integer(2))
ans
},
'UserLibraryPath' = function() {
	 ans = .COM(x, 'UserLibraryPath', .dispatch = as.integer(2))
ans
},
'AutoPercentEntry' = function() {
	 ans = .COM(x, 'AutoPercentEntry', .dispatch = as.integer(2))
ans
},
'LanguageSettings' = function() {
	 ans = .COM(x, 'LanguageSettings', .dispatch = as.integer(2))
	 as(ans, 'NA')
},
'Dummy101' = function() {
	 ans = .COM(x, 'Dummy101', .dispatch = as.integer(2))
	 ans
},
'AnswerWizard' = function() {
	 ans = .COM(x, 'AnswerWizard', .dispatch = as.integer(2))
	 as(ans, 'NA')
},
'CalculationVersion' = function() {
	 ans = .COM(x, 'CalculationVersion', .dispatch = as.integer(2))
ans
},
'ShowWindowsInTaskbar' = function() {
	 ans = .COM(x, 'ShowWindowsInTaskbar', .dispatch = as.integer(2))
ans
},
'FeatureInstall' = function() {
	 ans = .COM(x, 'FeatureInstall', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'Ready' = function() {
	 ans = .COM(x, 'Ready', .dispatch = as.integer(2))
ans
},
'FindFormat' = function() {
	 ans = .COM(x, 'FindFormat', .dispatch = as.integer(2))
	 as(ans, 'CellFormat')
},
'ReplaceFormat' = function() {
	 ans = .COM(x, 'ReplaceFormat', .dispatch = as.integer(2))
	 as(ans, 'CellFormat')
},
'UsedObjects' = function() {
	 ans = .COM(x, 'UsedObjects', .dispatch = as.integer(2))
	 as(ans, 'UsedObjects')
},
'CalculationState' = function() {
	 ans = .COM(x, 'CalculationState', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'CalculationInterruptKey' = function() {
	 ans = .COM(x, 'CalculationInterruptKey', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'Watches' = function() {
	 ans = .COM(x, 'Watches', .dispatch = as.integer(2))
	 as(ans, 'Watches')
},
'DisplayFunctionToolTips' = function() {
	 ans = .COM(x, 'DisplayFunctionToolTips', .dispatch = as.integer(2))
ans
},
'AutomationSecurity' = function() {
	 ans = .COM(x, 'AutomationSecurity', .dispatch = as.integer(2))
	 as(ans, 'Constants')
},
'DisplayPasteOptions' = function() {
	 ans = .COM(x, 'DisplayPasteOptions', .dispatch = as.integer(2))
ans
},
'DisplayInsertOptions' = function() {
	 ans = .COM(x, 'DisplayInsertOptions', .dispatch = as.integer(2))
ans
},
'GenerateGetPivotData' = function() {
	 ans = .COM(x, 'GenerateGetPivotData', .dispatch = as.integer(2))
ans
},
'AutoRecover' = function() {
	 ans = .COM(x, 'AutoRecover', .dispatch = as.integer(2))
	 as(ans, 'AutoRecover')
},
'Hwnd' = function() {
	 ans = .COM(x, 'Hwnd', .dispatch = as.integer(2))
ans
},
'Hinstance' = function() {
	 ans = .COM(x, 'Hinstance', .dispatch = as.integer(2))
ans
},
'ErrorCheckingOptions' = function() {
	 ans = .COM(x, 'ErrorCheckingOptions', .dispatch = as.integer(2))
	 as(ans, 'ErrorCheckingOptions')
},
'AutoFormatAsYouTypeReplaceHyperlinks' = function() {
	 ans = .COM(x, 'AutoFormatAsYouTypeReplaceHyperlinks', .dispatch = as.integer(2))
ans
},
'SmartTagRecognizers' = function() {
	 ans = .COM(x, 'SmartTagRecognizers', .dispatch = as.integer(2))
	 as(ans, 'SmartTagRecognizers')
},
'NewWorkbook' = function() {
	 ans = .COM(x, 'NewWorkbook', .dispatch = as.integer(2))
	 as(ans, 'NA')
},
'SpellingOptions' = function() {
	 ans = .COM(x, 'SpellingOptions', .dispatch = as.integer(2))
	 as(ans, 'SpellingOptions')
},
'Speech' = function() {
	 ans = .COM(x, 'Speech', .dispatch = as.integer(2))
	 as(ans, 'Speech')
},
'MapPaperSize' = function() {
	 ans = .COM(x, 'MapPaperSize', .dispatch = as.integer(2))
ans
},
'ShowStartupDialog' = function() {
	 ans = .COM(x, 'ShowStartupDialog', .dispatch = as.integer(2))
ans
},
'DecimalSeparator' = function() {
	 ans = .COM(x, 'DecimalSeparator', .dispatch = as.integer(2))
ans
},
'ThousandsSeparator' = function() {
	 ans = .COM(x, 'ThousandsSeparator', .dispatch = as.integer(2))
ans
},
'UseSystemSeparators' = function() {
	 ans = .COM(x, 'UseSystemSeparators', .dispatch = as.integer(2))
ans
},
'ThisCell' = function() {
	 ans = .COM(x, 'ThisCell', .dispatch = as.integer(2))
	 as(ans, 'Range')
},
'RTD' = function() {
	 ans = .COM(x, 'RTD', .dispatch = as.integer(2))
	 as(ans, 'RTD')
},
'DisplayDocumentActionTaskPane' = function() {
	 ans = .COM(x, 'DisplayDocumentActionTaskPane', .dispatch = as.integer(2))
ans
},
'ArbitraryXMLSupportAvailable' = function() {
	 ans = .COM(x, 'ArbitraryXMLSupportAvailable', .dispatch = as.integer(2))
ans
} )
'COM._Application.SetProperty'  = list('ActivePrinter' = function(value) {
value = as(value, 'character')
.COM(x, 'ActivePrinter', value, .dispatch = as.integer(4))
},
'AlertBeforeOverwriting' = function(value) {
value = as(value, 'logical')
.COM(x, 'AlertBeforeOverwriting', value, .dispatch = as.integer(4))
},
'AltStartupPath' = function(value) {
value = as(value, 'character')
.COM(x, 'AltStartupPath', value, .dispatch = as.integer(4))
},
'AskToUpdateLinks' = function(value) {
value = as(value, 'logical')
.COM(x, 'AskToUpdateLinks', value, .dispatch = as.integer(4))
},
'EnableAnimations' = function(value) {
value = as(value, 'logical')
.COM(x, 'EnableAnimations', value, .dispatch = as.integer(4))
},
'CalculateBeforeSave' = function(value) {
value = as(value, 'logical')
.COM(x, 'CalculateBeforeSave', value, .dispatch = as.integer(4))
},
'Calculation' = function(value) {
value = as(value, 'Constants')
.COM(x, 'Calculation', value, .dispatch = as.integer(4))
},
'Caption' = function(value) {
value = as(value, 'character')
.COM(x, 'Caption', value, .dispatch = as.integer(4))
},
'CellDragAndDrop' = function(value) {
value = as(value, 'logical')
.COM(x, 'CellDragAndDrop', value, .dispatch = as.integer(4))
},
'DisplayClipboardWindow' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayClipboardWindow', value, .dispatch = as.integer(4))
},
'ColorButtons' = function(value) {
value = as(value, 'logical')
.COM(x, 'ColorButtons', value, .dispatch = as.integer(4))
},
'CommandUnderlines' = function(value) {
value = as(value, 'Constants')
.COM(x, 'CommandUnderlines', value, .dispatch = as.integer(4))
},
'ConstrainNumeric' = function(value) {
value = as(value, 'logical')
.COM(x, 'ConstrainNumeric', value, .dispatch = as.integer(4))
},
'CopyObjectsWithCells' = function(value) {
value = as(value, 'logical')
.COM(x, 'CopyObjectsWithCells', value, .dispatch = as.integer(4))
},
'Cursor' = function(value) {
value = as(value, 'Constants')
.COM(x, 'Cursor', value, .dispatch = as.integer(4))
},
'CutCopyMode' = function(value) {
value = as(value, 'Constants')
.COM(x, 'CutCopyMode', value, .dispatch = as.integer(4))
},
'DataEntryMode' = function(value) {
value = as(value, 'integer')
.COM(x, 'DataEntryMode', value, .dispatch = as.integer(4))
},
'DefaultFilePath' = function(value) {
value = as(value, 'character')
.COM(x, 'DefaultFilePath', value, .dispatch = as.integer(4))
},
'DisplayAlerts' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayAlerts', value, .dispatch = as.integer(4))
},
'DisplayFormulaBar' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayFormulaBar', value, .dispatch = as.integer(4))
},
'DisplayFullScreen' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayFullScreen', value, .dispatch = as.integer(4))
},
'DisplayNoteIndicator' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayNoteIndicator', value, .dispatch = as.integer(4))
},
'DisplayCommentIndicator' = function(value) {
value = as(value, 'Constants')
.COM(x, 'DisplayCommentIndicator', value, .dispatch = as.integer(4))
},
'DisplayExcel4Menus' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayExcel4Menus', value, .dispatch = as.integer(4))
},
'DisplayRecentFiles' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayRecentFiles', value, .dispatch = as.integer(4))
},
'DisplayScrollBars' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayScrollBars', value, .dispatch = as.integer(4))
},
'DisplayStatusBar' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayStatusBar', value, .dispatch = as.integer(4))
},
'EditDirectlyInCell' = function(value) {
value = as(value, 'logical')
.COM(x, 'EditDirectlyInCell', value, .dispatch = as.integer(4))
},
'EnableAutoComplete' = function(value) {
value = as(value, 'logical')
.COM(x, 'EnableAutoComplete', value, .dispatch = as.integer(4))
},
'EnableCancelKey' = function(value) {
value = as(value, 'Constants')
.COM(x, 'EnableCancelKey', value, .dispatch = as.integer(4))
},
'EnableSound' = function(value) {
value = as(value, 'logical')
.COM(x, 'EnableSound', value, .dispatch = as.integer(4))
},
'EnableTipWizard' = function(value) {
value = as(value, 'logical')
.COM(x, 'EnableTipWizard', value, .dispatch = as.integer(4))
},
'FixedDecimal' = function(value) {
value = as(value, 'logical')
.COM(x, 'FixedDecimal', value, .dispatch = as.integer(4))
},
'FixedDecimalPlaces' = function(value) {
value = as(value, 'integer')
.COM(x, 'FixedDecimalPlaces', value, .dispatch = as.integer(4))
},
'Height' = function(value) {
value = as(value, 'numeric')
.COM(x, 'Height', value, .dispatch = as.integer(4))
},
'IgnoreRemoteRequests' = function(value) {
value = as(value, 'logical')
.COM(x, 'IgnoreRemoteRequests', value, .dispatch = as.integer(4))
},
'Interactive' = function(value) {
value = as(value, 'logical')
.COM(x, 'Interactive', value, .dispatch = as.integer(4))
},
'Iteration' = function(value) {
value = as(value, 'logical')
.COM(x, 'Iteration', value, .dispatch = as.integer(4))
},
'LargeButtons' = function(value) {
value = as(value, 'logical')
.COM(x, 'LargeButtons', value, .dispatch = as.integer(4))
},
'Left' = function(value) {
value = as(value, 'numeric')
.COM(x, 'Left', value, .dispatch = as.integer(4))
},
'MaxChange' = function(value) {
value = as(value, 'numeric')
.COM(x, 'MaxChange', value, .dispatch = as.integer(4))
},
'MaxIterations' = function(value) {
value = as(value, 'integer')
.COM(x, 'MaxIterations', value, .dispatch = as.integer(4))
},
'MoveAfterReturn' = function(value) {
value = as(value, 'logical')
.COM(x, 'MoveAfterReturn', value, .dispatch = as.integer(4))
},
'MoveAfterReturnDirection' = function(value) {
value = as(value, 'Constants')
.COM(x, 'MoveAfterReturnDirection', value, .dispatch = as.integer(4))
},
'ODBCTimeout' = function(value) {
value = as(value, 'integer')
.COM(x, 'ODBCTimeout', value, .dispatch = as.integer(4))
},
'OnCalculate' = function(value) {
value = as(value, 'character')
.COM(x, 'OnCalculate', value, .dispatch = as.integer(4))
},
'OnData' = function(value) {
value = as(value, 'character')
.COM(x, 'OnData', value, .dispatch = as.integer(4))
},
'OnDoubleClick' = function(value) {
value = as(value, 'character')
.COM(x, 'OnDoubleClick', value, .dispatch = as.integer(4))
},
'OnEntry' = function(value) {
value = as(value, 'character')
.COM(x, 'OnEntry', value, .dispatch = as.integer(4))
},
'OnSheetActivate' = function(value) {
value = as(value, 'character')
.COM(x, 'OnSheetActivate', value, .dispatch = as.integer(4))
},
'OnSheetDeactivate' = function(value) {
value = as(value, 'character')
.COM(x, 'OnSheetDeactivate', value, .dispatch = as.integer(4))
},
'OnWindow' = function(value) {
value = as(value, 'character')
.COM(x, 'OnWindow', value, .dispatch = as.integer(4))
},
'PivotTableSelection' = function(value) {
value = as(value, 'logical')
.COM(x, 'PivotTableSelection', value, .dispatch = as.integer(4))
},
'PromptForSummaryInfo' = function(value) {
value = as(value, 'logical')
.COM(x, 'PromptForSummaryInfo', value, .dispatch = as.integer(4))
},
'ReferenceStyle' = function(value) {
value = as(value, 'Constants')
.COM(x, 'ReferenceStyle', value, .dispatch = as.integer(4))
},
'RollZoom' = function(value) {
value = as(value, 'logical')
.COM(x, 'RollZoom', value, .dispatch = as.integer(4))
},
'ScreenUpdating' = function(value) {
value = as(value, 'logical')
.COM(x, 'ScreenUpdating', value, .dispatch = as.integer(4))
},
'SheetsInNewWorkbook' = function(value) {
value = as(value, 'integer')
.COM(x, 'SheetsInNewWorkbook', value, .dispatch = as.integer(4))
},
'ShowChartTipNames' = function(value) {
value = as(value, 'logical')
.COM(x, 'ShowChartTipNames', value, .dispatch = as.integer(4))
},
'ShowChartTipValues' = function(value) {
value = as(value, 'logical')
.COM(x, 'ShowChartTipValues', value, .dispatch = as.integer(4))
},
'StandardFont' = function(value) {
value = as(value, 'character')
.COM(x, 'StandardFont', value, .dispatch = as.integer(4))
},
'StandardFontSize' = function(value) {
value = as(value, 'numeric')
.COM(x, 'StandardFontSize', value, .dispatch = as.integer(4))
},
'StatusBar' = function(value) {

.COM(x, 'StatusBar', value, .dispatch = as.integer(4))
},
'ShowToolTips' = function(value) {
value = as(value, 'logical')
.COM(x, 'ShowToolTips', value, .dispatch = as.integer(4))
},
'Top' = function(value) {
value = as(value, 'numeric')
.COM(x, 'Top', value, .dispatch = as.integer(4))
},
'DefaultSaveFormat' = function(value) {
value = as(value, 'Constants')
.COM(x, 'DefaultSaveFormat', value, .dispatch = as.integer(4))
},
'TransitionMenuKey' = function(value) {
value = as(value, 'character')
.COM(x, 'TransitionMenuKey', value, .dispatch = as.integer(4))
},
'TransitionMenuKeyAction' = function(value) {
value = as(value, 'integer')
.COM(x, 'TransitionMenuKeyAction', value, .dispatch = as.integer(4))
},
'TransitionNavigKeys' = function(value) {
value = as(value, 'logical')
.COM(x, 'TransitionNavigKeys', value, .dispatch = as.integer(4))
},
'UserControl' = function(value) {
value = as(value, 'logical')
.COM(x, 'UserControl', value, .dispatch = as.integer(4))
},
'UserName' = function(value) {
value = as(value, 'character')
.COM(x, 'UserName', value, .dispatch = as.integer(4))
},
'Visible' = function(value) {
value = as(value, 'logical')
.COM(x, 'Visible', value, .dispatch = as.integer(4))
},
'Width' = function(value) {
value = as(value, 'numeric')
.COM(x, 'Width', value, .dispatch = as.integer(4))
},
'WindowState' = function(value) {
value = as(value, 'Constants')
.COM(x, 'WindowState', value, .dispatch = as.integer(4))
},
'UILanguage' = function(value) {
value = as(value, 'integer')
.COM(x, 'UILanguage', value, .dispatch = as.integer(4))
},
'DefaultSheetDirection' = function(value) {
value = as(value, 'integer')
.COM(x, 'DefaultSheetDirection', value, .dispatch = as.integer(4))
},
'CursorMovement' = function(value) {
value = as(value, 'integer')
.COM(x, 'CursorMovement', value, .dispatch = as.integer(4))
},
'ControlCharacters' = function(value) {
value = as(value, 'logical')
.COM(x, 'ControlCharacters', value, .dispatch = as.integer(4))
},
'EnableEvents' = function(value) {
value = as(value, 'logical')
.COM(x, 'EnableEvents', value, .dispatch = as.integer(4))
},
'DisplayInfoWindow' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayInfoWindow', value, .dispatch = as.integer(4))
},
'ExtendList' = function(value) {
value = as(value, 'logical')
.COM(x, 'ExtendList', value, .dispatch = as.integer(4))
},
'AutoPercentEntry' = function(value) {
value = as(value, 'logical')
.COM(x, 'AutoPercentEntry', value, .dispatch = as.integer(4))
},
'ShowWindowsInTaskbar' = function(value) {
value = as(value, 'logical')
.COM(x, 'ShowWindowsInTaskbar', value, .dispatch = as.integer(4))
},
'FeatureInstall' = function(value) {
value = as(value, 'Constants')
.COM(x, 'FeatureInstall', value, .dispatch = as.integer(4))
},
'FindFormat' = function(value) {
value = as(value, 'CellFormat')
.COM(x, 'FindFormat', value, .dispatch = as.integer(4))
},
'ReplaceFormat' = function(value) {
value = as(value, 'CellFormat')
.COM(x, 'ReplaceFormat', value, .dispatch = as.integer(4))
},
'CalculationInterruptKey' = function(value) {
value = as(value, 'Constants')
.COM(x, 'CalculationInterruptKey', value, .dispatch = as.integer(4))
},
'DisplayFunctionToolTips' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayFunctionToolTips', value, .dispatch = as.integer(4))
},
'AutomationSecurity' = function(value) {
value = as(value, 'Constants')
.COM(x, 'AutomationSecurity', value, .dispatch = as.integer(4))
},
'DisplayPasteOptions' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayPasteOptions', value, .dispatch = as.integer(4))
},
'DisplayInsertOptions' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayInsertOptions', value, .dispatch = as.integer(4))
},
'GenerateGetPivotData' = function(value) {
value = as(value, 'logical')
.COM(x, 'GenerateGetPivotData', value, .dispatch = as.integer(4))
},
'AutoFormatAsYouTypeReplaceHyperlinks' = function(value) {
value = as(value, 'logical')
.COM(x, 'AutoFormatAsYouTypeReplaceHyperlinks', value, .dispatch = as.integer(4))
},
'MapPaperSize' = function(value) {
value = as(value, 'logical')
.COM(x, 'MapPaperSize', value, .dispatch = as.integer(4))
},
'ShowStartupDialog' = function(value) {
value = as(value, 'logical')
.COM(x, 'ShowStartupDialog', value, .dispatch = as.integer(4))
},
'DecimalSeparator' = function(value) {
value = as(value, 'character')
.COM(x, 'DecimalSeparator', value, .dispatch = as.integer(4))
},
'ThousandsSeparator' = function(value) {
value = as(value, 'character')
.COM(x, 'ThousandsSeparator', value, .dispatch = as.integer(4))
},
'UseSystemSeparators' = function(value) {
value = as(value, 'logical')
.COM(x, 'UseSystemSeparators', value, .dispatch = as.integer(4))
},
'DisplayDocumentActionTaskPane' = function(value) {
value = as(value, 'logical')
.COM(x, 'DisplayDocumentActionTaskPane', value, .dispatch = as.integer(4))
} )
'COM._Application.Methods'  = list('Calculate' = function(  ){
	ans = .COM(x, 'Calculate', .dispatch = as.integer(1), .ids =279)
		 ans
},
'DDEExecute' = function( Channel, String ){
	if( missing( Channel )||missing( String ) ) {
	stop('You must specify a value for the argument(s)  Channel, String ')
	}
	Channel = as(Channel, 'integer')
	String = as(String, 'character')
	ans = .COM(x, 'DDEExecute', Channel, String, .dispatch = as.integer(1), .ids =333)
		 ans
},
'DDEInitiate' = function( App, Topic ){
	if( missing( App )||missing( Topic ) ) {
	stop('You must specify a value for the argument(s)  App, Topic ')
	}
	App = as(App, 'character')
	Topic = as(Topic, 'character')
	ans = .COM(x, 'DDEInitiate', App, Topic, .dispatch = as.integer(1), .ids =334)
	ans
},
'DDEPoke' = function( Channel, Item, Data ){
	if( missing( Channel )||missing( Item )||missing( Data ) ) {
	stop('You must specify a value for the argument(s)  Channel, Item, Data ')
	}
	Channel = as(Channel, 'integer')
	
	
	ans = .COM(x, 'DDEPoke', Channel, Item, Data, .dispatch = as.integer(1), .ids =335)
		 ans
},
'DDERequest' = function( Channel, Item ){
	if( missing( Channel )||missing( Item ) ) {
	stop('You must specify a value for the argument(s)  Channel, Item ')
	}
	Channel = as(Channel, 'integer')
	Item = as(Item, 'character')
	ans = .COM(x, 'DDERequest', Channel, Item, .dispatch = as.integer(1), .ids =336)
		 ans
},
'DDETerminate' = function( Channel ){
	if( missing( Channel ) ) {
	stop('You must specify a value for the argument(s)  Channel ')
	}
	Channel = as(Channel, 'integer')
	ans = .COM(x, 'DDETerminate', Channel, .dispatch = as.integer(1), .ids =337)
		 ans
},
'Evaluate' = function( Name ){
	if( missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Name ')
	}
	
	ans = .COM(x, 'Evaluate', Name, .dispatch = as.integer(1), .ids =1)
		 ans
},
'_Evaluate' = function( Name ){
	if( missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Name ')
	}
	
	ans = .COM(x, '_Evaluate', Name, .dispatch = as.integer(1), .ids =-5)
		 ans
},
'ExecuteExcel4Macro' = function( String ){
	if( missing( String ) ) {
	stop('You must specify a value for the argument(s)  String ')
	}
	String = as(String, 'character')
	ans = .COM(x, 'ExecuteExcel4Macro', String, .dispatch = as.integer(1), .ids =350)
		 ans
},
'Intersect' = function( Arg1, Arg2, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	if( missing( Arg1 )||missing( Arg2 ) ) {
	stop('You must specify a value for the argument(s)  Arg1, Arg2 ')
	}
	Arg1 = as(Arg1, 'Range')
	Arg2 = as(Arg2, 'Range')
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Intersect', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =766)
		 as(ans, 'Range')
},
'Run' = function( Macro = NA, Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Run', Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =259)
		 ans
},
'_Run2' = function( Macro = NA, Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, '_Run2', Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =806)
		 ans
},
'SendKeys' = function( Keys, Wait = NA ){
	if( missing( Keys ) ) {
	stop('You must specify a value for the argument(s)  Keys ')
	}
	
	
	ans = .COM(x, 'SendKeys', Keys, Wait, .dispatch = as.integer(1), .ids =383)
		 ans
},
'Union' = function( Arg1, Arg2, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	if( missing( Arg1 )||missing( Arg2 ) ) {
	stop('You must specify a value for the argument(s)  Arg1, Arg2 ')
	}
	Arg1 = as(Arg1, 'Range')
	Arg2 = as(Arg2, 'Range')
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Union', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =779)
		 as(ans, 'Range')
},
'ActivateMicrosoftApp' = function( Index ){
	if( missing( Index ) ) {
	stop('You must specify a value for the argument(s)  Index ')
	}
	Index = as(Index, 'Constants')
	ans = .COM(x, 'ActivateMicrosoftApp', Index, .dispatch = as.integer(1), .ids =1095)
		 ans
},
'AddChartAutoFormat' = function( Chart, Name, Description = NA ){
	if( missing( Chart )||missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Chart, Name ')
	}
	
	Name = as(Name, 'character')
	
	ans = .COM(x, 'AddChartAutoFormat', Chart, Name, Description, .dispatch = as.integer(1), .ids =216)
		 ans
},
'AddCustomList' = function( ListArray, ByRow = NA ){
	if( missing( ListArray ) ) {
	stop('You must specify a value for the argument(s)  ListArray ')
	}
	
	
	ans = .COM(x, 'AddCustomList', ListArray, ByRow, .dispatch = as.integer(1), .ids =780)
		 ans
},
'CentimetersToPoints' = function( Centimeters ){
	if( missing( Centimeters ) ) {
	stop('You must specify a value for the argument(s)  Centimeters ')
	}
	Centimeters = as(Centimeters, 'numeric')
	ans = .COM(x, 'CentimetersToPoints', Centimeters, .dispatch = as.integer(1), .ids =1086)
	ans
},
'CheckSpelling' = function( Word, CustomDictionary = NA, IgnoreUppercase = NA ){
	if( missing( Word ) ) {
	stop('You must specify a value for the argument(s)  Word ')
	}
	Word = as(Word, 'character')
	
	
	ans = .COM(x, 'CheckSpelling', Word, CustomDictionary, IgnoreUppercase, .dispatch = as.integer(1), .ids =505)
	ans
},
'ConvertFormula' = function( Formula, FromReferenceStyle, ToReferenceStyle = NA, ToAbsolute = NA, RelativeTo = NA ){
	if( missing( Formula )||missing( FromReferenceStyle ) ) {
	stop('You must specify a value for the argument(s)  Formula, FromReferenceStyle ')
	}
	
	FromReferenceStyle = as(FromReferenceStyle, 'Constants')
	
	
	
	ans = .COM(x, 'ConvertFormula', Formula, FromReferenceStyle, ToReferenceStyle, ToAbsolute, RelativeTo, .dispatch = as.integer(1), .ids =325)
		 ans
},
'Dummy1' = function( Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA ){
	
	
	
	
	ans = .COM(x, 'Dummy1', Arg1, Arg2, Arg3, Arg4, .dispatch = as.integer(1), .ids =1782)
		 ans
},
'Dummy2' = function( Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA ){
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Dummy2', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, .dispatch = as.integer(1), .ids =1783)
		 ans
},
'Dummy3' = function(  ){
	ans = .COM(x, 'Dummy3', .dispatch = as.integer(1), .ids =1784)
		 ans
},
'Dummy4' = function( Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA ){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Dummy4', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, .dispatch = as.integer(1), .ids =1785)
		 ans
},
'Dummy5' = function( Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA ){
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Dummy5', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, .dispatch = as.integer(1), .ids =1786)
		 ans
},
'Dummy6' = function(  ){
	ans = .COM(x, 'Dummy6', .dispatch = as.integer(1), .ids =1787)
		 ans
},
'Dummy7' = function(  ){
	ans = .COM(x, 'Dummy7', .dispatch = as.integer(1), .ids =1788)
		 ans
},
'Dummy8' = function( Arg1 = NA ){
	
	ans = .COM(x, 'Dummy8', Arg1, .dispatch = as.integer(1), .ids =1789)
		 ans
},
'Dummy9' = function(  ){
	ans = .COM(x, 'Dummy9', .dispatch = as.integer(1), .ids =1790)
		 ans
},
'Dummy10' = function( arg = NA ){
	
	ans = .COM(x, 'Dummy10', arg, .dispatch = as.integer(1), .ids =1791)
	ans
},
'Dummy11' = function(  ){
	ans = .COM(x, 'Dummy11', .dispatch = as.integer(1), .ids =1792)
		 ans
},
'DeleteChartAutoFormat' = function( Name ){
	if( missing( Name ) ) {
	stop('You must specify a value for the argument(s)  Name ')
	}
	Name = as(Name, 'character')
	ans = .COM(x, 'DeleteChartAutoFormat', Name, .dispatch = as.integer(1), .ids =217)
		 ans
},
'DeleteCustomList' = function( ListNum ){
	if( missing( ListNum ) ) {
	stop('You must specify a value for the argument(s)  ListNum ')
	}
	ListNum = as(ListNum, 'integer')
	ans = .COM(x, 'DeleteCustomList', ListNum, .dispatch = as.integer(1), .ids =783)
		 ans
},
'DoubleClick' = function(  ){
	ans = .COM(x, 'DoubleClick', .dispatch = as.integer(1), .ids =349)
		 ans
},
'_FindFile' = function(  ){
	ans = .COM(x, '_FindFile', .dispatch = as.integer(1), .ids =1068)
		 ans
},
'GetCustomListContents' = function( ListNum ){
	if( missing( ListNum ) ) {
	stop('You must specify a value for the argument(s)  ListNum ')
	}
	ListNum = as(ListNum, 'integer')
	ans = .COM(x, 'GetCustomListContents', ListNum, .dispatch = as.integer(1), .ids =786)
		 ans
},
'GetCustomListNum' = function( ListArray ){
	if( missing( ListArray ) ) {
	stop('You must specify a value for the argument(s)  ListArray ')
	}
	
	ans = .COM(x, 'GetCustomListNum', ListArray, .dispatch = as.integer(1), .ids =785)
	ans
},
'GetOpenFilename' = function( FileFilter = NA, FilterIndex = NA, Title = NA, ButtonText = NA, MultiSelect = NA ){
	
	
	
	
	
	ans = .COM(x, 'GetOpenFilename', FileFilter, FilterIndex, Title, ButtonText, MultiSelect, .dispatch = as.integer(1), .ids =1075)
		 ans
},
'GetSaveAsFilename' = function( InitialFilename = NA, FileFilter = NA, FilterIndex = NA, Title = NA, ButtonText = NA ){
	
	
	
	
	
	ans = .COM(x, 'GetSaveAsFilename', InitialFilename, FileFilter, FilterIndex, Title, ButtonText, .dispatch = as.integer(1), .ids =1076)
		 ans
},
'Goto' = function( Reference = NA, Scroll = NA ){
	
	
	ans = .COM(x, 'Goto', Reference, Scroll, .dispatch = as.integer(1), .ids =475)
		 ans
},
'Help' = function( HelpFile = NA, HelpContextID = NA ){
	
	
	ans = .COM(x, 'Help', HelpFile, HelpContextID, .dispatch = as.integer(1), .ids =354)
		 ans
},
'InchesToPoints' = function( Inches ){
	if( missing( Inches ) ) {
	stop('You must specify a value for the argument(s)  Inches ')
	}
	Inches = as(Inches, 'numeric')
	ans = .COM(x, 'InchesToPoints', Inches, .dispatch = as.integer(1), .ids =1087)
	ans
},
'InputBox' = function( Prompt, Title = NA, Default = NA, Left = NA, Top = NA, HelpFile = NA, HelpContextID = NA, Type = NA ){
	if( missing( Prompt ) ) {
	stop('You must specify a value for the argument(s)  Prompt ')
	}
	Prompt = as(Prompt, 'character')
	
	
	
	
	
	
	
	ans = .COM(x, 'InputBox', Prompt, Title, Default, Left, Top, HelpFile, HelpContextID, Type, .dispatch = as.integer(1), .ids =357)
		 ans
},
'MacroOptions' = function( Macro = NA, Description = NA, HasMenu = NA, MenuText = NA, HasShortcutKey = NA, ShortcutKey = NA, Category = NA, StatusBar = NA, HelpContextID = NA, HelpFile = NA ){
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'MacroOptions', Macro, Description, HasMenu, MenuText, HasShortcutKey, ShortcutKey, Category, StatusBar, HelpContextID, HelpFile, .dispatch = as.integer(1), .ids =1135)
		 ans
},
'MailLogoff' = function(  ){
	ans = .COM(x, 'MailLogoff', .dispatch = as.integer(1), .ids =945)
		 ans
},
'MailLogon' = function( Name = NA, Password = NA, DownloadNewMail = NA ){
	
	
	
	ans = .COM(x, 'MailLogon', Name, Password, DownloadNewMail, .dispatch = as.integer(1), .ids =943)
		 ans
},
'NextLetter' = function(  ){
	ans = .COM(x, 'NextLetter', .dispatch = as.integer(1), .ids =972)
		 as(ans, 'Workbook')
},
'OnKey' = function( Key, Procedure = NA ){
	if( missing( Key ) ) {
	stop('You must specify a value for the argument(s)  Key ')
	}
	Key = as(Key, 'character')
	
	ans = .COM(x, 'OnKey', Key, Procedure, .dispatch = as.integer(1), .ids =626)
		 ans
},
'OnRepeat' = function( Text, Procedure ){
	if( missing( Text )||missing( Procedure ) ) {
	stop('You must specify a value for the argument(s)  Text, Procedure ')
	}
	Text = as(Text, 'character')
	Procedure = as(Procedure, 'character')
	ans = .COM(x, 'OnRepeat', Text, Procedure, .dispatch = as.integer(1), .ids =769)
		 ans
},
'OnTime' = function( EarliestTime, Procedure, LatestTime = NA, Schedule = NA ){
	if( missing( EarliestTime )||missing( Procedure ) ) {
	stop('You must specify a value for the argument(s)  EarliestTime, Procedure ')
	}
	
	Procedure = as(Procedure, 'character')
	
	
	ans = .COM(x, 'OnTime', EarliestTime, Procedure, LatestTime, Schedule, .dispatch = as.integer(1), .ids =624)
		 ans
},
'OnUndo' = function( Text, Procedure ){
	if( missing( Text )||missing( Procedure ) ) {
	stop('You must specify a value for the argument(s)  Text, Procedure ')
	}
	Text = as(Text, 'character')
	Procedure = as(Procedure, 'character')
	ans = .COM(x, 'OnUndo', Text, Procedure, .dispatch = as.integer(1), .ids =770)
		 ans
},
'Quit' = function(  ){
	ans = .COM(x, 'Quit', .dispatch = as.integer(1), .ids =302)
		 ans
},
'RecordMacro' = function( BasicCode = NA, XlmCode = NA ){
	
	
	ans = .COM(x, 'RecordMacro', BasicCode, XlmCode, .dispatch = as.integer(1), .ids =773)
		 ans
},
'RegisterXLL' = function( Filename ){
	if( missing( Filename ) ) {
	stop('You must specify a value for the argument(s)  Filename ')
	}
	Filename = as(Filename, 'character')
	ans = .COM(x, 'RegisterXLL', Filename, .dispatch = as.integer(1), .ids =30)
	ans
},
'Repeat' = function(  ){
	ans = .COM(x, 'Repeat', .dispatch = as.integer(1), .ids =301)
		 ans
},
'ResetTipWizard' = function(  ){
	ans = .COM(x, 'ResetTipWizard', .dispatch = as.integer(1), .ids =928)
		 ans
},
'Save' = function( Filename = NA ){
	
	ans = .COM(x, 'Save', Filename, .dispatch = as.integer(1), .ids =283)
		 ans
},
'SaveWorkspace' = function( Filename = NA ){
	
	ans = .COM(x, 'SaveWorkspace', Filename, .dispatch = as.integer(1), .ids =212)
		 ans
},
'SetDefaultChart' = function( FormatName = NA, Gallery = NA ){
	
	
	ans = .COM(x, 'SetDefaultChart', FormatName, Gallery, .dispatch = as.integer(1), .ids =219)
		 ans
},
'Undo' = function(  ){
	ans = .COM(x, 'Undo', .dispatch = as.integer(1), .ids =303)
		 ans
},
'Volatile' = function( Volatile = NA ){
	
	ans = .COM(x, 'Volatile', Volatile, .dispatch = as.integer(1), .ids =788)
		 ans
},
'_Wait' = function( Time ){
	if( missing( Time ) ) {
	stop('You must specify a value for the argument(s)  Time ')
	}
	
	ans = .COM(x, '_Wait', Time, .dispatch = as.integer(1), .ids =393)
		 ans
},
'_WSFunction' = function( Arg1 = NA, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, '_WSFunction', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =169)
		 ans
},
'Wait' = function( Time ){
	if( missing( Time ) ) {
	stop('You must specify a value for the argument(s)  Time ')
	}
	
	ans = .COM(x, 'Wait', Time, .dispatch = as.integer(1), .ids =1770)
	ans
},
'GetPhonetic' = function( Text = NA ){
	
	ans = .COM(x, 'GetPhonetic', Text, .dispatch = as.integer(1), .ids =1795)
	ans
},
'Dummy12' = function( p1, p2 ){
	if( missing( p1 )||missing( p2 ) ) {
	stop('You must specify a value for the argument(s)  p1, p2 ')
	}
	p1 = as(p1, 'PivotTable')
	p2 = as(p2, 'PivotTable')
	ans = .COM(x, 'Dummy12', p1, p2, .dispatch = as.integer(1), .ids =1803)
		 ans
},
'CalculateFull' = function(  ){
	ans = .COM(x, 'CalculateFull', .dispatch = as.integer(1), .ids =1805)
		 ans
},
'FindFile' = function(  ){
	ans = .COM(x, 'FindFile', .dispatch = as.integer(1), .ids =1771)
	ans
},
'Dummy13' = function( Arg1, Arg2 = NA, Arg3 = NA, Arg4 = NA, Arg5 = NA, Arg6 = NA, Arg7 = NA, Arg8 = NA, Arg9 = NA, Arg10 = NA, Arg11 = NA, Arg12 = NA, Arg13 = NA, Arg14 = NA, Arg15 = NA, Arg16 = NA, Arg17 = NA, Arg18 = NA, Arg19 = NA, Arg20 = NA, Arg21 = NA, Arg22 = NA, Arg23 = NA, Arg24 = NA, Arg25 = NA, Arg26 = NA, Arg27 = NA, Arg28 = NA, Arg29 = NA, Arg30 = NA ){
	if( missing( Arg1 ) ) {
	stop('You must specify a value for the argument(s)  Arg1 ')
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	ans = .COM(x, 'Dummy13', Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30, .dispatch = as.integer(1), .ids =1933)
		 ans
},
'Dummy14' = function(  ){
	ans = .COM(x, 'Dummy14', .dispatch = as.integer(1), .ids =1944)
		 ans
},
'CalculateFullRebuild' = function(  ){
	ans = .COM(x, 'CalculateFullRebuild', .dispatch = as.integer(1), .ids =1945)
		 ans
},
'CheckAbort' = function( KeepAbort = NA ){
	
	ans = .COM(x, 'CheckAbort', KeepAbort, .dispatch = as.integer(1), .ids =1952)
		 ans
},
'DisplayXMLSourcePane' = function( XmlMap = NA ){
	
	ans = .COM(x, 'DisplayXMLSourcePane', XmlMap, .dispatch = as.integer(1), .ids =2252)
		 ans
},
'Support' = function( Object, ID, arg = NA ){
	if( missing( Object )||missing( ID ) ) {
	stop('You must specify a value for the argument(s)  Object, ID ')
	}
	
	ID = as(ID, 'integer')
	
	ans = .COM(x, 'Support', Object, ID, arg, .dispatch = as.integer(1), .ids =2255)
		 ans
} )
